package de.cranktheory.confluence.excel.export;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

public class DefaultLinkParser implements LinkParser
{
    private final UrlResolver _urlResolver;

    public DefaultLinkParser(UrlResolver urlResolved)
    {
        _urlResolver = urlResolved;
    }

    @Override
    public Link parseLink(XMLEventReader reader, XMLEvent event) throws XMLStreamException
    {
        // TODO: Add support for anchor-property in ac:link element
        if (StorageFormat.isHyperlink(event))
        {
            // Process html hyperlink
            Attribute href = event.asStartElement()
                .getAttributeByName(QName.valueOf("href"));
            String url = href.getValue();

            String label = XMLEvents.nextIsCharacters(reader)
                    ? reader.nextEvent()
                        .asCharacters()
                        .getData()
                    : url;

            return new Link(url, label);
        }
        else if (StorageFormat.isStorageFormatLink(event) && StorageFormat.isPageLink(reader.peek()))
        {
            StartElement pageElement = reader.nextEvent()
                .asStartElement();

            String pageTitle = StorageFormat.getPageTitle(pageElement);
            String spaceKey = StorageFormat.getSpaceKey(pageElement);

            String url = _urlResolver.toAbsoluteUrl(_urlResolver.resolvePageUrl(pageTitle, spaceKey));
            reader.nextEvent(); // navigate to the label
            String label = parseLabel(reader, pageTitle);
            return new Link(url, label);
        }
        else if (StorageFormat.isStorageFormatLink(event) && StorageFormat.isAttachmentLink(reader.peek()))
        {
            String attachmentFilename = StorageFormat.getAttachmentFilename(reader.nextEvent()
                .asStartElement());

            String pageTitle = null;
            String spaceKey = null;

            if (XMLEvents.nextIs(reader, "page"))
            {
                // attachment is on a different page
                StartElement pageElement = reader.nextEvent()
                    .asStartElement();
                pageTitle = StorageFormat.getPageTitle(pageElement);
                spaceKey = StorageFormat.getSpaceKey(pageElement);

                // Skip page end tag
                reader.nextEvent();
            }
            String url = _urlResolver.toAbsoluteUrl(_urlResolver.resolveAttachmentUrl(attachmentFilename, pageTitle,
                    spaceKey));
            reader.nextEvent(); // navigate to the label
            String label = parseLabel(reader, attachmentFilename);
            return new Link(url, label);
        }

        return LinkParser.UnsupportedLink;
    }

    private static String parseLabel(XMLEventReader reader, String defaultLabel) throws XMLStreamException
    {
        if (XMLEvents.nextIs(reader, "plain-text-link-body"))
        {
            // Navigate to the content
            reader.nextEvent();
            if (XMLEvents.nextIsCharacters(reader))
            {
                String label = reader.nextEvent()
                    .asCharacters()
                    .getData();
                return label;
            }
        }
        return defaultLabel;
    }
}
