package de.cranktheory.plugins.confluence.excel.export;

import java.io.IOException;
import java.io.InputStream;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.XMLEvent;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.atlassian.confluence.pages.Attachment;
import com.atlassian.confluence.pages.Page;
import com.atlassian.confluence.pages.PageManager;

public class ImageParser
{
    private static final Logger LOG = LoggerFactory.getLogger(ImageParser.class);

    private final Page _page;
    private final PageManager _pageManager;

    public ImageParser(Page page, PageManager pageManager)
    {
        _page = page;
        _pageManager = pageManager;
    }

    public void parseImage(XMLEventReader reader, WorksheetBuilder sheetBuilder) throws XMLStreamException
    {
        XMLEvent event = reader.nextEvent();
        if (XMLEvents.isStart(event, "attachment"))
        {
            Attribute attributeByName = event.asStartElement()
                .getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}filename"));

            String attachmentFilename = attributeByName.getValue();

            Attachment attachment = getAttachment(reader, attachmentFilename);

            if (attachment == null)
            {
                LOG.warn(String.format("Attachment %s on page %s not found.", attachmentFilename, _page.getId()));
                return;
            }

            try
            {
                InputStream contentsAsStream = attachment.getContentsAsStream();
                String contentType = attachment.getContentType();
                sheetBuilder.drawPictureToCell(StreamUtils.readAllBytes(contentsAsStream),
                        ImageUtils.mimeTypeToPoiFormat(contentType));
            }
            catch (PictureDrawingException e)
            {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            catch (IOException e)
            {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
    }

    private Attachment getAttachment(XMLEventReader reader, String attachmentFilename) throws XMLStreamException
    {
        Attachment attachment = _page.getAttachmentNamed(attachmentFilename);
        // If attachment is null then its probably a referenced attachment we'll handle in the page block below.
        if (attachment == null)
        {
            XMLEvent pageEvent = reader.nextEvent();
            if (XMLEvents.isStart(pageEvent, "page"))
            {
                // Attachment included from somewhere else
                Attribute pageTitleAttribute = pageEvent.asStartElement()
                    .getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}content-title"));
                String pageTitle = pageTitleAttribute.getValue();

                Attribute spaceKeyAttribute = pageEvent.asStartElement()
                    .getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}space-key"));

                // If no spaceKey is provided, then the referenced page is located in the same Space as _page
                String spaceKey = spaceKeyAttribute == null
                        ? _page.getSpaceKey()
                        : spaceKeyAttribute.getValue();
                Page referencedPage = _pageManager.getPage(spaceKey, pageTitle);

                // If the page is in trash (referencedPage == null), omit the attachment
                if (referencedPage != null)
                {
                    attachment = referencedPage.getAttachmentNamed(attachmentFilename);
                }
            }
        }

        return attachment;
    }
}
