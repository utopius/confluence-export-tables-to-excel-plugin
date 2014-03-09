package de.cranktheory.confluence.excel.export;

import javax.xml.namespace.QName;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

public final class StorageFormat
{
    private StorageFormat()
    {
        // Nope
    }

    public static boolean isStorageFormatLink(XMLEvent event)
    {
        return XMLEvents.isStart(event, "link");
    }

    public static boolean isHyperlink(XMLEvent event)
    {
        return event != null && XMLEvents.isStart(event, "a");
    }

    public static boolean isPageLink(XMLEvent event)
    {
        return event != null && XMLEvents.isStart(event, "page");
    }

    public static String getPageTitle(XMLEvent event)
    {
        Attribute pageTitleAttribute = event.asStartElement()
            .getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}content-title"));
        String pageTitle = pageTitleAttribute.getValue();

        return pageTitle;
    }

    public static String getSpaceKey(XMLEvent event)
    {
        Attribute spaceKeyAttribute = event.asStartElement()
            .getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}space-key"));

        return spaceKeyAttribute == null
                ? null
                : spaceKeyAttribute.getValue();
    }

    /**
     * Returns the spaceKey of the given link event or the key of the specified page, if not found on the event.
     *
     * @param event
     * @param page
     * @return
     */
    public static String getSpaceKey(XMLEvent event, String defaultKey)
    {
        String spaceKey = getSpaceKey(event);
        if (spaceKey == null)
        {
            spaceKey = defaultKey;
        }
        return spaceKey;
    }

    public static boolean isAttachmentLink(XMLEvent event)
    {
        return event != null && XMLEvents.isStart(event, "attachment");
    }

    public static String getAttachmentFilename(StartElement linkContent)
    {
        Attribute attributeByName = linkContent.asStartElement()
            .getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}filename"));

        String value = attributeByName.getValue();
        return value;
    }
}
