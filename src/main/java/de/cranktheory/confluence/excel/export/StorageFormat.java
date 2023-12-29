package de.cranktheory.confluence.excel.export;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.xml.namespace.QName;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.XMLEvent;

import com.google.common.base.Preconditions;

/**
 * Provides some utility methods to check storage format elements and to access element attributes / values.
 */
public final class StorageFormat
{
    private StorageFormat()
    {
        // Nope
    }

    public static boolean isStorageFormatLink(@Nonnull XMLEvent event)
    {
        Preconditions.checkNotNull(event, "event");

        return XMLEvents.isStart(event, "link");
    }

    public static boolean isHyperlink(@Nullable XMLEvent event)
    {
        return event != null && XMLEvents.isStart(event, "a");
    }

    public static boolean isPageLink(@Nullable XMLEvent event)
    {
        return event != null && XMLEvents.isStart(event, "page");
    }

    public static String getPageTitle(@Nonnull XMLEvent event)
    {
        Preconditions.checkNotNull(event, "event");

        Attribute pageTitleAttribute = event.asStartElement()
            .getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}content-title"));
        String pageTitle = pageTitleAttribute.getValue();

        return pageTitle;
    }

    public static String getSpaceKey(@Nonnull XMLEvent event)
    {
        Preconditions.checkNotNull(event, "event");

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
    public static String getSpaceKey(@Nonnull XMLEvent event, @Nonnull String defaultKey)
    {
        Preconditions.checkNotNull(event, "event");
        Preconditions.checkNotNull(defaultKey, "defaultKey");

        String spaceKey = getSpaceKey(event);
        if (spaceKey == null)
        {
            spaceKey = defaultKey;
        }
        return spaceKey;
    }

    public static boolean isAttachmentLink(@Nullable XMLEvent event)
    {
        return event != null && XMLEvents.isStart(event, "attachment");
    }

    public static String getAttachmentFilename(@Nonnull XMLEvent linkContent)
    {
        Preconditions.checkNotNull(linkContent, "linkContent");

        Attribute attributeByName = linkContent.asStartElement()
            .getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}filename"));

        String value = attributeByName.getValue();
        return value;
    }
}
