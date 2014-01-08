package de.cranktheory.confluence.excel.export;

import javax.xml.stream.events.XMLEvent;

import com.google.common.base.Preconditions;

public final class XMLEvents
{
    private XMLEvents()
    {
        // Nope
    }

    public static boolean isStart(XMLEvent event, String elementName)
    {
        Preconditions.checkNotNull(event, "event");
        Preconditions.checkNotNull(elementName, "elementName");

        return event.isStartElement() && elementName.equals(event.asStartElement()
            .getName()
            .getLocalPart());
    }

    public static boolean isEnd(XMLEvent event, String elementName)
    {
        Preconditions.checkNotNull(event, "event");
        Preconditions.checkNotNull(elementName, "elementName");

        return event.isEndElement() && elementName.equals(event.asEndElement()
            .getName()
            .getLocalPart());
    }
}
