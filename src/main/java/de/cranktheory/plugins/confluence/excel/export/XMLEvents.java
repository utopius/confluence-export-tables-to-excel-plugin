package de.cranktheory.plugins.confluence.excel.export;

import javax.xml.stream.events.XMLEvent;

public final class XMLEvents
{
    private XMLEvents()
    {
        //Nope
    }

    public static boolean isStart(XMLEvent event, String elementName)
    {
        return event.isStartElement() && elementName.equals(event.asStartElement()
            .getName()
            .getLocalPart());
    }

    public static boolean isEnd(XMLEvent event, String elementName)
    {
        return event.isEndElement() && elementName.equals(event.asEndElement()
            .getName()
            .getLocalPart());
    }
}
