package de.cranktheory.confluence.excel.export;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.XMLEvent;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;

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

    public static boolean isStartMacro(XMLEvent event, String name)
    {
        Preconditions.checkNotNull(event, "event");
        Preconditions.checkArgument(!Strings.isNullOrEmpty(name), "name is null or empty.");

        if (isStart(event, name))
        {
            Attribute macroNameAttribute = event.asStartElement()
                .getAttributeByName(QName.valueOf("{http://atlassian.com/content}name"));

            return name.equals(macroNameAttribute.getValue());
        }

        return false;
    }

    public static boolean nextIsParameter(XMLEventReader reader) throws XMLStreamException
    {
        Preconditions.checkNotNull(reader, "reader");

        XMLEvent next = reader.peek();
        boolean nextIsParameter = next != null && isStart(next, "parameter");
        return nextIsParameter;
    }
}
