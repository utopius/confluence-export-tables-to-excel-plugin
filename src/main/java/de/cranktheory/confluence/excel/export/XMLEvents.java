package de.cranktheory.confluence.excel.export;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
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

    public static boolean isStartExportMacro(XMLEvent event)
    {
        Preconditions.checkNotNull(event, "event");

        if (isStart(event, "structured-macro"))
        {
            Attribute macroNameAttribute = event.asStartElement()
                .getAttributeByName(QName.valueOf("{http://atlassian.com/content}name"));

            return "export-table".equals(macroNameAttribute.getValue());
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

    public static boolean nextIsCharacters(XMLEventReader reader) throws XMLStreamException
    {
        return reader.peek() != null && reader.peek()
            .isCharacters();
    }

    public static boolean nextIs(XMLEventReader reader, String searchedElement) throws XMLStreamException
    {
        Preconditions.checkNotNull(searchedElement, "searchedElement must not be null");

        boolean hasNextElement = reader.peek() != null && reader.peek()
            .isStartElement();

        if (hasNextElement)
        {
            String elementName = reader.peek()
                .asStartElement()
                .getName()
                .getLocalPart();
            return searchedElement.equalsIgnoreCase(elementName);
        }
        return false;
    }
}
