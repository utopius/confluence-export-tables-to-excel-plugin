package de.cranktheory.confluence.excel.export;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;

public interface LinkParser
{
    public static final Link UnsupportedLink = new Link("UNSUPPORTED LINK", "UNSUPPORTED LINK");

    /**
     * @param event
     * @return <code>true</code> if the specified <code>event</code> starts a link element which can be handled by
     *         the {@link LinkParser}.
     */
    boolean isLink(XMLEvent event);

    Link parseLink(XMLEventReader reader, XMLEvent event) throws XMLStreamException;

}
