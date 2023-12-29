package de.cranktheory.confluence.excel.export;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;

public interface LinkParser
{
    public static final Link UnsupportedLink = new Link("UNSUPPORTED LINK", "UNSUPPORTED LINK");

    /**
     * Tries to parse a valid {@link Link} if possible. Otherwise {@link #UnsupportedLink} is returned.
     *
     * @param reader
     * @param event
     * @return well, a link.
     * @throws XMLStreamException
     */
    Link parseLink(XMLEventReader reader, XMLEvent event) throws XMLStreamException;

}
