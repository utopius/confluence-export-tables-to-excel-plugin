package de.cranktheory.confluence.excel.export;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;

public interface ImageParserr
{

    void parseImage(XMLEventReader reader, WorksheetBuilder sheetBuilder) throws XMLStreamException;

}
