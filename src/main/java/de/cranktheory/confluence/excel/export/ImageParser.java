package de.cranktheory.confluence.excel.export;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;

public interface ImageParser
{

    void parseImage(XMLEventReader reader, WorksheetBuilder sheetBuilder) throws XMLStreamException;

}
