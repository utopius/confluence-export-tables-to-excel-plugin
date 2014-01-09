package ut.de.cranktheory.plugins.confluence.excel;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;

import de.cranktheory.confluence.excel.export.ImageParser;
import de.cranktheory.confluence.excel.export.WorksheetBuilder;

final class MockImageParser implements ImageParser
{
    @Override
    public void parseImage(XMLEventReader reader, WorksheetBuilder sheetBuilder) throws XMLStreamException
    {
    }
}