package de.cranktheory.confluence.excel.export;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;

public class TableParser
{
    private final ParserFactory _parserFactory;
    private final WorkbookBuilder _workbookBuilder;
    private final String _sheetname;

    public TableParser(ParserFactory parserFactory, WorkbookBuilder workbookBuilder, final String sheetname)
    {
        _parserFactory = parserFactory;
        _workbookBuilder = workbookBuilder;
        _sheetname = sheetname;
    }

    public void parseTable(XMLEventReader reader) throws XMLStreamException
    {
        WorksheetBuilder sheetBuilder = _workbookBuilder.createSheet(_sheetname);

        int rowIndex = 0;

        while (reader.hasNext())
        {
            XMLEvent event = reader.nextEvent();

            if (XMLEvents.isEnd(event, "table")) return;

            if (XMLEvents.isStart(event, "tr"))
            {
                sheetBuilder.createRow(rowIndex++);
                parseRow(reader, _workbookBuilder, sheetBuilder);
            }
        }
    }

    private void parseRow(XMLEventReader reader, WorkbookBuilder workbookBuilder, WorksheetBuilder sheetBuilder)
            throws XMLStreamException
    {
        int cellIndex = 0;

        CellParser parser = _parserFactory.newCellParser();
        while (reader.hasNext())
        {
            XMLEvent event = reader.nextEvent();

            if (XMLEvents.isEnd(event, "tr")) return;

            if (XMLEvents.isStart(event, "th") || XMLEvents.isStart(event, "td"))
            {
                sheetBuilder.createCell(cellIndex++);
                parser.parseCell(reader, sheetBuilder, event.asStartElement());
            }
        }
    }
}