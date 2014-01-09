package de.cranktheory.confluence.excel.export;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;

import com.google.common.base.Preconditions;

public class TableParser
{
    public static TableParser newInstance(ImageParser imageParser)
    {
        return new TableParser(Preconditions.checkNotNull(imageParser, "imageParser"));
    }

    private final ImageParser _imageParser;
    private int _nestedTableCount = 0;

    private TableParser(ImageParser imageParser)
    {
        _imageParser = imageParser;
    }

    public void parseTable(XMLEventReader reader, WorkbookBuilder workbookBuilder, String sheetname)
            throws XMLStreamException
    {
        _nestedTableCount = 0;

        internalParseTable(reader, workbookBuilder, sheetname);
    }

    private void internalParseTable(XMLEventReader reader, WorkbookBuilder workbookBuilder, String sheetname)
            throws XMLStreamException
    {
        WorksheetBuilder sheetBuilder = workbookBuilder.createSheet(sheetname);

        int rowIndex = 0;

        while (reader.hasNext())
        {
            XMLEvent event = reader.nextEvent();

            if (XMLEvents.isEnd(event, "table"))
            {
                return;
            }

            if (XMLEvents.isStart(event, "tr"))
            {
                sheetBuilder.createRow(rowIndex++);
                parseRow(reader, workbookBuilder, sheetBuilder);
            }
        }
    }

    private void parseRow(XMLEventReader reader, WorkbookBuilder workbookBuilder, WorksheetBuilder sheetBuilder)
            throws XMLStreamException
    {
        int cellIndex = 0;
        while (reader.hasNext())
        {
            XMLEvent event = reader.nextEvent();

            if (XMLEvents.isEnd(event, "tr")) return;

            if (XMLEvents.isStart(event, "th") || XMLEvents.isStart(event, "td"))
            {
                sheetBuilder.createCell(cellIndex++);
                parseCell(reader, workbookBuilder, sheetBuilder);
            }
        }
    }

    private void parseCell(XMLEventReader reader, WorkbookBuilder workbookBuilder, WorksheetBuilder sheetBuilder)
            throws XMLStreamException
    {
        while (reader.hasNext())
        {
            XMLEvent event = reader.nextEvent();

            if (XMLEvents.isEnd(event, "th") || XMLEvents.isEnd(event, "td"))
            {
                return;
            }

            if (XMLEvents.isStart(event, "table"))
            {
                // WOHOO, a nested table
                ++_nestedTableCount;
                String newSheetName = sheetBuilder.getCurrentSheetname() + " nested Table " + _nestedTableCount;
                internalParseTable(reader, workbookBuilder, newSheetName);
                sheetBuilder.setHyperlinkToSheet(newSheetName);
            }
            else if (event.isCharacters())
            {
                String data = "" + event.asCharacters()
                    .getData();
                sheetBuilder.setCellText(data.trim());
            }
            else if (XMLEvents.isStart(event, "image"))
            {
                _imageParser.parseImage(reader, sheetBuilder);
            }
        }
    }
}