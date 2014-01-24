package de.cranktheory.confluence.excel.export;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.EntityReference;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

import com.google.common.base.Strings;

public class CellParser
{
    private final ParserFactory _factory;
    private final ImageParser _imageParser;
    private int _nestedTableCount = 0;

    public CellParser(ParserFactory factory)
    {
        _factory = factory;
        _imageParser = factory.newImageParser();
    }

    public void parseCell(XMLEventReader reader, WorksheetBuilder sheetBuilder, StartElement startElement)
            throws XMLStreamException
    {
        StringBuilder builder = new StringBuilder();

        while (reader.hasNext())
        {
            XMLEvent event = reader.nextEvent();

            if (event == null) return;
            if (XMLEvents.isEnd(event, "th") || XMLEvents.isEnd(event, "td")) break;

            if (XMLEvents.isStart(event, "table"))
            {
                // WOHOO, a nested table
                ++_nestedTableCount;
                String newSheetName = sheetBuilder.getCurrentSheetname() + " nested Table " + _nestedTableCount;

                _factory.newTableParser(newSheetName)
                    .parseTable(reader);
                sheetBuilder.setHyperlinkToSheet(newSheetName);
            }
            else if (XMLEvents.isStart(event, "image"))
            {
                _imageParser.parseImage(reader, sheetBuilder);
            }
            else if (event.isCharacters() || event.isEntityReference())
            {
                parseText(event, builder);
            }
        }

        String cellValue = builder.toString();
        if (!Strings.isNullOrEmpty(cellValue))
        {
            sheetBuilder.setCellText(cellValue);
        }
    }

    private void parseText(XMLEvent event, StringBuilder builder) throws XMLStreamException
    {
        if (event.isEntityReference())
        {
            EntityReference ref = (EntityReference) event;
            builder.append(ref.getDeclaration()
                .getReplacementText());
        }
        else if (event.isCharacters())
        {
            builder.append(event.asCharacters()
                .getData());

        }
    }
}