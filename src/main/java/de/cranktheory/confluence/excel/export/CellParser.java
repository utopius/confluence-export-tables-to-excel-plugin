package de.cranktheory.confluence.excel.export;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.EntityReference;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

public class CellParser
{
    private final ParserFactory _factory;
    private final ImageParser _imageParser;
    private int _nestedTableCount = 0;
    private LinkParser _linkParser;

    public CellParser(ParserFactory factory, ImageParser imageParser, LinkParser linkParser)
    {
        _factory = factory;
        _imageParser = imageParser;
        _linkParser = linkParser;
    }

    public void parseCell(XMLEventReader reader, WorksheetBuilder sheetBuilder, CellBuilder cellBuilder,
            StartElement startElement) throws XMLStreamException
    {
        while (reader.hasNext())
        {
            XMLEvent event = reader.nextEvent();

            if (XMLEvents.isEnd(event, "th") || XMLEvents.isEnd(event, "td")) break;

            if (XMLEvents.isStart(event, "table"))
            {
                // WOHOO, a nested table
                ++_nestedTableCount;
                String newSheetName = sheetBuilder.getCurrentSheetname() + " nested Table " + _nestedTableCount;

                _factory.newTableParser(newSheetName)
                    .parseTable(reader);
                cellBuilder.setHyperlinkToSheet(newSheetName);
            }
            else if (XMLEvents.isStart(event, "image"))
            {
                _imageParser.parseImage(reader, sheetBuilder, cellBuilder);
            }
            else if (XMLEvents.isStart(event, "ol"))
            {
                cellBuilder.startOrderedList();
            }
            else if (XMLEvents.isStart(event, "ul"))
            {
                cellBuilder.startUnorderedList();
            }
            if (XMLEvents.isStart(event, "li"))
            {
                cellBuilder.startListItem();
            }
            else if (XMLEvents.isEnd(event, "li"))
            {
                cellBuilder.endListItem();
            }
            else if(_linkParser.isLink(event))
            {
                Link link = _linkParser.parseLink(reader, event);
                cellBuilder.setHyperlink(link);
            }
            else if (event.isCharacters() || event.isEntityReference())
            {
                cellBuilder.appendText(parseText(event));
            }
        }

        cellBuilder.build();
    }

    private String parseText(XMLEvent event) throws XMLStreamException
    {
        if (event.isEntityReference())
        {
            EntityReference ref = (EntityReference) event;
            return ref.getDeclaration()
                .getReplacementText();
        }
        else if (event.isCharacters())
        {
            return event.asCharacters()
                .getData();
        }
        return "";
    }
}
