package de.cranktheory.confluence.excel.export;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;

import org.apache.commons.lang3.StringEscapeUtils;

import com.google.common.base.Preconditions;

public class MacroParser
{
    private final String _sheetToExport;
    private final ParserFactory _parserFactory;

    public MacroParser(ParserFactory parserFactory, String sheetnameToExport)
    {
        _parserFactory = parserFactory;
        _sheetToExport = Preconditions.checkNotNull(sheetnameToExport, "sheetToExport");
    }

    public MacroParser(ParserFactory parserFactory)
    {
        _parserFactory = parserFactory;
        _sheetToExport = null;
    }

    public void parse(final XMLEventReader reader) throws XMLStreamException
    {
        int nonMacroTableCount = 0;
        while (reader.hasNext())
        {
            XMLEvent event = reader.nextEvent();

            if (XMLEvents.isStartExportMacro(event))
            {
                String sheetname = parseSheetname(reader);
                if (_sheetToExport != null && _sheetToExport.equalsIgnoreCase(sheetname))
                {
                    // Only parse the wanted macro and then stop
                    parseMacro(reader, sheetname);
                    return;
                }
                else
                {
                    // Parse all macros
                    parseMacro(reader, sheetname);
                }
            }
            else if (_sheetToExport == null && XMLEvents.isStart(event, "table"))
            {
                // Parse all tables if no specific macro is searched for.
                ++nonMacroTableCount;
                String sheetName = String.format("Table %s", nonMacroTableCount);
                _parserFactory.newTableParser(sheetName)
                    .parseTable(reader);
            }
        }
    }

    private String parseSheetname(XMLEventReader reader) throws XMLStreamException
    {
        while (XMLEvents.nextIsParameter(reader))
        {
            do
            {
                // There MUST be at least the sheetname parameter
                if ("sheetname".equalsIgnoreCase(reader.nextEvent()
                    .asStartElement()
                    .getAttributeByName(QName.valueOf("{http://atlassian.com/content}name"))
                    .getValue()))
                {
                    final String elementText = reader.getElementText();
                    final String sheetname = StringEscapeUtils.unescapeHtml4(elementText);

                    return sheetname;
                }
            } while (!XMLEvents.isEnd(reader.nextEvent(), "parameter"));
        }

        throw new IllegalStateException(
                String.valueOf("No Macro parameter elements found. To export a specific table you must set the parameter 'sheetname'."));
    }

    private void parseMacro(XMLEventReader reader, String sheetname) throws XMLStreamException
    {
        while (reader.hasNext())
        {
            XMLEvent event = reader.nextEvent();

            boolean endTagReached = XMLEvents.isEnd(event, "structured-macro");
            if (endTagReached)
            {
                break;
            }

            // final boolean foundNestedMacro = Util.isStart(event, "structured-macro");
            // if (foundNestedMacro)
            // {
            // MacroData nestedMacro = parseMacro(reader, event);
            // }

            if (XMLEvents.isStart(event, "table"))
            {
                _parserFactory.newTableParser(sheetname)
                    .parseTable(reader);
            }
        }
    }
}