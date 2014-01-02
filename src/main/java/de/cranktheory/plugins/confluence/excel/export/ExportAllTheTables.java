package de.cranktheory.plugins.confluence.excel.export;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.XMLEvent;

import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.base.Preconditions;

public class ExportAllTheTables implements WorkbookExporter
{
    private final WorkbookBuilder _builder;

    private TableParser _tableParser;

    public ExportAllTheTables(WorkbookBuilder builder, TableParser tableParser)
    {
        _builder = builder;
        _tableParser = tableParser;
    }

    @Override
    public Workbook export(final XMLEventReader reader) throws XMLStreamException
    {
        Preconditions.checkNotNull(reader, "reader");

        int nonMacroTableCount = 0;

        while (reader.hasNext())
        {
            XMLEvent event = reader.nextEvent();

            if (isExportTableMacro(event))
            {
                String sheetname = parseSheetname(reader);

                parseMacro(reader, sheetname, _tableParser);
            }
            else if (XMLEvents.isStart(event, "table"))
            {
                // Applies to all tables which are not enclosed in an export-table macro
                ++nonMacroTableCount;
                _tableParser.parseTable(reader, _builder, String.format("Table %s", nonMacroTableCount));
            }
        }

        return _builder.getWorkbook();
    }

    private boolean isExportTableMacro(XMLEvent event)
    {
        if (XMLEvents.isStart(event, "structured-macro"))
        {
            Attribute macroNameAttribute = event.asStartElement()
                .getAttributeByName(QName.valueOf("{http://atlassian.com/content}name"));

            return "export-table".equals(macroNameAttribute.getValue());
        }

        return false;
    }

    private String parseSheetname(XMLEventReader reader) throws XMLStreamException
    {
        while (nextIsParameter(reader))
        {
            do
            {
                // There MUST be at least the sheetname parameter
                Attribute parameterName = reader.nextEvent()
                    .asStartElement()
                    .getAttributeByName(QName.valueOf("{http://atlassian.com/content}name"));

                XMLEvent parameterContent = reader.nextEvent();

                if ("sheetname".equalsIgnoreCase(parameterName.getValue()))
                {
                    Preconditions.checkState(parameterContent.isCharacters(),
                            "Dafuq?! The parameter value is not a character event.");

                    String sheetname = parameterContent.asCharacters()
                        .getData();
                    return sheetname;
                }
            } while (!XMLEvents.isEnd(reader.nextEvent(), "parameter"));
        }

        throw new IllegalStateException(String.valueOf("No Macro parameter elements found."));
    }

    private void parseMacro(XMLEventReader reader, String sheetname, TableParser tableParser) throws XMLStreamException
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
                tableParser.parseTable(reader, _builder, sheetname);
            }
        }
    }

    private boolean nextIsParameter(XMLEventReader reader) throws XMLStreamException
    {
        XMLEvent peek = reader.peek();
        boolean nextIsParameter = peek != null && XMLEvents.isStart(peek, "parameter");
        return nextIsParameter;
    }

}