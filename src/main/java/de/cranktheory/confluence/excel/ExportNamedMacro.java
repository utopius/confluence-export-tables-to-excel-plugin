package de.cranktheory.confluence.excel;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;

import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.base.Preconditions;

import de.cranktheory.confluence.excel.export.MacroParser;
import de.cranktheory.confluence.excel.export.ParserFactory;
import de.cranktheory.confluence.excel.export.WorkbookBuilder;
import de.cranktheory.confluence.excel.export.WorkbookExporter;

/**
 * Exports tables from a page which are decorated with the {@link ExportTableMacro}.
 */
public class ExportNamedMacro implements WorkbookExporter
{
    private final WorkbookBuilder _builder;
    private final ParserFactory _parserFactory;
    private String _sheetToExport;

    public ExportNamedMacro(WorkbookBuilder builder, ParserFactory parserFactory, String sheetToExport)
    {
        _builder = builder;
        _parserFactory = parserFactory;
        _sheetToExport = sheetToExport;
    }

    @Override
    public Workbook export(final XMLEventReader reader) throws XMLStreamException
    {
        Preconditions.checkNotNull(reader, "reader");

        MacroParser parser = _parserFactory.newMacroParser(_sheetToExport);
        parser.parse(reader);

        return _builder.getWorkbook();
    }
}
