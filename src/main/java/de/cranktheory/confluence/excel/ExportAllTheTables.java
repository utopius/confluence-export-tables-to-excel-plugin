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
 * Exports all tables on a Page.
 */
public class ExportAllTheTables implements WorkbookExporter
{
    private final WorkbookBuilder _builder;
    private final ParserFactory _parserFactory;

    public ExportAllTheTables(WorkbookBuilder builder, ParserFactory parserFactory)
    {
        _builder = builder;
        _parserFactory = parserFactory;
    }

    @Override
    public Workbook export(final XMLEventReader reader) throws XMLStreamException
    {
        Preconditions.checkNotNull(reader, "reader");

        MacroParser parser = _parserFactory.newMacroParser();
        parser.parse(reader);

        return _builder.getWorkbook();
    }
}