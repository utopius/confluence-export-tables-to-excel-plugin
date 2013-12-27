package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookBuilder
{
    public abstract Workbook getWorkbook() throws IOException;

    public abstract WorksheetBuilder createSheet(String title);
}