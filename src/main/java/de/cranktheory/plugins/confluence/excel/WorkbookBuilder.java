package de.cranktheory.plugins.confluence.excel;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookBuilder
{
    public abstract Workbook getWorkbook();

    public abstract WorksheetBuilder createSheet(String title);
}