package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;

public class XSSFWorkbookBuilder implements WorkbookBuilder
{
    private XSSFWorkbook _workBook;

    public XSSFWorkbookBuilder()
    {
        _workBook = new XSSFWorkbook();
    }

    @Override
    public WorksheetBuilder createSheet(String sheetName)
    {
        Preconditions.checkState(_workBook != null, "You have to call createWorkbook first.");
        Preconditions.checkState(!Strings.isNullOrEmpty(sheetName), "sheetName is null or empty.");

        String safeSheetName = WorkbookUtil.createSafeSheetName(sheetName);
        XSSFSheet _currentSheet = _workBook.createSheet(safeSheetName);

        return new XSSFWorksheetBuilder(_workBook, _currentSheet);
    }

    @Override
    public Workbook getWorkbook() throws IOException
    {
        return _workBook;
    }
}