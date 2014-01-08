package de.cranktheory.confluence.excel.export.xssf;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;

import de.cranktheory.confluence.excel.export.WorkbookBuilder;
import de.cranktheory.confluence.excel.export.WorksheetBuilder;

public class XSSFWorkbookBuilder implements WorkbookBuilder
{
    private final XSSFWorkbook _workBook;

    public XSSFWorkbookBuilder()
    {
        _workBook = new XSSFWorkbook();
    }

    @Override
    public WorksheetBuilder createSheet(String sheetName)
    {
        Preconditions.checkState(!Strings.isNullOrEmpty(sheetName), "sheetName is null or empty.");

        String safeSheetName = WorkbookUtil.createSafeSheetName(sheetName);
        XSSFSheet _currentSheet = _workBook.createSheet(safeSheetName);

        return XSSFWorksheetBuilder.newInstance(_workBook, _currentSheet);
    }

    @Override
    public Workbook getWorkbook()
    {
        return _workBook;
    }
}