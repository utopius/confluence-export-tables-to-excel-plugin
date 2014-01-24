package de.cranktheory.confluence.excel.export.xssf;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Preconditions;

import de.cranktheory.confluence.excel.export.CellBuilder;
import de.cranktheory.confluence.excel.export.WorksheetBuilder;

public class XSSFWorksheetBuilder implements WorksheetBuilder
{
    public static XSSFWorksheetBuilder newInstance(XSSFWorkbook workBook, XSSFSheet workSheet)
    {
        return new XSSFWorksheetBuilder(Preconditions.checkNotNull(workBook, "workBook"), Preconditions.checkNotNull(
                workSheet, "workSheet"));
    }

    private final XSSFWorkbook _workBook;
    private final XSSFSheet _workSheet;
    private final XSSFDrawing _drawing;

    private XSSFRow _currentRow;
    private XSSFCell _currentCell;

    private XSSFWorksheetBuilder(XSSFWorkbook workBook, XSSFSheet currentSheet)
    {
        _workBook = workBook;
        _workSheet = currentSheet;
        _drawing = _workSheet.createDrawingPatriarch();
    }

    @Override
    public String getCurrentSheetname()
    {
        return _workSheet.getSheetName();
    }

    @Override
    public void createRow(int index)
    {
        _currentRow = _workSheet.createRow(index);
    }

    @Override
    public CellBuilder createCell(int index)
    {
        Preconditions.checkState(_currentRow != null, "You have to call createRow first.");

        _currentCell = _currentRow.createCell(index);
        return new XSSFCellBuilder(_workBook, _currentCell, _drawing);
    }
}
