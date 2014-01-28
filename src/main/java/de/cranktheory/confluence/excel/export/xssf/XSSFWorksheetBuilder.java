package de.cranktheory.confluence.excel.export.xssf;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
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

    private final XSSFWorkbook _workbook;
    private final XSSFSheet _sheet;
    private final XSSFDrawing _drawing;

    private XSSFRow _currentRow;
    private int _lastColumnIndex = -1;

    private XSSFWorksheetBuilder(XSSFWorkbook workBook, XSSFSheet currentSheet)
    {
        _workbook = workBook;
        _sheet = currentSheet;
        _drawing = _sheet.createDrawingPatriarch();
    }

    @Override
    public String getCurrentSheetname()
    {
        return _sheet.getSheetName();
    }

    @Override
    public void createRow(int index)
    {
        _currentRow = _sheet.createRow(index);
    }

    @Override
    public CellBuilder createCell(int index)
    {
        Preconditions.checkState(_currentRow != null, "You have to call createRow first.");

        XSSFCell cell = _currentRow.createCell(index);

        if (index > _lastColumnIndex)
        {
            _lastColumnIndex = index;
        }

        return new XSSFCellBuilder(_workbook, cell, _drawing);
    }

    @Override
    public void build()
    {
        for (int columnIndex = 0; columnIndex < _lastColumnIndex; ++columnIndex)
        {
            _sheet.autoSizeColumn(columnIndex);
        }
    }
}
