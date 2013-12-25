package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Preconditions;

public class WorkbookBuilder
{
    private XSSFWorkbook _workBook;
    private XSSFSheet _currentSheet;
    private XSSFDrawing _currentSheetDrawing;
    private XSSFRow _currentRow;
    private XSSFCell _currentCell;

    public WorkbookBuilder()
    {
    }

    public WorkbookBuilder createWorkbook()
    {
        _workBook = new XSSFWorkbook();
        return this;
    }

    public WorkbookBuilder createSheet(String title)
    {
        Preconditions.checkState(_workBook != null, "You have to call createWorkbook first.");

        _currentSheet = _workBook.createSheet(title);
        _currentSheetDrawing = _currentSheet.createDrawingPatriarch();
        return this;
    }

    public void createRow(int index)
    {
        Preconditions.checkState(_currentSheet != null, "You have to call createSheet first.");
        _currentRow = _currentSheet.createRow(index);
    }

    public void createCell(int i)
    {
        Preconditions.checkState(_currentRow != null, "You have to call createRow first.");
        _currentCell = _currentRow.createCell(i);
    }

    public void addTextToCell(String cellValue)
    {
        Preconditions.checkState(_currentCell != null, "You have to call createCell first.");
        _currentCell.setCellValue(cellValue);
    }

    public void drawPictureToCell(byte[] imageInByte, int imageType)
    {
        Preconditions.checkState(_currentCell != null, "You have to call createCell first.");
        Picture pict = _currentSheetDrawing.createPicture(createAnchor(_currentCell), _workBook.addPicture(imageInByte, imageType));
        try
        {
            pict.resize();
        }
        catch (Exception e)
        {
            // pict.resize() is ugly, throwing NPEs without contextual information about the real
            // error (FUCK YOU, Apache).
            // TODO Log possible error causes like ImageIO registered readers, etc.
            e.printStackTrace();
        }
    }

    private ClientAnchor createAnchor(XSSFCell cell)
    {
        ClientAnchor anchor = _workBook.getCreationHelper()
            .createClientAnchor();
        int columnIndex = cell.getColumnIndex();
        int rowIndex = cell.getRowIndex();
        anchor.setCol1(columnIndex);
        anchor.setRow1(rowIndex);
        anchor.setCol2(columnIndex + 1);
        anchor.setRow2(rowIndex + 1);
        return anchor;
    }

    public XSSFWorkbook build() throws IOException
    {
        return _workBook;
    }
}