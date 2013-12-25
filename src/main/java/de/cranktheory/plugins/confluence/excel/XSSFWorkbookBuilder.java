package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Preconditions;

public class XSSFWorkbookBuilder implements WorkbookBuilder
{
    private XSSFWorkbook _workBook;
    private XSSFSheet _currentSheet;
    private XSSFDrawing _currentSheetDrawing;
    private XSSFRow _currentRow;
    private XSSFCell _currentCell;

    public XSSFWorkbookBuilder()
    {
        _workBook = new XSSFWorkbook();
    }

    @Override
    public void createSheet(String title)
    {
        Preconditions.checkState(_workBook != null, "You have to call createWorkbook first.");

        _currentSheet = _workBook.createSheet(WorkbookUtil.createSafeSheetName(title));
        _currentSheetDrawing = _currentSheet.createDrawingPatriarch();
    }

    @Override
    public void createRow(int index)
    {
        Preconditions.checkState(_currentSheet != null, "You have to call createSheet first.");

        _currentRow = _currentSheet.createRow(index);
    }

    @Override
    public void createCell(int i)
    {
        Preconditions.checkState(_currentRow != null, "You have to call createRow first.");

        _currentCell = _currentRow.createCell(i);
    }

    @Override
    public void addTextToCell(String cellValue)
    {
        Preconditions.checkState(_currentCell != null, "You have to call createCell first.");

        _currentCell.setCellValue(cellValue);
    }

    @Override
    public void drawPictureToCell(byte[] imageInByte, int imageType) throws PictureDrawingException
    {
        Preconditions.checkState(_currentCell != null, "You have to call createCell first.");

        ClientAnchor anchor = createAnchor(_currentCell);
        int pictureIndex = _workBook.addPicture(imageInByte, imageType);
        Picture pict = _currentSheetDrawing.createPicture(anchor, pictureIndex);
        try
        {
            pict.resize();
        }
        catch (Exception e)
        {
            // pict.resize() is ugly, throwing NPEs without contextual information about the real
            // error (FUCK YOU, Apache). Therefore we catch ALL exceptions here.
            // TODO Add context info to exception like ImageIO registered readers, etc.
            throw new PictureDrawingException("Error during picture drawing - resize failed", e);
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

    @Override
    public Workbook getWorkbook() throws IOException
    {
        return _workBook;
    }
}