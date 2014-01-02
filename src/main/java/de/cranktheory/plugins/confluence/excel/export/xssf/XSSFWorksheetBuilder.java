package de.cranktheory.plugins.confluence.excel.export.xssf;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Preconditions;

import de.cranktheory.plugins.confluence.excel.export.PictureDrawingException;
import de.cranktheory.plugins.confluence.excel.export.WorksheetBuilder;

public class XSSFWorksheetBuilder implements WorksheetBuilder
{
    private final XSSFWorkbook _workBook;
    private final XSSFSheet _currentSheet;
    private final XSSFDrawing _currentSheetDrawing;

    private XSSFRow _currentRow;
    private XSSFCell _currentCell;

    public XSSFWorksheetBuilder(XSSFWorkbook workBook, XSSFSheet currentSheet)
    {
        _workBook = workBook;
        _currentSheet = currentSheet;
        _currentSheetDrawing = _currentSheet.createDrawingPatriarch();
    }

    @Override
    public String getCurrentSheetname()
    {
        return _currentSheet.getSheetName();
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
    public void setCellText(String cellValue)
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
    public void setHyperlinkToSheet(String sheetName)
    {
        XSSFCreationHelper creationHelper = _workBook.getCreationHelper();

        Hyperlink link = creationHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
        link.setAddress("'" + sheetName + "'!A1");
        _currentCell.setCellValue(sheetName);
        _currentCell.setHyperlink(link);
        _currentCell.setCellStyle(createHyperlinkStyle());
    }

    private CellStyle createHyperlinkStyle()
    {
        // cell style for hyperlinks
        // by default hyperlinks are blue and underlined
        CellStyle hlinkStyle = _workBook.createCellStyle();
        Font hlinkFont = _workBook.createFont();
        hlinkFont.setUnderline(Font.U_SINGLE);
        hlinkFont.setColor(IndexedColors.BLUE.getIndex());
        hlinkStyle.setFont(hlinkFont);
        return hlinkStyle;
    }
}
