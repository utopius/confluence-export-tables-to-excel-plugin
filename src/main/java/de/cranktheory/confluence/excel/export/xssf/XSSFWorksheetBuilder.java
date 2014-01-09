package de.cranktheory.confluence.excel.export.xssf;

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

import de.cranktheory.confluence.excel.export.PictureDrawingException;
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
    public void createCell(int index)
    {
        Preconditions.checkState(_currentRow != null, "You have to call createRow first.");

        _currentCell = _currentRow.createCell(index);
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
        Picture pict = _drawing.createPicture(anchor, pictureIndex);
        try
        {
            pict.resize();
        }
        catch (Exception e)
        {
            // pict.resize() is ugly, throwing NPEs without contextual information about the real
            // error (thanks@Apache). Therefore we catch ALL exceptions here.
            // TODO Add context info to exception like ImageIO registered readers, etc.
            throw new PictureDrawingException("Error during picture drawing - resize failed", e);
        }
    }

    private ClientAnchor createAnchor(XSSFCell cell)
    {
        ClientAnchor anchor = _workBook.getCreationHelper()
            .createClientAnchor();

        // TODO: Improve image anchoring
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
        link.setAddress(String.format("'%s'!A1", sheetName));
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
