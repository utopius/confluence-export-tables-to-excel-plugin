package de.cranktheory.confluence.excel.export.xssf;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;

import de.cranktheory.confluence.excel.export.CellBuilder;
import de.cranktheory.confluence.excel.export.PictureDrawingException;

public class XSSFCellBuilder implements CellBuilder
{
    private static final char BULLET_CHARACTER = '\u2022';

    private final StringBuilder _builder = new StringBuilder();
    private final XSSFWorkbook _workbook;
    private final XSSFCell _currentCell;
    private final XSSFDrawing _drawing;

    private int _number = -1;
    private boolean _orderedList = false;

    private XSSFCellStyle _cellStyle;

    public XSSFCellBuilder(XSSFWorkbook workbook, XSSFCell cell, XSSFDrawing drawing)
    {
        _workbook = workbook;
        _currentCell = cell;
        _drawing = drawing;
        _cellStyle = _workbook.createCellStyle();
        _currentCell.setCellStyle(_cellStyle);
    }

    @Override
    public void startOrderedList()
    {
        _number = 1;
        _orderedList = true;

        _cellStyle.setWrapText(true);
    }

    @Override
    public void startUnorderedList()
    {
        _number = -1;
        _orderedList = false;

        _cellStyle.setWrapText(true);
    }

    @Override
    public void startListItem()
    {
        if (_orderedList)
        {
            Preconditions.checkState(_orderedList && _number > 0, "You must call startOrderedList() first.");

            String listCharacter = String.valueOf(_number++) + ". ";
            _builder.append(listCharacter);
        }
        else
        {
            _builder.append(BULLET_CHARACTER + " ");
        }
    }

    @Override
    public void endListItem()
    {
        _builder.append("\n");
    }

    @Override
    public void appendText(String text)
    {
        _builder.append(text);
    }

    @Override
    public void build()
    {
        String cellValue = _builder.toString();
        if (!Strings.isNullOrEmpty(cellValue))
        {
            setCellText(cellValue);
        }
    }

    private void setCellText(String cellValue)
    {
        Preconditions.checkState(_currentCell != null, "You have to call createCell first.");

        _currentCell.setCellValue(new XSSFRichTextString(cellValue));
    }

    @Override
    public void drawPictureToCell(byte[] imageInByte, int imageType) throws PictureDrawingException
    {
        Preconditions.checkState(_currentCell != null, "You have to call createCell first.");

        ClientAnchor anchor = createAnchor(_currentCell);
        int pictureIndex = _workbook.addPicture(imageInByte, imageType);
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
        ClientAnchor anchor = _workbook.getCreationHelper()
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
        XSSFCreationHelper creationHelper = _workbook.getCreationHelper();

        Hyperlink link = creationHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
        link.setAddress(String.format("'%s'!A1", sheetName));
        _builder.append(" " + sheetName);
        _currentCell.setHyperlink(link);

        // cell style for hyperlinks
        // by default hyperlinks are blue and underlined
        Font hlinkFont = _workbook.createFont();
        hlinkFont.setUnderline(Font.U_SINGLE);
        hlinkFont.setColor(IndexedColors.BLUE.getIndex());
        _cellStyle.setFont(hlinkFont);

        _currentCell.setCellStyle(_cellStyle);
    }
}