package de.cranktheory.confluence.excel.export.xssf;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Preconditions;

import de.cranktheory.confluence.excel.export.CellBuilder;
import de.cranktheory.confluence.excel.export.Link;
import de.cranktheory.confluence.excel.export.PictureDrawingException;

public class XSSFCellBuilder implements CellBuilder
{
    private static final char BULLET_CHARACTER = '\u2022';

    private final StringBuilder _builder = new StringBuilder();
    private final XSSFWorkbook _workbook;
    private final XSSFDrawing _drawing;
    private final XSSFCell _cell;

    private int _number = -1;
    private boolean _orderedList = false;

    private XSSFCellStyle _cellStyle;

    public XSSFCellBuilder(XSSFWorkbook workbook, XSSFCell cell, XSSFDrawing drawing)
    {
        _workbook = workbook;
        _cell = cell;
        _drawing = drawing;
        _cellStyle = _workbook.createCellStyle();

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
        String cellValue = _builder.toString()
            .trim();
        if (!cellValue.isEmpty())
        {
            _cell.setCellValue(new XSSFRichTextString(cellValue));
            _cell.setCellStyle(_cellStyle);
        }
    }

    @Override
    public void drawPictureToCell(byte[] imageInByte, int imageType) throws PictureDrawingException
    {
        _drawing.createPicture(createAnchor(_cell), _workbook.addPicture(imageInByte, imageType));
    }

    private ClientAnchor createAnchor(XSSFCell cell)
    {
        ClientAnchor anchor = _workbook.getCreationHelper()
            .createClientAnchor();

        int columnIndex = cell.getColumnIndex();
        int rowIndex = cell.getRowIndex();

        anchor.setCol1(columnIndex);
        anchor.setRow1(rowIndex);
        anchor.setAnchorType(ClientAnchor.MOVE_DONT_RESIZE);
        anchor.setCol2(++columnIndex);
        anchor.setRow2(++rowIndex);

        return anchor;
    }

    @Override
    public void setHyperlinkToSheet(String sheetName)
    {
        XSSFCreationHelper creationHelper = _workbook.getCreationHelper();

        Hyperlink link = creationHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
        link.setAddress(String.format("'%s'!A1", sheetName));
        _builder.append(" " + sheetName);
        _cell.setHyperlink(link);

        // cell style for hyperlinks
        // by default hyperlinks are blue and underlined
        Font hlinkFont = _workbook.createFont();
        hlinkFont.setUnderline(Font.U_SINGLE);
        hlinkFont.setColor(IndexedColors.BLUE.getIndex());
        _cellStyle.setFont(hlinkFont);

        _cell.setCellStyle(_cellStyle);
    }

    @Override
    public void setHyperlink(Link link)
    {
        XSSFCreationHelper creationHelper = _workbook.getCreationHelper();
        Hyperlink hyperlink = creationHelper.createHyperlink(Hyperlink.LINK_URL);
        hyperlink.setAddress(link.getUrl());
        hyperlink.setLabel(link.getLabel());
        _builder.append(link.getLabel());
        _cell.setHyperlink(hyperlink);

        // cell style for hyperlinks
        // by default hyperlinks are blue and underlined
        Font hlinkFont = _workbook.createFont();
        hlinkFont.setUnderline(Font.U_SINGLE);
        hlinkFont.setColor(IndexedColors.BLUE.getIndex());
        _cellStyle.setFont(hlinkFont);

        _cell.setCellStyle(_cellStyle);
    }
}