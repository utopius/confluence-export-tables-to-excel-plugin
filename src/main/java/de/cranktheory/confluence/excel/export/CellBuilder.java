package de.cranktheory.confluence.excel.export;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

public interface CellBuilder
{
    /**
     * Start an unordered list in the {@link Cell}. Call {@link #startListItem()}, add the item with
     * {@link #appendText(String)} and finish the item with {@link #endListItem()}.
     */
    void startUnorderedList();

    /**
     * Start an ordered list in the {@link Cell}. Call {@link #startListItem()}, add the item with
     * {@link #appendText(String)} and finish the item with {@link #endListItem()}.
     */
    void startOrderedList();

    /**
     * Starts a list item. Calls to {@link #appendText(String)} will be added to the current list item.
     */
    void startListItem();

    /**
     * Ends a list item.
     */
    void endListItem();

    /**
     * Appends the given <code>text</code> to the {@link Cell}.
     *
     * @param text the text to append.
     */
    void appendText(String text);

    /**
     *
     * Draws the given {@code imageInByte} to the current cell using the given {@code imageType}.
     *
     * @param imageInByte
     *            the image in bytes
     * @param imageType
     *            either {@link Workbook#PICTURE_TYPE_JPEG} or {@link Workbook#PICTURE_TYPE_PNG}
     * @throws PictureDrawingException
     *             thrown when the drawing fails.
     */
    void drawPictureToCell(byte[] imageInByte, int imageType) throws PictureDrawingException;

    /**
     * Adds a hyperlink to the given {@code sheetName}.
     *
     * @param sheetName
     *            the name of the sheet to link to.
     */
    void setHyperlinkToSheet(String sheetName);

    /**
     * Writes the contents to the cell.
     */
    void build();
}
