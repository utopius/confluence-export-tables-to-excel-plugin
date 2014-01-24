package de.cranktheory.confluence.excel.export;

import org.apache.poi.ss.usermodel.Workbook;

public interface CellBuilder
{

    public abstract void build();

    public abstract void appendText(String text);

    public abstract void endListItem();

    public abstract void startListItem();

    public abstract void startUnorderedList();

    public abstract void startOrderedList();

    /**
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
}
