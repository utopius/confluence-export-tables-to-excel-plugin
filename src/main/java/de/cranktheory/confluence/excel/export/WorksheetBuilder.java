package de.cranktheory.confluence.excel.export;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorksheetBuilder
{

    /**
     * @return the name of the current sheet.
     */
    String getCurrentSheetname();

    /**
     * Creates a row at the given {@code index} in the current sheet.
     *
     * @param index
     *            Well, the index.
     */
    void createRow(int index);

    /**
     * Creates a cell at the given {@code index} in the current sheet.
     *
     * @param index
     *            Well, the index.
     */
    void createCell(int index);

    /**
     * Adds the given {@code cellValue} as text to the current cell.
     *
     * @param cellValue
     */
    void setCellText(String cellValue);

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
