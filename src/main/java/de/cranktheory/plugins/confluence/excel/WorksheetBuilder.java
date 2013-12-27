package de.cranktheory.plugins.confluence.excel;


public interface WorksheetBuilder
{

    public abstract void createRow(int index);

    public abstract void createCell(int i);

    public abstract void setCellText(String cellValue);

    public abstract void drawPictureToCell(byte[] imageInByte, int imageType) throws PictureDrawingException;

    String getCurrentSheetname();

    public abstract void setHyperlinkToSheet(String sheetName);
}
