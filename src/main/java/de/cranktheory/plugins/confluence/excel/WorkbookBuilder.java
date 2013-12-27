package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookBuilder
{

    public abstract void createSheet(String title);

    public abstract void createRow(int index);

    public abstract void createCell(int i);

    public abstract void addTextToCell(String cellValue);

    public abstract void drawPictureToCell(byte[] imageInByte, int imageType) throws PictureDrawingException;

    public abstract Workbook getWorkbook() throws IOException;

    String getCurrentSheetname();

}