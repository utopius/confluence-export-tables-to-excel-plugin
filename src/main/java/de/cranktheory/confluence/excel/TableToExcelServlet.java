package de.cranktheory.confluence.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.net.URL;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.base.Strings;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class TableToExcelServlet extends HttpServlet
{

    private static final Logger log = LoggerFactory.getLogger(TableToExcelServlet.class);

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException
    {
        String id = req.getParameter("id");
        String tableData = req.getParameter("tabledata");

        log.info("TableData: " + tableData);
        System.out.println("TableToExcel id: " + id);
        System.out.println("TableToExcel tabledata: " + tableData);

        if (Strings.isNullOrEmpty(tableData)) tableData = "Parameter tabledata was null or empty.";

        resp.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        resp.setHeader("Content-disposition", "attachment; filename=" + id + ".xlsx");

        JsonObject jsonTable = new JsonParser().parse(tableData)
            .getAsJsonObject();

        XSSFWorkbook workBook = new XSSFWorkbook();
        XSSFSheet sheet = workBook.createSheet(id);

        parseSheet(workBook, sheet, jsonTable);

        OutputStream out = resp.getOutputStream();
        workBook.write(out);
        out.flush();
    }

    public static void parseSheet(XSSFWorkbook workBook, XSSFSheet sheet, JsonObject jsonTable)
    {
        JsonArray rows = jsonTable.getAsJsonArray("rows");

        for (int rowIDX = 0; rowIDX < rows.size(); ++rowIDX)
        {
            JsonObject jsonRow = rows.get(rowIDX)
                .getAsJsonObject();
            System.out.println("TableToExcel found new row: " + jsonRow);

            parseRow(workBook, sheet, sheet.createRow(rowIDX), jsonRow);
        }
    }

    private static void parseRow(XSSFWorkbook workBook, XSSFSheet sheet, XSSFRow row, JsonObject jsonRow)
    {
        // boolean isHeader = jsonRow.get("isHeaderRow")
        // .getAsBoolean();

        JsonArray cells = jsonRow.get("cells")
            .getAsJsonArray();

        for (int i = 0; i < cells.size(); i++)
        {
            JsonObject cellJsonObject = cells.get(i)
                .getAsJsonObject();

            parseCell(workBook, sheet, row.createCell(i), cellJsonObject);
        }
    }

    private static void parseCell(XSSFWorkbook workBook, XSSFSheet sheet, XSSFCell cell, JsonObject jsonCell)
    {
        if (jsonCell.has("iconUrl"))
        {
            try
            {
                URL iconUrl = new URL(jsonCell.get("iconUrl")
                    .getAsString());

                String path = iconUrl.getPath();

                if (path.contains("."))
                {
                    String extension = path.substring(path.lastIndexOf("."));

                    if (!".png".equalsIgnoreCase(extension))
                    {
                        System.out.println("TableToExcel - extension '" + extension + "' is currently not supported...");
                        cell.setCellValue("Only png images supported");
                    }
                    else
                    {
                        byte[] bytes = loadImageFromURL(iconUrl);

                        int pictureIdx = workBook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);

                        ClientAnchor anchor = workBook.getCreationHelper()
                            .createClientAnchor();
                        anchor.setCol1(cell.getColumnIndex());
                        anchor.setRow1(cell.getRowIndex());

                        Picture pict = sheet.createDrawingPatriarch()
                            .createPicture(anchor, pictureIdx);
                        pict.resize();
                    }
                }
            }
            catch (MalformedURLException e)
            {
                e.printStackTrace();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
        else
        {
            String cellValue = jsonCell.get("text")
                .getAsString();
            cell.setCellValue(cellValue);
            System.out.println("TableToExcel found new cell: " + cellValue);
        }
    }

    private static byte[] loadImageFromURL(URL iconUrl) throws IOException
    {
        InputStream iconStream = iconUrl.openStream();
        byte[] bytes = IOUtils.toByteArray(iconStream);
        iconStream.close();
        return bytes;
    }
}
