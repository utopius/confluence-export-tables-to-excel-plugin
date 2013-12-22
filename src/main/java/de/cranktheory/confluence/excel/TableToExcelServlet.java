package de.cranktheory.confluence.excel;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.OutputStream;
import java.io.StringReader;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.base.Strings;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class TableToExcelServlet extends HttpServlet
{

    private static final Logger log = LoggerFactory.getLogger(TableToExcelServlet.class);

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException

    {
        String id = req.getParameter("id");
        String tableData = req.getParameter("tabledata");

        handleRequest(resp, id, tableData);
    }

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException
    {
        String id = req.getParameter("id");
        String tableData = req.getParameter("tabledata");

        handleRequest(resp, id, tableData);
    }

    private void handleRequest(HttpServletResponse resp, String id, String tableData) throws IOException
    {
        log.info("TableData: " + tableData);
        System.out.println("TableToExcel id: " + id);
        System.out.println("TableToExcel tabledata: " + tableData);

        if (Strings.isNullOrEmpty(tableData)) tableData = "Parameter tabledata was null or empty.";

        // You must tell the browser the file type you are going to send
        // for example application/pdf, text/plain, text/html, image/jpg
        resp.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

        // Make sure to show the download dialog
        resp.setHeader("Content-disposition", "attachment; filename=" + id + ".xlsx");

        // This should send the file to browser
        OutputStream out = resp.getOutputStream();

        // writeXlsxFromCSV(id, tableData, out);
        XSSFWorkbook workBook = createWorkbookFromJson(id, tableData, out);
        workBook.write(out);
        out.flush();
    }

    public static XSSFWorkbook createWorkbookFromJson(String title, String json, OutputStream output)
    {
        try
        {
            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet(title);

            JsonObject root = new JsonParser().parse(json)
                .getAsJsonObject();
            JsonArray rows = root.getAsJsonArray("rows");

            for (int rowIDX = 0; rowIDX < rows.size(); ++rowIDX)
            {
                JsonObject jsonRow = rows.get(rowIDX).getAsJsonObject();
                System.out.println("TableToExcel found new row: " + jsonRow);

                boolean isHeader = jsonRow.get("isHeaderRow")
                    .getAsBoolean();

                XSSFRow currentRow = sheet.createRow(rowIDX);

                JsonArray cells = jsonRow.get("cells")
                    .getAsJsonArray();

                for (int i = 0; i < cells.size(); i++)
                {
                    String cellValue = cells.get(i)
                        .getAsString();

                    System.out.println("TableToExcel found new cell: " + cellValue);

                    currentRow.createCell(i)
                        .setCellValue(cellValue);
                }
            }

            return workBook;
        }
        catch (Exception ex)
        {
            System.out.println(ex.getMessage() + "Exception in try");
            return new XSSFWorkbook();
        }
    }

    public static XSSFWorkbook createWorkbookFromCSV(String title, String csv, OutputStream output)
    {
        try
        {
            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet(title);
            String currentLine = null;
            int rowNum = 0;

            String csvWithLineBreaks = csv.replace("%0D%0A", "\r\n");

            BufferedReader br = new BufferedReader(new StringReader(csvWithLineBreaks));

            while ((currentLine = br.readLine()) != null)
            {
                System.out.println("TableToExcel found new line: " + currentLine);

                String[] headerCells = currentLine.split("||");
                boolean hasHeaders = headerCells.length > 0;

                String cells[] = currentLine.split("|");
                boolean hasCells = cells.length > 0;

                ++rowNum;
                XSSFRow currentRow = sheet.createRow(rowNum);

                if (hasHeaders)
                {
                    for (int i = 0; i < headerCells.length; i++)
                    {
                        currentRow.createCell(i)
                            .setCellValue(headerCells[i]);
                    }
                }
                else if (hasCells)
                {
                    for (int i = 0; i < cells.length; i++)
                    {
                        currentRow.createCell(i)
                            .setCellValue(cells[i]);
                    }
                }
            }

            return workBook;
        }
        catch (Exception ex)
        {
            System.out.println(ex.getMessage() + "Exception in try");
            return new XSSFWorkbook();
        }
    }
}
