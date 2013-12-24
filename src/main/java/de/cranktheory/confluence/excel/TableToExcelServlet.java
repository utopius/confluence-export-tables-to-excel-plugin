package de.cranktheory.confluence.excel;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.ImmutableMap.Builder;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class TableToExcelServlet extends HttpServlet
{
    private static final Logger log = LoggerFactory.getLogger(TableToExcelServlet.class);
    private static final ImmutableMap<String, Integer> mimeTypeToPoiImageFormat;

    static
    {
        Builder<String, Integer> builder = ImmutableMap.builder();
        mimeTypeToPoiImageFormat = builder.put("image/jpg", Workbook.PICTURE_TYPE_JPEG)
            .put("image/jpeg", Workbook.PICTURE_TYPE_JPEG)
            .put("image/png", Workbook.PICTURE_TYPE_PNG)
            .build();
    }

    private static boolean isImageTypeSupported(String mimeType)
    {
        return mimeTypeToPoiImageFormat.containsKey(mimeType);
    }

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException
    {
        String title = req.getParameter("id");
        Preconditions.checkArgument(!Strings.isNullOrEmpty(title), "Request parameter 'id' is null or empty.");

        String tableData = req.getParameter("tabledata");
        Preconditions.checkArgument(!Strings.isNullOrEmpty(tableData),
                "Request parameter 'tabledata' is null or empty.");

        resp.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        resp.setHeader("Content-disposition", "attachment; filename=" + title + ".xlsx");

        String sessionId = req.getSession(true)
            .getId();

        XSSFWorkbook workbook = parseWorkbook(tableData, sessionId, title);

        ServletOutputStream outputStream = resp.getOutputStream();
        workbook.write(outputStream);
        outputStream.flush();
    }

    private XSSFWorkbook parseWorkbook(String tableData, String sessionId, String title) throws IOException
    {
        // TODO: Sheet title should be a property in the JSON sheet.
        WorkbookBuilder builder = new WorkbookBuilder().createWorkbook()
            .createSheet(title);

        JsonObject sheet = new JsonParser().parse(tableData)
            .getAsJsonObject();

        parseSheet(builder, sessionId, sheet);

        return builder.build();
    }

    private void parseSheet(WorkbookBuilder builder, String sessionId, JsonObject sheet)
    {
        JsonArray rows = sheet.getAsJsonArray("rows");

        parseRows(builder, sessionId, rows);
    }

    private void parseRows(WorkbookBuilder builder, String sessionId, JsonArray rows)
    {
        for (int rowIDX = 0; rowIDX < rows.size(); ++rowIDX)
        {
            JsonObject row = rows.get(rowIDX)
                .getAsJsonObject();

            builder.createRow(rowIDX);

            parseRow(builder, sessionId, row);
        }
    }

    private void parseRow(WorkbookBuilder builder, String sessionId, JsonObject row)
    {
        JsonArray cells = row.get("cells")
            .getAsJsonArray();

        parseCells(builder, sessionId, cells);
    }

    private void parseCells(WorkbookBuilder builder, String sessionId, JsonArray cells)
    {
        for (int i = 0; i < cells.size(); i++)
        {
            builder.createCell(i);

            parseCell(builder, sessionId, cells.get(i)
                .getAsJsonObject());
        }
    }

    private void parseCell(WorkbookBuilder builder, String sessionId, JsonObject jsonCell)
    {
        // TODO: Improve cell type determination (type info should be provided in json?)
        if (!jsonCell.has("iconUrl"))
        {
            String cellValue = jsonCell.get("text")
                .getAsString();
            builder.addTextToCell(cellValue);
            System.out.println("TableToExcel found new cell: " + cellValue);
        }
        else
        {
            String url = jsonCell.get("iconUrl")
                .getAsString();
            Preconditions.checkArgument(!Strings.isNullOrEmpty(url), "url is null or empty.");

            try
            {
                URLConnection connection = new URL(url).openConnection();
                connection.setRequestProperty("Cookie", "JSESSIONID=" + sessionId);
                connection.connect();

                String mimeType = getMimeType(connection);

                if (isImageTypeSupported(mimeType))
                {
                    InputStream inputStream = connection.getInputStream();
                    byte[] imageInByte = readAllBytes(inputStream);
                    Integer imageFormat = mimeTypeToPoiImageFormat.get(mimeType);
                    int imageType = Preconditions.checkNotNull(imageFormat);

                    // if(mimeType.equalsIgnoreCase("image/png")
                    // {
                    // //Do it the ImageIO way for png
                    // byte[] bytes = loadImageFromURL(iconUrl);
                    // }
                    builder.drawPictureToCell(imageInByte, imageType);
                }
                else
                {
                    System.out.println("Table-To-Excel: " + mimeType + " is not supported.");
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
    }

    private static String getMimeType(URLConnection connection)
    {
        Preconditions.checkNotNull(connection, "connection");

        String contentType = connection.getContentType();
        if (Strings.isNullOrEmpty(contentType)) return null;

        int indexOf = contentType.indexOf(";");

        String mimeType = indexOf > -1
                ? contentType.substring(0, indexOf)
                : contentType;

        return mimeType;
    }

    private static byte[] readAllBytes(InputStream inputStream) throws IOException
    {
        Preconditions.checkNotNull(inputStream, "inputStream");

        // It was simply not possible to use ImageIO to load the image...
        // reading it manually using byte streams worked at least for pngs and jpgs
        BufferedInputStream in = new BufferedInputStream(inputStream);
        ByteArrayOutputStream baos = new ByteArrayOutputStream(4096);
        BufferedOutputStream out = new BufferedOutputStream(baos);
        int i;
        while ((i = in.read()) != -1)
        {
            out.write(i);
        }
        out.flush();
        out.close();
        // baos.flush();
        byte[] imageInByte = baos.toByteArray();
        baos.close();
        return imageInByte;
    }
}
