package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class TableToExcelServlet extends HttpServlet
{
    private static final Logger log = LoggerFactory.getLogger(TableToExcelServlet.class);

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException
    {
        // TODO Auto-generated method stub
        super.doGet(req, resp);
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

        JsonObject jsonSheet = new JsonParser().parse(tableData)
            .getAsJsonObject();

        ServletOutputStream outputStream = resp.getOutputStream();

        // TODO: Sheet title should be a property in the JSON sheet.
        SheetParser.newInstance(new XSSFWorkbookBuilder(), new CellParser(new ImageLoader(sessionId)))
            .parseSheet(title, jsonSheet)
            .getWorkbook()
            .write(outputStream);

        outputStream.flush();
    }
}
