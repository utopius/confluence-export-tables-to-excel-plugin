package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;
import java.util.ArrayList;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.atlassian.confluence.content.render.xhtml.DefaultConversionContext;
import com.atlassian.confluence.content.render.xhtml.XhtmlException;
import com.atlassian.confluence.pages.Page;
import com.atlassian.confluence.pages.PageManager;
import com.atlassian.confluence.xhtml.api.XhtmlContent;
import com.atlassian.confluence.xhtml.api.XhtmlVisitor;
import com.google.common.base.Preconditions;
import com.google.common.base.Strings;
import com.google.common.collect.Lists;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class TableToExcelServlet extends HttpServlet
{
    private static final Logger log = LoggerFactory.getLogger(TableToExcelServlet.class);

    private final PageManager _pageManager;

    private final XhtmlContent _xhtmlContent;

    public TableToExcelServlet(PageManager pageManager, XhtmlContent xhtmlContent)
    {
        _pageManager = pageManager;
        _xhtmlContent = xhtmlContent;
    }

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException
    {
        String pageId = req.getParameter("pageId");
        Preconditions.checkArgument(!Strings.isNullOrEmpty(pageId), "pageId parameter is null or empty.");

        Page page = _pageManager.getPage(Long.parseLong(pageId));
        try
        {
            String bodyAsString = page.getBodyAsString();

            // final String body = _xhtmlContent.convertStorageToView(bodyAsString, new
            // DefaultConversionContext(page.toPageContext()));
            DefaultConversionContext context = new DefaultConversionContext(page.toPageContext());

            XSSFWorkbookBuilder workbookBuilder = new XSSFWorkbookBuilder();
            ArrayList<? extends XhtmlVisitor> visitors = Lists.newArrayList(new PageVisitor(_pageManager, page,
                    workbookBuilder));
            _xhtmlContent.handleXhtmlElements(bodyAsString, context, visitors);

            resp.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            resp.setHeader("Content-disposition", "attachment; filename=foobar.xlsx");

            workbookBuilder.getWorkbook()
                .write(resp.getOutputStream());
        }
        // catch (XMLStreamException e)
        // {
        // // TODO Auto-generated catch block
        // e.printStackTrace();
        // }
        catch (XhtmlException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        System.out.println("TableToExcelServlet.doGet - pageId: " + pageId);
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
