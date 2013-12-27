package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;
import java.util.ArrayList;

import javax.servlet.ServletException;
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

public class TableToExcelServlet extends HttpServlet
{
    private static final long serialVersionUID = 4580469806768736781L;

    private static final Logger LOG = LoggerFactory.getLogger(TableToExcelServlet.class);

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
        String pageId = getParameter(req, "pageId");

        Page page = _pageManager.getPage(Long.parseLong(pageId));

        convertTable(resp, page, new PageVisitor(_pageManager, page, new XSSFWorkbookBuilder()));
    }

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException
    {
        String pageId = getParameter(req, "pageId");
        String sheetname = getParameter(req, "sheetname");

        Page page = _pageManager.getPage(Long.parseLong(pageId));

        convertTable(resp, page, new TableMacroVisitor(_pageManager, page, new XSSFWorkbookBuilder(), sheetname));
    }

    private static String getParameter(HttpServletRequest req, String name)
    {
        String value = req.getParameter(name);
        Preconditions.checkArgument(!Strings.isNullOrEmpty(value), "Request parameter '" + name + "' is null or empty.");

        return value;
    }

    private void convertTable(HttpServletResponse resp, Page page, TableConverter visitor) throws IOException
    {
        String bodyAsString = page.getBodyAsString();

        DefaultConversionContext context = new DefaultConversionContext(page.toPageContext());
        try
        {
            ArrayList<? extends XhtmlVisitor> visitors = Lists.newArrayList(visitor);
            _xhtmlContent.handleXhtmlElements(bodyAsString, context, visitors);

            resp.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            String filename = "Excel-Export-" + page.getTitle() + ".xlsx";

            resp.setHeader("Content-disposition", "attachment; filename=" + filename.replace(" ", "_"));
            visitor.getWorkbook()
                .write(resp.getOutputStream());
        }
        catch (XhtmlException e)
        {
            LOG.error("Error converting table.", e);
            throw new IOException(e);
        }
    }
}
