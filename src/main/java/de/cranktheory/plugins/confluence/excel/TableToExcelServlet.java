package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;
import java.io.StringReader;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.xml.stream.XMLStreamException;

import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.atlassian.confluence.content.render.xhtml.XhtmlConstants;
import com.atlassian.confluence.content.render.xhtml.XmlEventReaderFactory;
import com.atlassian.confluence.pages.Page;
import com.atlassian.confluence.pages.PageManager;
import com.google.common.base.Preconditions;
import com.google.common.base.Strings;

import de.cranktheory.plugins.confluence.excel.export.ExportAllTheTables;
import de.cranktheory.plugins.confluence.excel.export.ExportNamedMacro;
import de.cranktheory.plugins.confluence.excel.export.ImageParser;
import de.cranktheory.plugins.confluence.excel.export.TableParser;
import de.cranktheory.plugins.confluence.excel.export.WorkbookExporter;
import de.cranktheory.plugins.confluence.excel.export.xssf.XSSFWorkbookBuilder;

public class TableToExcelServlet extends HttpServlet
{
    private static final long serialVersionUID = 4580469806768736781L;

    private static final Logger LOG = LoggerFactory.getLogger(TableToExcelServlet.class);

    private final PageManager _pageManager;
    private final XmlEventReaderFactory _xmlEventReaderFactory;

    public TableToExcelServlet(PageManager pageManager, XmlEventReaderFactory xmlEventReaderFactory)
    {
        _pageManager = pageManager;
        _xmlEventReaderFactory = xmlEventReaderFactory;
    }

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException
    {
        String pageId = getParameter(req, "pageId");

        Page page = _pageManager.getPage(Long.parseLong(pageId));

        convertTable(resp, page, new ExportAllTheTables(new XSSFWorkbookBuilder(), new TableParser(new ImageParser(
                page, _pageManager))));
    }

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException
    {
        String pageId = getParameter(req, "pageId");
        String sheetname = getParameter(req, "sheetname");

        Page page = _pageManager.getPage(Long.parseLong(pageId));

        convertTable(resp, page, new ExportNamedMacro(new XSSFWorkbookBuilder(), new TableParser(new ImageParser(page,
                _pageManager)), sheetname));
    }

    private static String getParameter(HttpServletRequest req, String name)
    {
        String value = req.getParameter(name);
        Preconditions.checkArgument(!Strings.isNullOrEmpty(value), "Request parameter '" + name + "' is null or empty.");

        return value;
    }

    private void convertTable(HttpServletResponse resp, Page page, WorkbookExporter tableExporter)
    {
        try
        {
            Workbook workbook = tableExporter.export(_xmlEventReaderFactory.createXMLEventReader(new StringReader(
                    page.getBodyAsString()), XhtmlConstants.STORAGE_NAMESPACES, false));
            resp.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            String safePageTitle = page.getTitle()
                .replace(" ", "_");

            resp.setHeader("Content-disposition", String.format("attachment; filename=Excel-Export-%s.xlsx",
                    safePageTitle));

            workbook.write(resp.getOutputStream());
        }
        catch (XMLStreamException e)
        {
            LOG.error("Error while parsing Confluence Storage Format.", e);
        }
        catch (IOException e)
        {
            LOG.error("Error writing the response.", e);
        }
    }
}
