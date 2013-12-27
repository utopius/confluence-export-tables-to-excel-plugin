package de.cranktheory.plugins.confluence.excel;

import javax.xml.stream.events.XMLEvent;

import org.apache.poi.ss.usermodel.Workbook;

import com.atlassian.confluence.content.render.xhtml.ConversionContext;
import com.atlassian.confluence.pages.Page;
import com.atlassian.confluence.pages.PageManager;
import com.google.common.base.Preconditions;

public class PageVisitor implements VisitorParent, TableConverter
{
    private final PageManager _pageManager;
    private final Page _page;
    private final WorkbookBuilder _workbookBuilder;

    private TableVisitor _tableVisitor;
    private int _sheetNumber = 0;

    public PageVisitor(PageManager pageManager, Page page, WorkbookBuilder workbookBuilder)
    {
        _workbookBuilder = workbookBuilder;
        _page = page;
        _pageManager = pageManager;
    }

    @Override
    public boolean handle(XMLEvent xmlEvent, ConversionContext context)
    {
        if (_tableVisitor != null)
        {
            _tableVisitor.handle(xmlEvent, context);
            return true;
        }

        if (xmlEvent.isStartElement() && "table".equalsIgnoreCase(xmlEvent.asStartElement()
            .getName()
            .getLocalPart()))
        {
            ++_sheetNumber;
            String sheetName = "Table " + _sheetNumber;
            WorksheetBuilder sheetBuilder = createSheet(sheetName);
            _tableVisitor = new TableVisitor(_pageManager, _page, sheetBuilder, this, sheetName);
        }

        return true;
    }

    @Override
    public WorksheetBuilder createSheet(String sheetName)
    {
        return _workbookBuilder.createSheet(sheetName);
    }

    @Override
    public void giveBackControl(TableVisitor tableVisitor)
    {
        Preconditions.checkNotNull(tableVisitor, "tableVisitor");
        Preconditions.checkState(_tableVisitor == tableVisitor, "WTF?!?!1!");

        _tableVisitor = null;
    }

    @Override
    public Workbook getWorkbook()
    {
        return _workbookBuilder.getWorkbook();
    }
}
