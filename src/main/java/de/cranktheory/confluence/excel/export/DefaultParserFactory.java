package de.cranktheory.confluence.excel.export;

import com.atlassian.confluence.pages.Page;
import com.atlassian.confluence.pages.PageManager;

public class DefaultParserFactory implements ParserFactory
{
    private final PageManager _pageManager;
    private final Page _page;
    private WorkbookBuilder _workbookBuilder;

    public DefaultParserFactory(PageManager pageManager, Page page, WorkbookBuilder workbookBuilder)
    {
        _pageManager = pageManager;
        _page = page;
        _workbookBuilder = workbookBuilder;
    }

    @Override
    public CellParser newCellParser()
    {
        return new CellParser(this);
    }

    @Override
    public TableParser newTableParser(final String sheetname)
    {
        return new TableParser(this, _workbookBuilder, sheetname);
    }

    @Override
    public ImageParser newImageParser()
    {
        return ImageParserImpl.newInstance(_page, _pageManager);
    }

    @Override
    public MacroParser newMacroParser(String sheetnameToExport)
    {
        return new MacroParser(this, sheetnameToExport);
    }

    @Override
    public MacroParser newMacroParser()
    {
        return new MacroParser(this);
    }
}
