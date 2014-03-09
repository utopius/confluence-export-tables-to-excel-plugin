package de.cranktheory.confluence.excel.export;

import com.atlassian.confluence.pages.Page;
import com.atlassian.confluence.pages.PageManager;

public class DefaultParserFactory implements ParserFactory
{
    private final PageManager _pageManager;
    private final Page _page;
    private WorkbookBuilder _workbookBuilder;
    private String _baseUrl;

    public DefaultParserFactory(PageManager pageManager, Page page, WorkbookBuilder workbookBuilder, String baseUrl)
    {
        _pageManager = pageManager;
        _page = page;
        _workbookBuilder = workbookBuilder;
        _baseUrl = baseUrl;
    }

    @Override
    public CellParser newCellParser()
    {
        return new CellParser(this, newImageParser(), newLinkParser());
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

    @Override
    public LinkParser newLinkParser()
    {
        return new DefaultLinkParser(newUrlResolver());
    }

    @Override
    public UrlResolver newUrlResolver()
    {
        return new DefaultUrlResolver(_pageManager, _page, _baseUrl);
    }
}
