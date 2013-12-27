package de.cranktheory.plugins.confluence.excel;

import javax.xml.namespace.QName;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.XMLEvent;

import org.apache.poi.ss.usermodel.Workbook;

import com.atlassian.confluence.content.render.xhtml.ConversionContext;
import com.atlassian.confluence.pages.Page;
import com.atlassian.confluence.pages.PageManager;
import com.google.common.base.Preconditions;

public class TableMacroVisitor implements VisitorParent, TableConverter
{
    private final PageManager _pageManager;
    private final Page _page;
    private final WorkbookBuilder _workbookBuilder;

    private TableVisitor _tableVisitor;
    private String _sheetName;
    private boolean _inExportTableMacro;
    private boolean _inMacro;
    private boolean _extractSheetname;
    private String _macroSheetName;

    public TableMacroVisitor(PageManager pageManager, Page page, WorkbookBuilder workbookBuilder, String sheetName)
    {
        _workbookBuilder = workbookBuilder;
        _page = page;
        _pageManager = pageManager;
        _sheetName = sheetName;
    }

    @Override
    public boolean handle(XMLEvent xmlEvent, ConversionContext context)
    {
        if (_tableVisitor != null)
        {
            _tableVisitor.handle(xmlEvent, context);
            return true;
        }

        //FIXME: Finding tables with the default macro parameter "sheetname" does not work always
        if (xmlEvent.isStartElement() && "table".equalsIgnoreCase(xmlEvent.asStartElement()
            .getName()
            .getLocalPart()) && (_sheetName.equals(_macroSheetName) || _macroSheetName == null && _inExportTableMacro))
        {
            _tableVisitor = new TableVisitor(_pageManager, _page, _workbookBuilder.createSheet(_sheetName), this,
                    _macroSheetName);
        }
        if (xmlEvent.isStartElement() && "structured-macro".equalsIgnoreCase(xmlEvent.asStartElement()
            .getName()
            .getLocalPart()))
        {
            Attribute name = xmlEvent.asStartElement()
                .getAttributeByName(QName.valueOf("{http://atlassian.com/content}name"));
            System.out.println(name);

            _inMacro = true;
            if (name != null && "export-table".equals(name.getValue()))
            {
                _inExportTableMacro = true;
                System.out.println(_inExportTableMacro);
                // Now we need to get the sheetName parameter from the inner tag, if present.
            }
        }
        else if (xmlEvent.isStartElement() && "parameter".equalsIgnoreCase(xmlEvent.asStartElement()
            .getName()
            .getLocalPart()))
        {
            Attribute name = xmlEvent.asStartElement()
                .getAttributeByName(QName.valueOf("{http://atlassian.com/content}name"));

            _extractSheetname = name != null && "sheetname".equalsIgnoreCase(name.getValue());
        }
        else if (xmlEvent.isCharacters() && _extractSheetname)
        {
            _macroSheetName = xmlEvent.asCharacters()
                .getData();
        }
        else if (xmlEvent.isEndElement() && "parameter".equalsIgnoreCase(xmlEvent.asEndElement()
            .getName()
            .getLocalPart()) && _extractSheetname)
        {
            _extractSheetname = false;
        }
        else if (xmlEvent.isEndElement() && "structured-macro".equalsIgnoreCase(xmlEvent.asEndElement()
            .getName()
            .getLocalPart()) && _inMacro)
        {
            if (_inExportTableMacro)
            {
                _inExportTableMacro = false;
            }
            _inMacro = false;
            System.out.println(_inExportTableMacro);
            // Now we need to get the sheetName parameter from the inner tag, if present.
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
        _macroSheetName = null;
        _tableVisitor = null;
    }

    @Override
    public Workbook getWorkbook()
    {
        return _workbookBuilder.getWorkbook();
    }
}
