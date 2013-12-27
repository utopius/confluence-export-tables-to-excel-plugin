package de.cranktheory.plugins.confluence.excel;

import org.apache.poi.ss.usermodel.Workbook;

import com.atlassian.confluence.xhtml.api.XhtmlVisitor;

public interface TableConverter extends XhtmlVisitor
{

    Workbook getWorkbook();

}
