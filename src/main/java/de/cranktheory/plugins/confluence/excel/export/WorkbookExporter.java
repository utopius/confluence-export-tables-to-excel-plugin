package de.cranktheory.plugins.confluence.excel.export;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookExporter
{

    Workbook export(final XMLEventReader reader) throws XMLStreamException;

}
