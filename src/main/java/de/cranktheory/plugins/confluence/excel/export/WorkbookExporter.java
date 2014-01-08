package de.cranktheory.plugins.confluence.excel.export;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookExporter
{

    /**
     * Exports the contents of the given {@link XMLEventReader} as {@link Workbook}.
     *
     * @param reader
     *            {@link XMLEventReader} holding the data in Confluence Storage Format.
     * @return the exported {@link Workbook}.
     * @throws XMLStreamException
     */
    Workbook export(final XMLEventReader reader) throws XMLStreamException;

}
