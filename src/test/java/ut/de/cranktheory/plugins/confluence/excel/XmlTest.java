package ut.de.cranktheory.plugins.confluence.excel;

import java.io.StringReader;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;

import org.apache.poi.ss.usermodel.Workbook;

import com.atlassian.confluence.content.render.xhtml.DefaultXmlEventReaderFactory;
import com.atlassian.confluence.content.render.xhtml.XhtmlConstants;
import com.google.common.base.Charsets;
import com.google.common.base.Preconditions;
import com.google.common.io.Resources;

import de.cranktheory.confluence.excel.export.ExportAllTheTables;
import de.cranktheory.confluence.excel.export.WorkbookExporter;

public class XmlTest
{
    protected final XMLEventReader createXmlReaderFromFile(String filename)
    {
        Preconditions.checkNotNull(filename, "filename");

        try
        {
            StringReader reader = new StringReader(Resources.toString(getClass().getResource(filename), Charsets.UTF_8));

            // A bit ugly to instantiate DefaultXmlEventReaderFactory directly, but as the Storage Format is decorated
            // by the event reader/factory/thingies of Confluence with necessary doctype and namespace declarations,
            // this is just plain necessary.
            XMLEventReader xmlEventReader = new DefaultXmlEventReaderFactory().createXMLEventReader(reader,
                    XhtmlConstants.STORAGE_NAMESPACES, false);

            return xmlEventReader;
        }
        catch (Exception e)
        {
            throw new IllegalStateException("Could not create XMLEventReader for file " + filename, e);
        }
    }

    protected Workbook export(WorkbookExporter target, String filename)
    {
        Preconditions.checkNotNull(target, "target");
        Preconditions.checkNotNull(filename, "filename");

        try
        {
            return target.export(createXmlReaderFromFile(filename));
        }
        catch (XMLStreamException e)
        {
            throw new IllegalStateException("Export failed.", e);
        }
    }
}
