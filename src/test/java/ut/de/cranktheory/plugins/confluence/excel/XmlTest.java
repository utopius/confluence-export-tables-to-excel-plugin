package ut.de.cranktheory.plugins.confluence.excel;

import java.io.InputStreamReader;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;

import com.google.common.base.Preconditions;

public class XmlTest
{
    protected final XMLEventReader openXmlFile(String filename)
    {
        Preconditions.checkNotNull(filename, "filename");

        try
        {
            InputStreamReader reader = new InputStreamReader(getClass().getResource(filename)
                .openStream());
            XMLEventReader xmlEventReader = XMLInputFactory.newInstance()
                .createXMLEventReader(reader);

            return xmlEventReader;
        }
        catch (Exception e)
        {
            throw new IllegalStateException("Could not create XMLEventReader for file " + filename, e);
        }
    }
}
