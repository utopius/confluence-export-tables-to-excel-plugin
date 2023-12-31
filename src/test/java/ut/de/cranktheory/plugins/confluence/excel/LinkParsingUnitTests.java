package ut.de.cranktheory.plugins.confluence.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import de.cranktheory.confluence.excel.ExportAllTheTables;
import de.cranktheory.confluence.excel.export.DefaultParserFactory;
import de.cranktheory.confluence.excel.export.DefaultUrlResolver;
import de.cranktheory.confluence.excel.export.ImageParser;
import de.cranktheory.confluence.excel.export.ParserFactory;
import de.cranktheory.confluence.excel.export.UrlResolver;
import de.cranktheory.confluence.excel.export.xssf.XSSFWorkbookBuilder;

public class LinkParsingUnitTests extends XmlTest
{
    @Test
    public void Given_a_single_table_with_hyperlinks_Then_export_the_links_correctly()
    {
        XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
        ParserFactory parserFactory = createParserFactoryForLinkTests(builder);
        ExportAllTheTables target = new ExportAllTheTables(builder, parserFactory);

        Workbook export = export(target, "single_table_hyperlinks.xml");

        Assert.assertEquals("http://localhost/mylink", export.getSheetAt(0)
            .getRow(0)
            .getCell(0)
            .getStringCellValue());
        Assert.assertNotNull(export.getSheetAt(0)
            .getRow(0)
            .getCell(0)
            .getHyperlink());

        Assert.assertEquals("MyLink", export.getSheetAt(0)
            .getRow(1)
            .getCell(0)
            .getStringCellValue());
        Assert.assertNotNull(export.getSheetAt(0)
            .getRow(1)
            .getCell(0)
            .getHyperlink());
    }

    @Test
    public void Given_a_single_table_with_multiple_hyperlinks_in_one_cell_Then_the_last_link_wins()
    {
        XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
        ParserFactory parserFactory = createParserFactoryForLinkTests(builder);
        ExportAllTheTables target = new ExportAllTheTables(builder, parserFactory);

        Workbook export = export(target, "single_table_multiple_hyperlinks.xml");

        // Fails because both link texts are appended to the cell, but the first link is lost (in transition ;) )
        Assert.assertEquals("MySecondLink", export.getSheetAt(0)
            .getRow(0)
            .getCell(0)
            .getStringCellValue());
        Assert.assertNotNull(export.getSheetAt(0)
            .getRow(0)
            .getCell(0)
            .getHyperlink());
    }

    @Test
    public void Given_a_single_table_with_pagelinks_Then_export_the_links_correctly()
    {
        XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
        ParserFactory parserFactory = createParserFactoryForLinkTests(builder);
        ExportAllTheTables target = new ExportAllTheTables(builder, parserFactory);

        Workbook export = export(target, "single_table_pagelinks.xml");

        Assert.assertEquals("MyLinkTitle", export.getSheetAt(0)
            .getRow(0)
            .getCell(0)
            .getStringCellValue());
        Assert.assertNotNull(export.getSheetAt(0)
            .getRow(0)
            .getCell(0)
            .getHyperlink());
        Assert.assertEquals("ThisIsAPageTitle", export.getSheetAt(0)
            .getRow(1)
            .getCell(0)
            .getStringCellValue());
        Assert.assertNotNull(export.getSheetAt(0)
            .getRow(1)
            .getCell(0)
            .getHyperlink());
    }

    @Test
    public void Given_a_single_table_with_attachmentlinks_Then_export_the_links_correctly()
    {
        XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
        ParserFactory parserFactory = createParserFactoryForLinkTests(builder);
        ExportAllTheTables target = new ExportAllTheTables(builder, parserFactory);

        Workbook export = export(target, "single_table_attachmentlinks.xml");

        Assert.assertEquals("MyFileLabel", export.getSheetAt(0)
            .getRow(0)
            .getCell(0)
            .getStringCellValue());
        Assert.assertNotNull(export.getSheetAt(0)
            .getRow(0)
            .getCell(0)
            .getHyperlink());
        Assert.assertEquals("Filename.xml", export.getSheetAt(0)
            .getRow(1)
            .getCell(0)
            .getStringCellValue());
        Assert.assertNotNull(export.getSheetAt(0)
            .getRow(1)
            .getCell(0)
            .getHyperlink());
    }

    private ParserFactory createParserFactoryForLinkTests(XSSFWorkbookBuilder builder)
    {
        ParserFactory parserFactory = new DefaultParserFactory(null, null, builder, "")
        {
            public ImageParser newImageParser()
            {
                return new MockImageParser();
            }

            @Override
            public UrlResolver newUrlResolver()
            {
                return new DefaultUrlResolver(null, null, "http://localhost/confluence")
                {
                    @Override
                    public String resolvePageUrl(String pageTitle, String spaceKey)
                    {
                        return "/ThisIsAPageTitle";
                    }

                    @Override
                    public String resolveAttachmentUrl(String attachmentFilename, String pageTitle, String spaceKey)
                    {
                        return "/Filename.xml";
                    }
                };
            }
        };
        return parserFactory;
    }
}