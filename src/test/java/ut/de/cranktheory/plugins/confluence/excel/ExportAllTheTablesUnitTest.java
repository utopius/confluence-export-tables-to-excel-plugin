package ut.de.cranktheory.plugins.confluence.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import de.cranktheory.confluence.excel.ExportAllTheTables;
import de.cranktheory.confluence.excel.export.DefaultParserFactory;
import de.cranktheory.confluence.excel.export.ImageParser;
import de.cranktheory.confluence.excel.export.ParserFactory;
import de.cranktheory.confluence.excel.export.xssf.XSSFWorkbookBuilder;

public class ExportAllTheTablesUnitTest extends XmlTest
{
    @Test
    public void Given_a_single_table_Then_export_one_sheet_2_rows_And_4_cells()
    {
        XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
        ParserFactory parserFactory = new DefaultParserFactory(null, null, builder)
        {
            public ImageParser newImageParser()
            {
                return new MockImageParser();
            }
        };
        ExportAllTheTables target = new ExportAllTheTables(builder, parserFactory);

        Workbook export = export(target, "single_table.xml");

        Assert.assertEquals(1, export.getNumberOfSheets());
        Assert.assertEquals(1, export.getSheetAt(0)
            .getLastRowNum());
        Assert.assertEquals(2, export.getSheetAt(0)
            .getRow(0)
            .getLastCellNum());
    }

    @Test
    public void Given_two_tables_Then_export_them_as_two_sheets()
    {
        XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
        ParserFactory parserFactory = new DefaultParserFactory(null, null, builder)
        {
            public ImageParser newImageParser()
            {
                return new MockImageParser();
            }
        };
        ExportAllTheTables target = new ExportAllTheTables(builder, parserFactory);

        Workbook export = export(target, "two_tables.xml");

        Assert.assertEquals(2, export.getNumberOfSheets());
    }

    @Test
    public void Given_a_table_with_a_nested_table_Then_export_them_as_two_sheets_And_link_the_nested_table_to_the_respective_cell()
    {
        XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
        ParserFactory parserFactory = new DefaultParserFactory(null, null, builder)
        {
            public ImageParser newImageParser()
            {
                return new MockImageParser();
            }
        };
        ExportAllTheTables target = new ExportAllTheTables(builder, parserFactory);

        Workbook export = export(target, "nested_table.xml");

        Assert.assertEquals(2, export.getNumberOfSheets());

        Cell cell = export.getSheetAt(0).getRow(1).getCell(1);
        Hyperlink hyperlink = cell.getHyperlink();

        Assert.assertNotNull(hyperlink);
        Assert.assertEquals("'Table 1 nested Table 1'!A1", hyperlink.getAddress());
    }

    @Test
    public void Given_a_table_with_two_nested_tables_Then_export_them_as_three_sheets_And_link_the_nested_tables_to_the_respective_cells()
    {
        XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
        ParserFactory parserFactory = new DefaultParserFactory(null, null, builder)
        {
            public ImageParser newImageParser()
            {
                return new MockImageParser();
            }
        };
        ExportAllTheTables target = new ExportAllTheTables(builder, parserFactory);

        Workbook export = export(target, "two_nested_tables.xml");

        Assert.assertEquals(3, export.getNumberOfSheets());

        Cell cell1 = export.getSheetAt(0).getRow(1).getCell(0);
        Hyperlink hyperlink1 = cell1.getHyperlink();

        Assert.assertNotNull(hyperlink1);
        Assert.assertEquals("'Table 1 nested Table 1'!A1", hyperlink1.getAddress());

        Cell cell2 = export.getSheetAt(0).getRow(1).getCell(1);
        Hyperlink hyperlink2 = cell2.getHyperlink();

        Assert.assertNotNull(hyperlink2);
        Assert.assertEquals("'Table 1 nested Table 2'!A1", hyperlink2.getAddress());
    }
}