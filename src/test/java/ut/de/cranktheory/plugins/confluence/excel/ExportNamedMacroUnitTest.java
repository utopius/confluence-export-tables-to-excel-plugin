package ut.de.cranktheory.plugins.confluence.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import de.cranktheory.confluence.excel.ExportNamedMacro;
import de.cranktheory.confluence.excel.export.DefaultParserFactory;
import de.cranktheory.confluence.excel.export.ImageParser;
import de.cranktheory.confluence.excel.export.ParserFactory;
import de.cranktheory.confluence.excel.export.xssf.XSSFWorkbookBuilder;

public class ExportNamedMacroUnitTest extends XmlTest
{
    @Test
    public void Given_a_single_table_Then_export_one_sheet_2_rows_And_4_cells()
    {
        XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
        ParserFactory parserFactory = new DefaultParserFactory(null, null, builder, "")
        {
            public ImageParser newImageParser()
            {
                return new MockImageParser();
            }
        };
        ExportNamedMacro target = new ExportNamedMacro(builder, parserFactory, "mySheet");
        Workbook export = export(target, "macro_single_table.xml");

        Assert.assertEquals(1, export.getNumberOfSheets());
        Assert.assertEquals(1, export.getSheetAt(0)
            .getLastRowNum());
        Assert.assertEquals(2, export.getSheetAt(0)
            .getRow(0)
            .getLastCellNum());
    }

    @Test
    public void Given_a_single_table_with_an_Umlaut_sheetname_Then_export_one_sheet_2_rows_And_4_cells()
    {
        XSSFWorkbookBuilder builder = new XSSFWorkbookBuilder();
        ParserFactory parserFactory = new DefaultParserFactory(null, null, builder, "")
        {
            public ImageParser newImageParser()
            {
                return new MockImageParser();
            }
        };
        ExportNamedMacro target = new ExportNamedMacro(builder, parserFactory, "mySheet");

        Workbook export = export(target, "macro_single_table_umlaut.xml");

        Assert.assertEquals(1, export.getNumberOfSheets());
        Assert.assertEquals(1, export.getSheetAt(0)
            .getLastRowNum());
        Assert.assertEquals(2, export.getSheetAt(0)
            .getRow(0)
            .getLastCellNum());
    }
}