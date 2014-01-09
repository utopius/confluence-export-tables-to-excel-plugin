package ut.de.cranktheory.plugins.confluence.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import de.cranktheory.confluence.excel.export.ExportAllTheTables;
import de.cranktheory.confluence.excel.export.ExportNamedMacro;
import de.cranktheory.confluence.excel.export.TableParser;
import de.cranktheory.confluence.excel.export.xssf.XSSFWorkbookBuilder;

public class ExportNamedMacroUnitTest extends XmlTest
{
    @Test
    public void Given_a_single_table_Then_export_one_sheet_2_rows_And_4_cells()
    {
        ExportNamedMacro target = ExportNamedMacro.newInstance(new XSSFWorkbookBuilder(),
                TableParser.newInstance(new MockImageParser()), "mySheet");

        Workbook export = export(target, "macro_single_table.xml");

        Assert.assertEquals(1, export.getNumberOfSheets());
        Assert.assertEquals(1, export.getSheetAt(0)
            .getLastRowNum());
        Assert.assertEquals(2, export.getSheetAt(0)
            .getRow(0)
            .getLastCellNum());
    }
}