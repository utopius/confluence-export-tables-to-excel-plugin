package ut.de.cranktheory.confluence.excel;

import org.junit.Test;
import de.cranktheory.confluence.excel.MyPluginComponent;
import de.cranktheory.confluence.excel.MyPluginComponentImpl;

import static org.junit.Assert.assertEquals;

public class TableToExcelUnitTest
{
    @Test
    public void testMyName()
    {
        MyPluginComponent component = new MyPluginComponentImpl(null);
        assertEquals("names do not match!", "myComponent",component.getName());
    }
}