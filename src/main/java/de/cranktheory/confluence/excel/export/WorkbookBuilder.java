package de.cranktheory.confluence.excel.export;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookBuilder
{
    /**
     * Returns the {@link Workbook} of the Builder.
     *
     * @return the built {@link Workbook}.
     */
    Workbook getWorkbook();

    /**
     * Creates a new {@link Sheet} in the {@link Workbook} with the given {@code title}.
     *
     * @param title
     *            the title of the new {@link Sheet}
     * @return A {@link WorksheetBuilder} which can be used to create the contents of the {@link Sheet}.
     */
    WorksheetBuilder createSheet(String title);
}