package de.cranktheory.confluence.excel.export;

public interface WorksheetBuilder
{

    /**
     * @return the name of the current sheet.
     */
    String getCurrentSheetname();

    /**
     * Creates a row at the given {@code index} in the current sheet.
     *
     * @param index
     *            Well, the index.
     */
    void createRow(int index);

    /**
     * Creates a cell at the given {@code index} in the current sheet.
     *
     * @param index
     *            Well, the index.
     */
    CellBuilder createCell(int index);

    void build();
}
