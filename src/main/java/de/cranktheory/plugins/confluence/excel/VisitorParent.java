package de.cranktheory.plugins.confluence.excel;

public interface VisitorParent
{

    public void giveBackControl(TableVisitor tableVisitor);

    public abstract WorksheetBuilder createSheet(String sheetName);

}