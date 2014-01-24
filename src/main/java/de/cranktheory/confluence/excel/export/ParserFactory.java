package de.cranktheory.confluence.excel.export;

public interface ParserFactory
{
    CellParser newCellParser();

    TableParser newTableParser(String newSheetName);

    ImageParser newImageParser();

    MacroParser newMacroParser();

    MacroParser newMacroParser(String sheetnameToExport);
}
