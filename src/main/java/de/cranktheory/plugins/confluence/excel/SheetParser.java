package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;
import java.net.MalformedURLException;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

public class SheetParser
{
    public static SheetParser newInstance(WorkbookBuilder builder, CellParser cellParser)
    {
        return new SheetParser(builder, cellParser);
    }

    private CellParser _cellParser;
    private WorkbookBuilder _builder;
    private WorksheetBuilder _sheetBuilder;

    private SheetParser(WorkbookBuilder builder, CellParser cellParser)
    {
        _builder = builder;
        _cellParser = cellParser;
    }

    public WorkbookBuilder parseSheet(String title, JsonObject sheet)
    {
        _sheetBuilder = _builder.createSheet(title);
        JsonArray rows = sheet.getAsJsonArray("rows");

        for (int i = 0; i < rows.size(); ++i)
        {
            JsonObject row = rows.get(i)
                .getAsJsonObject();

            _sheetBuilder.createRow(i);

            parseRow(row);
        }

        return _builder;
    }

    private void parseRow(JsonObject row)
    {
        JsonArray cells = row.get("cells")
            .getAsJsonArray();

        for (int i = 0; i < cells.size(); i++)
        {
            _sheetBuilder.createCell(i);

            try
            {
                _cellParser.parseCell(_sheetBuilder, cells.get(i)
                    .getAsJsonObject());
            }
            catch (MalformedURLException e)
            {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            catch (IOException e)
            {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            catch (PictureDrawingException e)
            {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
    }
}
