package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;

import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.ImmutableMap.Builder;
import com.google.gson.JsonObject;

public class CellParser
{
    private ImageLoader _imageLoader;

    public CellParser(ImageLoader imageLoader)
    {
        _imageLoader = imageLoader;
    }

    public void parseCell(WorkbookBuilder builder, JsonObject jsonCell)
            throws MalformedURLException, IOException, PictureDrawingException
    {
        // TODO: Improve cell type determination (type info should be provided in json?)
        if (!jsonCell.has("iconUrl"))
        {
            String cellValue = jsonCell.get("text")
                .getAsString();
            builder.addTextToCell(cellValue);
            System.out.println("TableToExcel found new cell: " + cellValue);
            return;
        }
        String url = jsonCell.get("iconUrl")
            .getAsString();
        Preconditions.checkArgument(!Strings.isNullOrEmpty(url), "url is null or empty.");

        Image image = _imageLoader.load(new URL(url));

        String mimeType = image.getMimeType();
        if (!ImageUtils.isImageTypeSupported(mimeType))
        {
            System.out.println("Table-To-Excel: " + mimeType + " is not supported.");
            return;
        }

        Integer imageFormat = ImageUtils.mimeTypeToPoiFormat(mimeType);
        int imageType = Preconditions.checkNotNull(imageFormat);

        // if(mimeType.equalsIgnoreCase("image/png")
        // {
        // //Do it the ImageIO way for png
        // byte[] bytes = loadImageFromURL(iconUrl);
        // }
        builder.drawPictureToCell(image.getBytes(), imageType);
    }
}
