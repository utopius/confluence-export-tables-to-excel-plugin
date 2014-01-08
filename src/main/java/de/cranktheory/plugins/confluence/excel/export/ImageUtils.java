package de.cranktheory.plugins.confluence.excel.export;

import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.base.Preconditions;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.ImmutableMap.Builder;

public final class ImageUtils
{
    private static final ImmutableMap<String, Integer> mimeTypeToPoiImageFormat;

    private ImageUtils()
    {
        //Nope
    }

    static
    {
        Builder<String, Integer> builder = ImmutableMap.builder();
        mimeTypeToPoiImageFormat = builder.put("image/jpg", Workbook.PICTURE_TYPE_JPEG)
            .put("image/jpeg", Workbook.PICTURE_TYPE_JPEG)
            .put("image/png", Workbook.PICTURE_TYPE_PNG)
            .build();
    }

    public static boolean isImageTypeSupported(String mimeType)
    {
        return mimeTypeToPoiImageFormat.containsKey(mimeType);
    }

    public static int mimeTypeToPoiFormat(String mimeType)
    {
        Integer integer = mimeTypeToPoiImageFormat.get(mimeType);
        Preconditions.checkState(integer != null, mimeType + " is not supported.");

        return integer;
    }
}
