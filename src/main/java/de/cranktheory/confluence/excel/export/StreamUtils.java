package de.cranktheory.confluence.excel.export;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import com.google.common.base.Preconditions;

public final class StreamUtils
{
    private StreamUtils()
    {
        // Nope
    }

    //TODO: Check if this is really necessary...ImageIO should work.
    static byte[] readAllBytes(InputStream inputStream) throws IOException
    {
        Preconditions.checkNotNull(inputStream, "inputStream");

        // It was simply not possible to use ImageIO to load the image...
        // reading it manually using byte streams worked at least for pngs and jpgs
        BufferedInputStream in = new BufferedInputStream(inputStream);
        ByteArrayOutputStream baos = new ByteArrayOutputStream(4096);
        BufferedOutputStream out = new BufferedOutputStream(baos);
        int i;
        while ((i = in.read()) != -1)
        {
            out.write(i);
        }
        out.flush();
        out.close();
        // baos.flush();
        byte[] imageInByte = baos.toByteArray();
        baos.close();
        return imageInByte;
    }
}
