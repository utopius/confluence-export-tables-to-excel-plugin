package de.cranktheory.plugins.confluence.excel;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;

public class ImageLoader
{
    private String _sessionId;

    public ImageLoader(String sessionId)
    {
        _sessionId = sessionId;
    }

    public Image load(URL url) throws IOException
    {
        Preconditions.checkNotNull(url, "url");

        URLConnection connection = url.openConnection();
        connection.setRequestProperty("Cookie", "JSESSIONID=" + _sessionId);
        connection.connect();

        String mimeType = getMimeType(connection);

        InputStream inputStream = connection.getInputStream();
        byte[] imageInByte = readAllBytes(inputStream);

        return new Image(mimeType, imageInByte);
    }

    private static String getMimeType(URLConnection connection)
    {
        Preconditions.checkNotNull(connection, "connection");

        String contentType = connection.getContentType();
        if (Strings.isNullOrEmpty(contentType)) return null;

        int indexOf = contentType.indexOf(";");

        String mimeType = indexOf > -1
                ? contentType.substring(0, indexOf)
                : contentType;

        return mimeType;
    }

    private static byte[] readAllBytes(InputStream inputStream) throws IOException
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
