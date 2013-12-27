package de.cranktheory.plugins.confluence.excel;

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
        byte[] imageInByte = StreamUtils.readAllBytes(inputStream);

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
}
