package de.cranktheory.plugins.confluence.excel;

public class Image
{
    private final String _mimeType;
    private final byte[] _imageInByte;

    public Image(String mimeType, byte[] imageInByte)
    {
        _mimeType = mimeType;
        _imageInByte = imageInByte;
    }

    public String getMimeType()
    {
        return _mimeType;
    }

    public byte[] getBytes()
    {
        return _imageInByte;
    }

}
