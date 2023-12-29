package de.cranktheory.confluence.excel.export;

public class Link
{
    private final String _url;
    private final String _label;

    public Link(String url, String label)
    {
        _url = url;
        _label = label;
    }

    public String getUrl()
    {
        return _url;
    }

    public String getLabel()
    {
        return _label;
    }

}
