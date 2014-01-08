package de.cranktheory.confluence.excel.export;

public class PictureDrawingException extends Exception
{
    private static final long serialVersionUID = 589707395678401105L;

    public PictureDrawingException()
    {
    }

    public PictureDrawingException(String message)
    {
        super(message);
    }

    public PictureDrawingException(Throwable cause)
    {
        super(cause);
    }

    public PictureDrawingException(String message, Throwable cause)
    {
        super(message, cause);
    }

    public PictureDrawingException(String message, Throwable cause, boolean enableSuppression,
            boolean writableStackTrace)
    {
        super(message, cause, enableSuppression, writableStackTrace);
    }

}
