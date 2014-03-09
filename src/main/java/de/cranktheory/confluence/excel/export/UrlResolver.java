package de.cranktheory.confluence.excel.export;

import javax.annotation.CheckForNull;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;

public interface UrlResolver
{
    @CheckForNull
    String resolvePageUrl(@Nonnull String pageTitle, @Nullable String spaceKey);

    @CheckForNull
    String resolveAttachmentUrl(@Nonnull String attachmentFilename, @Nullable String pageTitle,
            @Nullable String spaceKey);

    public abstract String toAbsoluteUrl(String relativeUrl);

}