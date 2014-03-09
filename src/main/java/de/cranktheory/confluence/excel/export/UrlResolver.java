package de.cranktheory.confluence.excel.export;

import javax.annotation.CheckForNull;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;

public interface UrlResolver
{
    /**
     * Returns a relative url to a page. If the <code>spaceKey</code> is <code>null</code>, the current page context
     * will be used to resolve the url. Use {@link #toAbsoluteUrl(String)} to convert it to an absolute url.
     *
     * @param pageTitle
     *            the title of the page.
     * @param spaceKey
     *            the key of the space where the page is located (optional).
     * @return the relative url to a page with the given <code>pageTitle</code> or <code>null</code> if it cannot be
     *         resolved.
     */
    @CheckForNull
    String resolvePageUrl(@Nonnull String pageTitle, @Nullable String spaceKey);

    /**
     * Returns a relative url to an attachment. If the <code>pageTitle</code> and <code>spaceKey</code> are
     * <code>null</code>, the current page context is used. Use {@link #toAbsoluteUrl(String)} to convert it to an
     * absolute url.
     *
     * @param attachmentFilename
     * @param pageTitle
     * @param spaceKey
     * @return the relative url to an attachment with the given <code>attachmentFilename</code> or <code>null</code> if
     *         it cannot be resolved.
     */
    @CheckForNull
    String resolveAttachmentUrl(@Nonnull String attachmentFilename, @Nullable String pageTitle,
            @Nullable String spaceKey);

    /**
     * Converts the <code>relativeUrl</code> to an absolute one containing the Confluence base url.
     *
     * @param relativeUrl
     *            the relative url to convert.
     * @return an absolute url.
     */
    String toAbsoluteUrl(String relativeUrl);

}