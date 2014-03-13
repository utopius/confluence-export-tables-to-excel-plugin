package de.cranktheory.confluence.excel.export;

import com.atlassian.confluence.pages.Attachment;
import com.atlassian.confluence.pages.Page;
import com.atlassian.confluence.pages.PageManager;
import com.google.common.base.Preconditions;
import com.google.common.base.Strings;

public class DefaultUrlResolver implements UrlResolver
{
    private final PageManager _pageManager;
    private final Page _page;
    private final String _baseUrl;

    public DefaultUrlResolver(PageManager pageManager, Page page, String baseUrl)
    {
        _pageManager = pageManager;
        _page = page;
        _baseUrl = baseUrl;
    }

    @Override
    public String resolvePageUrl(String pageTitle, String spaceKey)
    {
        Preconditions.checkNotNull(pageTitle, "pageTitle");

        Page referencedPage = _pageManager.getPage(spaceKey == null
                ? _page.getSpaceKey()
                : spaceKey, pageTitle);

        if (referencedPage == null) return null;

        return referencedPage.getUrlPath();
    }

    @Override
    public String resolveAttachmentUrl(String attachmentFilename, String pageTitle, String spaceKey)
    {
        Preconditions.checkNotNull(attachmentFilename, "attachmentFilename");

        Attachment attachment = _page.getAttachmentNamed(attachmentFilename);

        // caller seems to know a location
        if (pageTitle != null)
        {
            String pageSpaceKey = spaceKey != null
                    ? spaceKey
                    : _page.getSpaceKey();

            Page page = _pageManager.getPage(pageSpaceKey, pageTitle);

            attachment = page == null
                    ? null
                    : page.getAttachmentNamed(attachmentFilename);
        }

        if(attachment == null) return null;

        return attachment.getUrlPath();
    }

    @Override
    public String toAbsoluteUrl(String relativeUrl)
    {
        return _baseUrl + Strings.nullToEmpty(relativeUrl);
    }
}
