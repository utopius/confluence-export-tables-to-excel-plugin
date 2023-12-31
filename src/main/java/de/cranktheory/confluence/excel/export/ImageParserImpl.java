package de.cranktheory.confluence.excel.export;

import java.io.IOException;
import java.io.InputStream;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.atlassian.confluence.pages.Attachment;
import com.atlassian.confluence.pages.Page;
import com.atlassian.confluence.pages.PageManager;
import com.google.common.base.Preconditions;

public class ImageParserImpl implements ImageParser
{
    private static final Logger LOG = LoggerFactory.getLogger(ImageParserImpl.class);

    public static ImageParser newInstance(Page page, PageManager pageManager)
    {
        return new ImageParserImpl(Preconditions.checkNotNull(page, "page"), Preconditions.checkNotNull(pageManager,
                "pageManager"));
    }

    private final Page _page;
    private final PageManager _pageManager;

    private ImageParserImpl(Page page, PageManager pageManager)
    {
        _page = page;
        _pageManager = pageManager;
    }

    @Override
    public void parseImage(XMLEventReader reader, WorksheetBuilder sheetBuilder, CellBuilder cellBuilder) throws XMLStreamException
    {
        XMLEvent event = reader.nextEvent();
        if (XMLEvents.isStart(event, "attachment"))
        {
            String attachmentFilename = StorageFormat.getAttachmentFilename(event);

            Attachment attachment = getAttachment(reader, attachmentFilename);

            if (attachment == null)
            {
                LOG.warn(String.format("Attachment %s on page %s not found.", attachmentFilename, _page.getId()));
                return;
            }

            try
            {
                InputStream contentsAsStream = attachment.getContentsAsStream();
                String contentType = attachment.getContentType();
                cellBuilder.drawPictureToCell(StreamUtils.readAllBytes(contentsAsStream),
                        ImageUtils.mimeTypeToPoiFormat(contentType));
            }
            catch (PictureDrawingException e)
            {
                LOG.warn(String.format("Cannot draw picture %s to sheet %s,", attachmentFilename,
                        sheetBuilder.getCurrentSheetname()), e);
            }
            catch (IOException e)
            {
                LOG.warn(String.format("IO error drawing picture %s to sheet %s.", attachmentFilename,
                        sheetBuilder.getCurrentSheetname()), e);
            }
        }
    }

    private Attachment getAttachment(XMLEventReader reader, String attachmentFilename) throws XMLStreamException
    {
        Attachment attachment = _page.getAttachmentNamed(attachmentFilename);

        if (attachment == null)
        {
            // If attachment is null then its probably a referenced attachment.
            attachment = getReferencedAttachment(reader, attachmentFilename, attachment);
        }

        return attachment;
    }

    private Attachment getReferencedAttachment(XMLEventReader reader, String attachmentFilename, Attachment attachment)
            throws XMLStreamException
    {
        XMLEvent pageEvent = reader.nextEvent();
        if (XMLEvents.isStart(pageEvent, "page"))
        {
            // Attachment included from somewhere else
            String pageTitle = StorageFormat.getPageTitle(pageEvent);
            String spaceKey = StorageFormat.getSpaceKey(pageEvent, _page.getSpaceKey());

            Page referencedPage = _pageManager.getPage(spaceKey, pageTitle);

            // If the page is in trash (referencedPage == null), omit the attachment
            if (referencedPage == null)
            {
                LOG.info(String.format(
                        "Attachment '%s' not found, it's probably in the trash and will not be exported.",
                        attachmentFilename));
            }
            else
            {
                attachment = referencedPage.getAttachmentNamed(attachmentFilename);
            }
        }
        return attachment;
    }
}
