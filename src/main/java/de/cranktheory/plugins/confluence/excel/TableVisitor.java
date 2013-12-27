package de.cranktheory.plugins.confluence.excel;

import java.io.IOException;
import java.io.InputStream;

import javax.xml.namespace.QName;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

import com.atlassian.confluence.content.render.xhtml.ConversionContext;
import com.atlassian.confluence.pages.Attachment;
import com.atlassian.confluence.pages.Page;
import com.atlassian.confluence.pages.PageManager;
import com.atlassian.confluence.xhtml.api.XhtmlVisitor;
import com.google.common.base.Preconditions;

public class TableVisitor implements XhtmlVisitor, VisitorParent
{
    private final PageManager _pageManager;
    private final Page _page;
    private final WorkbookBuilder _workbookBuilder;

    private final VisitorParent _parent;

    private TableVisitor _nestedVisitor;

    private int _currentRowIndex = -1;
    private int _currentColumnIndex = -1;

    private boolean _insideCell;
    private boolean _insideImage;
    private boolean _insideAttachment;
    private String _currentAttachmentFilename;
    private String _sheetName;
    private int _innerTableCount;

    protected TableVisitor(PageManager pageManager, Page page, WorkbookBuilder workbookBuilder, VisitorParent parent,
            String sheetName)
    {
        _workbookBuilder = workbookBuilder;
        _page = page;
        _pageManager = pageManager;
        _parent = parent;
        _sheetName = sheetName;
    }

    @Override
    public void giveBackControl(TableVisitor tableVisitor)
    {
        Preconditions.checkNotNull(tableVisitor, "tableVisitor");
        Preconditions.checkState(_nestedVisitor == tableVisitor, "WTF?!?!1!");

        _nestedVisitor = null;
    }

    @Override
    public boolean handle(XMLEvent xmlEvent, ConversionContext context)
    {
        if (_nestedVisitor != null)
        {
            _nestedVisitor.handle(xmlEvent, context);
            return true;
        }

        if (xmlEvent.isStartElement())
        {
            startElement(xmlEvent);
        }
        else if (xmlEvent.isEndElement())
        {
            endElement(xmlEvent);
        }
        else if (xmlEvent.isCharacters() && _insideCell)
        {
            String data = xmlEvent.asCharacters()
                .getData();
            _workbookBuilder.addTextToCell(data);
        }
        return true;
    }

    private void startElement(XMLEvent xmlEvent)
    {
        StartElement startElement = xmlEvent.asStartElement();
        String localPart = startElement.getName()
            .getLocalPart();

        if ("table".equalsIgnoreCase(localPart))
        {
            // Whoopsie, a nested table (grmbl)
            ++_innerTableCount;
            _nestedVisitor = createNestedVisitor();
        }
        else if ("tr".equalsIgnoreCase(localPart))
        {
            ++_currentRowIndex;

            _workbookBuilder.createRow(_currentRowIndex);
        }
        else if ("td".equalsIgnoreCase(localPart) || "th".equalsIgnoreCase(localPart))
        {
            _insideCell = true;
            ++_currentColumnIndex;

            _workbookBuilder.createCell(_currentColumnIndex);
        }
        else if ("image".equalsIgnoreCase(localPart) && _insideCell)
        {
            _insideImage = true;
        }
        else if ("attachment".equalsIgnoreCase(localPart) && _insideImage)
        {
            _insideAttachment = true;

            Attribute attributeByName = startElement.getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}filename"));
            _currentAttachmentFilename = attributeByName.getValue();

            Attachment attachment = _page.getAttachmentNamed(_currentAttachmentFilename);
            // If attachment is null then its probably a referenced attachment we'll handle in the page block below.
            if (attachment != null)
            {
                addPictureFromAttachment(attachment);
            }
        }
        else if ("page".equalsIgnoreCase(localPart) && _insideImage && _insideAttachment)
        {
            // Attachment included from somewhere else
            Attribute pageTitleAttribute = startElement.getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}content-title"));
            String pageTitle = pageTitleAttribute.getValue();

            Attribute spaceKeyAttribute = startElement.getAttributeByName(QName.valueOf("{http://atlassian.com/resource/identifier}space-key"));

            // If no spaceKey is provided, then the page is located in the same Space as _page
            String spaceKey = spaceKeyAttribute == null
                    ? _page.getSpaceKey()
                    : spaceKeyAttribute.getValue();
            Page referencedPage = _pageManager.getPage(spaceKey, pageTitle);

            // If the page is in trash (referencedPage == null), omit the attachment
            if (referencedPage != null)
            {
                Attachment attachment = referencedPage.getAttachmentNamed(_currentAttachmentFilename);
                addPictureFromAttachment(attachment);
            }
        }
    }

    private TableVisitor createNestedVisitor()
    {
        String sheetName = "Table " + _innerTableCount;
        createSheet(sheetName);
        return new TableVisitor(_pageManager, _page, _workbookBuilder, this, sheetName);
    }

    private void endElement(XMLEvent xmlEvent)
    {
        String localPart = xmlEvent.asEndElement()
            .getName()
            .getLocalPart();

        if ("table".equals(localPart))
        {
            // Give back control to the parent visitor
            _currentRowIndex = -1;
            _parent.giveBackControl(this);
        }
        else if ("tr".equals(localPart))
        {
            _currentColumnIndex = -1;
        }
        else if ("td".equals(localPart) || "th".equalsIgnoreCase(localPart))
        {
            _insideCell = false;
        }
        else if ("image".equalsIgnoreCase(localPart) && _insideCell)
        {
            _insideImage = false;
        }
        else if ("attachment".equalsIgnoreCase(localPart) && _insideImage)
        {
            _insideAttachment = false;
        }
    }

    private void addPictureFromAttachment(Attachment attachmentNamed)
    {
        Attachment latestVersion = (Attachment) attachmentNamed.getLatestVersion();

        try
        {
            InputStream contentsAsStream = latestVersion.getContentsAsStream();
            String contentType = latestVersion.getContentType();
            _workbookBuilder.drawPictureToCell(StreamUtils.readAllBytes(contentsAsStream),
                    ImageUtils.mimeTypeToPoiFormat(contentType));
        }
        catch (PictureDrawingException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        catch (IOException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    @Override
    public void createSheet(String sheetName)
    {
        //FIXME: ARGH...after the nested visitor created a new sheet, we will write in the new sheet, too, once we get control back!
        _parent.createSheet(_sheetName + " " + sheetName);
    }
}
