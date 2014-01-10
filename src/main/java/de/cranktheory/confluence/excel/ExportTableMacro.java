package de.cranktheory.confluence.excel;

import java.util.Map;

import org.apache.commons.lang.StringUtils;
import org.apache.velocity.VelocityContext;

import com.atlassian.confluence.content.render.xhtml.ConversionContext;
import com.atlassian.confluence.macro.Macro;
import com.atlassian.confluence.macro.MacroExecutionException;
import com.atlassian.confluence.renderer.radeox.macros.MacroUtils;
import com.atlassian.confluence.util.velocity.VelocityUtils;
import com.atlassian.renderer.RenderContext;
import com.atlassian.renderer.v2.RenderMode;
import com.atlassian.renderer.v2.macro.BaseMacro;
import com.atlassian.renderer.v2.macro.MacroException;

/**
 * Macro that allows to export a specific table on a Confluence page.
 */
public class ExportTableMacro extends BaseMacro implements Macro
{
    public static enum ButtonLocation {
        Above, Below, Both
    }

    // v2 Macro methods
    @Override
    public String execute(Map<String, String> parameters, String body, ConversionContext context)
            throws MacroExecutionException
    {
        VelocityContext contextMap = new VelocityContext(MacroUtils.defaultVelocityContext());
        contextMap.put("body", body);

        long id = context.getPageContext()
                .getEntity()
                .getId();
        contextMap.put("pageId", id);
        contextMap.put("sheetname", StringUtils.defaultString(parameters.get("sheetname"), "excel-export"));

        ButtonLocation buttonLocation = parameters.containsKey("button-location")
                ? Enum.valueOf(ButtonLocation.class, parameters.get("button-location"))
                : ButtonLocation.Below;

        contextMap.put("buttonAbove", buttonLocation == ButtonLocation.Above || buttonLocation == ButtonLocation.Both);
        contextMap.put("buttonBelow", buttonLocation == ButtonLocation.Below || buttonLocation == ButtonLocation.Both);

        String renderedTemplate = VelocityUtils.getRenderedTemplate("/templates/export-table-macro.vm", contextMap);
        return renderedTemplate;
    }

    @Override
    public BodyType getBodyType()
    {
        return BodyType.RICH_TEXT;
    }

    @Override
    public boolean suppressMacroRenderingDuringWysiwyg()
    {
        return true;
    }

    @Override
    public OutputType getOutputType()
    {
        return OutputType.BLOCK;
    }

    // v1 Macro methods

    @SuppressWarnings("unchecked")
    @Override
    public String execute(@SuppressWarnings("rawtypes") Map parameters, String body, RenderContext renderContext)
            throws MacroException
    {
        try
        {
            return execute(parameters, body, (ConversionContext) null);
        }
        catch (MacroExecutionException e)
        {
            throw new MacroException(e);
        }

    }

    @Override
    public RenderMode getBodyRenderMode()
    {
        return RenderMode.ALL;
    }

    @Override
    public boolean hasBody()
    {
        return true;
    }

}
