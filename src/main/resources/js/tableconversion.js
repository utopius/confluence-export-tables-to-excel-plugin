function exportTable(parent, pageId, sheetname, url)
{
    var table = jQuery(parent).find("table");

    var form = document.createElement("form");
    form.setAttribute("method", "post");
    form.setAttribute("action", url);
    form.setAttribute("target", "_blank");

    var pageIdField = document.createElement("input");
    pageIdField.setAttribute("type", "hidden");
    pageIdField.setAttribute("name", "pageId");
    pageIdField.setAttribute("value", pageId);
    form.appendChild(pageIdField);

    var sheetnameField = document.createElement("input");
    sheetnameField.setAttribute("type", "hidden");
    sheetnameField.setAttribute("name", "sheetname");
    sheetnameField.setAttribute("value", sheetname);
    form.appendChild(sheetnameField);

    document.body.appendChild(form);
    form.submit();

    return true;
}