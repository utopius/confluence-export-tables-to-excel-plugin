function exportTableToText (table, url) {

    var obj = { rows: [] };
    $('tr', table).not("td div table thead tr").not("td div table tbody tr").each(function()
    {
        currentRow = { isHeaderRow: false, cells: [] }

        jQuery(this).children().each(function ()
        {
            var isHeader = jQuery(this).is("th");
            var isCell = jQuery(this).is("td");
            if (isHeader || isCell)
            {
                if(isHeader) {
                    currentRow.isHeaderRow = true;
                }
                var content = jQuery(this).text();
                var currentCell = { text: content }
                currentRow.cells.push(currentCell);

                var imgChildren = jQuery(this).find('img');
                if(imgChildren.length > 0) {
                    currentCell.iconUrl = imgChildren[0].src;
                }
            }
        });
        obj.rows.push(currentRow);
    });
    var json = JSON.stringify(obj);

    return json;
}

function exportToEXCEL(button, title, url)
{
    var detected = false;

    var table = jQuery(button.parentNode).find("table");

    var data = exportTableToText(table, url);

    var form = document.createElement("form");
    form.setAttribute("method", "post");
    form.setAttribute("action", url);
    form.setAttribute("target", "_blank");

    var idField = document.createElement("input");
    idField.setAttribute("type", "hidden");
    idField.setAttribute("name", "id");
    idField.setAttribute("value", title);
    form.appendChild(idField);

    var dataField = document.createElement("input");
    dataField.setAttribute("type", "hidden");
    dataField.setAttribute("name", "tabledata");
    dataField.setAttribute("value", data);
    form.appendChild(dataField);

    document.body.appendChild(form);
    form.submit();

    return true;
}