function exportSheetsAsXML(title) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  v = sheet.getRange('A1:F1000').getValues();

  var xml = `\
<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>${title}</name>
`;

  var style = HtmlService.createTemplateFromFile('style.html');
  xml += style.evaluate().getContent();

  for (var i=3; ;i++) {
    var row = v[i];

    if (row[0] == '')
      break

    if (row[0] != title || row[3] == '' || row[4] == '')
      continue
    var placemark = HtmlService.createTemplateFromFile('placemark.html');
    placemark.name = row[1];
    placemark.description = row[5];
    placemark.url = row[2];
    placemark.coordinates = `${row[4]},${row[3]},0`;
    placemark.styleurl = title;
    xml += placemark.evaluate().getContent();
  }

  xml += `\
  </Document>
</kml>`;

  return xml;
}

function createExportedXml(cate) {
  var xml = exportSheetsAsXML(cate);
  Logger.log(xml);

  var file = DriveApp.createFile(`${cate}.kml`, xml);
  // Display a modal dialog box with custom HtmlService content.
  var template = HtmlService.createTemplateFromFile('dialog.html');
  template.content =   file.getBlob().getDataAsString();
  template.filename = `${cate}.kml`;
  
  var htmlOutput = template.evaluate()
     .setWidth(400)
     .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ダウンロード');
  file.setTrashed(true);
}

function downloadGurume() {
  createExportedXml("グルメ情報")
}

function downloadKanko() {
  createExportedXml("観光情報")
}

function downloadOmiyage() {
  createExportedXml("お土産情報")
}