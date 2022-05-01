// When you test this demonstration script, please copy and paste "demo.js" and "index.html" to the script editor of Google Document.
// And, run openSidebar(). By this, the sidebar is opened.
function getNamesToIds() {
  const doc = DocumentApp.getActiveDocument();
  const res = DocNamedRangeApp.getActiveDocument(doc).getNamesToIds();
  DocumentApp.getUi().alert(JSON.stringify(res, null, "  "));
}

function setNamedRangeFromSelection(name) {
  const doc = DocumentApp.getActiveDocument();
  const res =
    DocNamedRangeApp.getActiveDocument(doc).setNamedRangeFromSelection(name);
  DocumentApp.getUi().alert(res);
}

function selectNamedRangeById(id) {
  const doc = DocumentApp.getActiveDocument();
  DocNamedRangeApp.getActiveDocument(doc).selectNamedRangeById(id);
}

function checkNamedRangeOfCursorById(id) {
  const doc = DocumentApp.getActiveDocument();
  const res =
    DocNamedRangeApp.getActiveDocument(doc).checkNamedRangeOfCursorById(id);
  const msg = res
    ? `Cursor is <b>inside</b> range of '${id}'.`
    : `Cursor is <b>outside</b> range of '${id}'.`;
  DocumentApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(msg).setWidth(400).setHeight(50),
    "sample"
  );
}

// Please run this function.
function openSidebar() {
  DocumentApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile("index").setTitle("Demo")
  );
}
