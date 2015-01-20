
function onOpen() {
  DocumentApp.getUi().createMenu('Karibou')
  .addItem('Create invoices', 'createDocsFromJSON')  
  .addToUi();
}

//
// get the first table
function findFirstTable(body,cnt){
  // Define the search parameters.
  var searchType = DocumentApp.ElementType.TABLE;
  var searchResult = null;

  // Search until the paragraph is found.
  var it=0;while (searchResult = body.findElement(searchType, searchResult)) {
    var table = searchResult.getElement().asTable();
    if(!cnt|| ++it===cnt) return table;
  }
}
 
function createDocsFromJSON(){
  // filter fields
  var fields=['oid','shipping','customer','quantity','title','part','finalprice']
  
  
  // get id template - relevé d'activité
  var templateid = "1ZgUFFseN2DH1kfkTYemTgIPlYYp7j8Med8os0nVlm0M"; 
  // folder name of where to put completed invoices

  var FOLDER_NAME = "shop en ligne/Comptabilité/Karibou invoices"; 
  // get the data from karibou api (check auth)
  var response = UrlFetchApp.fetch('http://api.demo.karibou.ch/v1/orders/invoices/shops/12'); 
  var invoices = JSON.parse(response.getContentText()); //

  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.BOLD] = false;

  //
  // move file to right folder
  var folder = DocsList.getFolder(FOLDER_NAME);
  
  // invoices data structure
  // - invoices.from
  // - invoices.to
  // - invoices.products={count,title,amount}
  // - invoices.shops={slug:[items]}
  // - invoices.shops[slug]=[{oid,rank,shipping,customer,quantity,title,part,price,finaleprice}]        

  Object.keys(invoices.shops).forEach(function(slug){
    //
    // create one document by shop
    var docid = DocsList.getFileById(templateid).makeCopy().getId();

    //
    // move file in output folder
    // var newDoc = DocumentApp.create(FOLDER_NAME+" - "+username);
    DocsList.getFileById(docid).addToFolder(folder);
    //DocsList.getFileById(docid).rename()
    var doc = DocumentApp.openById(docid);
    var body = doc.getActiveSection();
    var table = findFirstTable(body,2)

    body.replaceText("var_date_from", Utilities.formatDate(new Date(invoices.from), "GMT", "dd.MM.yyyy"));
    body.replaceText("var_date_to", Utilities.formatDate(new Date(invoices.to), "GMT", "dd.MM.yyyy"));
    // body.replaceText("var_client_name", invoices.clientName);
    // body.replaceText("var_boutique_name", slug);
    // body.replaceText("var_address_1", invoices.);
    // body.replaceText("var_address_2", invoices.);
    // body.replaceText("var_date_report", invoices.);
    // body.replaceText("var_iban", invoices.);
    body.replaceText("var_count", invoices.shops[slug].length);
    // body.replaceText("var_fees", invoices.);
    // body.replaceText("var_tot", invoices.);
    // body.replaceText("var_subt", invoices.);

    //
    // replace global variables
    invoices.shops[slug].forEach(function(item){
      // construct items table
      var tr = table.appendTableRow();
      Object.keys(item).forEach(function(key){
        Logger.log(key+':'+item[key])
        if(fields.indexOf(key)!==-1){
          var td = tr.appendTableCell(item[key]);
          td.setAttributes(cellStyle);
        }
      })


    })    
    doc.saveAndClose();
    //appendToDoc(doc, newDoc); // add the filled in template to the students file
    //DocsList.getFileById(docid).setTrashed(true); // delete temporay template file

  })

}
 
// Taken from Johninio's code http://www.google.com/support/forum/p/apps-script/thread?tid=032262c2831acb66&hl=en
function appendToDoc(src, dst) {
  // iterate accross the elements in the source adding to the destination
  for (var i = 0; i < src.getNumChildren(); i++) {
    appendElementToDoc(dst, src.getChild(i));
  }
}
 
function appendElementToDoc(doc, object) {
  var type = object.getType(); // need to handle different types para, table etc differently
  var element = object.removeFromParent(); // need to remove or can't append
  Logger.log("Element type is "+type);
  if (type == "PARAGRAPH") {
    doc.appendParagraph(element);
  } else if (type == "TABLE") {
    doc.appendTable(element);
  } // else if ... I think you get the gist of it
}
