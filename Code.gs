function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Show in sidebar', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Anki export');
  DocumentApp.getUi().showSidebar(ui);
}

// https://stackoverflow.com/questions/47299478/google-apps-script-docs-convert-selected-element-to-html
var globalNestedListLevel = -1;
var listIds = [];
var listNestLevels = {};
var listCounters = {};

function extractHtmlCardsFromDocument() 
{
  var flashCards = [];
  var images = [];
  var cardsAndImages;

  writeStatusToCache("unknown", 0);
  cards = findFlashCardsInDocument();
  writeStatusToCache(cards.length, 0);
  
  // Walk through all found flash cards
  for (var i = 0; i < cards.length; i++) 
  {
    var fc = cards[i];
    var html = [];
    for(var j = 0; j < fc.backItems.length; j++)
    {
       html.push(processItem_V1(fc.backItems[j], listCounters, images));
    }
    var hash = hashCode(fc.front);
    flashCards.push({front: fc.front, html: html.join('')});
    writeStatusToCache(cards.length, i+1);
  }
  
  var imagesJson = base64EncodeImages(images);
  cardsAndImages = {deckName:DocumentApp.getActiveDocument().getName(), cards: flashCards, images: imagesJson};
  
  // images are sent together with the flash cards to the client so that we
  // won't have to process the document again when exporting
  // Ideally they should be stored in the CacheService. But since each key can only hold 100kB of data
  // a work around is needed to store larger images in several keys.
  return cardsAndImages;
}

function hashCode(s) 
{
  var h = 0, l = s.length, i = 0;
  if ( l > 0 )
    while (i < l)
      h = (h << 5) - h + s.charCodeAt(i++) | 0;
  return h;
};

// Write status to cache so that client can fetch it and update progress
function writeStatusToCache(total, current)
{
  var cache = CacheService.getDocumentCache();
  if(total != null)
    cache.put("total", total);
  if(current != null)
    cache.put("current", current);
}

// Called from client to get status
function readStatusFromCache()
{
  var cache = CacheService.getDocumentCache();
  var total = cache.get("total");
  var current = cache.get("current");
  return {total: total, current:current};
}

// Encodes an array of images to base64 and stores them together with name and type in an array and returns its' JSON representation
function base64EncodeImages(images)
{
  var base64Images = [];
  
  for(var i=0;i<images.length;i++)
  {
    var base64Img = Utilities.base64Encode(images[i].blob.getBytes());   
    base64Images.push({name: images[i].name, type: images[i].type, data:base64Img});
  }

  var json = JSON.stringify(base64Images);  
  return json;
}

// Takes an json string encoded by base64EncodeImages() and decodes it into an array of image blobs
function decodeBase64Images(json)
{
  var base64Images = [];
  var images = [];
  
  var base64Images = JSON.parse(json);
  
  for(var i=0;i<base64Images.length;i++)
  {
    var imgBlob = Utilities.newBlob("");
    imgBlob.setBytes(Utilities.base64Decode(base64Images[i].data));
    imgBlob.setName(base64Images[i].name);
    imgBlob.setContentType(base64Images[i].type);
    images.push(imgBlob);
  }
  return images;
}

// Cache can only store 100kB/key
// so as a workaround all images are sent to client :(
function storeImagesInCache(images)
{
  var base64Images = [];
  
  for(var i=0;i<images.length;i++)
  {
    var base64Img = Utilities.base64Encode(images[i].blob.getBytes());   
    base64Images.push({name: images[i].name, type: images[i].type, data:base64Img});
  }

  var json = JSON.stringify(base64Images);  
  var cache = CacheService.getDocumentCache();
  cache.put("images", json);
}

// Cache can only stor 100kB/key
// so as a workaround all images are sent to client :(
function readImagesFromCache()
{
  var base64Images = [];
  var images = [];
  
  var cache = CacheService.getDocumentCache();
  var json = cache.get("images");
  var base64Images = JSON.parse(json);
  
  for(var i=0;i<base64Images.length;i++)
  {
    var imgBlob = Utilites.newBlob("");
    imgBlob.setBytes() = Utilities.base64Decode(base64Images[i].data);
    imgBlob.setName(base64Images[i].name);
    imgBlob.setContentType(base64Images[i].type);
    images.push(imgBlob);
  }
  return images;
}

function createCsvFromCards(cards)
{
  var output = [];
  for(var i=0;i<cards.length;i++)
  {
    output.push(cards[i].front);
    output.push('§');
    output.push(cards[i].html);
    output.push("\n\n");
  }
  return output.join('');
}

/*

function exportToGoogleDrive(cardsAndImages)
{
  var images = decodeBase64Images(cardsAndImages.images);
  var csv = createCsvFromCards(cardsAndImages.cards);
  
  var folder = DriveApp.createFolder("Anki-export");
  folder.createFile("ankiexport.txt", csv);
  
  for (var j=0; j<images.length; j++)
  {
    folder.createFile(images[j]);
  }
}

function emailHtml(html, images) {
  var attachments = [];
  for (var j=0; j<images.length; j++) {
    attachments.push( {
      "fileName": images[j].name,
      "mimeType": images[j].type,
      "content": images[j].blob.getBytes() } );
  }

  var inlineImages = {};
  for (var j=0; j<images.length; j++) {
    inlineImages[[images[j].name]] = images[j].blob;
  }

  var name = DocumentApp.getActiveDocument().getName()+".txt";
  attachments.push({"fileName":name, "mimeType": "text/html", "content": html});
  MailApp.sendEmail({
     to: Session.getActiveUser().getEmail(),
     subject: name,
     htmlBody: name + " - konverterad till CSV",
     inlineImages: inlineImages,
     attachments: attachments
   });
}
*/

function findFlashCardsInDocument()
{
  var cards = [];
  var item = null;
  var doc = DocumentApp.getActiveDocument();
  
/* DEBUG CODE - sets a selection so that it can be debugged
var b = doc.getBody();
var rangeBuilder = doc.newRange();
for (var i = 0; i < 25; i++) 
{
  rangeBuilder.addElement(b.getChild(i));
}
doc.setSelection(rangeBuilder.build());*/
  
  var selection = doc.getSelection();
  if(selection)
  {
    var selectedItemCount = selection.getRangeElements().length;
    var selectedElements = selection.getRangeElements();
    var indexOfFirstElement = doc.getBody().getChildIndex(selectedElements[0].getElement());
    Logger.log(selectedItemCount);
    var nextElementIdx = 0;
    while(nextElementIdx < selectedItemCount)
    {
      var element = selection.getRangeElements()[nextElementIdx].getElement();
      var type = element.getType().toString();
      // The index of the first selected element needs to be subtracted from the returned index
      nextElementIdx = findFlashCardsInItem(element, cards) - indexOfFirstElement;
    }
  }
  else
  {
    var body = doc.getBody();
    var children = 0;
    if(body.getNumChildren)
      children = body.getNumChildren();

    var nextElementIdx = 0;
    while(nextElementIdx < children)
    {
      var element = body.getChild(nextElementIdx);
      var type = element.getType().toString();
      nextElementIdx = findFlashCardsInItem(element, cards);
    }
  }
  return cards;
}

function findFlashCardsInItem(item, cardCollection)
{
  var type = item.getType().toString();

  if(isHeader(item) && !hasSubHeader(item))
  {
    var front = item.getText();
    var backItems = [];
    
    do
    {
      var nextSibling = item.getNextSibling();
      var nextHeaderLevel = getHeaderLevel(nextSibling);
      if(nextHeaderLevel == 0)
      {
        backItems.push(nextSibling);
        item = nextSibling;
      }   
    }
    while(nextHeaderLevel == 0 && item.getNextSibling())
    
    // add card
    cardCollection.push({front: front, backItems: backItems});
  }
  // return index of next document element
  return item.getParent().getChildIndex(item)+1;
 
}

function hasSubHeader(paragraph)
{
  var type = paragraph.getType().toString();
  var headerLevel = getHeaderLevel(paragraph);
  
  var nextSibling = paragraph.getNextSibling();
  if(nextSibling == null)
    return false;
  
  var nextHeaderLevel = getHeaderLevel(nextSibling);
  while(nextHeaderLevel == 0)
  {
    var nextSibling = nextSibling.getNextSibling();
    if(nextSibling == null)
      break;
    var nextHeaderLevel = getHeaderLevel(nextSibling);
  }
  
  return headerLevel < nextHeaderLevel;
}

function getHeaderLevel(paragraph)
{
  if(paragraph == null)
    return -1;
  
  if(paragraph.getHeading)
  {
    var heading = paragraph.getHeading();
    switch (heading) 
    {
      case DocumentApp.ParagraphHeading.HEADING6: 
        return 6;
      case DocumentApp.ParagraphHeading.HEADING5: 
        return 5;
      case DocumentApp.ParagraphHeading.HEADING4:
        return 4;
      case DocumentApp.ParagraphHeading.HEADING3:
        return 3;
      case DocumentApp.ParagraphHeading.HEADING2:
        return 2;
      case DocumentApp.ParagraphHeading.HEADING1:
        return 1;
      default: 
        return 0;
    }
  }
  else
    return false;
}


function isHeader(item)
{  
  var text = ""
  if(item.getText)
  {
    text = item.getText();
    if(text.trim().length ==0)
      return false; // 0 length headers are considered invalid!
  }
  return getHeaderLevel(item) > 0;
}


function processImage(item, images, output)
{
  images = images || [];
  var blob = item.getBlob();
  var contentType = blob.getContentType();
  var extension = "";
  if (/\/png$/.test(contentType)) {
    extension = ".png";
  } else if (/\/gif$/.test(contentType)) {
    extension = ".gif";
  } else if (/\/jpe?g$/.test(contentType)) {
    extension = ".jpg";
  } else {
    throw "Unsupported image type: "+contentType;
  }
  var imagePrefix = DocumentApp.getActiveDocument().getName()+"_image_";
  var imageCounter = images.length;
  var name = imagePrefix + imageCounter + extension;
  blob.setName(name);
  imageCounter++;
  output.push('<br><img src="'+name+'" />');
  images.push( {
    "blob": blob,
    "type": contentType,
    "name": name});
}

function convertFlashCardToHtml(index)
{
  var html = [];
  
  //get card from cache
  var cache = CacheService.getPublicCache();
  var cardsJson = cache.get("cards");
  cards = JSON.parse(cardsJson);
  Logger.log(cards);
  
  for(var i = 0; i < cards[index].backItems.length; i++)
    html.push(processItem_V1(cards[index].backItems[i], listCounters, images));
  
  return html;
}


/**
 * @param {PositionedImage[]} positionedImages - https://developers.google.com/apps-script/reference/document/positioned-image
 * @param {{"blob": Blob,"type": string,"name": string}[]} images
 * @param {string[]} output
 */
function processPositionedImages(positionedImages, images, output, imagesOptions) {
  //TODO:
  //https://developers.google.com/apps-script/reference/document/positioned-image
}

/**
 * @param {string[]} atts
 */
function dumpAttributes(atts) {
  // Log the paragraph attributes.
  for (var att in atts) {
    if (atts[att]) Logger.log(att + ":" + atts[att]);
  }
}

function getAbsoluteListItemNestLevel(listItem)
{
  // get base nest level of list
  // (always depending on context of list)
  var listNestLevel = listNestLevels[listItem.getListId()] || 0;
  var itemNestLevel = listNestLevel + listItem.getNestingLevel();
  return itemNestLevel;
}

/**
* @param {Element} item - https://developers.google.com/apps-script/reference/document/element
* @param {Object} listCounters
* @param {{"blob": Blob,"type": string,"name": string, "height": number, "width": number}[]} images
* @returns {string}
*/
function processItem_V1(item, listCounters, images, imagesOptions, footnotes) {
  var output = [];
  var prefix = "",
      suffix = "";
  var style = "";
  
  var hasPositionedImages = false;
  if (item.getPositionedImages) 
  {
    positionedImages = item.getPositionedImages();
    hasPositionedImages = true;
  }
  
  var itemType = item.getType();
  
  if (itemType === DocumentApp.ElementType.PARAGRAPH) {
    //https://developers.google.com/apps-script/reference/document/paragraph
    
    if (item.getNumChildren() == 0) {
      return "<br />";
    }
    
    var p = "";
    
    if (item.getIndentStart() != null) {
      p += "margin-left:" + item.getIndentStart() + "; ";
    } else {
      // p += "margin-left: 0; "; // superfluous
    }
    
    // what does getIndentEnd actually do? the value is the same as in getIndentStart
    /*if (item.getIndentEnd() != null) {
    p += "margin-right:" + item.getIndentStart() + "; ";
    } else {
    // p += "margin-right: 0; "; // superfluous
    }*/
    
    //Text Alignment
    switch (item.getAlignment()) {
        // Add a # for each heading level. No break, so we accumulate the right number.
        //case DocumentApp.HorizontalAlignment.LEFT:
        //  p += "text-align: left;"; break;
      case DocumentApp.HorizontalAlignment.CENTER:
        p += "text-align: center;";
        break;
      case DocumentApp.HorizontalAlignment.RIGHT:
        p += "text-align: right;";
        break;
      case DocumentApp.HorizontalAlignment.JUSTIFY:
        p += "text-align: justify;";
        break;
      default:
        p += "";
    }
    
    //TODO: getLineSpacing(line-height), getSpacingBefore(margin-top), getSpacingAfter(margin-bottom),
    
    //TODO: 
    //INDENT_END      Enum  The end indentation setting in points, for paragraph elements.
    //INDENT_FIRST_LINE Enum  The first line indentation setting in points, for paragraph elements.
    //INDENT_START      Enum  The start indentation setting in points, for paragraph elements.
    
    if (p !== "") {
      style = 'style="' + p + '"';
    }
    
    //TODO: add DocumentApp.ParagraphHeading.TITLE, DocumentApp.ParagraphHeading.SUBTITLE
    
    //Heading or only paragraph
    switch (item.getHeading()) {
        // Add a # for each heading level. No break, so we accumulate the right number.
      case DocumentApp.ParagraphHeading.HEADING6:
        prefix = "<h6 " + style + ">", suffix = "</h6>";
        break;
      case DocumentApp.ParagraphHeading.HEADING5:
        prefix = "<h5 " + style + ">", suffix = "</h5>";
        break;
      case DocumentApp.ParagraphHeading.HEADING4:
        prefix = "<h4 " + style + ">", suffix = "</h4>";
        break;
      case DocumentApp.ParagraphHeading.HEADING3:
        prefix = "<h3 " + style + ">", suffix = "</h3>";
        break;
      case DocumentApp.ParagraphHeading.HEADING2:
        prefix = "<h2 " + style + ">", suffix = "</h2>";
        break;
      case DocumentApp.ParagraphHeading.HEADING1:
        prefix = "<h1 " + style + ">", suffix = "</h1>";
        break;
      default:
        prefix = "<p " + style + ">", suffix = "</p>";
    }
    
    var attr = item.getAttributes();
    
  } else if (itemType === DocumentApp.ElementType.INLINE_IMAGE) {
    processImage(item, images, output, imagesOptions);
  } else if (itemType === DocumentApp.ElementType.INLINE_DRAWING) {
    //TODO
    Logger.log("INLINE_DRAWING: " + JSON.stringify(item));
  } 
  else if (itemType === DocumentApp.ElementType.LIST_ITEM) {
    var listItem = item;
    var text = listItem.getText();
    var gt = listItem.getGlyphType();
    var key = listItem.getListId()// + '.' + listItem.getNestingLevel();
    
    if(listIds.indexOf(key) <0)
    {
      listIds.push(key);
      prefix += "<ul>"//;, suffix = '</ul>';
      globalNestedListLevel++;
      listNestLevels[key] = globalNestedListLevel;
    }
    
    // get base nest level of list
    // (always depending on context of list)
    var listNestLevel = listNestLevels[key] || 0;
    var itemNestLevel = getAbsoluteListItemNestLevel(listItem);//)listNestLevel + listItem.getNestingLevel();
    
    while(itemNestLevel > globalNestedListLevel)
    {
      prefix += "<ul>"//;, suffix = '</ul>';
      globalNestedListLevel++;
    }
    
    
    prefix += "<li>", suffix = '</li>';
    
    // list added - increase counter
    var counter = listCounters[key + "." + itemNestLevel] || 0;
    counter++;
    listCounters[key + "." + itemNestLevel] = counter; 
    
    // debug
    //prefix += ' ' + key + " - " + listItem.getNestingLevel().toString() + " - ";
    
    var nextSibling = listItem.getNextSibling();
    //var nextSiblingType = nextSibling.getType().toString();
    var isAtDocumentEnd = listItem.isAtDocumentEnd();
    var nextIsListItem = (nextSibling && (nextSibling.getType() == DocumentApp.ElementType.LIST_ITEM));
    var nextIsSameListId = false;
    if(nextIsListItem)
      nextIsSameListId = nextSibling.getListId() == listItem.getListId();
    var currentNestingLevel = listItem.getNestingLevel();
    
    var nextNestingLevel = false;
    if(nextIsListItem)
    {
      // known list? Get absolute nesting level!
      if(listIds.indexOf(nextSibling.getListId())>=0)
        nextNestingLevel = getAbsoluteListItemNestLevel(nextSibling);
      else // else global nesting level is base
        nextNestingLevel = globalNestedListLevel + nextSibling.getNestingLevel();
    }
    var nextIsLowerNestingLevel = (nextIsListItem && (nextNestingLevel < globalNestedListLevel));
    
    while( (isAtDocumentEnd || (!nextIsListItem) || nextIsLowerNestingLevel  ) && globalNestedListLevel >= 0) 
    {
      suffix += "</ul>";
      globalNestedListLevel--;
      
      nextSibling = listItem.getNextSibling();
    //  nextSiblingType = nextSibling.getType().toString();
      isAtDocumentEnd = listItem.isAtDocumentEnd();
      nextIsListItem = (nextSibling && (nextSibling.getType() == DocumentApp.ElementType.LIST_ITEM));
      nextIsSameListId = false;
      if(nextIsListItem)
        nextIsSameListId = nextSibling.getListId() != listItem.getListId();
      
      nextNestingLevel = false;
      if(nextIsListItem)
      {
        // known list? Get absolute nesting level!
        if(listIds.indexOf(nextSibling.getListId())>0)
          nextNestingLevel = getAbsoluteListItemNestLevel(nextSibling);
        else // else global nesting level is base
          nextNestingLevel = globalNestedListLevel + nextSibling.getNestingLevel();
      }
      var nextIsLowerNestingLevel = (nextIsListItem && (nextNestingLevel < globalNestedListLevel));
      
      nextIsLowerNestingLevel = (nextIsListItem && (getAbsoluteListItemNestLevel(nextSibling) < globalNestedListLevel));
    }
    //   else
    //   {
    //     counter++;
    //  listCounters[key] = counter;
    //  }
  } else if (itemType === DocumentApp.ElementType.TABLE) {
    var row = item.getRow(0)
    var numCells = row.getNumCells();
    var tableWidth = 0;
    
    for (var i = 0; i < numCells; i++) {
      tableWidth += item.getColumnWidth(i);
    }
    Logger.log("TABLE tableWidth: " + tableWidth);
    
    //https://stackoverflow.com/questions/339923/set-cellpadding-and-cellspacing-in-css
    var style = ' style="border-collapse: collapse; width:' + tableWidth + 'px; "';
    
    prefix = '<table' + style + '>', suffix = "</table>";
    //Logger.log("TABLE: " + JSON.stringify(item));
  } else if (itemType === DocumentApp.ElementType.TABLE_ROW) {
    
    var minimumHeight = item.getMinimumHeight();
    Logger.log("TABLE_ROW getMinimumHeight: " + minimumHeight);
    
    prefix = "<tr>", suffix = "</tr>";
    //Logger.log("TABLE_ROW: " + JSON.stringify(item));
  } else if (itemType === DocumentApp.ElementType.TABLE_CELL) {
    /*
    BACKGROUND_COLOR  Enum  The background color of an element (Paragraph, Table, etc) or document.
    BORDER_COLOR  Enum  The border color, for table elements.
    BORDER_WIDTH  Enum  The border width in points, for table elements.
    PADDING_BOTTOM  Enum  The bottom padding setting in points, for table cell elements.
    PADDING_LEFT  Enum  The left padding setting in points, for table cell elements.
    PADDING_RIGHT Enum  The right padding setting in points, for table cell elements.
    PADDING_TOP     Enum  The top padding setting in points, for table cell elements.
    VERTICAL_ALIGNMENT  Enum  The vertical alignment setting, for table cell elements.
    WIDTH         Enum  The width setting, for table cell and image elements.
    */
    
    //https://wiki.selfhtml.org/wiki/HTML/Tabellen/Zellen_verbinden
    var colSpan = item.getColSpan();
    Logger.log("TABLE_CELL getColSpan: " + colSpan);
    // colspan="3"
    
    var rowSpan = item.getRowSpan();
    Logger.log("TABLE_CELL getRowSpan: " + rowSpan);
    // rowspan ="3"
    
    //TODO: WIDTH must be recalculated in percent
    var atts = item.getAttributes();
    
    var style = ' style=" width:' + atts.WIDTH + 'px; border: 1px solid black; padding: 5px;"';
    
    prefix = '<td' + style + '>', suffix = "</td>";
    //Logger.log("TABLE_CELL: " + JSON.stringify(item));
  } else if (itemType === DocumentApp.ElementType.FOOTNOTE) {
    //TODO
    var note = item.getFootnoteContents();
    var counter = footnotes.length + 1;
    output.push("<sup><a name='link" + counter + "' href='#footnote" + counter + "'>[" + counter + "]</a></sup>");
    var newFootnote = "<aside class='footnote' epub:type='footnote' id='footnote" + counter + "'><a name='footnote" + counter + "' epub:type='noteref'>[" + counter + "]</a>";
    var numChildren = note.getNumChildren();
    for (var i = 0; i < numChildren; i++) {
      var child = note.getChild(i);
      newFootnote += processItem_V1(child, listCounters, images, imagesOptions, footnotes);
    }
    newFootnote += "<a href='#link" + counter + "' id='#link" + counter + "'>↩</a></aside>"
    footnotes.push(newFootnote);
    Logger.log("FOOTNOTE: " + JSON.stringify(item));
  } else if (itemType === DocumentApp.ElementType.HORIZONTAL_RULE) {
    output.push("<hr />");
    //Logger.log("HORIZONTAL_RULE: " + JSON.stringify(item));
  } else if (itemType === DocumentApp.ElementType.UNSUPPORTED) {
    Logger.log("UNSUPPORTED: " + JSON.stringify(item));
  }
  
  output.push(prefix);
  
  if (hasPositionedImages === true) {
    processPositionedImages(positionedImages, images, output, imagesOptions);
  }
  
  if (item.getType() == DocumentApp.ElementType.TEXT) {
    processText(item, output);
  } else {
    
    if (item.getNumChildren) {
      var numChildren = item.getNumChildren();
      
      // Walk through all the child elements of the doc.
      for (var i = 0; i < numChildren; i++) {
        var child = item.getChild(i);
        output.push(processItem_V1(child, listCounters, images, imagesOptions, footnotes));
      }
    }
    
  }
  
  output.push(suffix);
  return output.join('');
}

//points = pixel * 72 / 96
//1em = 16px (Browser Default wert) 
//1px = 1/16 = 0.0625em 

function pointsToPixel(points) {
  return points * 96 / 72;
}

function pixelToPoints(pixel) {
  return pixel * 72 / 96;
}

function pixelToEm(pixel) {
  return pixel / 16;
}

function emToPixel(em) {
  return em * 16;
}


/**
* @param {Text} item - https://developers.google.com/apps-script/reference/document/text
* @param {string[]} output
*/
function processText(item, output) {
  var text = item.getText();
  var indices = item.getTextAttributeIndices();
  
  if (text === '\r') 
  {
    Logger.log("\\r: ");
    return;
  }
  
  for (var i = 0; i < indices.length; i++)
  {
    var partAtts = item.getAttributes(indices[i]);
    var startPos = indices[i];
    var endPos = i + 1 < indices.length ? indices[i + 1] : text.length;
    var partText = text.substring(startPos, endPos);
    
    partText = partText.replace(new RegExp("(\r)", 'g'), "<br/>");
    //Logger.log(partText);
    dumpAttributes(partAtts);
    
    //TODO if only ITALIC use: <blockquote></blockquote>
    
    //TODO: change html tags to css (i, strong, u)
    
    //css font-style:italic;
    if (partAtts.ITALIC) {
      output.push('<i>');
    }
    //css font-weight: bold;
    if (partAtts.BOLD) {
      output.push('<strong>');
    }
    //css text-decoration: underline
    if (partAtts.UNDERLINE) {
      output.push('<u>');
    }
    
    var style = "";
    
    // font family, color and size changes disabled
    /*if (partAtts.FONT_FAMILY) {
    style = style + 'font-family: ' + partAtts.FONT_FAMILY + '; ';
    }
    if (partAtts.FONT_SIZE) {
    var pt = partAtts.FONT_SIZE;
    var em = pixelToEm(pointsToPixel(pt));
    style = style + 'font-size: ' + pt + 'pt;  font-size: ' + em + 'em; ';
    }
    if (partAtts.FOREGROUND_COLOR) {
    style = style + 'color: ' + partAtts.FOREGROUND_COLOR + '; '; //partAtts.FOREGROUND_COLOR
    }
    if (partAtts.BACKGROUND_COLOR) {
    style = style + 'background-color: ' + partAtts.BACKGROUND_COLOR + '; ';
    }*/
    if (partAtts.STRIKETHROUGH) {
      style = style + 'text-decoration: line-through; ';
    }
    
    var a = item.getTextAlignment(startPos);
    if (a !== DocumentApp.TextAlignment.NORMAL && a !== null) {
      if (a === DocumentApp.TextAlignment.SUBSCRIPT) {
        style = style + 'vertical-align : sub; font-size : 60%; ';
      } else if (a === DocumentApp.TextAlignment.SUPERSCRIPT) {
        style = style + 'vertical-align : super; font-size : 60%; ';
      }
    }
    
    // If someone has written [xxx] and made this whole text some special font, like superscript
    // then treat it as a reference and make it superscript.
    // Unfortunately in Google Docs, there's no way to detect superscript
    if (partText.indexOf('[') == 0 && partText[partText.length - 1] == ']') {
      if (style !== "") {
        style = ' style="' + style + '"';
      }
      output.push('<sup' + style + '>' + partText + '</sup>');
    } else if (partText.trim().indexOf('http://') == 0 || partText.trim().indexOf('https://') == 0) {
      if (style !== "") {
        style = ' style="' + style + '"';
      }
      output.push('<a' + style + ' href="' + partText + '" rel="nofollow">' + partText + '</a>');
    } else if (partAtts.LINK_URL) {
      if (style !== "") {
        style = ' style="' + style + '"';
      }
      output.push('<a' + style + ' href="' + partAtts.LINK_URL + '" rel="nofollow">' + partText + '</a>');
    } else {
      if (style !== "") {
        partText = '<span style="' + style + '">' + partText + '</span>';
      }
      output.push(partText);
    }
    
    if (partAtts.ITALIC) {
      output.push('</i>');
    }
    if (partAtts.BOLD) {
      output.push('</strong>');
    }
    if (partAtts.UNDERLINE) {
      output.push('</u>');
    }
    
  }
  //}
}
