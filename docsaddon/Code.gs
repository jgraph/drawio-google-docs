/**
 * diagrams.net Diagrams Docs add-on v2.4
 * Copyright (c) 2020, JGraph Ltd
 */
var EXPORT_URL = "https://convert.diagrams.net/node/export";
var DIAGRAMS_URL = "https://app.diagrams.net/";
var DRAW_URL = "https://www.draw.io/";

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen()
{
  DocumentApp.getUi().createAddonMenu()
      .addItem("Insert Diagrams...", "insertDiagrams")
      .addSeparator()
      .addItem("Update Selected", "updateSelected")
      .addItem("Update All", "updateAll")
      .addSeparator()
      .addItem("Edit Selected...", "editSelected")
      .addItem("New Diagram...", "newDiagram")
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall()
{
  onOpen();
}

/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  DriveApp.getFolders();
  return ScriptApp.getOAuthToken();
}

/**
 * Shows a picker and lets the user select multiple diagrams to insert.
 */
function insertDiagrams()
{
  var html = HtmlService.createHtmlOutputFromFile('Picker.html')
      .setWidth(640).setHeight(480)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showModalDialog(html, 'Select files');
}

/**
 * Inserts an image for the given diagram.
 */
function pickerHandler(items)
{
  if (items != null && items.length > 0)
  {
      var insertedElts = [];
      var inserted = 0;
      var errors = [];
    
      for (var i = 0; i < items.length; i++)
      {
        try
        {
          var ins = insertDiagram(items[i].id, items[i].page); 
        
          if (ins != null)
          {
          	insertedElts.push(ins);
            inserted++;
          }
	      else
	      {
	    	errors.push("- " + items[i].name + ": Unknown error");
	      }
        }
        catch (e)
        {
          errors.push("- " + items[i].name + ": " + e.message);
        }
      }
    
      // Shows message only in case of errors
      if (errors.length > 0)
      {
        var msg = "";

        if (errors.length > 0)
        {
          msg += errors.length + " insert" + ((errors.length > 1) ? "s" : "") + " failed:\n";
        }
        
        msg += errors.join("\n");
        DocumentApp.getUi().alert(msg);
      }
      else if (insertedElts.length > 0)
      {
      	var doc = DocumentApp.getActiveDocument();
		var rangeBuilder = doc.newRange();
		
		for (var i = 0; i < insertedElts.length; i++)
		{
  			rangeBuilder.addElement(insertedElts[i]);
		}
		
		doc.setSelection(rangeBuilder.build());
      }
  }
}

/**
 * Inserts the given diagram at the given position.
 */
function insertDiagram(id, page)
{
  var result = fetchImage(id, page, 'auto');
  var blob = result[0];
  var img = null;
  
  if (blob != null)
  {
	  var doc = DocumentApp.getActiveDocument();
	  var pos = doc.getCursor();
	  
	  img = (pos != null) ? pos.insertInlineImage(blob) : doc.getBody().appendImage(blob);
	  img.setLinkUrl(createLink(id, page, (result.length > 3) ? result[3] : null));
	  
	  var wmax = 2 * doc.getBody().getPageWidth() / 3;
	  var hmax = doc.getBody().getPageHeight();
	  
	  if (wmax > 0 && hmax > 0)
	  {
	  	  // Scales to document width if not placeholder
		  var style = img.getAttributes();
		  var w = style[DocumentApp.Attribute.WIDTH];
		  var h = style[DocumentApp.Attribute.HEIGHT];
		  
	      var minscale = Math.min(1, Math.min(wmax / w, hmax / h));
 	      style[DocumentApp.Attribute.WIDTH] = w * minscale;
	      style[DocumentApp.Attribute.HEIGHT] = h * minscale;
          img.setAttributes(style);
	  }
  }
  else
  {
    throw new Error("Invalid image " + id);
  }
  
  return img;
}

/**
 * Updates the selected diagrams in-place.
 */
function updateSelected()
{
  var selection = DocumentApp.getActiveDocument().getSelection();
    
  if (selection)
  {
    var selected = selection.getSelectedElements();
    var elts = [];
    
    // Unwraps selection elements
    for (var i = 0; i < selected.length; i++)
    {
      elts.push(selected[i].getElement());
    }
    
    updateElements(elts);
  }
  else
  {
    DocumentApp.getUi().alert("Selection is empty");
  }
}

/**
 * Updates all diagrams in the document.
 */
function updateAll()
{
  updateElements(DocumentApp.getActiveDocument().getBody().getImages(), true);
}

/**
 * Updates all diagrams in the document.
 */
function updateElements(elts, ignoreMissingLinks)
{
  if (elts != null)
  {
    var updatedElts = [];
    var updated = 0;
    var errors = [];
    
    for (var i = 0; i < elts.length; i++)
    {
      try
      {
      	var upd = updateElement(elts[i], ignoreMissingLinks);
      
        if (upd != null)
        {
          updatedElts.push(upd);
          updated++;
        }
      }
      catch (e)
      {
        errors.push("- " + e.message);
      }
    }
    
    // Shows status in case of errors or multiple updates
    if (errors.length > 0 || updated > 1)
    {
      var msg = "";
      
      if (updated > 0)
      {
        msg += updated + " diagram" + ((updated > 1) ? "s" : "") + " updated\n";
      }
      
      if (errors.length > 0)
      {
        msg += errors.length + " update" + ((errors.length > 1) ? "s" : "") + " failed:\n";
      }
      
      msg += errors.join("\n");
      DocumentApp.getUi().alert(msg);
    }
    else if (updatedElts.length > 0)
    {
  	  var doc = DocumentApp.getActiveDocument();
	  var rangeBuilder = doc.newRange();
	
	  for (var i = 0; i < updatedElts.length; i++)
	  {
	    rangeBuilder.addElement(updatedElts[i]);
	  }
	
  	  doc.setSelection(rangeBuilder.build());
    }
  }
}

/**
 * Returns true if the given URL points to draw.io
 */
function createLink(id, page, pageId)
{
  var params = [];
  
  if (pageId != null)
  {
  	params.push('page-id=' + pageId);
  }
  else if (page != null && page != "0")
  {
    params.push('page=' + page);
  }
  
  params.push('scale=auto');
  
  return DIAGRAMS_URL + ((params.length > 0) ? "?" + params.join("&") : "") + "#G" + id;
}

/**
 * Returns true if the given URL points to draw.io
 */
function isValidLink(url)
{
  return url != null && (url.substring(0, DRAW_URL.length) == DRAW_URL ||
  	url.substring(0, DIAGRAMS_URL.length) == DIAGRAMS_URL ||
  	url.substring(0, 22) == "https://drive.draw.io/");
}

/**
 * Returns the diagram ID for the given URL.
 */
function getDiagramId(url)
{
  return url.substring(url.lastIndexOf("#G") + 2);
}

/**
 * Returns the diagram ID for the given URL.
 */
function getUrlParams(url)
{
  var result = {};
  var idx = url.indexOf("?");
  
  if (idx > 0)
  {
    var idx2 = url.indexOf("#", idx + 1);
    
    if (idx2 < 0)
    {
      idx2 = url.length;
    }
    
    if (idx2 > idx)
    {
      var search = url.substring(idx + 1, idx2);
      var tokens = search.split("&");
      
      for (var i = 0; i < tokens.length; i++)
      {
         var idx3 = tokens[i].indexOf('=');
        
         if (idx3 > 0)
         {
           result[tokens[i].substring(0, idx3)] = tokens[i].substring(idx3 + 1);
         }
      }
    }
  }
  
  return result;
}

/**
 * Updates the diagram in the given inline image and returns the new inline image.
 */
function updateElement(elt, ignoreMissingLinks)
{
  var result = null;
  
  if (elt.getType() == DocumentApp.ElementType.INLINE_IMAGE)
  {
    var url = elt.getLinkUrl();
    
    if (url == null)
    {
      if (!ignoreMissingLinks)
      {
        throw new Error("Missing link")
      }
    }
    else if (isValidLink(url))
    {
      var id = getDiagramId(url);
      var params = getUrlParams(url);
      
      if (id != null && id.length > 0)
      {
        result = updateDiagram(id, params["page"], params["scale"] || 1, elt, params["page-id"]);
      }
      else
      {
        throw new Error("Invalid link " + url);
      }
    }
  }
  
  return result;
}

/**
 * Updates the diagram in the given inline image and returns the new inline image.
 */
function updateDiagram(id, page, scale, elt, pageId)
{
  var img = null;
  var blob = fetchImage(id, page, scale, pageId)[0];
  
  if (blob != null)
  {
	  var par = elt.getParent();
	  var idx = par.getChildIndex(elt);
	  img = par.insertInlineImage(idx + 1, blob);
	  img.setLinkUrl(elt.getLinkUrl());
	  
	  var style = elt.getAttributes();
	  var w = style[DocumentApp.Attribute.WIDTH];
	  
	  if (w > 1)
	  {
		  var style2 = img.getAttributes();
		  var w2 = style2[DocumentApp.Attribute.WIDTH];
		  var aspect = w2 / style2[DocumentApp.Attribute.HEIGHT];
		  
		  // Keeps width, aspect and link
		  style[DocumentApp.Attribute.WIDTH] = w;
		  style[DocumentApp.Attribute.HEIGHT] = w / aspect;
		  
		  img.setAttributes(style);
	  }
	  
	  elt.removeFromParent();
  }
  else
  {
    throw new Error("Invalid image " + id);
  }
  
  return img;
}

/**
 * Fetches the diagram for the given document ID.
 */
function fetchImage(id, page, scale, pageId)
{
    var file = DriveApp.getFileById(id);

    if (file != null && file.getSize() > 0)
    {
        var isPng = file.getMimeType() == "image/png";
      
        var fileData = isPng? Utilities.base64Encode(file.getBlob().getBytes()) : encodeURIComponent(file.getBlob().getDataAsString());
      
    	var data =
    	{
		  "format": "png",
          "scale": scale || "1",
		  "xml": fileData,
          "extras": "{\"pageWidth\": " + Math.round(DocumentApp.getActiveDocument().getBody().getPageWidth() * 1.5) + ", \"isPng\": " + isPng + ", \"isGoogleApp\": true}" 
		};
		  
		if (pageId != null)
		{
			data.pageId = pageId;
		}
		else
		{
			data.from = page || "0";
		}
    
	    var response = UrlFetchApp.fetch(EXPORT_URL,
	    {
		  "method": "post",
		  "payload": data
		});
		
		var headers = response.getHeaders();
		
		return [response.getBlob(), headers["content-ex-width"] || 0, headers["content-ex-height"] || 0, headers["content-page-id"]];
    }
    else
    {
    	// Returns an empty, transparent 1x1 PNG image as a placeholder
    	return [Utilities.newBlob(Utilities.base64Decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+M9QDwADgQF/e5IkGQAAAABJRU5ErkJggg=="), "image/png")];
    }
}

/**
 * Edits the selected diagram.
 */
function editSelected()
{
  var selection = DocumentApp.getActiveDocument().getSelection();
    
  if (selection)
  {
    var selected = selection.getSelectedElements();
    
    // Unwraps selection elements
    for (var i = 0; i < selected.length; i++)
    {
      var elt = selected[i].getElement();
  
	  if (elt.getType() == DocumentApp.ElementType.INLINE_IMAGE)
	  {
	    var url = elt.getLinkUrl();
	    
	    if (isValidLink(url))
	    {
	      var id = getDiagramId(url);
	      
	      if (id != null && id.length > 0)
	      {
	      	openUrl('https://www.draw.io/#G' + id);
	      	break;
	      }
	    }
	  }
    }
  }
  else
  {
    DocumentApp.getUi().alert("Could not open diagram for editing");
  }
}

/**
 * Creates a new diagram.
 */
function newDiagram()
{
  openUrl('https://app.diagrams.net/?mode=google');
}

/**
 * Open a URL in a new tab.
 */
function openUrl(url)
{
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth(90).setHeight(1);
  DocumentApp.getUi().showModalDialog(html, "Opening...");
}
