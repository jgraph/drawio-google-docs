/**
 * diagrams.net Diagrams Slides add-on v2.4
 * Copyright (c) 2020, JGraph Ltd
 */
var EXPORT_URL = "https://convert.diagrams.net/node/export";
var DIAGRAMS_URL = "https://app.diagrams.net/";
var DRAW_URL = "https://www.draw.io/";
var SCALING_VALUE = 0.8; // Google Slides seem to be downscaling all images by this amount
var BORDER = 20; // Offset and padding for images

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen()
{
  SlidesApp.getUi().createAddonMenu()
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
  SlidesApp.getUi().showModalDialog(html, 'Select files');
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
    
      // if there are selected items in the slides, assume they are going to be replaced
      // by the newly inserted images
      var selectionCoordinates = getSelectionCoordinates();
      var offsetX = Math.max(BORDER, selectionCoordinates[0]);
      var offsetY = Math.max(BORDER, selectionCoordinates[1]);
    
      deleteSelectedElements();
    
      var step = 10;
    
      for (var i = 0; i < items.length; i++)
      {
        try
        {
          var ins = insertDiagram(items[i].id, items[i].page, offsetX, offsetY); 
        
          if (ins != null)
          { 
            inserted++;
            insertedElts.push(ins);
            offsetX = offsetX + step;
            offsetY = offsetY + step;
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
        SlidesApp.getUi().alert(msg);
      }
      else
	  {
	    for (var i = 0; i < insertedElts.length; i++)
	  	{
	  		insertedElts[i].select(i == 0);
	   	}
	  }
  }
}

/**
  Finds left-most and top-most coordinates of selected page elements; (0,0) by default
  @return left-most and top-most coordinates in an array
**/
function getSelectionCoordinates() 
{
  var selection = SlidesApp.getActivePresentation().getSelection();
  switch (selection.getSelectionType()) 
  {
    case SlidesApp.SelectionType.PAGE_ELEMENT:
    {
      // only interested if selection is containing page elements
      var elements = selection.getPageElementRange();
      var top = 1000000;
      var left = 1000000;
      if (elements) 
      {
        // find the left-most, top-most coordinate of selected elements
        var pageElements = elements.getPageElements();
        for (var i = 0; i < pageElements.length; i++) 
        {
          var element = pageElements[i];
          var elementTop = element.getTop();
          var elementLeft = element.getLeft();
          if (top > elementTop)
            top = elementTop;
          if (left > elementLeft)
            left = elementLeft;
        }
        return [left, top];
      }
    }
  }
  return [0, 0];
}

/**
  Deletes selected elements
**/
function deleteSelectedElements() 
{
  var selection = SlidesApp.getActivePresentation().getSelection();
  switch (selection.getSelectionType()) 
  {
    case SlidesApp.SelectionType.PAGE_ELEMENT: 
      {
      // only interested if selection is containing page elements
      var elements = selection.getPageElementRange();
      if (elements) {
        var pageElements = elements.getPageElements();
        // find the left-most, top-most coordinate of selected elements
        for (var i = 0; i < pageElements.length; i++) 
        {
          // delete all selected page elements
          var element = pageElements[i];
          element.remove();
        }
      }
    } 
  }
}

/**
 * Inserts the given diagram at the given position.
 */
function insertDiagram(id, page, offsetX, offsetY)
{
  var result = fetchImage(id, page, 'auto');
  
  var blob = result[0];
  var w = result[1] * SCALING_VALUE;
  var h = result[2] * SCALING_VALUE;
  var img = null;
  
  if (blob != null)
  {
      var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide();
      img = slide.insertImage(blob);
      img.setLeft(offsetX);
      img.setTop(offsetY);

      var link = createLink(id, page, (result.length > 3) ? result[3] : null);
	  img.setLinkUrl(link);

	  var wmax = SlidesApp.getActivePresentation().getPageWidth() - 2 * BORDER;
      var hmax = SlidesApp.getActivePresentation().getPageHeight() - 2 * BORDER;
    
      if (wmax > 0 && hmax > 0)
      {    
		  // Scales to document width if not placeholder
	      if (w == 0 && h == 0)
	      {
		      w = img.getWidth();
	          h = img.getHeight();
	      }

	      var minscale = Math.min(1, Math.min(wmax / w, hmax / h));
	      img.setWidth(w * minscale);
	      img.setHeight(h * minscale);
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
  var selection = SlidesApp.getActivePresentation().getSelection();
    
  if (selection)
  {
    switch (selection.getSelectionType()) 
    {
      case SlidesApp.SelectionType.PAGE_ELEMENT:
      {
        var selected = selection.getPageElementRange();
        if (!selected)
          return;
        
        selected = selected.getPageElements();
        
        var elts = [];
        
        // Unwraps selection elements
        for (var i = 0; i < selected.length; i++)
        {
          var pageElement = selected[i];
          
          switch (pageElement.getPageElementType())
          {
            case SlidesApp.PageElementType.IMAGE:
            {
              elts.push(selected[i].asImage());
            }
          }
        }
        
        updateElements(elts);
      }
    }
  }
  else
  {
    SlidesApp.getUi().alert("No selection");
  }
}

/**
 * Updates all diagrams in the document.
 */
function updateAll()
{
  // collect all slides
  var slides = SlidesApp.getActivePresentation().getSlides();
  var images = [];
  
  for (var i = 0; i < slides.length; i++)
  {
    // collect all images on all slides
    var slide = slides[i];
    var slideImages = slide.getImages();
    images = images.concat(slideImages);
  }
  
  updateElements(images, true);
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
      SlidesApp.getUi().alert(msg);
    }
    else
    {
        for (var i = 0; i < updatedElts.length; i++)
    	{
    		updatedElts[i].select(i == 0);
    	}
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
  
  if (elt.getPageElementType() == SlidesApp.PageElementType.IMAGE)
  {
    var url = elt.getLink();
    
    if (url != null)
    {
      url = url.getUrl();
    }
    
    if (url == null)
    {
      if (!ignoreMissingLinks)
      {
        throw new Error("Missing link");
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
        // commenting this out as well - invalid link might indicate image is not coming from draw.io
        // throw new Error("Invalid link " + url);
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
  var result = fetchImage(id, page, scale, pageId);
  
  var isOK = false;
  
  if (result != null) 
  {
    var blob = result[0];
    var w = result[1];
    var h = result[2];
    
    if (blob != null)
    {
      isOK = true;
      // There doesn't seem to be a natural way to replace images in SlidesApp
      // Slides API seems to only provide means to get a page and a group associated with the image
      // Groups only allow removal of elements though, not insertions (seems like a half-baked API)
      
      // This code just adds a new image to page to the same position as the old image and removes the old image
      // TODO: No group information will be preserved right now
      var page = elt.getParentPage();
      var left = elt.getLeft();
      var top = elt.getTop();
      var wmax = elt.getWidth();
      var hmax = elt.getHeight();
      
      var minscale = Math.min(1, Math.min(wmax / w, hmax / h));

      // replace image with the same link
      var img = page.insertImage(blob, left, top, w * minscale, h * minscale);
      var link = createLink(id, page, result[3]);
      img.setLinkUrl(link);
      
      elt.remove();
    }
  }
  if (!isOK)
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
      
    	var data = {
		  "format": "png",
          "scale": scale || "1",
		  "xml": fileData,
		  "extras": "{\"pageWidth\": " + Math.round(SlidesApp.getActivePresentation().getPageWidth() * 1.5) + ", \"isPng\": " + isPng + ", \"isGoogleApp\": true}"
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
  var selection = SlidesApp.getActivePresentation().getSelection();
    
  if (selection)
  {
    switch (selection.getSelectionType()) 
    {
      case SlidesApp.SelectionType.PAGE_ELEMENT:
      {
        var selected = selection.getPageElementRange();

        if (!selected)
       	{
          return;
        }

        selected = selected.getPageElements();
        
        // Unwraps selection elements
        for (var i = 0; i < selected.length; i++)
        {
          var pageElement = selected[i];
          
          switch (pageElement.getPageElementType())
          {
            case SlidesApp.PageElementType.IMAGE:
            {
            	var elt = selected[i].asImage();
            	var url = elt.getLink();
    
    			if (url != null)
    			{
      				url = url.getUrl();
    			}
            	
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
      }
    }
  }
  else
  {
    SlidesApp.getUi().alert("No selection");
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
  SlidesApp.getUi().showModalDialog(html, "Opening...");
}
