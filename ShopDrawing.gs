// ─────────────────────────────────────────────────────────────────
// SHOP DRAWING AUTOMATION — version# 04/01-01:00AM by Claude Sonnet 4.6
// ─────────────────────────────────────────────────────────────────

/**
 * Called from handleEditMove_ when col D = "Quote Sent" on Leads or F/U.
 * Duplicates the Shop Drawing Slides template into the customer's Drive folder,
 * renames it, and fills footer placeholders. Skips silently if already exists.
 */
function m_createShopDrawing_(sheet, row) {
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    var templateId  = scriptProps.getProperty('SHOP_DRAWING_TEMPLATE_ID');
    if (!templateId) {
      Logger.log('m_createShopDrawing_: SHOP_DRAWING_TEMPLATE_ID not set.');
      SpreadsheetApp.getActive().toast('Shop Drawing failed: Template ID not set in Script Properties.', 'Shop Drawing Error', 6);
      return;
    }

    var colF  = sheet.getRange(row, 6);
    var colE  = sheet.getRange(row, 5).getValue();   // Customer Name
    var colH  = sheet.getRange(row, 8).getValue();   // Phone
    var colJ  = sheet.getRange(row, 10).getValue();  // Address
    var colR  = String(sheet.getRange(row, 18).getValue() || '').trim(); // Job Type
    var colAA = sheet.getRange(row, 27).getValue();  // Frame Type
    var colAB = sheet.getRange(row, 28).getValue();  // Fabric

    var displayName = colF.getDisplayValue() || 'Unknown';
    var fileName    = 'Shop Drawing - ' + displayName;

    // Extract folder from col F hyperlink
    var folderUrl = m_getFolderUrlFromCell_(colF);
    if (!folderUrl) {
      Logger.log('m_createShopDrawing_: No folder URL found in col F, row ' + row);
      SpreadsheetApp.getActive().toast('Shop Drawing failed: No folder link in col F.', 'Shop Drawing Error', 6);
      return;
    }
    var folderId = m_extractFolderIdFromUrl_(folderUrl);
    if (!folderId) {
      Logger.log('m_createShopDrawing_: Could not parse folder ID from URL: ' + folderUrl);
      SpreadsheetApp.getActive().toast('Shop Drawing failed: Could not parse folder ID.', 'Shop Drawing Error', 6);
      return;
    }

    var folder = DriveApp.getFolderById(folderId);

    // Skip silently if file already exists
    var existing = folder.getFilesByName(fileName);
    if (existing.hasNext()) {
      Logger.log('m_createShopDrawing_: Already exists, skipping — ' + fileName);
      SpreadsheetApp.getActive().toast('Shop Drawing already exists — skipped.', 'Shop Drawing', 4);
      return;
    }

    // ── Search for inline 3D render image from Proposal Review email ──
    var renderBlob = null;
    try {
      var searchQuery = 'in:sent subject:"Proposal Review: ' + displayName + '"';
      var threads = GmailApp.search(searchQuery, 0, 1);
      if (threads.length > 0) {
        var messages = threads[0].getMessages();
        if (messages.length > 0) {
          var message = messages[messages.length - 1]; // most recent
          var inlineImages = message.getAttachments({
            includeInlineImages: true,
            includeAttachments:  false
          });
          // Filter for image types only
          var imageAttachments = inlineImages.filter(function(a) {
            return a.getContentType() && a.getContentType().indexOf('image/') === 0;
          });
          if (imageAttachments.length > 0) {
            renderBlob = imageAttachments[0].copyBlob();
            Logger.log('m_createShopDrawing_: Found inline render image.');
          } else {
            Logger.log('m_createShopDrawing_: No inline images in email — skipping image.');
          }
        }
      } else {
        Logger.log('m_createShopDrawing_: No Proposal Review email found — skipping image.');
      }
    } catch (emailErr) {
      Logger.log('m_createShopDrawing_: Email search error — ' + emailErr.message);
    }

    // ── Copy template into customer folder ──
    var templateFile = DriveApp.getFileById(templateId);
    var copy         = templateFile.makeCopy(fileName, folder);
    var presentation = SlidesApp.openById(copy.getId());
    var slide        = presentation.getSlides()[0];
    var today        = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');

    // ── Insert 3D render image above footer if found ──
    if (renderBlob) {
      try {
        // Get slide dimensions
        var slideWidth  = presentation.getPageWidth();   // points
        var slideHeight = presentation.getPageHeight();  // points

        // Footer occupies roughly bottom 1.5 inches (108 pts)
        // Image fills the remaining space with padding
        var imgPadding   = 18;   // 0.25 inch padding
        var footerHeight = 108;
        var maxWidth     = (slideWidth  - (imgPadding * 2)) * 0.65;
        var maxHeight    = (slideHeight - footerHeight - (imgPadding * 2)) * 0.65;

        // Insert first to get natural dimensions, then scale with locked aspect ratio
        var insertedImage = slide.insertImage(renderBlob);
        var naturalWidth  = insertedImage.getWidth();
        var naturalHeight = insertedImage.getHeight();
        var scale         = Math.min(maxWidth / naturalWidth, maxHeight / naturalHeight);
        var imgWidth      = naturalWidth  * scale;
        var imgHeight     = naturalHeight * scale;

        // Center on page
        var imgLeft = (slideWidth  - imgWidth)  / 2;
        var imgTop  = (slideHeight - imgHeight) / 2;

        insertedImage.setWidth(imgWidth);
        insertedImage.setHeight(imgHeight);
        insertedImage.setLeft(imgLeft);
        insertedImage.setTop(imgTop);

        Logger.log('m_createShopDrawing_: Render image inserted into slide.');
      } catch (imgErr) {
        Logger.log('m_createShopDrawing_: Image insert error — ' + imgErr.message);
      }
    }

    // ── Fill footer placeholders ──
    var isAwardedSheet = sheet.getName() === 'Awarded';

    var colRLower = colR.toLowerCase();
    var notesText = '';
    if (colRLower === 'aluminum canopy') {
      notesText =
        'NOTES:\n' +
        'Structures designed with FL BLDG Code 2023 8E-1620 & not to exceed 170 MPH Vasd\n' +
        'ASCE 7 - 22CH 6, 29 Exp C-3 sec. Gust Cat 2.\n' +
        'STRUCTURAL ALUMINUM:\n' +
        'These specifications shall apply to the design of aluminum alloy load carrying members. ' +
        'Computation of forces, moments, stresses and deflections shall be in accordance with accepted methods of elastic structural analysis and engineering design.\n' +
        '1. All elements shall be aluminum alloy 6051-T6 or alloy 6063-T52 unless noted otherwise\n' +
        '2. All welds to comply with A.W.S code (latest addition)\n' +
        '3. Cover welds with corrosion resistance coating\n' +
        '4. Structures are designed in accordance with the following codes: - Aluminum design manual\n' +
        '5. All frames have designed using rational analysis\n' +
        '6. All connections shall be fully welded with the structural aluminum alloy fillet weld: 5356 unless noted elsewhere\n' +
        '7. Aluminum members in contact with concrete shall be protected with alkali-resistant coating such as heavy bodied bituminous paint or water methacrylate lacquer';
    } else if (colRLower === 'complete build' || colRLower === 'demo + complete') {
      notesText =
        'NOTES:\n' +
        'Clear Height:\n' +
        'Mounting to:\n' +
        'Pitch:';
    }
    // Re-cover and anything else → notesText stays ''

    var replacements = {
      '{{CLIENT_NAME_PHONE}}': (displayName || '') + (colH ? '  |  ' + colH : ''),
      '{{ADDRESS}}':           String(colJ  || '-'),
      '{{FABRIC}}':            isAwardedSheet ? '' : String(colAB || '-'),
      '{{FRAME}}':             isAwardedSheet ? '' : String(colAA || '-'),
      '{{DATE}}':              today,
      '{{NOTES}}':             notesText
    };

    slide.getShapes().forEach(function(shape) {
      if (!shape.getText) return;
      var tf = shape.getText();
      Object.keys(replacements).forEach(function(key) {
        tf.replaceAllText(key, replacements[key]);
      });
    });

    presentation.saveAndClose();
    Logger.log('m_createShopDrawing_: Created "' + fileName + '" → folder ' + folderId);
    SpreadsheetApp.getActive().toast(
      'Shop Drawing created' + (renderBlob ? ' with 3D render' : ' (no render found)') + ': ' + fileName,
      'Shop Drawing',
      5
    );

  } catch (e) {
    Logger.log('m_createShopDrawing_ error: ' + e.message);
    SpreadsheetApp.getActive().toast('Shop Drawing failed: ' + e.message, 'Shop Drawing Error', 6);
  }
}

/**
 * Reads the hyperlink URL out of a cell.
 * Supports: =HYPERLINK() formula, rich-text link, plain URL string.
 */
function m_getFolderUrlFromCell_(cell) {
  // 1. HYPERLINK formula
  var formula = cell.getFormula();
  if (formula) {
    var match = formula.match(/=HYPERLINK\(\s*"([^"]+)"/i);
    if (match) return match[1];
  }
  // 2. Rich-text link
  try {
    var runs = cell.getRichTextValue().getRuns();
    for (var i = 0; i < runs.length; i++) {
      var url = runs[i].getLinkUrl();
      if (url) return url;
    }
  } catch (e) {}
  // 3. Plain text URL fallback
  var val = cell.getValue();
  if (typeof val === 'string' && val.indexOf('http') === 0) return val;
  return null;
}

/**
 * Parses a Google Drive folder ID from a Drive URL.
 * Handles /drive/folders/ID and ?id=ID formats.
 */
function m_extractFolderIdFromUrl_(url) {
  if (!url) return null;
  var patterns = [
    /\/folders\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/
  ];
  for (var i = 0; i < patterns.length; i++) {
    var m = url.match(patterns[i]);
    if (m) return m[1];
  }
  return null;
}

/**
 * ONE-TIME RUNNER — Run once from the Apps Script editor.
 * Creates the master Shop Drawing Slides template in your Drive root.
 * Automatically saves its file ID to Script Properties as SHOP_DRAWING_TEMPLATE_ID.
 *
 * ACTION REQUIRED AFTER RUNNING:
 * Open the file "Shop Drawing Template" in Google Slides →
 * File → Page Setup → Custom → enter 11 × 17 in → Apply.
 */
function m_createShopDrawingTemplate_() {
  var presentation = SlidesApp.create('Shop Drawing Template');
  var slide        = presentation.getSlides()[0];
  slide.getBackground().setSolidFill('#FFFFFF');

  // Layout in points (1pt = 1/72"). Default Slides canvas ≈ 720 × 540pt.
  // We treat the canvas as representing 11"×17" — user sets actual page size manually.
  var W = 720;
  var H = 1100;  // approx 17" proportional

  var footerTop    = H - 190;
  var footerHeight = 190;

  // Helper: insert a labeled placeholder box
  var addBox = function(left, top, width, height, label, placeholder, valueFontSize) {
    valueFontSize = valueFontSize || 7;

    var labelBox = slide.insertTextBox(label, left, top, width, 11);
    var lStyle   = labelBox.getText().getTextStyle();
    lStyle.setFontSize(5.5).setBold(false).setForegroundColor('#666666');
    labelBox.getFill().setTransparent();
    labelBox.getBorder().setTransparent();

    var valBox = slide.insertTextBox(placeholder, left, top + 12, width, height - 13);
    var vStyle = valBox.getText().getTextStyle();
    vStyle.setFontSize(valueFontSize).setBold(false).setForegroundColor('#000000');
    valBox.getFill().setTransparent();
    valBox.getBorder().setTransparent();
  };

  // ── Footer border line ──────────────────────────────────────────
  var borderLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, footerTop - 2, W, 2);
  borderLine.getFill().setSolidFill('#000000');
  borderLine.getBorder().setTransparent();

  // ── Branding block (left) ──────────────────────────────────────
  var brandBox = slide.insertTextBox(
    'WALKER AWNING\n5190 NW 10th Terrace\nFort Lauderdale, FL 33309\n' +
    '954-772-1951\nteam@walkerawning.com\nwalkerawning.com\nCCC1516477',
    2, footerTop + 2, 128, footerHeight - 4
  );
  brandBox.getText().getTextStyle().setFontSize(6).setForegroundColor('#000000');
  brandBox.getFill().setTransparent();
  brandBox.getBorder().setTransparent();

  // ── Vertical divider after branding ────────────────────────────
  var div1 = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 132, footerTop, 1, footerHeight);
  div1.getFill().setSolidFill('#000000');
  div1.getBorder().setTransparent();

  // ── Row 1 — top half of footer ─────────────────────────────────
  var r1Top = footerTop + 4;
  var r1H   = 88;
  var x     = 136;

  addBox(x,   r1Top, 220, r1H, 'JOB / CLIENT NAME & PHONE NUMBER', '{{CLIENT_NAME_PHONE}}', 8);
  x += 224;
  addBox(x,   r1Top, 130, r1H, 'FABRIC',               '{{FABRIC}}',  8);
  x += 134;
  addBox(x,   r1Top, 118, r1H, 'STEEL / ALUM (FRAME)', '{{FRAME}}',   8);
  x += 122;
  addBox(x,   r1Top, 104, r1H, 'YARDAGE',              '-',           8);

  // ── Horizontal mid-divider ──────────────────────────────────────
  var midDiv = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 132, footerTop + 95, W - 132, 1);
  midDiv.getFill().setSolidFill('#000000');
  midDiv.getBorder().setTransparent();

  // ── Row 2 — bottom half of footer ──────────────────────────────
  var r2Top = footerTop + 100;
  var r2H   = 86;
  x = 136;

  addBox(x,   r2Top, 220, r2H, 'ADDRESS',                    '{{ADDRESS}}',     8);
  x += 224;
  addBox(x,   r2Top,  72, r2H, 'PROJECT MGR',                '{{PROJECT_MGR}}', 8);
  x += 76;
  addBox(x,   r2Top,  60, r2H, 'DATE',                       '{{DATE}}',        8);
  x += 64;
  addBox(x,   r2Top,  56, r2H, 'SCALLOP #',                  '-',               8);
  x += 60;
  addBox(x,   r2Top,  60, r2H, 'VALANCE',                    '{{VALANCE}}',     8);
  x += 64;
  addBox(x,   r2Top,  72, r2H, 'FRAME COLOR',                '-',               8);
  x += 76;
  addBox(x,   r2Top,  88, r2H, 'DATE ORDERED / ORDER INFO',  '-',               7);

  // ── Content area placeholder label ─────────────────────────────
  var contentLabel = slide.insertTextBox(
    '[ Paste Shop Drawing / 3D Render Here ]',
    W * 0.1, H * 0.3, W * 0.8, 36
  );
  contentLabel.getText().getTextStyle()
    .setFontSize(13)
    .setForegroundColor('#CCCCCC')
    .setItalic(true);
  contentLabel.getFill().setTransparent();
  contentLabel.getBorder().setTransparent();

  presentation.saveAndClose();

  // ── Save template ID to Script Properties ──────────────────────
  var fileId = DriveApp.getFileById(presentation.getId()).getId();
  PropertiesService.getScriptProperties().setProperty('SHOP_DRAWING_TEMPLATE_ID', fileId);

  Logger.log('m_createShopDrawingTemplate_: Template created. ID = ' + fileId);
  SpreadsheetApp.getUi().alert(
    '✅ Shop Drawing Template created in your Drive root.\n\n' +
    'File name: "Shop Drawing Template"\n' +
    'ID saved to Script Properties as SHOP_DRAWING_TEMPLATE_ID.\n\n' +
    '⚠️ ACTION REQUIRED:\n' +
    'Open the file → File → Page Setup → Custom\n' +
    '→ set width to 17 in, height to 11 in (landscape)\n' +
    '   OR width 11 in, height 17 in (portrait) → Apply.'
  );
}