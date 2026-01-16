/**
 * QBO ESTIMATES - QuickBooks Online Integration
 * Version: 1/16 9am EST by Claude Sonnet 4.5
 * by Claude Opus 4.1
 *
 * Functions:
 * - Send estimates from Leads sheet to QuickBooks (with approval dialog)
 * - Convert estimates to invoices when awarded
 * - Debug and configuration helpers
 */

/**
 * Convert estimate to invoice for awarded jobs
 * Run this from Awarded sheet with active cell on the row to convert
 */
function convertEstimateToInvoice() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Awarded");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("No 'Awarded' sheet found");
    return;
  }
  
  const row = sheet.getActiveCell().getRow();
  
  // Skip header row
  if (row === 1) {
    SpreadsheetApp.getUi().alert("Please select a data row, not the header.");
    return;
  }
  
  // Get estimate link from column P
  const estimateCell = sheet.getRange(row, 16); // Column P
  const estimateCellValue = estimateCell.getRichTextValue();
  const estimateUrl = estimateCellValue ? estimateCellValue.getLinkUrl() : "";
  
  if (!estimateUrl || !estimateUrl.includes("txnId=")) {
    SpreadsheetApp.getUi().alert("No valid estimate link found in column P");
    return;
  }
  
  // Extract estimate ID from URL
  const estimateId = estimateUrl.split("txnId=")[1].split("&")[0];
  
  // Check OAuth
  const service = getOAuthService();
  if (!service.hasAccess()) {
    SpreadsheetApp.getUi().alert("QuickBooks not authorized. Run 'Authorize QuickBooks' first.");
    return;
  }
  
  // Config
  const props = PropertiesService.getScriptProperties();
  const realmId = props.getProperty('QBO_REALM_ID');
  const environment = props.getProperty('QBO_ENVIRONMENT') || 'sandbox';
  const baseUrl = environment === 'production'
    ? `https://quickbooks.api.intuit.com/v3/company/${realmId}`
    : `https://sandbox-quickbooks.api.intuit.com/v3/company/${realmId}`;
  
  try {
    // First, get the estimate details
    let resp = UrlFetchApp.fetch(
      `${baseUrl}/estimate/${estimateId}`,
      { 
        headers: { 
          Authorization: "Bearer " + service.getAccessToken(), 
          Accept: "application/json" 
        }, 
        muteHttpExceptions: true 
      }
    );
    
    if (resp.getResponseCode() !== 200) {
      throw new Error("Failed to retrieve estimate: " + resp.getContentText());
    }
    
    const estimate = JSON.parse(resp.getContentText()).Estimate;
    
    // Build invoice payload from estimate
    const invoicePayload = {
      CustomerRef: estimate.CustomerRef,
      BillEmail: estimate.BillEmail,
      BillAddr: estimate.BillAddr,
      ShipAddr: estimate.ShipAddr,
      Line: estimate.Line, // Copy all line items
      GlobalTaxCalculation: "TaxExcluded",
      // Reference the original estimate
      LinkedTxn: [{
        TxnId: estimateId,
        TxnType: "Estimate"
      }]
    };
    
    // Create invoice
    resp = UrlFetchApp.fetch(baseUrl + "/invoice", {
      method: "post",
      headers: { 
        Authorization: "Bearer " + service.getAccessToken(), 
        Accept: "application/json", 
        "Content-Type": "application/json" 
      },
      payload: JSON.stringify(invoicePayload),
      muteHttpExceptions: true
    });
    
    if (resp.getResponseCode() !== 200) {
      const error = JSON.parse(resp.getContentText());
      throw new Error(`Failed to create invoice: ${error.Fault?.Error?.[0]?.Message || 'Unknown error'}`);
    }
    
    const invoice = JSON.parse(resp.getContentText()).Invoice;
    const invoiceId = invoice.Id;
    const invoiceNumber = invoice.DocNumber;
    const totalAmount = invoice.TotalAmt;
    
    // Try to send invoice to get share link
    let shareUrl = "";
    try {
      const customerEmail = invoice.BillEmail?.Address || sheet.getRange(row, 9).getValue(); // Column I for email
      if (customerEmail) {
        const sendResp = UrlFetchApp.fetch(`${baseUrl}/invoice/${invoiceId}/send`, {
          method: "post",
          headers: { 
            Authorization: "Bearer " + service.getAccessToken(), 
            Accept: "application/json", 
            "Content-Type": "application/json" 
          },
          payload: JSON.stringify({ SendTo: customerEmail }),
          muteHttpExceptions: true
        });
        
        if (sendResp.getResponseCode() === 200) {
          const sendResult = JSON.parse(sendResp.getContentText());
          shareUrl = sendResult?.Invoice?.InvoiceLink || 
                    sendResult?.Invoice?.PublicLink || 
                    "";
        }
      }
    } catch (e) {
      Logger.log("Could not send invoice: " + e);
    }
    
    // If no share link, use internal link
    if (!shareUrl) {
      shareUrl = environment === 'production'
        ? `https://qbo.intuit.com/app/invoice?txnId=${invoiceId}`
        : `https://sandbox.qbo.intuit.com/app/invoice?txnId=${invoiceId}`;
    }
    
    // Format amount as currency
    const formattedAmount = "$" + totalAmount.toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
    
    // Put invoice link in column N with amount as display text
    sheet.getRange(row, 14).setRichTextValue(
      SpreadsheetApp.newRichTextValue()
        .setText(formattedAmount)
        .setLinkUrl(shareUrl)
        .build()
    );
    
    SpreadsheetApp.getActive().toast(
      `✅ Invoice #${invoiceNumber} created for ${formattedAmount}`,
      "Success",
      5
    );
    
  } catch (err) {
    Logger.log("Error converting estimate to invoice: " + err.toString());
    sheet.getRange(row, 14).setValue("ERROR: " + err.message);
    SpreadsheetApp.getActive().toast("❌ Failed: " + err.message, "Error", 10);
  }
}

/**
 * Send estimate for current row in Leads sheet to QuickBooks
 * Shows approval dialog with customer info before creating
 */
function sendEstimateCurrentRow_() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (sheetName !== 'Leads') {
    ui.alert('Please run this from the Leads sheet');
    return;
  }
  
  const row = sheet.getActiveCell().getRow();
  
  if (row === 1) {
    ui.alert('Please select a data row, not the header');
    return;
  }
  
  // Check OAuth first
  const service = getOAuthService();
  if (!service.hasAccess()) {
    ui.alert('QuickBooks not authorized. Run "Authorize QuickBooks" first.');
    return;
  }
  
  // Get customer data from row for approval display
  const customerName = sheet.getRange(row, 5).getDisplayValue(); // Column E
  const displayName = sheet.getRange(row, 6).getDisplayValue(); // Column F
  const customerEmail = sheet.getRange(row, 9).getValue(); // Column I
  const address = sheet.getRange(row, 10).getValue(); // Column J
  const jobDescription = sheet.getRange(row, 11).getValue(); // Column K
  const quotePrice = parseFloat(sheet.getRange(row, 14).getValue()) || 0; // Column N
  const jobType = sheet.getRange(row, 18).getDisplayValue(); // Column R
  
  // Validate required fields before showing approval
  if (!customerName) {
    ui.alert('❌ Missing Data', 'Customer name (Column E) is required.', ui.ButtonSet.OK);
    return;
  }
  
  if (quotePrice <= 0) {
    ui.alert('❌ Missing Data', 'Quote price (Column N) must be greater than 0.', ui.ButtonSet.OK);
    return;
  }
  
  // Check if estimate already exists in column P
  const existingEstimate = sheet.getRange(row, 16).getRichTextValue();
  const existingUrl = existingEstimate ? existingEstimate.getLinkUrl() : '';
  if (existingUrl && existingUrl.includes('qbo.intuit.com')) {
    const overwriteResponse = ui.alert(
      '⚠️ Estimate Already Exists',
      'This row already has a QuickBooks estimate link in column P.\n\n' +
      'Do you want to create a NEW estimate anyway?\n' +
      '(This will overwrite the link in column P)',
      ui.ButtonSet.YES_NO
    );
    if (overwriteResponse !== ui.Button.YES) {
      SpreadsheetApp.getActive().toast('Estimate creation cancelled.', 'Cancelled', 3);
      return;
    }
  }
  
  // Format price for display
  const formattedPrice = '$' + quotePrice.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  
  // Build approval message with all customer info
  const approvalMessage = 
    '═══════════════════════════════════\n' +
    '       QUICKBOOKS ESTIMATE PREVIEW\n' +
    '═══════════════════════════════════\n\n' +
    '📋 CUSTOMER INFORMATION:\n' +
    '─────────────────────────────────────\n' +
    '  Name:          ' + (customerName || '(not provided)') + '\n' +
    '  Display Name:  ' + (displayName || '(not provided)') + '\n' +
    '  Email:         ' + (customerEmail || '(not provided)') + '\n' +
    '  Address:       ' + (address || '(not provided)') + '\n\n' +
    '📝 JOB DETAILS:\n' +
    '─────────────────────────────────────\n' +
    '  Job Type:      ' + (jobType || '(not provided)') + '\n' +
    '  Description:   ' + (jobDescription || '(not provided)') + '\n\n' +
    '💰 QUOTE AMOUNT:\n' +
    '─────────────────────────────────────\n' +
    '  Price:         ' + formattedPrice + '\n\n' +
    '═══════════════════════════════════\n' +
    'Row: ' + row + ' | Sheet: ' + sheetName + '\n' +
    '═══════════════════════════════════\n\n' +
    'Do you want to CREATE this estimate in QuickBooks?';
  
  // Show approval dialog
  const response = ui.alert(
    '🔍 Confirm Estimate Creation',
    approvalMessage,
    ui.ButtonSet.YES_NO
  );
  
  // Only proceed if user clicks YES
  if (response !== ui.Button.YES) {
    SpreadsheetApp.getActive().toast('Estimate creation cancelled by user.', 'Cancelled', 3);
    return;
  }
  
  // User approved - proceed with estimate creation
  try {
    // Get QuickBooks configuration
    const props = PropertiesService.getScriptProperties();
    const realmId = props.getProperty('QBO_REALM_ID');
    const environment = props.getProperty('QBO_ENVIRONMENT') || 'sandbox';
    const baseUrl = environment === 'production'
      ? `https://quickbooks.api.intuit.com/v3/company/${realmId}`
      : `https://sandbox-quickbooks.api.intuit.com/v3/company/${realmId}`;
    
    // Show progress
    SpreadsheetApp.getActive().toast('Creating estimate in QuickBooks...', '⏳ Please wait', 30);
    
    // First, find or create customer in QuickBooks
    const customerId = findOrCreateCustomer_(service, baseUrl, customerName, customerEmail, address);
    
    // Create estimate
    const estimatePayload = {
      CustomerRef: {
        value: customerId
      },
      Line: [{
        DetailType: "SalesItemLineDetail",
        Amount: quotePrice,
        Description: jobDescription || "Awning installation",
        SalesItemLineDetail: {
          Qty: 1,
          UnitPrice: quotePrice,
          ItemRef: {
            value: "1", // Use default service item - adjust as needed
            name: "Services"
          }
        }
      }],
      GlobalTaxCalculation: "TaxExcluded"
    };
    
    if (customerEmail) {
      estimatePayload.BillEmail = { Address: customerEmail };
    }
    
    const resp = UrlFetchApp.fetch(`${baseUrl}/estimate`, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(estimatePayload),
      muteHttpExceptions: true
    });
    
    if (resp.getResponseCode() !== 200) {
      const error = JSON.parse(resp.getContentText());
      throw new Error(`QuickBooks API error: ${error.Fault?.Error?.[0]?.Message || 'Unknown error'}`);
    }
    
    const estimate = JSON.parse(resp.getContentText()).Estimate;
    const estimateId = estimate.Id;
    const estimateNumber = estimate.DocNumber;
    
    // Build estimate URL
    const estimateUrl = environment === 'production'
      ? `https://qbo.intuit.com/app/estimate?txnId=${estimateId}`
      : `https://sandbox.qbo.intuit.com/app/estimate?txnId=${estimateId}`;
    
    // Write estimate link to column P as hyperlink with "QB" as display text
    sheet.getRange(row, 16).setRichTextValue(
      SpreadsheetApp.newRichTextValue()
        .setText('QB')
        .setLinkUrl(estimateUrl)
        .build()
    );
    
    // Show success with estimate details
    ui.alert(
      '✅ Estimate Created Successfully',
      'Estimate #' + estimateNumber + ' has been created in QuickBooks!\n\n' +
      'Customer: ' + customerName + '\n' +
      'Amount: ' + formattedPrice + '\n\n' +
      'The QB link has been added to column P (row ' + row + ').',
      ui.ButtonSet.OK
    );
    
  } catch (err) {
    Logger.log('Error sending estimate: ' + err.toString());
    ui.alert('❌ Error Creating Estimate', 'Error: ' + err.message, ui.ButtonSet.OK);
  }
}

/**
 * Helper: Find existing customer or create new one in QuickBooks
 */
function findOrCreateCustomer_(service, baseUrl, customerName, email, address) {
  // Search for existing customer
  const searchQuery = `SELECT * FROM Customer WHERE DisplayName = '${customerName.replace(/'/g, "\\'")}'`;
  const searchUrl = `${baseUrl}/query?query=${encodeURIComponent(searchQuery)}`;
  
  let resp = UrlFetchApp.fetch(searchUrl, {
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  });
  
  if (resp.getResponseCode() === 200) {
    const result = JSON.parse(resp.getContentText());
    if (result.QueryResponse?.Customer?.length > 0) {
      return result.QueryResponse.Customer[0].Id;
    }
  }
  
  // Customer not found, create new one
  const customerPayload = {
    DisplayName: customerName,
    PrimaryEmailAddr: email ? { Address: email } : undefined,
    BillAddr: address ? parseAddress_(address) : undefined
  };
  
  resp = UrlFetchApp.fetch(`${baseUrl}/customer`, {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(customerPayload),
    muteHttpExceptions: true
  });
  
  if (resp.getResponseCode() !== 200) {
    throw new Error('Failed to create customer: ' + resp.getContentText());
  }
  
  const customer = JSON.parse(resp.getContentText()).Customer;
  return customer.Id;
}

/**
 * Helper: Parse address string into QuickBooks address object
 */
function parseAddress_(addressString) {
  if (!addressString) return undefined;
  
  // Simple address parsing - you may want to enhance this
  const lines = addressString.split(',').map(s => s.trim());
  
  return {
    Line1: lines[0] || '',
    City: lines[1] || '',
    CountrySubDivisionCode: lines[2] || 'FL',
    PostalCode: lines[3] || '',
    Country: 'USA'
  };
}

/**
 * Configure QuickBooks environment (production vs sandbox)
 */
function configureEnvironment() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Configure QuickBooks Environment',
    'Which environment are you connecting to?\n\n' +
    'Choose YES for Production (live data)\n' +
    'Choose NO for Sandbox (test data)',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response === ui.Button.YES) {
    PropertiesService.getScriptProperties().setProperty('QBO_ENVIRONMENT', 'production');
    ui.alert('Environment set to PRODUCTION');
  } else if (response === ui.Button.NO) {
    PropertiesService.getScriptProperties().setProperty('QBO_ENVIRONMENT', 'sandbox');
    ui.alert('Environment set to SANDBOX');
  }
}

/**
 * Debug: List all QuickBooks items (for line item configuration)
 */
function listQuickBooksItems() {
  const service = getOAuthService();
  
  if (!service.hasAccess()) {
    SpreadsheetApp.getUi().alert('QuickBooks not authorized');
    return;
  }
  
  const props = PropertiesService.getScriptProperties();
  const realmId = props.getProperty('QBO_REALM_ID');
  const environment = props.getProperty('QBO_ENVIRONMENT') || 'sandbox';
  const baseUrl = environment === 'production'
    ? `https://quickbooks.api.intuit.com/v3/company/${realmId}`
    : `https://sandbox-quickbooks.api.intuit.com/v3/company/${realmId}`;
  
  try {
    const resp = UrlFetchApp.fetch(`${baseUrl}/query?query=SELECT * FROM Item`, {
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (resp.getResponseCode() === 200) {
      const result = JSON.parse(resp.getContentText());
      const items = result.QueryResponse?.Item || [];
      
      let message = 'QuickBooks Items:\n\n';
      items.forEach(item => {
        message += `ID: ${item.Id} - Name: ${item.Name} (${item.Type})\n`;
      });
      
      Logger.log(message);
      SpreadsheetApp.getUi().alert('Items Listed', 
        `Found ${items.length} items. See Execution log for details.`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      throw new Error('Failed to list items: ' + resp.getContentText());
    }
  } catch (e) {
    Logger.log('Error listing items: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}