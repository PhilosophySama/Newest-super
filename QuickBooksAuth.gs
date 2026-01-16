/**
 * QUICKBOOKS AUTH (OAuth2 setup)
 * Version: 1/16 9am EST by Claude Sonnet 4.5
 * by Claude Opus 4.1
 *
 * Handles authentication between Google Apps Script and QuickBooks Online.
 *
 * Requires OAuth2 library:
 * Library ID: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 * 
 * Required Script Properties (set these first!):
 * - QBO_CLIENT_ID: Your QuickBooks app client ID
 * - QBO_CLIENT_SECRET: Your QuickBooks app client secret
 * - QBO_REALM_ID: Your QuickBooks company ID
 * - QBO_ENVIRONMENT: production (or sandbox)
 * - QBO_REDIRECT_URI: Get this from getScriptUrl() after deploying as web app
 */

/**
 * Get the current web app URL - MUST run this after deploying as web app
 */
function getScriptUrl() {
  const url = ScriptApp.getService().getUrl();
  Logger.log('Your Web App URL (use this as QBO_REDIRECT_URI): ' + url);
  SpreadsheetApp.getUi().alert(
    'Web App URL',
    'Copy this URL and set it as QBO_REDIRECT_URI in Script Properties:\n\n' + url + '\n\nAlso add this EXACT URL to your Intuit app Redirect URIs.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return url;
}

/**
 * Initialize OAuth2 service for QuickBooks
 */
function getOAuthService() {
  const props = PropertiesService.getScriptProperties();
  const redirectUri = props.getProperty('QBO_REDIRECT_URI');
  
  if (!redirectUri) {
    throw new Error('QBO_REDIRECT_URI not set. Deploy as web app, run getScriptUrl(), then set the property.');
  }
  
  const clientId = props.getProperty('QBO_CLIENT_ID');
  const clientSecret = props.getProperty('QBO_CLIENT_SECRET');
  
  if (!clientId || !clientSecret) {
    throw new Error('QBO_CLIENT_ID and QBO_CLIENT_SECRET must be set in Script Properties');
  }
  
  return OAuth2.createService('QuickBooks')
    .setAuthorizationBaseUrl('https://appcenter.intuit.com/connect/oauth2')
    .setTokenUrl('https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope('com.intuit.quickbooks.accounting')
    .setParam('response_type', 'code')
    .setParam('redirect_uri', redirectUri);
}

/**
 * Clear stored authentication credentials
 */
function resetAuth() {
  try {
    const service = getOAuthService();
    service.reset();
    
    // Also clear realm ID
    PropertiesService.getScriptProperties().deleteProperty('QBO_REALM_ID');
    
    SpreadsheetApp.getActive().toast("Authorization reset successfully. Please authorize again.", "✅ Reset Complete", 3);
  } catch (e) {
    SpreadsheetApp.getActive().toast("Reset error: " + e.message, "❌ Error", 5);
  }
}

/**
 * Triggered from menu → "Authorize QuickBooks"
 * Opens authorization in a modal with a clickable link
 */
function authorize() {
  try {
    const service = getOAuthService();
    
    if (service.hasAccess()) {
      SpreadsheetApp.getActive().toast("QuickBooks already authorized!", "✅ Success", 3);
      return;
    }
    
    const authorizationUrl = service.getAuthorizationUrl();
    
    // Log for debugging
    Logger.log('Authorization URL: ' + authorizationUrl);
    
    // Create HTML with instructions and a big button
    const html = HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_blank">
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              text-align: center;
            }
            .button {
              display: inline-block;
              background-color: #2CA01C;
              color: white;
              padding: 15px 30px;
              text-decoration: none;
              border-radius: 5px;
              font-size: 18px;
              margin: 20px 0;
              font-weight: bold;
            }
            .button:hover {
              background-color: #248517;
            }
            .instructions {
              color: #666;
              font-size: 14px;
              margin-top: 20px;
            }
            .note {
              background: #fff3cd;
              border: 1px solid #ffc107;
              padding: 10px;
              border-radius: 4px;
              margin-top: 15px;
              font-size: 13px;
            }
            .debug {
              background: #f0f0f0;
              border: 1px solid #ccc;
              padding: 10px;
              border-radius: 4px;
              margin-top: 15px;
              font-size: 11px;
              text-align: left;
              word-break: break-all;
            }
          </style>
        </head>
        <body>
          <h2>Connect to QuickBooks</h2>
          <p>Click the button below to authorize QuickBooks access:</p>
          
          <a href="${authorizationUrl}" class="button" target="_blank">
            Authorize QuickBooks
          </a>
          
          <div class="instructions">
            <p><strong>Steps:</strong></p>
            <ol style="text-align: left; display: inline-block;">
              <li>Click the button above (opens in new window)</li>
              <li>Sign in to QuickBooks</li>
              <li>Select your company</li>
              <li>Click "Connect"</li>
              <li>Wait for success message</li>
              <li>Close this dialog</li>
            </ol>
          </div>
          
          <div class="note">
            <strong>Note:</strong> The authorization window will redirect back to this script automatically.
            If you see an error about "undefined", check that your Redirect URI matches in both:
            <br>• Script Properties (QBO_REDIRECT_URI)
            <br>• Intuit Developer App Settings
          </div>
          
          <div class="debug">
            <strong>Debug Info:</strong><br>
            Authorization URL: ${authorizationUrl}
          </div>
        </body>
      </html>
    `).setWidth(600).setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'QuickBooks Authorization');
    
  } catch (e) {
    Logger.log('Authorization error: ' + e.toString());
    SpreadsheetApp.getUi().alert('Authorization Error', 'Error: ' + e.message + '\n\nCheck Execution log for details.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * OAuth2 callback handler
 */
function authCallback(request) {
  const service = getOAuthService();
  
  Logger.log("OAuth callback received");
  Logger.log("Request parameters: " + JSON.stringify(request.parameter));
  
  const isAuthorized = service.handleCallback(request);
  
  if (isAuthorized) {
    // Save realmId if QuickBooks provides it
    if (request.parameter.realmId) {
      PropertiesService.getScriptProperties().setProperty('QBO_REALM_ID', request.parameter.realmId);
      Logger.log("Saved realmId: " + request.parameter.realmId);
    }
    
    Logger.log("Authorization successful!");
    
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <style>
            body {
              font-family: Arial, sans-serif;
              text-align: center;
              padding: 50px;
              background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
              color: white;
            }
            .container {
              background: white;
              color: #333;
              padding: 40px;
              border-radius: 10px;
              box-shadow: 0 4px 6px rgba(0,0,0,0.1);
              display: inline-block;
            }
            h1 { color: #2CA01C; margin-bottom: 10px; }
            .checkmark {
              font-size: 60px;
              animation: scaleIn 0.5s ease-in-out;
            }
            @keyframes scaleIn {
              from { transform: scale(0); }
              to { transform: scale(1); }
            }
          </style>
          <script>
            setTimeout(function() {
              window.close();
            }, 3000);
          </script>
        </head>
        <body>
          <div class="container">
            <div class="checkmark">✅</div>
            <h1>Success!</h1>
            <p style="font-size: 18px; margin: 20px 0;">QuickBooks is now connected to your Google Sheet.</p>
            <p style="color: #666;">This window will close automatically...</p>
          </div>
        </body>
      </html>
    `);
  } else {
    const error = service.getLastError() || 'Unknown error';
    Logger.log("Authorization failed: " + error);
    
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <style>
            body {
              font-family: Arial, sans-serif;
              text-align: center;
              padding: 50px;
              background: #f44336;
              color: white;
            }
            .container {
              background: white;
              color: #333;
              padding: 40px;
              border-radius: 10px;
              box-shadow: 0 4px 6px rgba(0,0,0,0.1);
              display: inline-block;
            }
            h1 { color: #f44336; }
            pre {
              background: #f5f5f5;
              padding: 15px;
              border-radius: 5px;
              text-align: left;
              overflow-x: auto;
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>❌ Authorization Failed</h1>
            <p>Error: ${error}</p>
            <p style="color: #666; margin-top: 20px;">
              Common causes:
              <ul style="text-align: left;">
                <li>Redirect URI mismatch</li>
                <li>Invalid Client ID or Secret</li>
                <li>App not published in Intuit Developer Portal</li>
              </ul>
            </p>
            <p><a href="https://script.google.com" style="color: #2CA01C;">Open Apps Script Console for Logs</a></p>
          </div>
        </body>
      </html>
    `);
  }
}

/**
 * Web app entry point - MUST be deployed as web app
 */
function doGet(e) {
  Logger.log("doGet called with parameters: " + JSON.stringify(e.parameter));
  
  // If this is an OAuth callback
  if (e.parameter.code || e.parameter.error) {
    return authCallback(e);
  }
  
  // Otherwise show status page
  const redirectUri = PropertiesService.getScriptProperties().getProperty('QBO_REDIRECT_URI');
  const hasClientId = !!PropertiesService.getScriptProperties().getProperty('QBO_CLIENT_ID');
  const hasClientSecret = !!PropertiesService.getScriptProperties().getProperty('QBO_CLIENT_SECRET');
  
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 40px;
            max-width: 800px;
            margin: 0 auto;
          }
          h3 { color: #2CA01C; }
          pre {
            background: #f5f5f5;
            padding: 15px;
            border-radius: 5px;
            overflow-x: auto;
            border-left: 4px solid #2CA01C;
          }
          .status {
            padding: 10px;
            margin: 10px 0;
            border-radius: 4px;
          }
          .ok { background: #d4edda; border-left: 4px solid #28a745; }
          .error { background: #f8d7da; border-left: 4px solid #dc3545; }
          .warning { background: #fff3cd; border-left: 4px solid #ffc107; }
        </style>
      </head>
      <body>
        <h3>✅ OAuth Callback Endpoint Active</h3>
        <p>This endpoint is ready to receive OAuth callbacks from QuickBooks.</p>
        
        <h4>Configuration Status:</h4>
        <div class="status ${redirectUri ? 'ok' : 'error'}">
          <strong>Redirect URI:</strong> ${redirectUri || '❌ NOT SET'}
        </div>
        <div class="status ${hasClientId ? 'ok' : 'error'}">
          <strong>Client ID:</strong> ${hasClientId ? '✅ Set' : '❌ Not Set'}
        </div>
        <div class="status ${hasClientSecret ? 'ok' : 'error'}">
          <strong>Client Secret:</strong> ${hasClientSecret ? '✅ Set' : '❌ Not Set'}
        </div>
        
        <div class="status warning">
          <strong>⚠️ Important:</strong> The Redirect URI above must EXACTLY match what's configured in your Intuit Developer app settings.
        </div>
        
        <h4>Setup Steps:</h4>
        <ol>
          <li>Deploy this script as a Web App (Deploy → New Deployment → Web App)</li>
          <li>Run <code>getScriptUrl()</code> from Script Editor to get your redirect URI</li>
          <li>Set all required Script Properties (Tools → Script Properties):
            <ul>
              <li>QBO_CLIENT_ID</li>
              <li>QBO_CLIENT_SECRET</li>
              <li>QBO_REDIRECT_URI (from step 2)</li>
              <li>QBO_ENVIRONMENT (production or sandbox)</li>
            </ul>
          </li>
          <li>Add the redirect URI to your Intuit app's Redirect URIs</li>
          <li>Use menu: Setup (QuickBooks) → Authorize QuickBooks</li>
        </ol>
      </body>
    </html>
  `);
}

/**
 * Test QuickBooks connection
 */
function testQuickBooksConnection_() {
  try {
    const service = getOAuthService();
    
    if (!service.hasAccess()) {
      SpreadsheetApp.getUi().alert(
        "QuickBooks Not Authorized",
        "Please run 'Authorize QuickBooks' first.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const props = PropertiesService.getScriptProperties();
    const realmId = props.getProperty('QBO_REALM_ID');
    
    if (!realmId) {
      SpreadsheetApp.getUi().alert(
        "Missing Company ID",
        "QBO_REALM_ID not found in Script Properties. This should be set automatically during authorization.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const environment = props.getProperty('QBO_ENVIRONMENT') || 'sandbox';
    const baseUrl = environment === 'production'
      ? 'https://quickbooks.api.intuit.com'
      : 'https://sandbox-quickbooks.api.intuit.com';
    
    const url = `${baseUrl}/v3/company/${realmId}/companyinfo/${realmId}`;
    
    Logger.log('Testing connection to: ' + url);
    
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('Response code: ' + responseCode);
    Logger.log('Response: ' + responseText);
    
    if (responseCode === 200) {
      const companyInfo = JSON.parse(responseText);
      const companyName = companyInfo.CompanyInfo?.CompanyName || 'Unknown';
      
      SpreadsheetApp.getActive().toast(
        `Connected to: ${companyName}\nEnvironment: ${environment}`,
        "✅ Connection Successful",
        5
      );
    } else {
      SpreadsheetApp.getActive().toast(
        `Error ${responseCode}. Check Execution log for details.`,
        "❌ Connection Failed",
        5
      );
    }
    
  } catch (e) {
    Logger.log('Connection test error: ' + e.toString());
    SpreadsheetApp.getActive().toast(
      "Connection test failed: " + e.message,
      "❌ Error",
      5
    );
  }
}

/**
 * Helper function to show current redirect URI and setup instructions
 */
function showRedirectUri_() {
  const props = PropertiesService.getScriptProperties();
  const redirectUri = props.getProperty('QBO_REDIRECT_URI');
  const clientId = props.getProperty('QBO_CLIENT_ID');
  const clientSecret = props.getProperty('QBO_CLIENT_SECRET');
  const realmId = props.getProperty('QBO_REALM_ID');
  const environment = props.getProperty('QBO_ENVIRONMENT');
  
  let message = '=== QuickBooks OAuth Configuration ===\n\n';
  
  message += 'Redirect URI: ' + (redirectUri || '❌ NOT SET') + '\n\n';
  message += 'Client ID: ' + (clientId ? '✅ Set' : '❌ Not Set') + '\n';
  message += 'Client Secret: ' + (clientSecret ? '✅ Set' : '❌ Not Set') + '\n';
  message += 'Realm ID: ' + (realmId || 'Not yet authorized') + '\n';
  message += 'Environment: ' + (environment || 'Not Set') + '\n\n';
  
  message += '⚠️ IMPORTANT:\n';
  message += '1. Deploy this script as a Web App first\n';
  message += '2. Run getScriptUrl() to get your redirect URI\n';
  message += '3. Add that EXACT URL to Script Properties as QBO_REDIRECT_URI\n';
  message += '4. Add that EXACT URL to your Intuit app\'s Redirect URIs\n';
  message += '5. Any mismatch will cause "undefined" errors';
  
  SpreadsheetApp.getUi().alert('QuickBooks Configuration', message, SpreadsheetApp.getUi().ButtonSet.OK);
}