/**
 * ============================================================
 * FILE: qbo-api-core.gs
 * ============================================================
 * This is the foundation of the QBO integration.
 * It handles everything needed to securely connect to the
 * QuickBooks Online API and make authenticated requests.
 *
 * What lives here:
 *  - Reading credentials safely from Script Properties
 *  - OAuth2 setup (authorization, token exchange, token refresh)
 *  - A reusable fetch helper used by all report functions
 *  - Menu setup when the spreadsheet opens
 *
 * What does NOT live here:
 *  - Any report or data logic → see qbo-reports.gs
 *
 * Requires:
 *  - apps-script-oauth2 library (ID: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF)
 *  - Script Properties set via Extensions > Apps Script > Project Settings:
 *      QBO_CLIENT_ID     → from your Intuit Developer app
 *      QBO_CLIENT_SECRET → from your Intuit Developer app
 * ============================================================
 */


// ── CONFIGURATION ─────────────────────────────────────────────────────────────

/**
 * Read credentials from Script Properties at runtime.
 * Never hardcode credentials in source code — this keeps secrets out of
 * version control and allows credentials to be rotated without code changes.
 */
const _props = PropertiesService.getScriptProperties();

const QBO_CONFIG = {
  clientId:     _props.getProperty("QBO_CLIENT_ID"),
  clientSecret: _props.getProperty("QBO_CLIENT_SECRET"),
  scopes:       ["com.intuit.quickbooks.accounting"],

  // Switch between "sandbox" (for testing) and "production" (live data).
  // The base API URL changes depending on this value — see qboApiBaseUrl_().
  environment: "production",
};


// ── SPREADSHEET MENU ──────────────────────────────────────────────────────────

/**
 * Runs automatically when the spreadsheet is opened.
 * Adds the "QBO Sync" menu so users can trigger actions without
 * needing to open the Apps Script editor.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("QBO Sync")
    .addItem("1) Show Redirect URI", "showRedirectUri")
    .addItem("2) Connect to QuickBooks", "connectToQuickBooks")
    .addSeparator()
    .addItem("Check connection", "checkConnection")
    .addItem("Disconnect", "disconnect")
    .addSeparator()
    // ← Add menu items for your reports here (defined in qbo-reports.gs)
    .addItem("Pull Customers", "pullCustomersToSheet")
    .addItem("Pull P&L by Customer", "pullPnLByCustomer")
    .addItem("Pull Open POs & Bills", "pullOpenPOsAndBills")
    .addToUi();
}


// ── OAUTH2 SETUP ──────────────────────────────────────────────────────────────

/**
 * Builds and returns the OAuth2 service object.
 * This is called internally during authorization and token refresh.
 *
 * The trailing underscore (_) is a naming convention in Apps Script
 * that marks a function as private — it cannot be called from triggers
 * or the Run menu, only from other functions in this project.
 */
function getQboService_() {
  return OAuth2.createService("qbo")
    // Intuit's authorization page — where the user approves access
    .setAuthorizationBaseUrl("https://appcenter.intuit.com/connect/oauth2")
    // Intuit's token endpoint — where we exchange codes for tokens
    .setTokenUrl("https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer")
    .setClientId(QBO_CONFIG.clientId)
    .setClientSecret(QBO_CONFIG.clientSecret)
    // The function that Intuit will redirect back to after the user approves
    .setCallbackFunction("authCallback")
    // Store tokens per-user so multiple users can connect independently
    .setPropertyStore(PropertiesService.getUserProperties())
    .setCache(CacheService.getUserCache())
    .setLock(LockService.getUserLock())
    // Intuit requires Basic auth (base64 of clientId:clientSecret) on the token endpoint
    .setTokenHeaders({
      Authorization:
        "Basic " +
        Utilities.base64Encode(
          QBO_CONFIG.clientId + ":" + QBO_CONFIG.clientSecret
        ),
      Accept: "application/json",
    })
    .setParam("response_type", "code")
    .setParam("scope", QBO_CONFIG.scopes.join(" "));
}

/**
 * Step 1 of connecting: Show the Redirect URI.
 * This URI must be registered in your Intuit Developer app settings
 * before authorization will work. Run this once when setting up.
 */
function showRedirectUri() {
  const service = getQboService_();
  const uri = service.getRedirectUri();
  SpreadsheetApp.getUi().alert(
    "Copy this Redirect URI into your Intuit Developer app settings:\n\n" + uri
  );
}

/**
 * Step 2 of connecting: Generate and show the authorization URL.
 * The user opens this URL, logs into QuickBooks, and approves access.
 * Intuit then redirects back to authCallback() with a temporary code.
 */
function connectToQuickBooks() {
  const service = getQboService_();
  const url = service.getAuthorizationUrl();
  SpreadsheetApp.getUi().alert(
    "Open this URL in your browser to authorize QuickBooks:\n\n" + url
  );
}


// ── OAUTH2 CALLBACK & TOKEN EXCHANGE ──────────────────────────────────────────

/**
 * Intuit redirects here after the user approves access.
 * The URL contains ?code=...&realmId=...
 *
 * This function:
 *  1. Extracts the authorization code and company ID (realmId)
 *  2. Exchanges the code for access + refresh tokens
 *  3. Saves all tokens to User Properties for later use
 *
 * @param {Object} request - The redirect request object from Intuit
 */
function authCallback(request) {
  try {
    const code    = request?.parameter?.code;
    const realmId = request?.parameter?.realmId;

    if (!code) {
      return HtmlService.createHtmlOutput(
        "No authorization code returned. Did you cancel?"
      );
    }

    // Save the realmId — this is the unique ID of the connected QBO company.
    // Every API call needs it in the URL path.
    if (realmId) {
      PropertiesService.getUserProperties().setProperty("QBO_REALMID", realmId);
    }

    // Exchange the temporary authorization code for long-lived tokens
    const token = exchangeAuthCodeForToken_(code);

    // Persist tokens so we don't need to re-authorize on every call
    const props = PropertiesService.getUserProperties();
    props.setProperty("QBO_ACCESS_TOKEN",  token.access_token);
    props.setProperty("QBO_REFRESH_TOKEN", token.refresh_token);
    // Store expiry as an absolute timestamp (ms) so we can check it later
    props.setProperty(
      "QBO_EXPIRES_AT",
      String(Date.now() + token.expires_in * 1000)
    );

    return HtmlService.createHtmlOutput(
      "✅ Connected to QuickBooks. You can close this tab."
    );
  } catch (err) {
    return HtmlService.createHtmlOutput(
      "Auth error: " + (err?.message || err)
    );
  }
}

/**
 * Exchanges a temporary authorization code for OAuth2 tokens.
 * Called once during initial setup from authCallback().
 *
 * @param {string} code - The authorization code from Intuit's redirect
 * @returns {Object} Parsed token response containing access_token, refresh_token, expires_in
 */
function exchangeAuthCodeForToken_(code) {
  const tokenUrl     = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer";
  const redirectUri  = getQboService_().getRedirectUri();
  const basic        = Utilities.base64Encode(
    QBO_CONFIG.clientId + ":" + QBO_CONFIG.clientSecret
  );

  const res = UrlFetchApp.fetch(tokenUrl, {
    method: "post",
    muteHttpExceptions: true,
    headers: {
      Authorization: "Basic " + basic,
      Accept: "application/json",
    },
    // Apps Script automatically form-encodes plain objects passed as payload
    payload: {
      grant_type:   "authorization_code",
      code:         code,
      redirect_uri: redirectUri,
    },
  });

  const status = res.getResponseCode();
  const body   = res.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(`Token exchange failed (${status}): ${body}`);
  }

  return JSON.parse(body);
}


// ── TOKEN MANAGEMENT ──────────────────────────────────────────────────────────

/**
 * Returns a valid access token, automatically refreshing it if it's
 * close to expiring (within 2 minutes).
 *
 * QBO access tokens expire after 1 hour. Instead of forcing the user
 * to reconnect every hour, we use the refresh token to silently get
 * a new access token in the background.
 *
 * @returns {string} A valid OAuth2 access token
 */
function getValidAccessToken_() {
  const props     = PropertiesService.getUserProperties();
  const access    = props.getProperty("QBO_ACCESS_TOKEN");
  const refresh   = props.getProperty("QBO_REFRESH_TOKEN");
  const expiresAt = Number(props.getProperty("QBO_EXPIRES_AT") || "0");

  if (!access || !refresh || !expiresAt) {
    throw new Error("Not connected. Run 'Connect to QuickBooks' from the menu.");
  }

  // If token is still valid for more than 2 minutes, use it as-is
  const TWO_MINUTES_MS = 2 * 60 * 1000;
  if (Date.now() < expiresAt - TWO_MINUTES_MS) return access;

  // Otherwise, refresh and save the new tokens
  const token = refreshAccessToken_(refresh);
  props.setProperty("QBO_ACCESS_TOKEN",  token.access_token);
  // Intuit may or may not return a new refresh token — keep the old one if not
  props.setProperty("QBO_REFRESH_TOKEN", token.refresh_token || refresh);
  props.setProperty(
    "QBO_EXPIRES_AT",
    String(Date.now() + token.expires_in * 1000)
  );

  return token.access_token;
}

/**
 * Uses the refresh token to get a new access token from Intuit.
 *
 * @param {string} refreshToken - The stored refresh token
 * @returns {Object} New token response with access_token, refresh_token, expires_in
 */
function refreshAccessToken_(refreshToken) {
  const tokenUrl = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer";
  const basic    = Utilities.base64Encode(
    QBO_CONFIG.clientId + ":" + QBO_CONFIG.clientSecret
  );

  const res = UrlFetchApp.fetch(tokenUrl, {
    method: "post",
    muteHttpExceptions: true,
    headers: {
      Authorization: "Basic " + basic,
      Accept: "application/json",
    },
    payload: {
      grant_type:    "refresh_token",
      refresh_token: refreshToken,
    },
  });

  const status = res.getResponseCode();
  const body   = res.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(`Token refresh failed (${status}): ${body}`);
  }

  return JSON.parse(body);
}


// ── CONNECTION UTILITIES ───────────────────────────────────────────────────────

/**
 * Shows a summary of the current connection state.
 * Useful for debugging — confirms that tokens and realmId are stored correctly.
 */
function checkConnection() {
  const props     = PropertiesService.getUserProperties();
  const access    = props.getProperty("QBO_ACCESS_TOKEN");
  const refresh   = props.getProperty("QBO_REFRESH_TOKEN");
  const expiresAt = Number(props.getProperty("QBO_EXPIRES_AT") || "0");
  const realmId   = props.getProperty("QBO_REALMID");

  SpreadsheetApp.getUi().alert(
    `Access token saved:  ${!!access}\n` +
    `Refresh token saved: ${!!refresh}\n` +
    `Expires at:          ${expiresAt ? new Date(expiresAt).toISOString() : "(none)"}\n` +
    `Company ID (realmId): ${realmId || "(none)"}`
  );
}

/**
 * Removes all stored tokens and resets the connection.
 * The user will need to re-authorize via 'Connect to QuickBooks'.
 */
function disconnect() {
  const props = PropertiesService.getUserProperties();
  ["QBO_ACCESS_TOKEN", "QBO_REFRESH_TOKEN", "QBO_EXPIRES_AT", "QBO_REALMID"]
    .forEach(key => props.deleteProperty(key));
  SpreadsheetApp.getUi().alert("Disconnected from QuickBooks.");
}


// ── BASE API HELPER ────────────────────────────────────────────────────────────

/**
 * Returns the correct API base URL based on the configured environment.
 * Sandbox and production point to different domains — using the wrong one
 * causes 403 errors even with valid credentials.
 *
 * @returns {string} Base URL for QBO API requests
 */
function qboApiBaseUrl_() {
  return QBO_CONFIG.environment === "sandbox"
    ? "https://sandbox-quickbooks.api.intuit.com"
    : "https://quickbooks.api.intuit.com";
}

/**
 * Returns the stored company ID (realmId).
 * Every QBO API URL includes the realmId to identify which company to query.
 *
 * @returns {string} The QBO company ID
 */
function getRealmId_() {
  const rid = PropertiesService.getUserProperties().getProperty("QBO_REALMID");
  if (!rid) throw new Error("Missing company ID. Reconnect via 'Connect to QuickBooks'.");
  return rid;
}

/**
 * The core API fetch helper. All report functions call this instead of
 * building their own HTTP requests.
 *
 * It automatically:
 *  - Builds the full QBO API URL (base + company ID + path)
 *  - Appends query parameters if provided
 *  - Attaches a valid Bearer token (refreshing if needed)
 *  - Throws a clear error if the request fails
 *
 * @param {string} path        - API path after /v3/company/{realmId}/ (e.g. "query", "reports/ProfitAndLoss")
 * @param {Object} queryParams - Optional key/value pairs to append as URL query string
 * @returns {Object} Parsed JSON response from QBO
 */
function qboFetch_(path, queryParams) {
  const realmId     = getRealmId_();
  const accessToken = getValidAccessToken_();

  // Build the full request URL
  let url = `${qboApiBaseUrl_()}/v3/company/${realmId}/${path}`;

  // Append query parameters if any were provided (e.g. ?query=SELECT...&minorversion=65)
  if (queryParams) {
    const qs = Object.keys(queryParams)
      .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(queryParams[k])}`)
      .join("&");
    url += "?" + qs;
  }

  const res = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true, // Prevents Apps Script from throwing on non-2xx — we handle errors below
    headers: {
      Authorization: "Bearer " + accessToken,
      Accept: "application/json",
    },
  });

  const status = res.getResponseCode();
  const body   = res.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(`QBO API error (${status}): ${body}`);
  }

  return JSON.parse(body);
}
