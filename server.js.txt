require('dotenv').config();
const express = require('express');
const session = require('express-session');
const cookieParser = require('cookie-parser');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 3000;

const BASE_URL = process.env.RENDER_EXTERNAL_URL
  ? 'https://' + process.env.RENDER_EXTERNAL_URL
  : 'http://localhost:' + PORT;

/* ── Google OAuth Setup ── */
const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
  BASE_URL + '/auth/google/callback'
);

const SCOPES = [
  'openid', 'profile', 'email',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.readonly'
];

/* ── Middleware ── */
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(session({
  secret: process.env.SESSION_SECRET || 'change-this-in-production',
  resave: false,
  saveUninitialized: false,
  cookie: {
    maxAge: 7 * 24 * 60 * 60 * 1000,
    secure: BASE_URL.startsWith('https'),
    sameSite: 'lax'
  }
}));
app.use(express.static('public'));

/* ── Auth Guard ── */
function requireAuth(req, res, next) {
  if (!req.session.tokens) return res.status(401).json({ error: 'Not authenticated' });
  next();
}

/* ── Refresh token if expired ── */
async function getAuth(req) {
  oauth2Client.setCredentials(req.session.tokens);
  const expiry = new Date(req.session.tokens.expiry_date);
  const now = new Date();
  if (now >= expiry || (expiry - now) < 300000) {
    try {
      const { credentials } = await oauth2Client.refreshAccessToken();
      oauth2Client.setCredentials(credentials);
      req.session.tokens = Object.assign({}, req.session.tokens, credentials);
    } catch (e) {
      req.session.destroy();
      throw new Error('Session expired');
    }
  }
  return oauth2Client;
}

/* ── Find or create spreadsheet ── */
async function ensureSpreadsheet(auth) {
  const drive = google.drive({ version: 'v3', auth: auth });
  const sheets = google.sheets({ version: 'v4', auth: auth });
  const NAME = 'Smart Expense Tracker';

  // Search for existing spreadsheet
  const list = await drive.files.list({
    q: "name='" + NAME + "' and mimeType='application/vnd.google-apps.spreadsheet'",
    fields: 'files(id, webViewLink)',
    pageSize: 1
  });

  if (list.data.files.length > 0) {
    return { id: list.data.files[0].id, url: list.data.files[0].webViewLink, created: false };
  }

  // Create new spreadsheet
  const ss = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: NAME },
      sheets: [{ properties: { title: 'Transactions' } }]
    }
  });

  const sheetId = ss.data.sheets[0].properties.sheetId;

  // Add headers + formatting
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: ss.data.spreadsheetId,
    requestBody: {
      requests: [
        {
          updateCells: {
            range: { sheetId: sheetId },
            rows: [{
              values: [
                { userEnteredValue: { stringValue: 'Date' } },
                { userEnteredValue: { stringValue: 'Type' } },
                { userEnteredValue: { stringValue: 'Category' } },
                { userEnteredValue: { stringValue: 'Amount' } },
                { userEnteredValue: { stringValue: 'Note' } },
                { userEnteredValue: { stringValue: 'Created At' } }
              ]
            }],
            fields: 'userEnteredValue'
          }
        },
        {
          repeatCell: {
            range: { sheetId: sheetId, startRowIndex: 0, endRowIndex: 1 },
            cell: {
              userEnteredFormat: {
                textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
                backgroundColor: { red: 0.059, green: 0.463, blue: 0.431 }
              }
            },
            fields: 'userEnteredFormat.textFormat,userEnteredFormat.backgroundColor'
          }
        },
        {
          updateSheetProperties: {
            properties: { sheetId: sheetId, gridProperties: { frozenRowCount: 1 } },
            fields: 'gridProperties.frozenRowCount'
          }
        }
      ]
    }
  });

  return { id: ss.data.spreadsheetId, url: ss.data.spreadsheetUrl, created: true };
}

/* ══════════════════════════════════
   AUTH ROUTES
   ══════════════════════════════════ */

app.get('/auth/google', (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: SCOPES
  });
  res.redirect(url);
});

app.get('/auth/google/callback', async (req, res) => {
  try {
    const { tokens } = await oauth2Client.getToken(req.query.code);
    oauth2Client.setCredentials(tokens);

    const userinfo = await google.oauth2({ auth: oauth2Client }).userinfo.get();
    req.session.tokens = tokens;
    req.session.user = userinfo.data;

    const ss = await ensureSpreadsheet(oauth2Client);
    req.session.spreadsheetId = ss.id;
    req.session.spreadsheetUrl = ss.url;

    res.redirect('/');
  } catch (err) {
    console.error('OAuth callback error:', err);
    res.redirect('/?error=1');
  }
});

app.get('/auth/me', requireAuth, (req, res) => {
  res.json({ user: req.session.user, spreadsheetUrl: req.session.spreadsheetUrl });
});

app.get('/auth/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/');
});

/* ══════════════════════════════════
   API ROUTES
   ══════════════════════════════════ */

// Get all transactions
app.get('/api/transactions', requireAuth, async (req, res) => {
  try {
    const auth = await getAuth(req);
    const sheets = google.sheets({ version: 'v4', auth: auth });
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: req.session.spreadsheetId,
      range: 'Transactions!A:F'
    });
    const rows = resp.data.values || [];
    const transactions = [];
    for (let i = 1; i < rows.length; i++) {
      transactions.push({
        row: i + 1,
        date: rows[i][0], type: rows[i][1], category: rows[i][2],
        amount: Number(rows[i][3]), note: rows[i][4], createdAt: rows[i][5]
      });
    }
    res.json({ success: true, transactions: transactions.reverse(), spreadsheetUrl: req.session.spreadsheetUrl });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Add transaction
app.post('/api/transactions', requireAuth, async (req, res) => {
  try {
    const auth = await getAuth(req);
    const sheets = google.sheets({ version: 'v4', auth: auth });
    const { date, type, category, amount, note } = req.body;
    await sheets.spreadsheets.values.append({
      spreadsheetId: req.session.spreadsheetId,
      range: 'Transactions!A:F',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[date, type, category, amount, note, new Date().toISOString()]] }
    });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Update transaction
app.put('/api/transactions/:row', requireAuth, async (req, res) => {
  try {
    const auth = await getAuth(req);
    const sheets = google.sheets({ version: 'v4', auth: auth });
    const { date, type, category, amount, note } = req.body;
    await sheets.spreadsheets.values.update({
      spreadsheetId: req.session.spreadsheetId,
      range: 'Transactions!A' + req.params.row + ':E' + req.params.row,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[date, type, category, amount, note]] }
    });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Delete transaction
app.delete('/api/transactions/:row', requireAuth, async (req, res) => {
  try {
    const auth = await getAuth(req);
    const sheets = google.sheets({ version: 'v4', auth: auth });
    const row = parseInt(req.params.row);
    const meta = await sheets.spreadsheets.get({ spreadsheetId: req.session.spreadsheetId });
    const sheetId = meta.data.sheets[0].properties.sheetId;
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: req.session.spreadsheetId,
      requestBody: { requests: [{ deleteDimension: { range: { sheetId: sheetId, dimension: 'ROWS', startIndex: row - 1, endIndex: row } } }] }
    });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Clear all data
app.delete('/api/transactions', requireAuth, async (req, res) => {
  try {
    const auth = await getAuth(req);
    const sheets = google.sheets({ version: 'v4', auth: auth });
    const resp = await sheets.spreadsheets.values.get({ spreadsheetId: req.session.spreadsheetId, range: 'Transactions!A:F' });
    const lastRow = (resp.data.values || []).length;
    if (lastRow > 1) {
      const meta = await sheets.spreadsheets.get({ spreadsheetId: req.session.spreadsheetId });
      const sheetId = meta.data.sheets[0].properties.sheetId;
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: req.session.spreadsheetId,
        requestBody: { requests: [{ deleteDimension: { range: { sheetId: sheetId, dimension: 'ROWS', startIndex: 1, endIndex: lastRow } } }] }
      });
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Get spreadsheet URL
app.get('/api/sheet-url', requireAuth, (req, res) => {
  res.json({ url: req.session.spreadsheetUrl });
});

// SPA fallback
app.get('*', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

app.listen(PORT, () => console.log('Server running on port ' + PORT));