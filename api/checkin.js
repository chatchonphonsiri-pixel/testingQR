const { google } = require("googleapis");

const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const CHECKED_IN_STATUS = "checked_in";
const REQUIRED_HEADERS = Object.freeze({
  token: "token",
  status: "status",
  checkedInBy: "checkedInBy"
});

function normalizeHeader(value) {
  return value?.toString().trim().toLowerCase() || "";
}

function findUniqueHeaderIndex(headers, headerName) {
  const expected = normalizeHeader(headerName);
  const matches = headers
    .map((header, index) => normalizeHeader(header) === expected ? index : -1)
    .filter(index => index !== -1);

  if (matches.length !== 1) {
    throw new Error(`Expected exactly one "${headerName}" column, found ${matches.length}`);
  }

  return matches[0];
}

function getColumnIndexes(headers) {
  return Object.fromEntries(
    Object.entries(REQUIRED_HEADERS).map(([key, headerName]) => [
      key,
      findUniqueHeaderIndex(headers, headerName)
    ])
  );
}

function toColumnLetter(zeroBasedIndex) {
  let index = zeroBasedIndex + 1;
  let column = "";

  while (index > 0) {
    const remainder = (index - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    index = Math.floor((index - 1) / 26);
  }

  return column;
}

function quoteSheetTitle(title) {
  return `'${title.replace(/'/g, "''")}'`;
}

module.exports = async function handler(request, response) {
  response.setHeader("Cache-Control", "no-store");

  if (request.method !== "POST") {
    response.setHeader("Allow", "POST");
    return response.status(405).json({ error: "รองรับเฉพาะ POST" });
  }

  const token = request.body?.token?.toString().trim();
  if (!token || token.length > 500) {
    return response.status(400).json({ error: "QR token ไม่ถูกต้อง" });
  }

  const checkedInBy = request.body?.checkedInBy?.toString().trim();
  if (!checkedInBy || checkedInBy.length > 100) {
    return response.status(400).json({ error: "กรุณาระบุชื่อผู้ตรวจ" });
  }

  const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n");
  const spreadsheetId = process.env.GOOGLE_SHEET_ID;
  const sheetGid = Number(process.env.GOOGLE_SHEET_GID || "453783738");

  if (!clientEmail || !privateKey || !spreadsheetId) {
    console.error("Missing Google service account environment variables");
    return response.status(500).json({ error: "เซิร์ฟเวอร์ยังตั้งค่า Google Sheet ไม่ครบ" });
  }

  try {
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: clientEmail,
        private_key: privateKey
      },
      scopes: [SHEETS_SCOPE]
    });
    const sheets = google.sheets({ version: "v4", auth });

    const metadata = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets.properties(sheetId,title)"
    });
    const targetSheet = metadata.data.sheets?.find(
      sheet => Number(sheet.properties?.sheetId) === sheetGid
    );

    if (!targetSheet?.properties?.title) {
      return response.status(500).json({ error: "ไม่พบแท็บที่กำหนดใน Google Sheet" });
    }

    const rangePrefix = quoteSheetTitle(targetSheet.properties.title);
    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${rangePrefix}!1:1`
    });
    const headers = headerResponse.data.values?.[0] || [];
    const columnIndexes = getColumnIndexes(headers);
    const columns = Object.fromEntries(
      Object.entries(columnIndexes).map(([key, index]) => [key, toColumnLetter(index)])
    );

    const columnValuesResponse = await sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges: [
        `${rangePrefix}!${columns.token}2:${columns.token}`,
        `${rangePrefix}!${columns.status}2:${columns.status}`,
        `${rangePrefix}!${columns.checkedInBy}2:${columns.checkedInBy}`
      ]
    });
    const [tokenRows = [], statusRows = [], checkedInByRows = []] =
      (columnValuesResponse.data.valueRanges || []).map(valueRange => valueRange.values || []);
    const rowOffset = tokenRows.findIndex(
      row => row[0]?.toString().trim() === token
    );

    if (rowOffset === -1) {
      return response.status(404).json({ error: "ไม่พบข้อมูลสำหรับ QR นี้" });
    }

    const sheetRow = rowOffset + 2;
    const existingStatus = statusRows[rowOffset]?.[0]?.toString().trim();
    const existingCheckedInBy = checkedInByRows[rowOffset]?.[0]?.toString().trim() || null;
    const alreadyCheckedIn = existingStatus === CHECKED_IN_STATUS;

    if (!alreadyCheckedIn) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: "RAW",
          data: [
            {
              range: `${rangePrefix}!${columns.status}${sheetRow}`,
              values: [[CHECKED_IN_STATUS]]
            },
            {
              range: `${rangePrefix}!${columns.checkedInBy}${sheetRow}`,
              values: [[checkedInBy]]
            }
          ]
        }
      });
    }

    return response.status(200).json({
      ok: true,
      status: CHECKED_IN_STATUS,
      alreadyCheckedIn,
      checkedInBy: alreadyCheckedIn ? existingCheckedInBy : checkedInBy
    });
  } catch (error) {
    console.error("Google Sheets check-in failed:", error);
    return response.status(500).json({
      error: "เชื่อมต่อ Google Sheet ไม่สำเร็จ โปรดตรวจสอบ Service Account"
    });
  }
};
