const { google } = require("googleapis");

const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const CHECKED_IN_STATUS = "checked_in";
const MILLISECONDS_PER_DAY = 86_400_000;
const GOOGLE_SHEETS_EPOCH_OFFSET = 25_569;
const THAILAND_UTC_OFFSET_MS = 7 * 60 * 60 * 1000;
const REQUIRED_HEADERS = Object.freeze({
  token: "token",
  status: "status",
  checkedInAt: "checkedInAt",
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

function toThailandDateSerial(date) {
  return (
    (date.getTime() + THAILAND_UTC_OFFSET_MS) / MILLISECONDS_PER_DAY
    + GOOGLE_SHEETS_EPOCH_OFFSET
  );
}

function createCellUpdate({ sheetId, rowIndex, columnIndex, value, numberFormat }) {
  const userEnteredValue = typeof value === "number"
    ? { numberValue: value }
    : { stringValue: value };
  const cell = { userEnteredValue };

  if (numberFormat) {
    cell.userEnteredFormat = { numberFormat };
  }

  return {
    updateCells: {
      range: {
        sheetId,
        startRowIndex: rowIndex,
        endRowIndex: rowIndex + 1,
        startColumnIndex: columnIndex,
        endColumnIndex: columnIndex + 1
      },
      rows: [{ values: [cell] }],
      fields: numberFormat
        ? "userEnteredValue,userEnteredFormat.numberFormat"
        : "userEnteredValue"
    }
  };
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
        `${rangePrefix}!${columns.checkedInAt}2:${columns.checkedInAt}`,
        `${rangePrefix}!${columns.checkedInBy}2:${columns.checkedInBy}`
      ]
    });
    const [tokenRows = [], statusRows = [], checkedInAtRows = [], checkedInByRows = []] =
      (columnValuesResponse.data.valueRanges || []).map(valueRange => valueRange.values || []);
    const rowOffset = tokenRows.findIndex(
      row => row[0]?.toString().trim() === token
    );

    if (rowOffset === -1) {
      return response.status(404).json({ error: "ไม่พบข้อมูลสำหรับ QR นี้" });
    }

    const sheetRow = rowOffset + 2;
    const existingStatus = statusRows[rowOffset]?.[0]?.toString().trim();
    const existingCheckedInAt = checkedInAtRows[rowOffset]?.[0]?.toString().trim() || null;
    const existingCheckedInBy = checkedInByRows[rowOffset]?.[0]?.toString().trim() || null;
    const alreadyCheckedIn = existingStatus === CHECKED_IN_STATUS;
    const checkedInDate = new Date();
    const checkedInAt = checkedInDate.toISOString();

    if (!alreadyCheckedIn) {
      const zeroBasedRowIndex = sheetRow - 1;
      const sheetId = targetSheet.properties.sheetId;

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            createCellUpdate({
              sheetId,
              rowIndex: zeroBasedRowIndex,
              columnIndex: columnIndexes.status,
              value: CHECKED_IN_STATUS
            }),
            createCellUpdate({
              sheetId,
              rowIndex: zeroBasedRowIndex,
              columnIndex: columnIndexes.checkedInAt,
              value: toThailandDateSerial(checkedInDate),
              numberFormat: {
                type: "DATE_TIME",
                pattern: "dd/mm/yyyy hh:mm:ss"
              }
            }),
            createCellUpdate({
              sheetId,
              rowIndex: zeroBasedRowIndex,
              columnIndex: columnIndexes.checkedInBy,
              value: checkedInBy
            })
          ]
        }
      });
    }

    return response.status(200).json({
      ok: true,
      status: CHECKED_IN_STATUS,
      alreadyCheckedIn,
      checkedInAt: alreadyCheckedIn ? existingCheckedInAt : checkedInAt,
      checkedInBy: alreadyCheckedIn ? existingCheckedInBy : checkedInBy
    });
  } catch (error) {
    console.error("Google Sheets check-in failed:", error);
    return response.status(500).json({
      error: "เชื่อมต่อ Google Sheet ไม่สำเร็จ โปรดตรวจสอบ Service Account"
    });
  }
};
