const { google } = require("googleapis");

const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";

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

    const escapedTitle = targetSheet.properties.title.replace(/'/g, "''");
    const rangePrefix = `'${escapedTitle}'`;
    const rowsResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${rangePrefix}!A:O`
    });
    const rows = rowsResponse.data.values || [];
    const rowIndex = rows.findIndex(
      (row, index) => index > 0 && row[7]?.toString().trim() === token
    );

    if (rowIndex === -1) {
      return response.status(404).json({ error: "ไม่พบข้อมูลสำหรับ QR นี้" });
    }

    const alreadyCheckedIn = rows[rowIndex][5]?.toString().trim() === "checked_in";

    if (!alreadyCheckedIn) {
      const updates = [
        {
          range: `${rangePrefix}!F${rowIndex + 1}`,
          values: [["checked_in"]]
        },
        {
          range: `${rangePrefix}!O${rowIndex + 1}`,
          values: [[checkedInBy]]
        }
      ];

      if (rows[0]?.[14]?.toString().trim() !== "checkedInBy") {
        updates.unshift({
          range: `${rangePrefix}!O1`,
          values: [["checkedInBy"]]
        });
      }

      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: "RAW",
          data: updates
        }
      });
    }

    return response.status(200).json({
      ok: true,
      status: "checked_in",
      alreadyCheckedIn,
      checkedInBy: alreadyCheckedIn
        ? rows[rowIndex][14]?.toString().trim() || null
        : checkedInBy
    });
  } catch (error) {
    console.error("Google Sheets check-in failed:", error);
    return response.status(500).json({
      error: "เชื่อมต่อ Google Sheet ไม่สำเร็จ โปรดตรวจสอบ Service Account"
    });
  }
};
