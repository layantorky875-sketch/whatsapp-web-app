const fs = require("fs");
const XLSX = require("xlsx");
const qrcode = require("qrcode-terminal");
const { Client, LocalAuth } = require("whatsapp-web.js");
const readline = require("readline");

// ================== SETTINGS ==================
const PASSWORD = "58975";
const EXCEL_FILE = "WhatsApp Business.xlsm";
const SHEET_NAME = "Send";
const HEADER_ROW_INDEX = 5; // الصف السادس (0-based)

// ================== PASSWORD INPUT ==================
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

rl.question("🔐 Enter password:\n> ", (pass) => {
  if (pass !== PASSWORD) {
    console.log("❌ Wrong password");
    process.exit(0);
  }
  rl.close();
  startBot();
});

// ================== MAIN ==================
function startBot() {
  if (!fs.existsSync(EXCEL_FILE)) {
    console.log("❌ Excel file not found");
    process.exit(0);
  }

  // ===== Read Excel (xlsm OK) =====
  const workbook = XLSX.readFile(EXCEL_FILE);
  const sheet = workbook.Sheets[SHEET_NAME];

  if (!sheet) {
    console.log("❌ Sheet 'Send' not found");
    process.exit(0);
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  if (rows.length <= HEADER_ROW_INDEX) {
    console.log("❌ No data rows found");
    process.exit(0);
  }

  // ===== Read header names from row 6 =====
  const headers = rows[HEADER_ROW_INDEX].map(h =>
    h.toString().trim().toLowerCase()
  );

  console.log("📄 Columns found:", headers);

  // ===== Detect columns by NAME =====
  const phoneCol = headers.findIndex(h =>
    h.includes("phone") || h.includes("mobile") || h.includes("رقم")
  );

  const messageCol = headers.findIndex(h =>
    h.includes("message") || h.includes("msg") || h.includes("رسالة")
  );

  const nameCol = headers.findIndex(h =>
    h.includes("name") || h.includes("اسم")
  );

  if (phoneCol === -1 || messageCol === -1) {
    console.log("❌ Phone or Message column not found");
    process.exit(0);
  }

  // ===== Read data starting from row 7 =====
  const dataRows = rows.slice(HEADER_ROW_INDEX + 1);

  const contacts = dataRows
    .filter(r => r[phoneCol] && r[messageCol])
    .map(r => ({
      phone: r[phoneCol].toString().replace(/\D/g, ""),
      message: r[messageCol].toString(),
      name: nameCol !== -1 ? r[nameCol].toString() : ""
    }));

  if (contacts.length === 0) {
    console.log("❌ No valid data to send");
    process.exit(0);
  }

  console.log(`📊 Loaded ${contacts.length} contacts`);

  // ===== WhatsApp Client =====
  const client = new Client({
    authStrategy: new LocalAuth({ clientId: "torky" }),
    puppeteer: {
      headless: false,
      args: ["--no-sandbox", "--disable-setuid-sandbox"]
    }
  });

  client.on("qr", (qr) => {
    console.log("🟢 Scan QR");
    qrcode.generate(qr, { small: true });
  });

  client.on("ready", async () => {
    console.log("✅ WhatsApp Ready");

    for (const c of contacts) {
      try {
        let finalMsg = c.message.replace(/{{name}}/gi, c.name);
        await client.sendMessage(`${c.phone}@c.us`, finalMsg);
        console.log(`📤 Sent → ${c.phone}`);
        await delay(4000); // أمان ضد الحظر
      } catch (err) {
        console.log(`❌ Failed → ${c.phone}`);
        console.log(err.message);
        break;
      }
    }

    console.log("🎉 Finished sending");
  });

  client.initialize();
}

// ================== DELAY ==================
function delay(ms) {
  return new Promise(res => setTimeout(res, ms));
}
