const fs = require("fs");
const XLSX = require("xlsx");
const qrcode = require("qrcode-terminal");
const { Client, LocalAuth } = require("whatsapp-web.js");

// ================== PASSWORD ==================
const PASSWORD = "58975";

// ================== PASSWORD CHECK ==================
const readline = require("readline").createInterface({
  input: process.stdin,
  output: process.stdout
});

readline.question("🔐 Enter password:\n> ", (input) => {
  if (input !== PASSWORD) {
    console.log("❌ Wrong password");
    process.exit(0);
  }
  readline.close();
  startBot();
});

// ================== MAIN ==================
function startBot() {
  const EXCEL_FILE = "WhatsApp Business.xlsm";

  if (!fs.existsSync(EXCEL_FILE)) {
    console.log("❌ Excel file not found");
    process.exit(0);
  }

  // ===== Read Excel =====
  const workbook = XLSX.readFile(EXCEL_FILE);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  const HEADER_ROW = 5; // الصف السادس
  const headers = rows[HEADER_ROW].map(h =>
    h.toString().trim().toLowerCase()
  );

  const phoneCol = headers.findIndex(h =>
    h.includes("phone") || h.includes("رقم")
  );
  const messageCol = headers.findIndex(h =>
    h.includes("message") || h.includes("رسالة")
  );
  const nameCol = headers.findIndex(h =>
    h.includes("name") || h.includes("اسم")
  );

  if (phoneCol === -1 || messageCol === -1) {
    console.log("❌ Phone or Message column not found");
    process.exit(0);
  }

  const data = rows.slice(HEADER_ROW + 1);

  const contacts = data
    .filter(r => r[phoneCol] && r[messageCol])
    .map(r => ({
      phone: r[phoneCol].toString().replace(/\D/g, ""),
      message: r[messageCol].toString(),
      name: nameCol !== -1 ? r[nameCol].toString() : ""
    }));

  if (contacts.length === 0) {
    console.log("❌ No data found");
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
        await client.sendMessage(`${c.phone}@c.us`, c.message);
        console.log(`📤 Sent → ${c.phone}`);
        await delay(4000); // تأخير لتجنب الحظر
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
