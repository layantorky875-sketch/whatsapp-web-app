const fs = require("fs");
const XLSX = require("xlsx");
const qrcode = require("qrcode-terminal");
const { Client, LocalAuth } = require("whatsapp-web.js");
const readline = require("readline");

const PASSWORD = "58975";
const EXCEL_FILE = "WhatsApp Business.xlsm";
const SHEET_NAME = "Send";

// ================= PASSWORD =================
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

function startBot() {
  if (!fs.existsSync(EXCEL_FILE)) {
    console.log("❌ Excel file not found");
    process.exit(0);
  }

  const wb = XLSX.readFile(EXCEL_FILE, { cellText: true });
  const sheet = wb.Sheets[SHEET_NAME];
  if (!sheet) {
    console.log("❌ Sheet Send not found");
    process.exit(0);
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  let headerRowIndex = -1;
  let phoneCol = -1;
  let messageCol = -1;
  let nameCol = -1;

  // 🔍 Find header row automatically
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].map(c => c.toString().trim().toLowerCase());

    row.forEach((cell, idx) => {
      if (cell.includes("phone") || cell.includes("رقم")) phoneCol = idx;
      if (cell.includes("message") || cell.includes("رسالة")) messageCol = idx;
      if (cell.includes("name") || cell.includes("اسم")) nameCol = idx;
    });

    if (phoneCol !== -1 && messageCol !== -1) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) {
    console.log("❌ Could not detect header row automatically");
    process.exit(0);
  }

  console.log("✅ Header found at row:", headerRowIndex + 1);
  console.log("📌 Phone column:", phoneCol + 1);
  console.log("📌 Message column:", messageCol + 1);
  if (nameCol !== -1) console.log("📌 Name column:", nameCol + 1);

  const data = rows.slice(headerRowIndex + 1)
    .filter(r => r[phoneCol] && r[messageCol])
    .map(r => ({
      phone: r[phoneCol].toString().replace(/\D/g, ""),
      message: r[messageCol].toString(),
      name: nameCol !== -1 ? r[nameCol].toString() : ""
    }));

  if (data.length === 0) {
    console.log("❌ No data rows found");
    process.exit(0);
  }

  console.log(`📊 Loaded ${data.length} messages`);

  const client = new Client({
    authStrategy: new LocalAuth({ clientId: "torky" }),
    puppeteer: {
      headless: false,
      args: ["--no-sandbox", "--disable-setuid-sandbox"]
    }
  });

  client.on("qr", qr => {
    console.log("🟢 Scan QR");
    qrcode.generate(qr, { small: true });
  });

  client.on("ready", async () => {
    console.log("✅ WhatsApp Ready");

    for (const c of data) {
      try {
        const msg = c.message.replace(/{{name}}/gi, c.name);
        await client.sendMessage(`${c.phone}@c.us`, msg);
        console.log("📤 Sent →", c.phone);
        await sleep(4000);
      } catch (e) {
        console.log("❌ Error:", e.message);
        break;
      }
    }

    console.log("🎉 Finished");
  });

  client.initialize();
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}
