const fs = require("fs");
const path = require("path");
const os = require("os");
const XLSX = require("xlsx");
const readline = require("readline");
const { Client, LocalAuth } = require("whatsapp-web.js");

/* =============== PASSWORD =============== */
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

function askPassword() {
  return new Promise((resolve) => {
    rl.question("🔐 Enter password: ", (p) => resolve(p.trim()));
  });
}

/* =============== FIND CHROME =============== */
function findChrome() {
  const locations = [
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
    path.join(os.homedir(), "AppData\\Local\\Google\\Chrome\\Application\\chrome.exe"),
  ];
  return locations.find(p => fs.existsSync(p)) || null;
}

/* =============== LOAD EXCEL (NO HEADER) =============== */
function loadMessages() {
  const file = "WhatsApp Business.xlsm";
  if (!fs.existsSync(file)) {
    console.log("❌ Excel file not found");
    process.exit();
  }

  const wb = XLSX.readFile(file, { cellText: true, cellDates: false });
  const ws = wb.Sheets["Send"];
  if (!ws) {
    console.log("❌ Sheet 'Send' not found");
    process.exit();
  }

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const messages = [];

  // نبدأ من الصف السادس (row index = 5)
  for (let r = 5; r <= range.e.r; r++) {
    let phone = "";
    let message = "";
    let name = "";

    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      const cell = ws[cellRef];
      if (!cell || !cell.v) continue;

      const val = String(cell.v).trim();

      // رقم دولي
      if (!phone && /^\d{10,15}$/.test(val.replace(/\D/g, ""))) {
        phone = val.replace(/\D/g, "");
        continue;
      }

      // رسالة (نص مش رقم)
      if (!message && val.length > 3 && !/^\d+$/.test(val)) {
        message = val;
        continue;
      }
    }

    if (!phone || !message) continue;

    messages.push({ phone, name, message });
  }

  console.log(`📊 Loaded ${messages.length} messages`);
  return messages;
}

/* =============== MAIN =============== */
(async () => {
  const pass = await askPassword();
  if (pass !== "58975") {
    console.log("❌ Wrong password");
    process.exit();
  }
  rl.close();

  const chromePath = findChrome();
  if (!chromePath) {
    console.log("❌ Chrome not found");
    process.exit();
  }

  const messages = loadMessages();
  if (messages.length === 0) {
    console.log("⚠️ No messages to send");
    process.exit();
  }

  const client = new Client({
    authStrategy: new LocalAuth({ clientId: "torky" }),
    puppeteer: {
      headless: false,
      executablePath: chromePath,
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    },
  });

  client.on("qr", () => {
    console.log("🟢 Scan QR (first time only)");
  });

  client.on("ready", async () => {
    console.log("✅ WhatsApp Ready");

    for (const m of messages) {
      try {
        await client.sendMessage(m.phone + "@c.us", m.message);
        console.log("📤 Sent:", m.phone);
        await new Promise(r => setTimeout(r, 20000));
      } catch {
        console.log("❌ Failed:", m.phone);
      }
    }

    console.log("🎉 Finished");
    process.exit();
  });

  client.initialize();
})();
