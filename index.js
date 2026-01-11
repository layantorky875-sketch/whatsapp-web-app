const fs = require("fs");
const path = require("path");
const os = require("os");
const XLSX = require("xlsx");
const readline = require("readline");
const { Client, LocalAuth } = require("whatsapp-web.js");

/* ================= PASSWORD ================= */
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

function askPassword() {
  return new Promise((resolve) => {
    rl.question("🔐 Enter password: ", (pass) => {
      resolve(pass.trim());
    });
  });
}

/* ================= FIND CHROME ================= */
function findChrome() {
  const paths = [
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
    path.join(
      os.homedir(),
      "AppData\\Local\\Google\\Chrome\\Application\\chrome.exe"
    ),
  ];

  for (const p of paths) {
    if (fs.existsSync(p)) return p;
  }
  return null;
}

/* ================= LOAD EXCEL (SMART) ================= */
function loadMessages() {
  const file = "WhatsApp Business.xlsm";
  if (!fs.existsSync(file)) {
    console.log("❌ Excel file not found");
    process.exit();
  }

  const wb = XLSX.readFile(file);
  const ws = wb.Sheets["Send"];
  if (!ws) {
    console.log("❌ Sheet 'Send' not found");
    process.exit();
  }

  // اقرأ الشيت كله Array
  const sheet = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
  });

  // الهيدر في الصف الخامس (index 4)
  const headerRow = sheet[4].map(h =>
    String(h).trim().toLowerCase()
  );

  const phoneCol = headerRow.indexOf("phone");
  const nameCol = headerRow.indexOf("name");
  const messageCol = headerRow.indexOf("message");

  if (phoneCol === -1 || messageCol === -1) {
    console.log("❌ Phone or Message column not found");
    console.log("📄 Columns found:", headerRow);
    process.exit();
  }

  const messages = [];

  // البيانات من الصف السادس (index 5)
  for (let i = 5; i < sheet.length; i++) {
    const row = sheet[i];
    if (!row) continue;

    const phone = String(row[phoneCol] || "").replace(/\D/g, "");
    const name = nameCol !== -1 ? String(row[nameCol] || "") : "";
    const message = String(row[messageCol] || "");

    // لو صف فاضي → تجاهل
    if (!phone || !message) continue;

    messages.push({
      phone,
      name,
      message,
    });
  }

  console.log(`📊 Loaded ${messages.length} messages`);
  return messages;
}

/* ================= MAIN ================= */
(async () => {
  const pass = await askPassword();
  if (pass !== "58975") {
    console.log("❌ Wrong password");
    process.exit();
  }
  rl.close();

  const chromePath = findChrome();
  if (!chromePath) {
    console.log("❌ Chrome not found on this PC");
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
    console.log("🟢 First time only: Scan QR");
  });

  client.on("ready", async () => {
    console.log("✅ WhatsApp Ready");

    for (const m of messages) {
      const chatId = m.phone + "@c.us";
      const text = m.message.replace("{{name}}", m.name);

      try {
        await client.sendMessage(chatId, text);
        console.log("📤 Sent to", m.phone);
        await new Promise(r => setTimeout(r, 20000));
      } catch (e) {
        console.log("❌ Failed:", m.phone);
      }
    }

    console.log("🎉 Finished sending");
    process.exit();
  });

  client.initialize();
})();
