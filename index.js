const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const XLSX = require('xlsx');
const prompt = require('prompt-sync')();
const fs = require('fs');

const config = require('./config.json');

console.log("🔐 Enter password:");
const pass = prompt("> ");

if (pass !== config.password) {
    console.log("❌ Wrong password");
    process.exit(1);
}

if (!fs.existsSync(config.excel_file)) {
    console.log("❌ Excel file not found");
    process.exit(1);
}

const workbook = XLSX.readFile(config.excel_file);
const sheet = workbook.Sheets[config.sheet_name];
const data = XLSX.utils.sheet_to_json(sheet);

const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: {
        headless: true
    }
});

client.on('qr', qr => {
    console.log("🟢 Scan QR:");
    qrcode.generate(qr, { small: true });
});

client.on('ready', async () => {
    console.log("✅ WhatsApp Ready");

    const delay = Math.floor(3600000 / config.rate_per_hour);

    for (let row of data) {
        if (!row.Phone || !row.Message) continue;

        const phone = row.Phone.toString().replace(/\D/g, '') + "@c.us";
        let msg = row.Message;

        if (row.Name) {
            msg = msg.replace("{{name}}", row.Name);
        }

        try {
            await client.sendMessage(phone, msg);
            console.log("📤 Sent to", row.Phone);
            await new Promise(r => setTimeout(r, delay));
        } catch (e) {
            console.log("❌ Failed:", row.Phone);
        }
    }

    console.log("🎉 Finished sending");
    process.exit(0);
});

client.initialize();
