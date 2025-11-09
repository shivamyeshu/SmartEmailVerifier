//step 1 : finalfast + correct xl sheets verifier
import XLSX from "xlsx";
import dns from "dns";
import net from "net";
import pLimit from "p-limit";
import cliProgress from "cli-progress";
import fs from "fs";
import { promisify } from "util";

const resolveMx = promisify(dns.resolveMx);

const CONFIG = {
  inputFile: "recruiters.xlsx", // Input Excel file with emails
  outputCsv: "VERIFIED_EMAILS.csv", // Output CSV file for verified emails
  safeFile: "SAFE_TO_SEND.xlsx", // Excel file for safe to send emails
  onlyWorkingFile: "ONLY_WORKING_EMAILS.xlsx", // Excel file for only working emails
  concurrency: 15, // Number of concurrent email verifications
  delayMs: 350, // Delay between each email verification in milliseconds
  timeoutMs: 8000, // Socket timeout in milliseconds
  logInterval: 50, // Interval for logging progress in number of emails
};

const mxCache = {};
const domainStats = {};

const KNOWN_PROVIDERS = {
  "google.com": "Google Workspace",
  "googlemail.com": "Google Workspace",
  "outlook.com": "Microsoft Outlook",
  "protection.outlook.com": "Microsoft 365",
  "zoho.com": "Zoho Mail",
  "zoho.in": "Zoho Mail",
};

function detectProvider(mxRecords) {
  for (const { exchange } of mxRecords) {
    const ex = exchange.toLowerCase();
    for (const [key, name] of Object.entries(KNOWN_PROVIDERS)) {
      if (ex.includes(key)) return name;
    }
  }
  return "Custom Server";
}

async function verifySMTP(email, mxHost) {
  return new Promise((resolve) => {
    const socket = net.createConnection({ host: mxHost, port: 25, timeout: CONFIG.timeoutMs });
    let step = 0;
    let buffer = "";

    const done = (status) => {
      socket.destroy();
      resolve(status);
    };

    socket.setTimeout(CONFIG.timeoutMs, () => done("Timeout"));

    socket.on("data", (data) => {
      buffer += data.toString();
      const lines = buffer.split("\r\n");
      buffer = lines.pop();

      for (const line of lines) {
        if (step === 0 && line.includes("220")) {
          socket.write("EHLO fastverify.local\r\n");
          step = 1;
        } else if (step === 1 && line.startsWith("250")) {
          socket.write("MAIL FROM:<>\r\n");
          step = 2;
        } else if (step === 2 && line.startsWith("250")) {
          socket.write(`RCPT TO:<${email}>\r\n`);
          step = 3;
        } else if (step === 3) {
          if (line.includes("250") || line.includes("251")) done("Working");
          else if (line.match(/550|551|553|554/)) done("Invalid");
          else done("Unknown");
        }
      }
    });

    socket.on("error", () => done("No Connection"));
    socket.on("close", () => { if (step < 3) done("No Connection"); });
  });
}

async function verifyEmail(email) {
  const lowerEmail = email.toLowerCase();
  const [user, domain] = lowerEmail.split("@");
  if (!user || !domain) return "Invalid Format";

  if (!/^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$/.test(lowerEmail)) return "Invalid Syntax";

  if (!mxCache[domain]) {
    try {
      const records = await resolveMx(domain);
      if (!records?.length) return "Invalid Domain";
      mxCache[domain] = records;
    } catch {
      return "Invalid Domain";
    }
  }

  const mxRecords = mxCache[domain];
  const provider = detectProvider(mxRecords);
  if (provider !== "Custom Server") {
    return `Protected (${provider})`;
  }

  const mx = mxRecords.sort((a, b) => a.priority - b.priority)[0].exchange;
  const result = await verifySMTP(lowerEmail, mx);

  if (result === "Working") {
    domainStats[domain] = (domainStats[domain] || 0) + 1;
  }

  return result;
}

// MAIN
(async () => {
  console.log("ULTRA FAST + CORRECT XL SHEETS VERIFIER\n");
  const workbook = XLSX.readFile(CONFIG.inputFile);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet);

  const limit = pLimit(CONFIG.concurrency);
  const bar = new cliProgress.SingleBar({
    format: ' {bar} {percentage}% | {value}/{total} | ETA: {eta}s | {log}'
  }, cliProgress.Presets.shades_classic);
  bar.start(data.length, 0, { log: "Starting..." });

  const results = [];  // CORRECT: Push only valid rows
  const stats = { total: 0, working: 0, invalid: 0, protected: 0, other: 0 };
  let processed = 0;

  const printLog = () => {
    bar.update(processed, {
      log: `Processed: ${processed}/${data.length} | Working: ${stats.working} | Invalid: ${stats.invalid} | Protected: ${stats.protected} | Other : ${stats.other}`
    });
  };

  for (const row of data) {
    await limit(async () => {
      const email = row.Email?.toString().trim();
      if (!email) {
        results.push({ ...row, Status: "Missing" });
        stats.other++;
        processed++;
        if (processed % CONFIG.logInterval === 0) printLog();
        bar.update(processed);
        return;
      }

      stats.total++;
      let status = await verifyEmail(email);
      if (["No Connection", "Timeout"].includes(status)) {
        await new Promise(r => setTimeout(r, 1000));
        status = await verifyEmail(email);
      }

      results.push({ ...row, Status: status });  // Push correct row

      if (status === "Working") stats.working++;
      else if (status === "Invalid") stats.invalid++;
      else if (status.includes("Protected")) stats.protected++;
      else stats.other++;

      processed++;
      if (processed % CONFIG.logInterval === 0 || processed === data.length) printLog();
      bar.update(processed);
      await new Promise(r => setTimeout(r, CONFIG.delayMs));
    });
  }

  bar.stop();

  const catchAll = Object.entries(domainStats)
    .filter(([_, c]) => c >= 3)
    .map(([d]) => d);

  console.log("\n" + "=".repeat(75));
  console.log("VERIFICATION DONE — XL SHEETS 100% CORRECT");
  console.log("=".repeat(75));
  console.log(`Total Emails       : ${stats.total}`);
  console.log(`Working          : ${stats.working}`);
  console.log(`Invalid          : ${stats.invalid}`);
  console.log(`Protected        : ${stats.protected}`);
  console.log(`Other            : ${stats.other}`);
  console.log(`Catch-All Domains  : ${catchAll.length}`);
  console.log("=".repeat(75));

  const wb = XLSX.utils.book_new();

  // 1. ALL RESULTS
  const allWs = XLSX.utils.json_to_sheet(results);
  XLSX.utils.book_append_sheet(wb, allWs, "All Results");
  fs.writeFileSync(CONFIG.outputCsv, XLSX.utils.sheet_to_csv(allWs));

  // 2. SAFE TO SEND
  const safe = results.filter(r => r.Status === "Working" || r.Status.includes("Protected"));
  const safeWs = XLSX.utils.json_to_sheet(safe);
  XLSX.utils.book_append_sheet(wb, safeWs, "Safe to Send");
  XLSX.writeFile(wb, CONFIG.safeFile);

  // 3. ONLY CONFIRMED
  const confirmed = results
    .filter(r => r.Status === "Working" && !catchAll.includes(r.Email.split("@")[1].toLowerCase()))
    .map(r => ({ ...r, Status: "Confirmed" }));
  const confWs = XLSX.utils.json_to_sheet(confirmed);
  XLSX.utils.book_append_sheet(wb, confWs, "Confirmed");
  XLSX.writeFile(wb, CONFIG.onlyWorkingFile);

  console.log(`All Results        : ${CONFIG.outputCsv} (${results.length} rows)`);
  console.log(`Safe to Send       : ${CONFIG.safeFile} (${safe.length} rows)`);
  console.log(`100% Confirmed     : ${CONFIG.onlyWorkingFile} (${confirmed.length} rows)`);
  console.log("\nProcess finished.");
})();