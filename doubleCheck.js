// step 2 : again check + correct xl sheets verifier
import XLSX from "xlsx";
import dns from "dns";
import net from "net";
import pLimit from "p-limit";
import cliProgress from "cli-progress";
import fs from "fs";
import { promisify } from "util";

const resolveMx = promisify(dns.resolveMx);


// Configuration
const CONFIG = {
  // change these as needed
  inputFile: "ONLY_WORKING_EMAILS.xlsx", // Input Excel file with emails
  sheetName: "Confirmed", // Sheet name containing emails to check
  outputFile: "SMART_TRIPLE_CONFIRMED.xlsx", // Output Excel file for results
  concurrency: 5,    // Number of concurrent email checks     
  delayMs: 1200,    // Delay between each email check     
  timeoutMs: 10000, // Socket timeout
  maxRetries: 3, // Number of retries for each email        
  retryDelay: 3000, // Delay between retries
 };

// Cache
const mxCache = {};
const failedDomains = new Set();

async function verifySMTP(email, mxHost, attempt = 1) {
  return new Promise((resolve) => {
    const socket = net.createConnection({ host: mxHost, port: 25, timeout: CONFIG.timeoutMs });
    let step = 0;
    let buffer = "";

    const done = (status) => {
      socket.destroy();
      resolve({ status, attempt });
    };

    socket.setTimeout(CONFIG.timeoutMs, () => done("Timeout"));

    socket.on("data", (data) => {
      buffer += data.toString();
      const lines = buffer.split("\r\n");
      buffer = lines.pop();

      for (const line of lines) {
        if (step === 0 && line.includes("220")) {
          socket.write("EHLO smartcheck.local\r\n");
          step = 1;
        } else if (step === 1 && line.startsWith("250")) {
          socket.write("MAIL FROM:<>\r\n");
          step = 2;
        } else if (step === 2 && line.startsWith("250")) {
          socket.write(`RCPT TO:<${email}>\r\n`);
          step = 3;
        } else if (step === 3) {
          if (line.includes("250") || line.includes("251")) done(" TRIPLE CONFIRMED ");
          else if (line.match(/550|551|553|554/)) done(" NOW INVALID ");
          else done("UNKNOWN");
        }
      }
    });

    socket.on("error", () => done(" No Connection "));
    socket.on("close", () => { if (step < 3) done(" No Connection "); });
  });
}

// MAIN
(async () => {
  console.log(`SMART TRIPLE CHECK STARTED — "${CONFIG.sheetName}" Sheet\n`);
  
  const workbook = XLSX.readFile(CONFIG.inputFile);
  if (!workbook.SheetNames.includes(CONFIG.sheetName)) {
    console.log(`ERROR: "${CONFIG.sheetName}" sheet not found!`);
    return;
  }

  const sheet = workbook.Sheets[CONFIG.sheetName];
  const data = XLSX.utils.sheet_to_json(sheet);

  console.log(`Found ${data.length} emails. Starting with retry logic...\n`);

  const limit = pLimit(CONFIG.concurrency);
  const bar = new cliProgress.SingleBar({}, cliProgress.Presets.shades_classic);
  bar.start(data.length, 0);

  const results = [];
  let tripleConfirmed = 0;
  let nowInvalid = 0;
  let noConnectionCount = 0;

  for (let i = 0; i < data.length; i++) {
    await limit(async () => {
      const row = data[i];
      const email = row.Email?.toString().trim() || row.email?.toString().trim();
      if (!email) {
        results.push({ ...row, Triple_Status: "Missing" });
        bar.update(i + 1);
        return;
      }
      
      // Extract domain from email
      const domain = email.split("@")[1].toLowerCase();

      // Skip if domain failed 3 times
      if (failedDomains.has(domain)) {
        results.push({ ...row, Triple_Status: "SKIPPED (Domain Blocked)" });
        bar.update(i + 1);
        return;
      }

      // Initialize status variables  
      let status = " No Connection ";
      let finalStatus = " No Connection ";
      let attempt = 0;

      for (attempt = 1; attempt <= CONFIG.maxRetries; attempt++) {
        if (!mxCache[domain]) {
          try {
            const records = await resolveMx(domain);
            if (!records?.length) {
              finalStatus = "Invalid Domain";
              break;
            }
            mxCache[domain] = records.sort((a, b) => a.priority - b.priority);
          } catch {
            finalStatus = "Invalid Domain";
            break;
          }
        }

        const mx = mxCache[domain][0].exchange;
        const result = await verifySMTP(email, mx, attempt);
        status = result.status;

        if (status === " TRIPLE CONFIRMED " || status === " NOW INVALID ") {
          finalStatus = status;
          break;
        }

        if (attempt < CONFIG.maxRetries) {
          await new Promise(r => setTimeout(r, CONFIG.retryDelay));
        }
      }

      // Mark domain as failed after 3 attempts
      if (finalStatus === " No Connection ") {
        failedDomains.add(domain);
        noConnectionCount++;
      }

      results.push({ ...row, Triple_Status: finalStatus, Attempts: attempt });

      if (finalStatus === " TRIPLE CONFIRMED ") tripleConfirmed++;
      else if (finalStatus === " NOW INVALID ") nowInvalid++;

      console.log(`${finalStatus.padEnd(20)} ${email} (Attempt ${attempt})`);
      bar.update(i + 1);
      await new Promise(r => setTimeout(r, CONFIG.delayMs));
    });
  }

  bar.stop();

  console.log("\n");
  console.log("\nProcessing complete.\n");
  console.log("\n" + "=".repeat(75));
  console.log(" SMART TRIPLE CHECK COMPLETE ");
  console.log("=".repeat(75));
  console.log(`Total Checked          : ${data.length}`);
  console.log(`TRIPLE CONFIRMED       : ${tripleConfirmed}`);
  console.log(`NOW INVALID            : ${nowInvalid}`);
  console.log(`NO CONNECTION (Retry)  : ${noConnectionCount}`);
  console.log(`BLOCKED DOMAINS        : ${failedDomains.size}`);
  console.log("=".repeat(75));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(results);
  XLSX.utils.book_append_sheet(wb, ws, "Smart Checked");
  XLSX.writeFile(wb, CONFIG.outputFile);

  console.log(`FINAL LIST: ${CONFIG.outputFile}`);
  console.log("\nProcess finished.");
})();