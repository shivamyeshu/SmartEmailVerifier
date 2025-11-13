// block.js = SMTP Server Connectivity Checker to know if SMTP server is reachable or not

import net from "net";
// Configuration variables
const TEST_SMTP = "smtp.gmail.com";   // Change to any SMTP server you want (e.g., Outlook, Yahoo, etc.)
const TEST_PORT = 25;
const TOTAL_ATTEMPTS = 100;
const DELAY_MS = 250;  // Delay between each check

let success = 0;
let failure = 0;

function checkSMTP(attempt = 1) {
    if (attempt > TOTAL_ATTEMPTS) {
        console.log(`=== RESULT ===`);
        console.log(`Attempts: ${TOTAL_ATTEMPTS}`);
        console.log(`Success: ${success}`);
        console.log(`Failure: ${failure}`);
        return;
    }

    const socket = net.createConnection({host: TEST_SMTP, port: TEST_PORT, timeout: 5000});
    let status = "error";

    socket.on("connect", () => {
        success++;
        status = "success";
        console.log(`${attempt}: Connection success`);
        socket.destroy();
    });

    socket.on("timeout", () => {
        failure++;
        status = "timeout";
        console.log(`${attempt}:  Timeout`);
        socket.destroy();
    });

    socket.on("error", (err) => {
        failure++;
        status = "error";
        console.log(`${attempt}:  Error: ${err.code || err.message}`);
        socket.destroy();
    });

    socket.on("close", () => {
        setTimeout(() => checkSMTP(attempt + 1), DELAY_MS);
    });
}

checkSMTP();
