const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const cheerio = require("cheerio");
const xlsx = require("xlsx");
const readline = require("readline");


const SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"];
const TOKEN_PATH = "token.json";
const CREDENTIALS_PATH = "credentials.json";

// Load credentials
async function loadCredentials() {
  const content = fs.readFileSync(CREDENTIALS_PATH);
  return JSON.parse(content);
}

// Authorize using OAuth2 client
async function authorize() {
  const open = (await import("open")).default;
  const credentials = await loadCredentials();
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  if (fs.existsSync(TOKEN_PATH)) {
    const token = fs.readFileSync(TOKEN_PATH);
    oAuth2Client.setCredentials(JSON.parse(token));
    return oAuth2Client;
  }

  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
  });

  console.log("🔗 Authorize this app by visiting this URL:\n", authUrl);
  await open(authUrl);

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  const code = await new Promise((resolve) => {
    rl.question("Enter the code from that page here: ", (code) => {
      rl.close();
      resolve(code);
    });
  });

  const tokenResponse = await oAuth2Client.getToken(code);
  oAuth2Client.setCredentials(tokenResponse.tokens);
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokenResponse.tokens));
  console.log("✅ Token stored to", TOKEN_PATH);
  return oAuth2Client;
}
// Parse invoice email
function parseInvoice(html, subject) {
  const $ = cheerio.load(html);
  const text = $("body").text();

  // const orderMatch = subject.match(/Order#\s*(\d+)/i); 
  // const orderNumber = orderMatch ? orderMatch[1] : "N/A";
  const orderNumber = subject || "N/A";

  const totalMatch = text.match(/Total:\s*\$([\d,\.]+)/i);
  const totalAmount = totalMatch ? totalMatch[1].replace(",", "") : "N/A";

  const items = [];

  $("table")
    .find("tr")
    .each((_, el) => {
      const tds = $(el).find("td");
      if (tds.length >= 5) {
        const itemCode = $(tds[0]).text().trim();
        const name = $(tds[1]).text().trim();
        const priceText = $(tds[3]).text().trim().replace("$", "");
        const quantityText = $(tds[4]).text().trim();

        const price = parseFloat(priceText) || 0;
        const quantity = parseInt(quantityText, 10) || 0;

        if (itemCode && name && price && quantity) {
          items.push({ itemCode, name, price, quantity });
        }
      }
    });

  return {
    orderNumber,
    totalAmount,
    items,
  };
}


// async function main() {
//   const auth = await authorize();
//   const gmail = google.gmail({ version: "v1", auth });

//   const allMessages = [];
//   let pageToken = null;

//   console.log("Fetching messages...");

//   // Fetch all messages with pagination
//   do {
//     const res = await gmail.users.messages.list({
//       userId: "me",
//       q: 'from:"WS Display" subject:invoice after:2024/02/09 before:2025/05/01',
//       maxResults: 100,
//       pageToken,
//     });

//     const messages = res.data.messages || [];
//     allMessages.push(...messages);
//     pageToken = res.data.nextPageToken;

//     console.log(`Fetched ${allMessages.length} messages so far...`);
//   } while (pageToken);

//   const result = [];

//   for (let i = 0; i < allMessages.length; i++) {
//     const msg = allMessages[i];
//     console.log(`Processing ${i + 1}/${allMessages.length}: ${msg.id}`);

//     const detail = await gmail.users.messages.get({ userId: "me", id: msg.id });
//     const parts = detail.data.payload.parts || [];
//     const htmlPart = parts.find((p) => p.mimeType === "text/html");

//     if (htmlPart && htmlPart.body && htmlPart.body.data) {
//       const htmlData = Buffer.from(htmlPart.body.data, "base64").toString("utf-8");

//       // Extract subject
//       const headers = detail.data.payload.headers;
//       const subjectHeader = headers.find((h) => h.name === "Subject");
//       const subject = subjectHeader ? subjectHeader.value : "";

//       const parsed = parseInvoice(htmlData);
//       parsed.orderNumber = subject;
//       result.push(parsed);
//     }
//   }

//   fs.writeFileSync("invoices_ws_display.json", JSON.stringify(result, null, 2));

//   const allItems = result.flatMap((r) =>
//     r.items.map((item) => ({
//       orderNumber: r.orderNumber,
//       totalAmount: r.totalAmount,
//       ...item,
//     }))
//   );

//   const worksheet = xlsx.utils.json_to_sheet(allItems);
//   const workbook = xlsx.utils.book_new();
//   xlsx.utils.book_append_sheet(workbook, worksheet, "Invoices");
//   xlsx.writeFile(workbook, "invoices_ws_display.xlsx");

//   console.log("Export complete.");
// }

// async function main() {
//   const auth = await authorize();
//   const gmail = google.gmail({ version: "v1", auth });

//   let allMessages = [];
//   let nextPageToken = null;

//   console.log("Fetching emails...");

//   do {
//     const res = await gmail.users.messages.list({
//       userId: "me",
//       q: 'from:"WS Display" subject:invoice after:2024/02/09 before:2025/05/01',
//       maxResults: 100,
//       pageToken: nextPageToken,
//     });

//     const messages = res.data.messages || [];
//     allMessages.push(...messages);
//     nextPageToken = res.data.nextPageToken;

//     console.log(`Fetched ${allMessages.length} messages so far...`);
//   } while (nextPageToken);

//   console.log(`Total messages to process: ${allMessages.length}`);

//   const result = [];

//   for (let i = 0; i < allMessages.length; i++) {
//     const msg = allMessages[i];
//     console.log(`Processing ${i + 1}/${allMessages.length}: ${msg.id}`);

//     const detail = await gmail.users.messages.get({ userId: "me", id: msg.id });
//     const headers = detail.data.payload.headers;
//     const subjectHeader = headers.find((h) => h.name === "Subject");
//     const subject = subjectHeader ? subjectHeader.value : "";

//     const parts = detail.data.payload.parts || [];
//     const htmlPart = parts.find((p) => p.mimeType === "text/html");
//     if (!htmlPart || !htmlPart.body || !htmlPart.body.data) continue;

//     const htmlData = Buffer.from(htmlPart.body.data, "base64").toString("utf-8");
//     const $ = cheerio.load(htmlData);

//     const orderMatch = subject.match(/Order#\s*(SO-\d+)/i);
//     const orderNumber = orderMatch ? orderMatch[1] : subject;

//     const totalMatch = $('td:contains("Total:")').next().text().trim().replace(/[^0-9.]/g, "");
//     const totalAmount = totalMatch || "N/A";

//     $("table")
//       .find("tr")
//       .each((_, el) => {
//         const tds = $(el).find("td");
//         if (tds.length >= 5) {
//           const itemCode = $(tds[0]).text().trim();
//           const name = $(tds[1]).text().trim();
//           const price = $(tds[3]).text().trim().replace(/[^0-9.]/g, "");
//           const quantity = $(tds[4]).text().trim().replace(/[^0-9]/g, "");

//           if (itemCode && name && price) {
//             result.push({
//               orderNumber,
//               totalAmount,
//               itemCode,
//               name,
//               price,
//               quantity,
//             });
//           }
//         }
//       });
//   }

//   const worksheet = xlsx.utils.json_to_sheet(result);
//   const workbook = xlsx.utils.book_new();
//   xlsx.utils.book_append_sheet(workbook, worksheet, "Invoices");
//   xlsx.writeFile(workbook, "invoices_ws_display.xlsx");

//   console.log("Export complete.");
// }

// Main function
async function main() {
  const auth = await authorize();
  const gmail = google.gmail({ version: "v1", auth });

  let allMessages = [];
  let nextPageToken = null;

  console.log("Fetching emails...");

  do {
    const res = await gmail.users.messages.list({
      userId: "me",
      q: 'from:"WS Display" subject:invoice after:2024/02/09 before:2025/05/01',
      maxResults: 100,
      pageToken: nextPageToken,
    });

    const messages = res.data.messages || [];
    allMessages.push(...messages);
    nextPageToken = res.data.nextPageToken;

    console.log(`Fetched ${allMessages.length} messages so far...`);
  } while (nextPageToken);

  console.log(`Total messages to process: ${allMessages.length}`);

  const result = [];

  for (let i = 0; i < allMessages.length; i++) {
    const msg = allMessages[i];
    console.log(`Processing ${i + 1}/${allMessages.length}: ${msg.id}`);

    const detail = await gmail.users.messages.get({ userId: "me", id: msg.id });
    const headers = detail.data.payload.headers;
    const subjectHeader = headers.find((h) => h.name === "Subject");
    const subject = subjectHeader ? subjectHeader.value : "";

    const parts = detail.data.payload.parts || [];
    const htmlPart = parts.find((p) => p.mimeType === "text/html");
    if (!htmlPart || !htmlPart.body || !htmlPart.body.data) continue;

    const htmlData = Buffer.from(htmlPart.body.data, "base64").toString("utf-8");

    // Use your parser
    const parsed = parseInvoice(htmlData);

    // Keep orderNumber as parsed from invoice
    parsed.subjectLine = subject; // just add subject as a new field

    result.push(parsed);
  }

  const allItems = result.flatMap((r) =>
    r.items.map((item) => ({
      orderNumber: r.orderNumber,
      totalAmount: r.totalAmount,
      subjectLine: r.subjectLine,
      ...item,
    }))
  );

  const worksheet = xlsx.utils.json_to_sheet(allItems);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Invoices");
  xlsx.writeFile(workbook, "invoices_ws_display.xlsx");

  console.log("Export complete.");
}



main().catch(console.error);
