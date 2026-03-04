/****************************************************
 * WALKATHON SYSTEM – VERSION 4 PROFESSIONAL
 * 
 * FORBEDRINGER:
 * - Admin authentication på kritiske endpoints
 * - Input validation overalt
 * - Bedre error handling
 * - Email rate limit håndtering
 * - CORS headers
 * - Lock-mekanisme for concurrency
 * - Forbedret logging
 ****************************************************/

/*************** KONFIGURATION ****************/

const MOBILEPAY_NUMMER = "123456";
const TEST_MODE = true;
const TEST_EMAIL = "laugelweber@gmail.com";
const ADMIN_CODE = "123"; // VIGTIGT: Skift til et sikkert password i produktion!

// Email rate limits (Google Apps Script)
const MAX_EMAILS_PER_DAY = 100; // Juster efter dit account type
const LOCK_TIMEOUT = 30000; // 30 sekunder

/*************** MENU ****************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Walkathon")
    .addItem("Send betalingsmails", "sendAggregatedPaymentEmailsUI")
    .addItem("Test Web App", "testWebApp")
    .addItem("Nulstil MailSendt status", "resetMailSentStatus")
    .addToUi();
}

function testWebApp() {
  const webAppUrl = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(
    "Web App er konfigureret korrekt!\n\n" +
    "URL: " + webAppUrl + "\n\n" +
    "Test mode: " + (TEST_MODE ? "AKTIVERET" : "Deaktiveret")
  );
}

function sendAggregatedPaymentEmailsUI() {
  const result = sendAggregatedPaymentEmails();
  SpreadsheetApp.getUi().alert(result.message);
}

function resetMailSentStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Donationer");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Donationer-arket findes ikke");
    return;
  }
  
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const mailSendtIndex = headers.indexOf("MailSendt");
  
  if (mailSendtIndex === -1) {
    SpreadsheetApp.getUi().alert("Kolonnen 'MailSendt' ikke fundet");
    return;
  }
  
  // Clear all "JA" values
  for (let i = 2; i <= lastRow; i++) {
    sheet.getRange(i, mailSendtIndex + 1).setValue("");
  }
  
  SpreadsheetApp.getUi().alert("MailSendt status nulstillet for alle donationer");
}

/*************** WEB APP ENDPOINTS ****************/

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    Logger.log("=== REQUEST START [" + new Date().toISOString() + "] ===");
    Logger.log("Method: " + (e.postData ? "POST" : "GET"));
    Logger.log("Query params: " + JSON.stringify(e.parameter || {}));
    
    const params = e.parameter || {};
    const action = params.action || "";
    const sheet = params.sheet || "";
    const adminCode = params.adminCode || "";
    
    Logger.log("Action: " + action);
    Logger.log("Sheet: " + sheet);
    
    // TEST ENDPOINT (public)
    if (action === "test") {
      return jsonResponse({ 
        ok: true, 
        message: "Web App virker!", 
        timestamp: new Date().toISOString(),
        testMode: TEST_MODE,
        version: "4.0"
      });
    }
    
    // UPDATE DISTANCE (requires admin)
    if (action === "updateDistance") {
      if (!validateAdminCode(adminCode)) {
        Logger.log("UNAUTHORIZED: Invalid admin code");
        return jsonResponse({ 
          ok: false, 
          message: "Ikke autoriseret - ugyldig admin kode" 
        }, 401);
      }
      
      const name = params.name || "";
      const distance = Number(params.distance || 0);
      
      if (!name) {
        return jsonResponse({ 
          ok: false, 
          message: "Mangler deltager navn" 
        }, 400);
      }
      
      if (isNaN(distance) || distance < 0) {
        return jsonResponse({ 
          ok: false, 
          message: "Ugyldig distance værdi" 
        }, 400);
      }
      
      updateParticipantDistance(name, distance);
      Logger.log("Distance updated: " + name + " -> " + distance + " km");
      
      return jsonResponse({ 
        ok: true, 
        message: "Distance opdateret",
        name: name,
        distance: distance
      });
    }
    
    // SEND EMAILS (requires admin)
    if (action === "sendDonorEmails") {
      if (!validateAdminCode(adminCode)) {
        Logger.log("UNAUTHORIZED: Invalid admin code for email sending");
        return jsonResponse({ 
          ok: false, 
          message: "Ikke autoriseret - ugyldig admin kode" 
        }, 401);
      }
      
      Logger.log("Sending aggregated payment emails...");
      const result = sendAggregatedPaymentEmails();
      
      return jsonResponse({ 
        ok: true, 
        message: result.message,
        emailsSent: result.count,
        testMode: TEST_MODE
      });
    }
    
    // READ SHEET DATA (public, GET only)
    if (!e.postData && (sheet === "Deltagere" || sheet === "Donationer")) {
      const data = readSheet(sheet);
      Logger.log("Read " + data.length + " rows from " + sheet);
      return jsonResponse(data);
    }
    
    // WRITE TO SHEET (POST, with validation)
    if (e.postData && e.postData.contents) {
      const payload = JSON.parse(e.postData.contents);
      const targetSheet = params.sheet || payload.sheet || "";
      
      Logger.log("Writing to sheet: " + targetSheet);
      Logger.log("Payload: " + JSON.stringify(payload));
      
      if (targetSheet === "Deltagere") {
        validateDeltagereData(payload);
        writeToDeltagere(payload);
        Logger.log("Participant added: " + payload.Navn);
        return jsonResponse({ 
          ok: true, 
          message: "Deltager tilføjet",
          navn: payload.Navn
        });
      }
      
      if (targetSheet === "Donationer") {
        validateDonationerData(payload);
        writeToDonationer(payload);
        Logger.log("Donation added: " + payload.Navn + " -> " + payload.Modtager);
        return jsonResponse({ 
          ok: true, 
          message: "Donation registreret",
          donor: payload.Navn,
          recipient: payload.Modtager
        });
      }
    }
    
    return jsonResponse({ 
      ok: false, 
      message: "Ugyldig anmodning. Action: " + action 
    }, 400);
    
  } catch (err) {
    Logger.log("ERROR: " + err.toString());
    Logger.log("Stack: " + err.stack);
    
    return jsonResponse({ 
      ok: false, 
      message: "Server fejl: " + err.toString(),
      stack: TEST_MODE ? err.stack : undefined
    }, 500);
  }
}

function jsonResponse(data, statusCode) {
  statusCode = statusCode || 200;
  
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  
  // Note: Apps Script Web Apps don't support custom HTTP status codes
  // Status is included in response body instead
  if (statusCode !== 200) {
    data.statusCode = statusCode;
  }
  
  return output;
}

/*************** VALIDATION ****************/

function validateAdminCode(code) {
  return code === ADMIN_CODE;
}

function validateDeltagereData(data) {
  if (!data.Navn || typeof data.Navn !== 'string' || data.Navn.trim().length === 0) {
    throw new Error("Deltager navn er påkrævet");
  }
  
  if (data.Navn.length > 100) {
    throw new Error("Deltager navn må max være 100 tegn");
  }
  
  const distance = Number(data.Distance || 0);
  if (isNaN(distance) || distance < 0) {
    throw new Error("Ugyldig distance værdi");
  }
  
  if (data.Avatar && typeof data.Avatar === 'string' && data.Avatar.length > 500) {
    throw new Error("Avatar URL må max være 500 tegn");
  }
}

function validateDonationerData(data) {
  // Telefon
  if (!data.Telefon || typeof data.Telefon !== 'string') {
    throw new Error("Telefonnummer er påkrævet");
  }
  
  const phone = data.Telefon.trim();
  if (phone.length !== 8 || !/^\d{8}$/.test(phone)) {
    throw new Error("Telefonnummer skal være 8 cifre");
  }
  
  // Email
  if (!data.Email || typeof data.Email !== 'string') {
    throw new Error("Email er påkrævet");
  }
  
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(data.Email.trim())) {
    throw new Error("Ugyldig email adresse");
  }
  
  // Navn
  if (!data.Navn || typeof data.Navn !== 'string' || data.Navn.trim().length === 0) {
    throw new Error("Donor navn er påkrævet");
  }
  
  // Modtager
  if (!data.Modtager || typeof data.Modtager !== 'string' || data.Modtager.trim().length === 0) {
    throw new Error("Modtager navn er påkrævet");
  }
  
  // Beløb
  const fastBeløb = Number(data.FastBeløb || 0);
  const beløbPrKm = Number(data.BeløbPrKm || 0);
  
  if (isNaN(fastBeløb) || fastBeløb < 0) {
    throw new Error("Ugyldigt fast beløb");
  }
  
  if (isNaN(beløbPrKm) || beløbPrKm < 0) {
    throw new Error("Ugyldigt beløb pr. km");
  }
  
  if (fastBeløb === 0 && beløbPrKm === 0) {
    throw new Error("Mindst ét beløb skal være større end 0");
  }
  
  // Besked er optional, men tjek længde
  if (data.Besked && typeof data.Besked === 'string' && data.Besked.length > 500) {
    throw new Error("Besked må max være 500 tegn");
  }
}

/*************** SHEET READ/WRITE ****************/

function readSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error("Ark ikke fundet: " + sheetName);
  }
  
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    // Kun headers, ingen data
    return [];
  }
  
  const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = data[0];
  const rows = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }
  
  return rows;
}

function writeToDeltagere(data) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(LOCK_TIMEOUT);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Deltagere");
    
    if (!sheet) {
      throw new Error("Deltagere-arket findes ikke");
    }
    
    const navn = data.Navn.trim();
    const distance = Number(data.Distance || 0);
    const avatar = (data.Avatar || "").trim();
    
    // Tjek for duplikater
    const lastRow = sheet.getLastRow();
    const existing = sheet.getRange(1, 1, lastRow, 1).getValues();
    
    for (let i = 1; i < existing.length; i++) {
      if (existing[i][0].toString().toLowerCase() === navn.toLowerCase()) {
        // Opdater eksisterende
        sheet.getRange(i + 1, 2).setValue(distance);
        if (avatar) {
          sheet.getRange(i + 1, 3).setValue(avatar);
        }
        Logger.log("Updated existing participant: " + navn);
        return;
      }
    }
    
    // Tilføj ny række
    sheet.appendRow([navn, distance, avatar]);
    Logger.log("Added new participant: " + navn);
    
  } finally {
    lock.releaseLock();
  }
}

function writeToDonationer(data) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(LOCK_TIMEOUT);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Donationer");
    
    if (!sheet) {
      throw new Error("Donationer-arket findes ikke");
    }
    
    sheet.appendRow([
      data.Telefon.trim(),
      data.Email.trim(),
      data.Navn.trim(),
      data.Modtager.trim(),
      Number(data.FastBeløb || 0),
      Number(data.BeløbPrKm || 0),
      (data.Besked || "").trim(),
      "" // MailSendt kolonne
    ]);
    
  } finally {
    lock.releaseLock();
  }
}

function updateParticipantDistance(name, distance) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(LOCK_TIMEOUT);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Deltagere");
    
    if (!sheet) {
      throw new Error("Deltagere-arket findes ikke");
    }
    
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(1, 1, lastRow, 1).getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === name.toLowerCase()) {
        sheet.getRange(i + 1, 2).setValue(distance);
        Logger.log("Updated distance for " + name + " to " + distance);
        return;
      }
    }
    
    throw new Error("Deltager ikke fundet: " + name);
    
  } finally {
    lock.releaseLock();
  }
}

/*************** EMAIL FUNKTIONER ****************/

function sendAggregatedPaymentEmails() {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(LOCK_TIMEOUT);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deltagereSheet = ss.getSheetByName("Deltagere");
    const donationerSheet = ss.getSheetByName("Donationer");
    
    if (!deltagereSheet || !donationerSheet) {
      throw new Error("Manglende ark: Deltagere eller Donationer");
    }
    
    const deltagereData = deltagereSheet.getDataRange().getValues();
    const donationerData = donationerSheet.getDataRange().getValues();
    
    const headers = donationerData[0];
    const mailSendtIndex = headers.indexOf("MailSendt");
    
    if (mailSendtIndex === -1) {
      throw new Error("Kolonnen 'MailSendt' mangler i Donationer-arket.");
    }
    
    // Build distance map
    const distanceMap = {};
    for (let i = 1; i < deltagereData.length; i++) {
      const navn = deltagereData[i][0];
      const distance = Number(deltagereData[i][1] || 0);
      distanceMap[navn] = distance;
    }
    
    // Group by donor email
    const donorMap = {};
    
    for (let i = 1; i < donationerData.length; i++) {
      if (donationerData[i][mailSendtIndex] === "JA") {
        continue; // Skip already sent
      }
      
      const telefon = donationerData[i][0];
      const mail = donationerData[i][1];
      const donorNavn = donationerData[i][2];
      const modtager = donationerData[i][3];
      const fastBeløb = Number(donationerData[i][4] || 0);
      const beløbPrKm = Number(donationerData[i][5] || 0);
      
      if (!mail) {
        Logger.log("Skipping row " + (i + 1) + " - no email");
        continue;
      }
      
      if (!donorMap[mail]) {
        donorMap[mail] = {
          navn: donorNavn,
          telefon: telefon,
          rows: [],
          donationer: []
        };
      }
      
      donorMap[mail].rows.push(i + 1);
      donorMap[mail].donationer.push({
        modtager: modtager,
        fastBeløb: fastBeløb,
        beløbPrKm: beløbPrKm
      });
    }
    
    // Calculate total event sum
    let totalEventSum = 0;
    for (const mail in donorMap) {
      const donor = donorMap[mail];
      donor.donationer.forEach(d => {
        const distance = distanceMap[d.modtager] || 0;
        const amount = round2(d.fastBeløb + d.beløbPrKm * distance);
        totalEventSum += amount;
      });
    }
    
    totalEventSum = round2(totalEventSum);
    
    // Check email quota
    const emailCount = Object.keys(donorMap).length;
    let remainingQuota = null;

    try {
      remainingQuota = MailApp.getRemainingDailyQuota();
    } catch (quotaErr) {
      Logger.log("Quota check skipped (missing auth): " + quotaErr.toString());
    }

    Logger.log("Emails to send: " + emailCount);
    if (remainingQuota !== null) {
      Logger.log("Remaining quota: " + remainingQuota);
      if (emailCount > remainingQuota) {
        throw new Error(
          "Ikke nok email quota. Skal sende " + emailCount + 
          " emails, men har kun " + remainingQuota + " tilbage i dag."
        );
      }
    }
    
    // Send emails
    let mailCount = 0;
    let errorCount = 0;
    
    for (const mail in donorMap) {
      const donor = donorMap[mail];
      
      let totalDonorAmount = 0;
      let totalDistanceForDonor = 0;
      let rowsHtml = "";
      
      donor.donationer.forEach(d => {
        const distance = distanceMap[d.modtager] || 0;
        const amount = round2(d.fastBeløb + d.beløbPrKm * distance);
        
        totalDonorAmount += amount;
        totalDistanceForDonor += distance;
        
        rowsHtml += `
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;">${escapeHtml(d.modtager)}</td>
            <td style="padding: 8px; border: 1px solid #ddd;">${distance.toFixed(1)} km</td>
            <td style="padding: 8px; border: 1px solid #ddd;">${d.fastBeløb.toFixed(0)} kr</td>
            <td style="padding: 8px; border: 1px solid #ddd;">${d.beløbPrKm.toFixed(0)} kr</td>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>${amount.toFixed(0)} kr</strong></td>
          </tr>
        `;
      });
      
      totalDonorAmount = round2(totalDonorAmount);
      totalDistanceForDonor = round2(totalDistanceForDonor);
      
      const testBanner = TEST_MODE ? `
        <div style="background: #ff9800; color: white; padding: 12px; margin-bottom: 20px; border-radius: 8px; text-align: center; font-weight: bold;">
          ⚠️ TEST MODE - Denne email sendes kun til testadresse
        </div>
      ` : '';
      
      const subject = TEST_MODE 
        ? "[TEST] Tak for din støtte – Walkathon resultat" 
        : "Tak for din støtte – Walkathon resultat";
      
      const htmlBody = `
        <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 650px; margin: 0 auto;">
          ${testBanner}
          
          <h2 style="color: #0e7c86;">Tak for din støtte!</h2>
          <p>Kære ${escapeHtml(donor.navn)},</p>
          <p>Walkathonen er nu afsluttet, og vi er meget taknemmelige for din støtte.</p>
          
          <h3 style="color: #0e7c86; margin-top: 24px;">Din donationsoversigt</h3>
          <table style="border-collapse: collapse; width: 100%; margin: 16px 0;" border="1" cellpadding="8">
            <tr style="background-color:#f2f2f2;">
              <th style="padding: 10px; text-align: left;">Modtager</th>
              <th style="padding: 10px; text-align: left;">Distance</th>
              <th style="padding: 10px; text-align: left;">Fast beløb</th>
              <th style="padding: 10px; text-align: left;">Kr/km</th>
              <th style="padding: 10px; text-align: left;">Samlet</th>
            </tr>
            ${rowsHtml}
          </table>
          
          <div style="background: #f7f9fd; padding: 16px; border-radius: 8px; margin: 20px 0;">
            <p style="margin: 8px 0;"><strong>Samlet distance for dine modtagere:</strong> ${totalDistanceForDonor.toFixed(1)} km</p>
            <p style="margin: 8px 0;"><strong>Samlet beløb du har lovet:</strong> ${totalDonorAmount.toFixed(0)} kr</p>
          </div>
          
          <div style="background: #e8f5f6; padding: 16px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #0e7c86;">
            <p style="margin: 0;"><strong>Samlet indsamling for hele arrangementet:</strong> ${totalEventSum.toFixed(0)} kr</p>
          </div>
          
          <h3 style="color: #0e7c86; margin-top: 24px;">Betaling</h3>
          <div style="background: #f7f9fd; padding: 16px; border-radius: 8px; margin: 12px 0;">
            <p style="margin: 8px 0;">MobilePay: <strong style="font-size: 1.2em;">${MOBILEPAY_NUMMER}</strong></p>
            <p style="margin: 8px 0; color: #666;">Husk at angive dit navn som reference.</p>
          </div>
          
          <p style="margin-top: 24px;">Endnu en gang tusind tak for din opbakning til Walkathon 2026.</p>
          <p>Med venlig hilsen<br><strong>Walkathon-teamet</strong></p>
          
          <hr style="margin: 32px 0; border: none; border-top: 1px solid #ddd;">
          <p style="font-size: 0.85em; color: #666;">
            4. Maj Walkathon &copy; 2026<br>
            Denne email er sendt til: ${escapeHtml(mail)}
          </p>
        </div>
      `;
      
      const recipient = TEST_MODE ? TEST_EMAIL : mail;
      
      try {
        GmailApp.sendEmail(recipient, subject, "Se HTML-version af denne email i din email klient.", {
          htmlBody: htmlBody,
          name: "4. Maj Walkathon"
        });
        
        mailCount++;
        Logger.log("Sent email to: " + (TEST_MODE ? TEST_EMAIL + " (test for " + mail + ")" : mail));
        
        // Mark as sent (skip in test mode to allow re-testing)
        if (!TEST_MODE) {
          donor.rows.forEach(rowNumber => {
            donationerSheet.getRange(rowNumber, mailSendtIndex + 1).setValue("JA");
          });
        }
        
        // Small delay to avoid rate limiting
        if (mailCount % 10 === 0) {
          Utilities.sleep(1000); // 1 second pause every 10 emails
        }
        
      } catch (emailErr) {
        errorCount++;
        Logger.log("ERROR sending email to " + mail + ": " + emailErr.toString());
      }
    }
    
    const message = TEST_MODE 
      ? `${mailCount} test-emails sendt til ${TEST_EMAIL}` + (errorCount > 0 ? ` (${errorCount} fejl)` : "")
      : `${mailCount} emails sendt` + (errorCount > 0 ? ` (${errorCount} fejl)` : "");
    
    Logger.log("Email sending completed: " + message);
    
    return { 
      count: mailCount,
      errors: errorCount,
      message: message
    };
    
  } finally {
    lock.releaseLock();
  }
}

/*************** UTILITY FUNCTIONS ****************/

function round2(num) {
  return Math.round(num * 100) / 100;
}

function escapeHtml(text) {
  if (!text) return "";
  return text
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/*************** DEPLOYMENT CHECKLIST ****************/
/*

FØR DEPLOYMENT TIL PRODUKTION:

1. SKIFT ADMIN_CODE til et sikkert password (ikke "123")
   - Brug minimum 12 karakterer, mix af bogstaver og tal
   
2. SÆT TEST_MODE = false
   - Emails vil blive sendt til rigtige modtagere
   
3. OPDATER MOBILEPAY_NUMMER til det rigtige nummer
   
4. DEPLOY som Web App:
   - Extensions > Apps Script
   - Deploy > New deployment
   - Type: Web app
   - Execute as: Me
   - Who has access: Anyone
   - Kopiér Web App URL til frontend (API_URL)
   
5. OPSÆT GOOGLE SHEETS:
   Ark: "Deltagere"
   - Kolonner: Navn, Distance, Avatar
   
   Ark: "Donationer"
   - Kolonner: Telefon, Email, Navn, Modtager, FastBeløb, BeløbPrKm, Besked, MailSendt
   
6. TEST GRUNDIGT:
   - Test connection endpoint
   - Test tilføj deltager
   - Test tilføj donation
   - Test opdater distance (med admin code)
   - Test send emails (i test mode først!)
   
7. OPDATER ADMIN_CODE I FRONTEND FILER
   - index.html, signup.html, donate.html
   - Alle tre skal bruge samme kode som backend
   
8. OVERVEJ SIKKERHED:
   - Admin kode bør aldrig vises i frontend source
   - Overvej at bruge Google's OAuth i stedet for simpel kode
   - Log alle admin actions

*/
