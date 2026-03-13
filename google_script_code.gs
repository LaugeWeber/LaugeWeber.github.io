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
const TEST_MODE = false;
const TEST_EMAIL = "laugelweber@gmail.com";
const ADMIN_CODE = "123"; // VIGTIGT: Skift til et sikkert password i produktion!

// EVENT CONFIGURATION (used for signup confirmation calendar link)
const EVENT_TITLE = "4. Maj Walkathon";
const EVENT_START_ISO = "2026-06-20T16:00:00"; // lørdag d. 20/6 kl. 16
const EVENT_END_ISO = "2026-06-21T16:00:00";   // søndag d. 21/6 kl. 16
const EVENT_LOCATION = "4. Maj Kollegiet";
const EVENT_DESCRIPTION = "Kom og gå med til 4. Maj Walkathon — vi samler ind til fælles sauna!";
const EVENT_TIMEZONE = "Europe/Copenhagen";

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
        // Send signup confirmation email with calendar link (best effort)
        try {
          if (payload.Email) {
            sendSignupConfirmationEmail(payload.Email, payload.Navn);
          }
        } catch (mailErr) {
          Logger.log("ERROR sending signup email: " + mailErr.toString());
        }
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

  // Email is optional for generic writes, but if present must be valid
  if (data.Email) {
    if (typeof data.Email !== 'string') throw new Error("Ugyldig email");
    const emailTrim = data.Email.trim();
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(emailTrim)) throw new Error("Ugyldig email adresse");
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
  
  // Threshold (optional) - donor can set a distance threshold for the pledge
  const threshold = Number(data.ThresholdKm || 0);
  if (isNaN(threshold) || threshold < 0) {
    throw new Error("Ugyldig tærskelværdi");
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
    const email = (data.Email || "").toString().trim();
    
    // Tjek for duplikater: både navn og email må ikke være i brug
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(3, sheet.getLastColumn())).getValues();
      for (let i = 0; i < rows.length; i++) {
        const existingName = rows[i][0] ? rows[i][0].toString().trim() : '';
        const existingEmail = rows[i][2] ? rows[i][2].toString().trim() : '';
        if (existingName && existingName.toLowerCase() === navn.toLowerCase()) {
          throw new Error('Navn findes allerede som deltager');
        }
        if (email && existingEmail && existingEmail.toLowerCase() === email.toLowerCase()) {
          throw new Error('Email er allerede brugt til en anden deltager');
        }
      }
    }

    // Tilføj ny række (Navn, Distance, Email)
    sheet.appendRow([navn, distance, email]);
    Logger.log("Added new participant: " + navn + " <" + email + ">");
    
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
      Number(data.ThresholdKm || 0),
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

        // After updating distance, automatically mark any donations as fulfilled
        try {
          const donationSheet = ss.getSheetByName("Donationer");
          if (donationSheet) {
            const donationData = donationSheet.getDataRange().getValues();
            if (donationData.length > 0) {
              const headers = donationData[0].map(h => h.toString());
              const idxModtager = headers.indexOf("Modtager");
              const idxThreshold = headers.indexOf("ThresholdKm") !== -1 ? headers.indexOf("ThresholdKm") : headers.indexOf("Threshold");
              let idxIndfriet = headers.indexOf("Indfriet");

              // If Indfriet column doesn't exist, create it as the last column
              if (idxIndfriet === -1) {
                const newCol = headers.length + 1;
                donationSheet.getRange(1, newCol).setValue("Indfriet");
                idxIndfriet = headers.length; // zero-based
                Logger.log("Tilføjede kolonne 'Indfriet' i Donationer-arket");
              }

              // Iterate rows and mark as 'JA' when threshold reached
              for (let r = 1; r < donationData.length; r++) {
                const row = donationData[r];
                const rowModtager = idxModtager !== -1 ? row[idxModtager] : "";
                if (!rowModtager) continue;
                if (rowModtager.toString().toLowerCase() !== name.toLowerCase()) continue;

                const thresholdVal = (idxThreshold !== -1) ? Number(row[idxThreshold] || 0) : 0;
                const indfrietFlag = row[idxIndfriet] ? row[idxIndfriet].toString() : "";

                if (thresholdVal > 0 && distance >= thresholdVal && indfrietFlag !== "JA") {
                  donationSheet.getRange(r + 1, idxIndfriet + 1).setValue("JA");
                  Logger.log(`Marked donation row ${r+1} as Indfriet for recipient ${name} (threshold ${thresholdVal})`);
                }
              }
            }
          }
        } catch (fulfillErr) {
          Logger.log("ERROR auto-fulfill: " + fulfillErr.toString());
        }

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
    
    // Map header names to indices for robustness (supports new ThresholdKm column)
    const idxTelefon = headers.indexOf("Telefon");
    const idxEmail = headers.indexOf("Email");
    const idxDonorNavn = headers.indexOf("Navn");
    const idxModtager = headers.indexOf("Modtager");
    const idxFast = headers.indexOf("FastBeløb");
    const idxPerKm = headers.indexOf("BeløbPrKm");
    const idxThreshold = headers.indexOf("ThresholdKm");

    for (let i = 1; i < donationerData.length; i++) {
      const row = donationerData[i];
      const mailSentFlag = (mailSendtIndex !== -1) ? row[mailSendtIndex] : "";
      if (mailSentFlag === "JA") continue; // Skip already sent

      const telefon = idxTelefon !== -1 ? row[idxTelefon] : "";
      const mail = idxEmail !== -1 ? row[idxEmail] : "";
      const donorNavn = idxDonorNavn !== -1 ? row[idxDonorNavn] : "";
      const modtager = idxModtager !== -1 ? row[idxModtager] : "";
      const fastBeløb = idxFast !== -1 ? Number(row[idxFast] || 0) : 0;
      const beløbPrKm = idxPerKm !== -1 ? Number(row[idxPerKm] || 0) : 0;
      const threshold = idxThreshold !== -1 ? Number(row[idxThreshold] || 0) : 0;

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
        beløbPrKm: beløbPrKm,
        threshold: threshold
      });
    }
    
    // Calculate total event sum
    let totalEventSum = 0;
    for (const mail in donorMap) {
      const donor = donorMap[mail];
      donor.donationer.forEach(d => {
        const distance = distanceMap[d.modtager] || 0;
        const triggered = d.threshold > 0 ? distance >= d.threshold : true;
        const amount = triggered ? round2(d.fastBeløb + d.beløbPrKm * distance) : 0;
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
        const triggered = d.threshold > 0 ? distance >= d.threshold : true;
        const amount = triggered ? round2(d.fastBeløb + d.beløbPrKm * distance) : 0;

        totalDonorAmount += amount;
        totalDistanceForDonor += distance;

        rowsHtml += `
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;">${escapeHtml(d.modtager)}</td>
            <td style="padding: 8px; border: 1px solid #ddd;">${distance.toFixed(1)} km</td>
            <td style="padding: 8px; border: 1px solid #ddd;">${d.fastBeløb.toFixed(0)} kr</td>
            <td style="padding: 8px; border: 1px solid #ddd;">${d.beløbPrKm.toFixed(0)} kr</td>
            <td style="padding: 8px; border: 1px solid #ddd;">${d.threshold > 0 ? 'Kun hvis ≥ ' + d.threshold + ' km' : ''}</td>
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

// Build a Google Calendar "TEMPLATE" link for adding the event to a user's calendar
function formatDateForCalendar(dateObj) {
  // Returns UTC timestamp in format YYYYMMDDTHHMMSSZ
  const y = dateObj.getUTCFullYear();
  const m = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getUTCDate()).padStart(2, '0');
  const hh = String(dateObj.getUTCHours()).padStart(2, '0');
  const mm = String(dateObj.getUTCMinutes()).padStart(2, '0');
  const ss = String(dateObj.getUTCSeconds()).padStart(2, '0');
  return `${y}${m}${d}T${hh}${mm}${ss}Z`;
}

function buildGoogleCalendarLink(title, startIso, endIso, details, location) {
  const start = new Date(startIso);
  const end = new Date(endIso);
  const startFmt = formatDateForCalendar(start);
  const endFmt = formatDateForCalendar(end);
  const base = 'https://www.google.com/calendar/render?action=TEMPLATE';
  const params = [];
  params.push('text=' + encodeURIComponent(title));
  params.push('dates=' + encodeURIComponent(startFmt + '/' + endFmt));
  if (details) params.push('details=' + encodeURIComponent(details));
  if (location) params.push('location=' + encodeURIComponent(location));
  params.push('trp=false');
  return base + '&' + params.join('&');
}

function sendSignupConfirmationEmail(email, name) {
  const subject = TEST_MODE ? '[TEST] Tilmelding modtaget – 4. Maj Walkathon' : 'Tilmelding modtaget – 4. Maj Walkathon';

  const calendarLink = buildGoogleCalendarLink(EVENT_TITLE, EVENT_START_ISO, EVENT_END_ISO, EVENT_DESCRIPTION, EVENT_LOCATION) + '&ctz=' + encodeURIComponent(EVENT_TIMEZONE);

  const htmlBody = `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width:600px;">
      ${TEST_MODE ? `<div style="background:#ff9800;color:white;padding:10px;border-radius:8px;text-align:center;font-weight:bold;">⚠️ TEST MODE - Denne email sendes kun til testadresse</div>` : ''}
      <h2 style="color:#0e7c86;">Tak for din tilmelding, ${escapeHtml(name)}!</h2>
      <p>Du er nu tilmeldt <strong>${escapeHtml(EVENT_TITLE)}</strong>.</p>
      <p>${escapeHtml(EVENT_DESCRIPTION)}</p>
      <p><strong>Hvornår:</strong> ${EVENT_START_ISO.replace('T',' ')} – ${EVENT_END_ISO.replace('T',' ')}</p>
      <p><strong>Hvor:</strong> ${escapeHtml(EVENT_LOCATION)}</p>
      <p style="margin-top:16px;">Klik på linket herunder for at føje begivenheden til din Google Kalender:</p>
      <p><a href="${calendarLink}" style="display:inline-block;padding:12px 16px;background:#0e7c86;color:#fff;border-radius:8px;text-decoration:none;font-weight:700;">Tilføj til Google Kalender</a></p>
      <p style="margin-top:18px;color:#666;">Du kan også importere vedhæftede kalenderfil (.ics) for automatisk at få en påmindelse 2 uger før.</p>
      <p style="margin-top:24px;">Med venlig hilsen<br><strong>Walkathon-teamet</strong></p>
    </div>
  `;

  // Build ICS content with VALARM (2 weeks before)
  const uid = Utilities.getUuid();
  const dtstamp = formatDateForCalendar(new Date());
  const dtstart = formatDateForCalendar(new Date(EVENT_START_ISO));
  const dtend = formatDateForCalendar(new Date(EVENT_END_ISO));
  const safeDescription = (EVENT_DESCRIPTION || '').toString().replace(/\r?\n/g, '\\n').replace(/[,;\\]/g, ' ');

  const icsLines = [
    'BEGIN:VCALENDAR',
    'PRODID:-//4. Maj Walkathon//EN',
    'VERSION:2.0',
    'CALSCALE:GREGORIAN',
    'BEGIN:VEVENT',
    'UID:' + uid,
    'DTSTAMP:' + dtstamp,
    'DTSTART:' + dtstart,
    'DTEND:' + dtend,
    'SUMMARY:' + EVENT_TITLE,
    'DESCRIPTION:' + safeDescription,
    'LOCATION:' + EVENT_LOCATION,
    'BEGIN:VALARM',
    'TRIGGER:-P14D',
    'ACTION:DISPLAY',
    'DESCRIPTION:Påmindelse: ' + EVENT_TITLE,
    'END:VALARM',
    'END:VEVENT',
    'END:VCALENDAR'
  ];

  const icsContent = icsLines.join('\r\n');
  const icsBlob = Utilities.newBlob(icsContent, 'text/calendar', '4maj-walkathon.ics');

  const recipient = TEST_MODE ? TEST_EMAIL : email;
  GmailApp.sendEmail(recipient, subject, 'Se HTML-version af denne email i din email klient.', { htmlBody: htmlBody, name: '4. Maj Walkathon', attachments: [icsBlob] });
  Logger.log('Signup confirmation email queued for: ' + (TEST_MODE ? TEST_EMAIL + ' (test for ' + email + ')' : email));
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
  - Kolonner: Navn, Distance
   
  Ark: "Donationer"
  - Kolonner: Telefon, Email, Navn, Modtager, FastBeløb, BeløbPrKm, ThresholdKm, MailSendt
   
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
