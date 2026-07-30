const WEEKLY_API_URL = "https://gateway.uptoskills.com/api/public/leaderboard/pro/past-top-performers?period=weekly&since2026=true&top=3";
const MONTHLY_API_URL = "https://gateway.uptoskills.com/api/public/leaderboard/pro/past-top-performers?period=monthly&months=4&top=3";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, X-Requested-With",
  "Access-Control-Max-Age": "86400",
  "Content-Type": "application/json"
};

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "";
  if (action) {
    return handleRequest(e);
  }
  const output = HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("UpToSkills - Winner Claim Portal")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
  output.setHeaders(CORS_HEADERS);
  return output;
}

function doPost(e) {
  let payload = {};
  if (e.postData && e.postData.contents) {
    try {
      payload = JSON.parse(e.postData.contents);
    } catch(err) {
      // Non-JSON body; proceed with query params only
    }
  }
  return handleRequest(
    Object.assign({}, e, {
      parameter: Object.assign({}, e.parameter, payload),
      postData: e.postData
    })
  );
}

function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders(CORS_HEADERS);
}

function handleRequest(e) {
  const action = e.parameter.action;
  try {
    switch (action) {
      case "getWinnersData":
        return jsonResponse({ success: true, data: getWinnersData() });
      case "checkUserClaimStatus":
        return jsonResponse(
          checkUserClaimStatus(
            e.parameter.userId,
            e.parameter.periodType,
            e.parameter.periodLabel,
            e.parameter.rank
          )
        );
      case "submitWinnerDetails": {
        let formData;
        if (e.postData && e.postData.contents) {
          formData = JSON.parse(e.postData.contents);
        } else if (e.parameter.postData) {
          formData = JSON.parse(e.parameter.postData);
        } else {
          return jsonResponse({ success: false, message: "No form data received" });
        }
        return jsonResponse(submitWinnerDetails(formData));
      }
      case "getAddressFromPinCode":
        return jsonResponse(getAddressFromPinCode(e.parameter.pin));
      default:
        return jsonResponse({ success: false, message: "Invalid action" });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  if (CORS_HEADERS) {
    output.setHeaders(CORS_HEADERS);
  }
  return output;
}

function getWinnersData() {
  try {
    const winnersList = [];

    const weeklyRes = UrlFetchApp.fetch(WEEKLY_API_URL, { muteHttpExceptions: true });
    if (weeklyRes.getResponseCode() === 200) {
      const weeklyJson = JSON.parse(weeklyRes.getContentText());
      if (weeklyJson.success && Array.isArray(weeklyJson.data)) {
        weeklyJson.data.forEach(period => {
          if (!Array.isArray(period.winners)) return;
          period.winners.forEach(w => {
            const fullName = (w.user.name || "").trim();
            const nameParts = fullName.split(" ");
            winnersList.push({
              userId: String(w.user.id),
              fullName,
              firstName: nameParts[0] || "",
              lastName: nameParts.slice(1).join(" ") || "",
              rank: String(w.rank),
              periodType: "Weekly",
              periodLabel: String(period.monthLabel),
              xp: w.total_xp
            });
          });
        });
      }
    }

    const monthlyRes = UrlFetchApp.fetch(MONTHLY_API_URL, { muteHttpExceptions: true });
    if (monthlyRes.getResponseCode() === 200) {
      const monthlyJson = JSON.parse(monthlyRes.getContentText());
      if (monthlyJson.success && Array.isArray(monthlyJson.data)) {
        monthlyJson.data.forEach(period => {
          if (!Array.isArray(period.winners)) return;
          period.winners.forEach(w => {
            const fullName = (w.user.name || "").trim();
            const nameParts = fullName.split(" ");
            winnersList.push({
              userId: String(w.user.id),
              fullName,
              firstName: nameParts[0] || "",
              lastName: nameParts.slice(1).join(" ") || "",
              rank: String(w.rank),
              periodType: "Monthly",
              periodLabel: String(period.monthLabel),
              xp: w.total_xp
            });
          });
        });
      }
    }

    return winnersList;
  } catch (err) {
    Logger.log("Error fetching API data: " + err);
    return [];
  }
}

function checkUserClaimStatus(userId, periodType, periodLabel, rank) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(periodType === "Weekly" ? "Weekly" : "Monthly");
    if (!sheet) {
      return { alreadySubmitted: false, specificClaimExists: false, previousData: null };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { alreadySubmitted: false, specificClaimExists: false, previousData: null };
    }

    let specificClaimExists = false;
    let previousData = null;

    const targetUserId = String(userId).trim();
    const targetType = String(periodType).trim();
    const targetPeriod = String(periodLabel).trim();
    const targetRank = String(rank).trim();

    for (let i = data.length - 1; i >= 1; i--) {
      let rowUserId, rowType, rowPeriod, rowRank;

      if (periodType === "Weekly") {
        rowUserId = String(data[i][12]).trim();
        rowType = String(data[i][13]).trim();
        rowPeriod = String(data[i][14]).trim();
        rowRank = String(data[i][15]).trim();

        if (rowUserId === targetUserId && !previousData) {
          previousData = {
            mobile: data[i][1],
            firstName: data[i][2],
            lastName: data[i][3],
            email: data[i][4],
            addressLine: data[i][5],
            pinCode: data[i][6],
            district: data[i][7],
            state: data[i][8],
            photoUrl: data[i][10]
          };
        }
      } else {
        rowUserId = String(data[i][10]).trim();
        rowType = String(data[i][11]).trim();
        rowPeriod = String(data[i][12]).trim();
        rowRank = String(data[i][13]).trim();

        if (rowUserId === targetUserId && !previousData) {
          previousData = {
            firstName: data[i][1],
            lastName: data[i][2],
            email: data[i][3],
            bankName: data[i][4],
            accountNumber: data[i][5],
            ifsc: data[i][6],
            upiId: data[i][7],
            photoUrl: data[i][8]
          };
        }
      }

      const rowPeriods = String(rowPeriod)
        .split(" | ")
        .map(s => s.trim())
        .filter(Boolean);
      const rowRanks = String(rowRank)
        .split(" | ")
        .map(s => s.trim().replace(/^#/, ""))
        .filter(Boolean);

      if (
        rowUserId === targetUserId &&
        rowType === targetType &&
        rowPeriods.includes(targetPeriod) &&
        rowRanks.includes(targetRank)
      ) {
        specificClaimExists = true;
      }
    }

    return {
      alreadySubmitted: previousData !== null,
      specificClaimExists: specificClaimExists,
      previousData: previousData
    };
  } catch (err) {
    return { alreadySubmitted: false, specificClaimExists: false, previousData: null };
  }
}

function saveFileToDrive(fileData) {
  try {
    const folderName = "Winner_Photos";
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    const contentType = fileData.mimeType;
    const bytes = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(bytes, contentType, fileData.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return file.getUrl();
  } catch (err) {
    Logger.log("File Upload Error: " + err);
    return "Upload Failed: " + err.toString();
  }
}

function getAddressFromPinCode(pin) {
  try {
    const apiUrl = `https://api.postalpincode.in/p/${pin}`;
    const response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { Accept: "application/json" }
    });
    const data = JSON.parse(response.getContentText());
    if (data && Array.isArray(data) && data[0] && data[0].Status === "Success" && data[0].PostOffice && data[0].PostOffice[0]) {
      const po = data[0].PostOffice[0];
      return { district: po.District || "", state: po.State || "" };
    }

    const fallbackUrl = `https://api.zippopotam.us/in/${pin}`;
    const fallbackResponse = UrlFetchApp.fetch(fallbackUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { Accept: "application/json" }
    });
    const fallbackData = JSON.parse(fallbackResponse.getContentText());
    if (fallbackData && fallbackData.places && fallbackData.places[0]) {
      return {
        district: fallbackData.places[0]["place name"] || "",
        state: fallbackData.places[0].state || ""
      };
    }

    return { district: "", state: "" };
  } catch (err) {
    Logger.log("PIN lookup error: " + err);
    return { district: "", state: "" };
  }
}

function buildWeeklyHeaders() {
  return [
    "Serial Number", "Mobile No.", "First Name", "Last Name", "Email",
    "Shipping Complete Address", "Shipping Address Pincode", "Shipping Address City (District)",
    "Shipping Address State", "Quantity", "Professional Photo", "Timestamp",
    "User ID", "Winner Type", "Winning Period", "Rank"
  ];
}

function buildMonthlyHeaders() {
  return [
    "Sl No.", "First Name", "Last Name", "Email", "Bank Name",
    "Account Number", "IFSC", "UPI ID", "Professional Photo", "Timestamp",
    "User ID", "Winner Type", "Winning Period", "Rank"
  ];
}

function buildWeeklyRow(formData, slNo) {
  const completeAddress = [
    formData.addressLine1,
    formData.addressLine2,
    formData.postOffice ? "P.O: " + formData.postOffice : ""
  ].filter(Boolean).join(", ");

  return [
    slNo,
    formData.mobile,
    formData.firstName,
    formData.lastName,
    formData.email,
    completeAddress,
    formData.pinCode,
    formData.district,
    formData.state,
    1,
    formData.photoUrl || "",
    new Date(),
    formData.userId,
    formData.periodType,
    formData.periodLabel,
    formData.rank
  ];
}

function buildMonthlyRow(formData, slNo) {
  return [
    slNo,
    formData.firstName,
    formData.lastName,
    formData.email,
    formData.bankName,
    formData.accountNumber,
    formData.ifsc,
    formData.upiId,
    formData.photoUrl || "",
    new Date(),
    formData.userId,
    formData.periodType,
    formData.periodLabel,
    formData.rank
  ];
}

function submitWinnerDetails(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const isWeekly = formData.periodType === "Weekly";
    const sheetName = isWeekly ? "Weekly" : "Monthly";
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    const expectedHeaders = isWeekly
      ? buildWeeklyHeaders()
      : buildMonthlyHeaders();

    const firstRow = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
    const headersMatch = expectedHeaders.every((h, i) => String(firstRow[i]).trim() === h);
    if (!headersMatch) {
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    }

    const data = sheet.getDataRange().getValues();
    const targetUserId = String(formData.userId).trim();
    const targetType = String(formData.periodType).trim();
    const targetPeriod = String(formData.periodLabel).trim();
    const targetRank = String(formData.rank).trim();

    let existingRowIndex = -1;
    let existingQuantity = 1;

    for (let i = 1; i < data.length; i++) {
      let rowUserId, rowType, rowPeriod, rowRank;

      if (isWeekly) {
        rowUserId = String(data[i][12]).trim();
        rowType = String(data[i][13]).trim();
        rowPeriod = String(data[i][14]).trim();
        rowRank = String(data[i][15]).trim();
      } else {
        rowUserId = String(data[i][10]).trim();
        rowType = String(data[i][11]).trim();
        rowPeriod = String(data[i][12]).trim();
        rowRank = String(data[i][13]).trim();
      }

      const rowPeriods = String(rowPeriod)
        .split(" | ")
        .map(s => s.trim())
        .filter(Boolean);
      const rowRanks = String(rowRank)
        .split(" | ")
        .map(s => s.trim().replace(/^#/, ""))
        .filter(Boolean);

      if (
        rowUserId === targetUserId &&
        rowType === targetType &&
        rowPeriods.includes(targetPeriod) &&
        rowRanks.includes(targetRank)
      ) {
        return {
          success: false,
          message: `Already submitted! You have already claimed the reward for ${formData.periodType} (${formData.periodLabel}) - Rank #${formData.rank}.`
        };
      }

      if (isWeekly && rowUserId === targetUserId) {
        existingRowIndex = i + 1;
        existingQuantity = Number(data[i][9]) || 1;
      }
    }

    let photoUrl = formData.existingPhotoUrl || "";
    if (formData.photoFile && formData.photoFile.data) {
      photoUrl = saveFileToDrive(formData.photoFile);
    }

    const submissionFormData = Object.assign({}, formData, { photoUrl });

    if (isWeekly) {
      if (existingRowIndex > 1 && formData.useExistingMode) {
        const newQty = existingQuantity + 1;
        sheet.getRange(existingRowIndex, 10).setValue(newQty);

        const currentPeriodStr = String(sheet.getRange(existingRowIndex, 15).getValue());
        const currentRankStr = String(sheet.getRange(existingRowIndex, 16).getValue());

        sheet.getRange(existingRowIndex, 15).setValue(currentPeriodStr + " | " + formData.periodLabel);
        sheet.getRange(existingRowIndex, 16).setValue(currentRankStr + " | #" + formData.rank);
        sheet.getRange(existingRowIndex, 12).setValue(new Date());
      } else {
        const slNo = data.length;
        sheet.appendRow(buildWeeklyRow(submissionFormData, slNo));
      }
    } else {
      if (formData.useExistingMode) {
        return { success: true, message: "Claim submitted successfully with existing details!" };
      }
      const slNo = data.length;
      sheet.appendRow(buildMonthlyRow(submissionFormData, slNo));
    }

    return { success: true, message: "Claim submitted successfully!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}