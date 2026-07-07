function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Business Efficiency Score (BES) Audit')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getQuestions() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("1_Framework_Setup");
    
    if (!sheet) {
      throw new Error("Tab bernama '1_Framework_Setup' tidak ditemukan di spreadsheet ini!");
    }
    
    var data = sheet.getDataRange().getValues();
    var questions = [];
    
    console.log("Header kolom setup sheet:", JSON.stringify(data[0]));
    console.log("Total baris data (termasuk header):", data.length);
    
    for (var i = 1; i < data.length; i++) {
      var rowId   = data[i][0];
      var rowText = data[i][3];
      
      if (rowId && String(rowId).trim() !== "" && rowText && String(rowText).trim() !== "") {
        questions.push({
          id:        String(rowId).trim(),
          dimension: data[i][1] ? String(data[i][1]).trim() : "General Control",
          indicator: data[i][2] ? String(data[i][2]).trim() : "",
          text:      String(rowText).trim()
        });
      }
    }
    
    console.log("Total pertanyaan dimuat:", questions.length);
    return questions;

  } catch (err) {
    console.error("ERROR di getQuestions:", err.toString());
    return { error: true, message: err.toString() };
  }
}

// -------------------------------------------------------
// Inisialisasi sheet 4_Leads_Record jika belum ada
// -------------------------------------------------------
function getOrCreateLeadsSheet(ss) {
  var sheet = ss.getSheetByName("4_Leads_Record");
  if (!sheet) {
    sheet = ss.insertSheet("4_Leads_Record");
    sheet.appendRow([
      "Timestamp",
      "Nama Perusahaan",
      "Email",
      "Nomor WA Business",
      "Bidang / Jenis Usaha",
      "Lama Operasional (Tahun)",
      "Jumlah Total Karyawan",
      "Overall BES Score",
      "Kategori Overall"
    ]);
    // Format header: bold + freeze baris pertama
    sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 160);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 120);
    sheet.setColumnWidth(9, 160);
    console.log("Sheet '4_Leads_Record' berhasil dibuat.");
  }
  return sheet;
}

// -------------------------------------------------------
// submitAudit — menerima data profil + jawaban audit
// profileJsonString: { companyName, email, whatsapp, industry, yearsOp, totalEmployees }
// -------------------------------------------------------
function submitAudit(profileJsonString, answersJsonString) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet      = ss.getSheetByName("2_Audit_Form");
    var setupSheet     = ss.getSheetByName("1_Framework_Setup");
    var dashboardSheet = ss.getSheetByName("3_Dashboard_Output");
    var leadsSheet     = getOrCreateLeadsSheet(ss);

    if (!formSheet)      throw new Error("Tab '2_Audit_Form' tidak ditemukan!");
    if (!setupSheet)     throw new Error("Tab '1_Framework_Setup' tidak ditemukan!");
    if (!dashboardSheet) throw new Error("Tab '3_Dashboard_Output' tidak ditemukan!");

    var timestamp = new Date();
    var setupData = setupSheet.getDataRange().getValues();

    // Parse profil perusahaan
    var profile = JSON.parse(profileJsonString);
    var companyName    = profile.companyName    || "";
    var email          = profile.email          || "";
    var whatsapp       = profile.whatsapp       || "";
    var industry       = profile.industry       || "";
    var yearsOp        = profile.yearsOp        || "";
    var totalEmployees = profile.totalEmployees || "";

    console.log("Profil diterima:", JSON.stringify(profile));

    // Parse jawaban audit
    var answers = JSON.parse(answersJsonString);
    console.log("Jumlah jawaban diterima:", answers.length);

    if (!answers || answers.length === 0) {
      throw new Error("Data jawaban kosong setelah di-parse.");
    }

    // -------------------------------------------------------
    // Bangun lookup map dari setupData
    // -------------------------------------------------------
    var questionMap = {};
    for (var i = 1; i < setupData.length; i++) {
      var sid = String(setupData[i][0]).trim();
      if (!sid) continue;

      var rawWeight = setupData[i][4];
      var weight = 0;
      if (rawWeight !== undefined && rawWeight !== null && rawWeight !== "") {
        weight = parseFloat(String(rawWeight).replace('%', '').trim()) || 0;
        if (weight > 1) { weight = weight / 100; }
      }

      questionMap[sid] = {
        dim:    setupData[i][1] ? String(setupData[i][1]).trim() : "General Control",
        text:   setupData[i][3] ? String(setupData[i][3]).trim() : "",
        weight: weight
      };
    }

    console.log("Total pertanyaan di questionMap:", Object.keys(questionMap).length);

    // -------------------------------------------------------
    // Fallback bobot merata jika semua bobot nol
    // -------------------------------------------------------
    var allWeightsZero = true;
    for (var key in questionMap) {
      if (questionMap[key].weight > 0) { allWeightsZero = false; break; }
    }

    if (allWeightsZero) {
      console.warn("Semua bobot 0 — pakai equal weight sebagai fallback.");
      var equalWeight = answers.length > 0 ? (1 / answers.length) : 1;
      for (var key in questionMap) {
        questionMap[key].weight = equalWeight;
      }
    }

    // -------------------------------------------------------
    // Hitung skor per dimensi
    // -------------------------------------------------------
    var scoresPerDimension      = {};
    var totalWeightPerDimension = {};

    answers.forEach(function(ans) {
      var targetId = String(ans.id).trim();
      var qData    = questionMap[targetId];

      if (!qData) {
        console.warn("ID tidak ditemukan di setup:", targetId);
        return;
      }

      var qDim    = qData.dim;
      var qText   = qData.text;
      var qWeight = qData.weight;

      formSheet.appendRow([timestamp, companyName, ans.id, qDim, qText, ans.value]);

      var multiplier = 0;
      if (ans.value === "Yes")         multiplier = 1.0;
      else if (ans.value === "Partial") multiplier = 0.5;
      else if (ans.value === "No")      multiplier = 0.0;

      if (!scoresPerDimension[qDim]) {
        scoresPerDimension[qDim]      = 0;
        totalWeightPerDimension[qDim] = 0;
      }

      scoresPerDimension[qDim]      += (qWeight * multiplier);
      totalWeightPerDimension[qDim] += qWeight;

      console.log("ID:", targetId, "| Dim:", qDim, "| Weight:", qWeight, "| Answer:", ans.value);
    });

    // -------------------------------------------------------
    // Tulis ke Dashboard
    // -------------------------------------------------------
    dashboardSheet.clearContents();
    dashboardSheet.appendRow(["Perusahaan / Klien:", companyName]);
    dashboardSheet.appendRow(["Email:", email]);
    dashboardSheet.appendRow(["Nomor WA Business:", whatsapp]);
    dashboardSheet.appendRow(["Bidang Usaha:", industry]);
    dashboardSheet.appendRow(["Lama Operasional:", yearsOp + " Tahun"]);
    dashboardSheet.appendRow(["Jumlah Karyawan:", totalEmployees]);
    dashboardSheet.appendRow(["Tanggal Audit:", timestamp]);
    dashboardSheet.appendRow(["", ""]);
    dashboardSheet.appendRow(["Dimensi Bisnis", "Skor Efisiensi (%)", "Kategori Kesehatan", "Rekomendasi Strategis"]);

    var totalFinalScore = 0;
    var dimensionCount  = 0;
    var summaryResults  = [];

    for (var dim in scoresPerDimension) {
      var maxWeight       = totalWeightPerDimension[dim];
      var finalDimPercent = maxWeight > 0 ? (scoresPerDimension[dim] / maxWeight) * 100 : 0;

      console.log("Dimensi:", dim, "| Final %:", finalDimPercent);

      var healthStatus   = "";
      var recommendation = "";

      if (finalDimPercent > 80) {
        healthStatus   = "Excellent / Sehat";
        recommendation = "Sistem operasional kokoh. Siap diintegrasikan ke ERP Otomatis.";
      } else if (finalDimPercent >= 50) {
        healthStatus   = "Needs Improvement / Bocor Halus";
        recommendation = "Benahi standardisasi SOP dan disiplin input data sebelum beli software.";
      } else {
        healthStatus   = "Critical / Risiko Tinggi / Bocor Parah";
        recommendation = "Kebocoran sistemik. Stop rencana IT, bereskan kontrol manual dulu!";
      }

      dashboardSheet.appendRow([dim, parseFloat(finalDimPercent.toFixed(2)), healthStatus, recommendation]);

      summaryResults.push({
        dimension: dim,
        score:     finalDimPercent.toFixed(1) + "%",
        status:    healthStatus,
        rec:       recommendation
      });

      totalFinalScore += finalDimPercent;
      dimensionCount++;
    }

    if (dimensionCount === 0) {
      throw new Error("Tidak ada dimensi yang berhasil dihitung. Periksa kolom Dimensi (B) di sheet '1_Framework_Setup'.");
    }

    var grandTotalScore = totalFinalScore / dimensionCount;
    dashboardSheet.appendRow(["", ""]);
    dashboardSheet.appendRow(["OVERALL BUSINESS EFFICIENCY SCORE", parseFloat(grandTotalScore.toFixed(2)) + "%"]);

    // -------------------------------------------------------
    // Tulis record ke 4_Leads_Record
    // -------------------------------------------------------
    var overallCategory = "";
    if (grandTotalScore > 80)       overallCategory = "Excellent / Sehat";
    else if (grandTotalScore >= 50) overallCategory = "Needs Improvement / Bocor Halus";
    else                             overallCategory = "Critical / Risiko Tinggi / Bocor Parah";

    leadsSheet.appendRow([
      timestamp,
      companyName,
      email,
      whatsapp,
      industry,
      yearsOp,
      totalEmployees,
      grandTotalScore.toFixed(1) + "%",
      overallCategory
    ]);

    console.log("Record leads berhasil disimpan. Grand Total Score:", grandTotalScore);

    return {
      grandScore: grandTotalScore.toFixed(1) + "%",
      breakdown:  summaryResults
    };

  } catch (err) {
    console.error("ERROR di submitAudit:", err.toString());
    return { error: true, message: err.toString() };
  }
}
