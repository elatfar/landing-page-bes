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
      throw new Error("Tab bernama '1_Framework_Setup' tidak ditemukan!");
    }
    
    var data = sheet.getDataRange().getValues();
    var questions = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][3]) { 
        questions.push({
          id: String(data[i][0]).trim(),
          dimension: data[i][1] ? String(data[i][1]).trim() : "General Control",
          indicator: data[i][2] ? String(data[i][2]).trim() : "",
          text: String(data[i][3]).trim()
        });
      }
    }
    return questions;
  } catch (err) {
    return { error: true, message: err.toString() };
  }
}

function submitAudit(companyName, answersJsonString) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName("2_Audit_Form");
    var setupSheet = ss.getSheetByName("1_Framework_Setup");
    var dashboardSheet = ss.getSheetByName("3_Dashboard_Output");
    
    var timestamp = new Date();
    var setupData = setupSheet.getDataRange().getValues();
    
    // BONGKAR JSON STRING MENJADI ARRAY OBJEK ASLI
    var answers = JSON.parse(answersJsonString);
    
    if (!answers || answers.length === 0) {
      throw new Error("Backend menerima data jawaban kosong setelah di-parse.");
    }
    
    var scoresPerDimension = {};
    var totalWeightPerDimension = {};
    
    // A. Simpan Jawaban Mentah & Hitung Bobot
    answers.forEach(function(ans) {
      var qText = "";
      var qDim = "";
      var qWeight = 0;
      
      var targetId = String(ans.id).trim();
      
      for (var i = 1; i < setupData.length; i++) {
        var currentSetupId = String(setupData[i][0]).trim();
        
        if (currentSetupId === targetId) {
          qDim = setupData[i][1];
          qText = setupData[i][3];
          
          var rawWeight = String(setupData[i][4]).replace('%', '').trim();
          qWeight = parseFloat(rawWeight) || 0;
          if (qWeight > 1) { qWeight = qWeight / 100; }
          break;
        }
      }
      
      if (!qDim) qDim = "General Control";
      
      // Tulis baris ke sheet respons (DI DALAM LOOP, sehingga terisi pasti ada datanya)
      formSheet.appendRow([timestamp, companyName, ans.id, qDim, qText, ans.value]);
      
      var multiplier = 0;
      if (ans.value === "Yes") multiplier = 1.0;
      else if (ans.value === "Partial") multiplier = 0.5;
      else if (ans.value === "No") multiplier = 0.0;
      
      if (!scoresPerDimension[qDim]) {
        scoresPerDimension[qDim] = 0;
        totalWeightPerDimension[qDim] = 0;
      }
      
      scoresPerDimension[qDim] += (qWeight * multiplier);
      totalWeightPerDimension[qDim] += qWeight;
    });
    
    // B. Tulis ke Dashboard & Buat Rangkuman Kembalian
    dashboardSheet.clearContents();
    dashboardSheet.appendRow(["Perusahaan / Klien:", companyName]);
    dashboardSheet.appendRow(["Tanggal Audit:", timestamp]);
    dashboardSheet.appendRow([]);
    dashboardSheet.appendRow(["Dimensi Bisnis", "Skor Efisiensi", "Kategori Kesehatan", "Rekomendasi Strategis"]);
    
    var totalFinalScore = 0;
    var dimensionCount = 0;
    var summaryResults = [];
    
    for (var dim in scoresPerDimension) {
      var maxWeight = totalWeightPerDimension[dim];
      var finalDimPercent = maxWeight > 0 ? (scoresPerDimension[dim] / maxWeight) * 100 : 0;
      
      var healthStatus = "";
      var recommendation = "";
      
      if (finalDimPercent > 80) {
        healthStatus = "Excellent / Sehat";
        recommendation = "Sistem operasional kokoh. Siap diintegrasikan ke ERP Otomatis.";
      } else if (finalDimPercent >= 50) {
        healthStatus = "Needs Improvement / Bocor Halus";
        recommendation = "Benahi standardisasi SOP dan disiplin input data sebelum beli software.";
      } else {
        healthStatus = "Critical / Risiko Tinggi / Bocor Parah";
        recommendation = "Kebocoran sistemik. Stop rencana IT, bereskan kontrol manual dulu!";
      }
      
      dashboardSheet.appendRow([dim, finalDimPercent.toFixed(2) + "%", healthStatus, recommendation]);
      
      summaryResults.push({
        dimension: dim,
        score: finalDimPercent.toFixed(1) + "%",
        status: healthStatus,
        rec: recommendation
      });
      
      totalFinalScore += finalDimPercent;
      dimensionCount++;
    }
    
    var grandTotalScore = dimensionCount > 0 ? (totalFinalScore / dimensionCount) : 0;
    dashboardSheet.appendRow([]);
    dashboardSheet.appendRow(["OVERALL BUSINESS EFFICIENCY SCORE", grandTotalScore.toFixed(2) + "%"]);
    
    return {
      grandScore: grandTotalScore.toFixed(1) + "%",
      breakdown: summaryResults
    };

  } catch (err) {
    return { error: true, message: err.toString() };
  }
}