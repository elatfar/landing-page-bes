function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Business Efficiency Score (BES) Audit')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getQuestions() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("1_Framework_Setup");

    if (!sheet) throw new Error("Tab '1_Framework_Setup' tidak ditemukan!");

    var data      = sheet.getDataRange().getValues();
    var questions = [];

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
    console.error("ERROR getQuestions:", err.toString());
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
      "Audit ID",
      "Nama Perusahaan",
      "Email",
      "Nomor WA Business",
      "Bidang / Jenis Usaha",
      "Lama Operasional (Tahun)",
      "Jumlah Total Karyawan",
      "Estimasi Omzet / Bulan",
      "Overall BES Score",
      "Kategori Overall",
      "Analisis AI"
    ]);
    sheet.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,  160);
    sheet.setColumnWidth(2,  200);
    sheet.setColumnWidth(3,  200);
    sheet.setColumnWidth(4,  200);
    sheet.setColumnWidth(5,  160);
    sheet.setColumnWidth(6,  180);
    sheet.setColumnWidth(7,  120);
    sheet.setColumnWidth(8,  120);
    sheet.setColumnWidth(9,  180);
    sheet.setColumnWidth(10, 120);
    sheet.setColumnWidth(11, 160);
    sheet.setColumnWidth(12, 400);
  }
  return sheet;
}

// -------------------------------------------------------
// submitAudit
// Semua kalkulasi dilakukan di memori dan dikembalikan
// langsung ke frontend — TIDAK ada shared state antar user.
//
// 3_Dashboard_Output DIHAPUS dari alur ini untuk mencegah
// race condition saat user bersamaan (clearContents tabrakan).
//
// 4_Leads_Record hanya appendRow — aman untuk concurrent.
// Setiap baris diberi auditId unik untuk lookup AI insight.
// -------------------------------------------------------
function submitAudit(profileJsonString, answersJsonString) {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName("2_Audit_Form");
    var setupSheet= ss.getSheetByName("1_Framework_Setup");
    var leadsSheet= getOrCreateLeadsSheet(ss);

    if (!formSheet)  throw new Error("Tab '2_Audit_Form' tidak ditemukan!");
    if (!setupSheet) throw new Error("Tab '1_Framework_Setup' tidak ditemukan!");

    // ID unik per sesi audit — pakai timestamp + random suffix
    var timestamp = new Date();
    var auditId   = "BES-" + timestamp.getTime() + "-" + Math.floor(Math.random() * 9000 + 1000);

    var setupData = setupSheet.getDataRange().getValues();
    var profile   = JSON.parse(profileJsonString);
    var answers   = JSON.parse(answersJsonString);

    if (!answers || answers.length === 0) {
      throw new Error("Data jawaban kosong.");
    }

    // ── Build question map ──
    var questionMap = {};
    for (var i = 1; i < setupData.length; i++) {
      var sid = String(setupData[i][0]).trim();
      if (!sid) continue;
      var rawWeight = setupData[i][4];
      var weight    = 0;
      if (rawWeight !== undefined && rawWeight !== null && rawWeight !== "") {
        weight = parseFloat(String(rawWeight).replace('%', '').trim()) || 0;
        if (weight > 1) weight = weight / 100;
      }
      questionMap[sid] = {
        dim:    setupData[i][1] ? String(setupData[i][1]).trim() : "General Control",
        text:   setupData[i][3] ? String(setupData[i][3]).trim() : "",
        weight: weight
      };
    }

    // Fallback equal weight
    var allZero = Object.keys(questionMap).every(function(k) { return questionMap[k].weight === 0; });
    if (allZero) {
      var ew = answers.length > 0 ? 1 / answers.length : 1;
      Object.keys(questionMap).forEach(function(k) { questionMap[k].weight = ew; });
    }

    // ── Hitung skor (semua di memori, tidak tulis ke shared sheet dulu) ──
    var scoresPerDim = {};
    var weightPerDim = {};

    answers.forEach(function(ans) {
      var id    = String(ans.id).trim();
      var qData = questionMap[id];
      if (!qData) return;

      // Simpan jawaban mentah ke 2_Audit_Form (append = aman concurrent)
      formSheet.appendRow([timestamp, auditId, profile.companyName || "", ans.id, qData.dim, qData.text, ans.value]);

      var mult = ans.value === "Yes" ? 1.0 : ans.value === "Partial" ? 0.5 : 0.0;
      if (!scoresPerDim[qData.dim]) { scoresPerDim[qData.dim] = 0; weightPerDim[qData.dim] = 0; }
      scoresPerDim[qData.dim] += qData.weight * mult;
      weightPerDim[qData.dim] += qData.weight;
    });

    // ── Rangkuman per dimensi ──
    var summaryResults = [];
    var totalScore     = 0;
    var dimCount       = 0;

    for (var dim in scoresPerDim) {
      var pct    = weightPerDim[dim] > 0 ? (scoresPerDim[dim] / weightPerDim[dim]) * 100 : 0;
      var status = pct > 80
        ? "Excellent / Sehat"
        : pct >= 50
          ? "Needs Improvement / Bocor Halus"
          : "Critical / Risiko Tinggi / Bocor Parah";
      var rec = pct > 80
        ? "Sistem operasional kokoh. Siap diintegrasikan ke ERP Otomatis."
        : pct >= 50
          ? "Benahi standardisasi SOP dan disiplin input data sebelum beli software."
          : "Kebocoran sistemik. Stop rencana IT, bereskan kontrol manual dulu!";

      summaryResults.push({ dimension: dim, score: pct.toFixed(1) + "%", status: status, rec: rec });
      totalScore += pct;
      dimCount++;
    }

    if (dimCount === 0) throw new Error("Tidak ada dimensi yang berhasil dihitung.");

    var grandScore = totalScore / dimCount;
    var category   = grandScore > 80
      ? "Excellent / Sehat"
      : grandScore >= 50
        ? "Needs Improvement / Bocor Halus"
        : "Critical / Risiko Tinggi / Bocor Parah";

    // ── Simpan leads (appendRow — aman concurrent) ──
    // Kolom Analisis AI dikosongkan dulu, diisi saat getAiInsight selesai
    leadsSheet.appendRow([
      timestamp,
      auditId,
      profile.companyName    || "",
      profile.email          || "",
      profile.whatsapp       || "",
      profile.industry       || "",
      profile.yearsOp        || "",
      profile.totalEmployees || "",
      profile.revenue        || "",
      grandScore.toFixed(1) + "%",
      category,
      "" // Analisis AI — diisi terpisah oleh getAiInsight
    ]);

    console.log("submitAudit selesai. auditId:", auditId, "| Score:", grandScore);

    // Kembalikan auditId ke frontend agar bisa diteruskan ke getAiInsight
    return {
      auditId:    auditId,
      grandScore: grandScore.toFixed(1) + "%",
      breakdown:  summaryResults
    };

  } catch (err) {
    console.error("ERROR submitAudit:", err.toString());
    return { error: true, message: err.toString() };
  }
}

// -------------------------------------------------------
// getAiInsight — kirim payload ke n8n, simpan hasil ke
// baris yang TEPAT berdasarkan auditId (bukan lastRow).
// Aman untuk concurrent karena pakai lookup by auditId.
// -------------------------------------------------------
function getAiInsight(insightPayloadJson) {
  try {
    var payload = JSON.parse(insightPayloadJson);
    var auditId = payload.auditId || "";

    var N8N_WEBHOOK_URL = "https://bse-atmam.app.n8n.cloud/webhook/bes-ai-insight";

    var options = {
      method:             "post",
      contentType:        "application/json",
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response   = UrlFetchApp.fetch(N8N_WEBHOOK_URL, options);
    var statusCode = response.getResponseCode();
    var body       = response.getContentText();

    console.log("n8n status:", statusCode, "| auditId:", auditId);

    if (statusCode !== 200) {
      return { error: true, message: "Webhook error " + statusCode + ": " + body };
    }

    var result = JSON.parse(body);
    var aiText = Array.isArray(result)
      ? ((result[0] && (result[0].output || result[0].insight || result[0].text)) || "")
      : (result.insight || result.output || result.text || "");

    if (!aiText || aiText.trim() === "") {
      return { error: true, message: "AI tidak menghasilkan teks. Periksa workflow n8n." };
    }

    // ── Simpan AI insight ke baris yang TEPAT via auditId ──
    if (auditId) {
      try {
        var ss          = SpreadsheetApp.getActiveSpreadsheet();
        var leadsSheet  = ss.getSheetByName("4_Leads_Record");

        if (leadsSheet) {
          var data    = leadsSheet.getDataRange().getValues();
          var header  = data[0];
          var idCol   = header.indexOf("Audit ID");      // kolom B (index 1)
          var aiCol   = header.indexOf("Analisis AI");   // kolom L (index 11)

          if (idCol !== -1 && aiCol !== -1) {
            for (var r = 1; r < data.length; r++) {
              if (String(data[r][idCol]).trim() === auditId) {
                leadsSheet.getRange(r + 1, aiCol + 1).setValue(aiText);
                console.log("AI insight disimpan di baris", r + 1, "untuk auditId:", auditId);
                break;
              }
            }
          }
        }
      } catch (saveErr) {
        console.error("Gagal simpan AI insight:", saveErr.toString());
      }
    }

    return { success: true, insight: aiText };

  } catch (err) {
    console.error("ERROR getAiInsight:", err.toString());
    return { error: true, message: err.toString() };
  }
}
