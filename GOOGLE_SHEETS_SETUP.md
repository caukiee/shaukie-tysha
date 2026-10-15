# Panduan Sambungkan Kad Kahwin ke Google Sheets (100% Percuma)

Anda **TIDAK PERLU MEMBAYAR APA-APA**. Google Sheets & Google Apps Script adalah perkhidmatan percuma seumur hidup yang disediakan oleh Google melalui akaun Gmail anda.

---

## 🚀 Langkah 1: Buka Google Sheets Baru
1. Buka pelayar dan pergi ke: [https://sheets.google.com](https://sheets.google.com)
2. Log masuk menggunakan **Akaun Google anda**.
3. Cipta fail Spreadsheet baharu dan namakannya: **RSVP & Ucapan Tysha & Shaukie**.

---

## 📜 Langkah 2: Buka Apps Script
1. Di bahagian menu atas Google Sheets, klik **Extensions** (atau **Sambungan**) > **Apps Script**.
2. Padam semua kod sedia ada di dalamnya (`function myFunction() { ... }`).
3. Salin (*copy*) kod lengkap di bawah dan tampal (*paste*) ke dalam editor Apps Script:

```javascript
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "get_ucapan";
  
  if (action === "get_ucapan") {
    var sheet = ss.getSheetByName("Ucapan");
    if (!sheet) {
      sheet = ss.insertSheet("Ucapan");
      sheet.appendRow(["Nama", "Ucapan", "Tarikh", "Likes"]);
      sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#f7e6e9");
      sheet.setFrozenRows(1);
    }
    
    var data = sheet.getDataRange().getValues();
    var list = [];
    // Abaikan baris header pertama
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        list.push({
          row: i + 1, // Baris fizikal dalam sheet
          nama: String(data[i][0]),
          teks: String(data[i][1]),
          masa: data[i][2] ? String(data[i][2]) : "Baru tadi",
          likes: Number(data[i][3]) || 1
        });
      }
    }
    // Susun daripada yang terkini ke terdahulu
    list.reverse();
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: list }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Sokongan kemaskini Like melalui GET
  if (action === "like_ucapan" || action === "like") {
    return handleLike(ss, e.parameter);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var contents = JSON.parse(e.postData.contents);
    var action = contents.action;
    
    if (action === "rsvp") {
      var sheet = ss.getSheetByName("RSVP");
      if (!sheet) {
        sheet = ss.insertSheet("RSVP");
        sheet.appendRow(["Nama", "No Telefon", "Kehadiran", "Bilangan Pax", "Sesi", "Wakil Pilihan", "Ucapan", "Tarikh"]);
        // Format cantik untuk header tab RSVP
        sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#f7e6e9");
        sheet.setFrozenRows(1);
      }
      sheet.appendRow([
        contents.nama || "",
        contents.tel || "",
        contents.kehadiran || "",
        contents.pax || "-",
        contents.sesi || "-",
        contents.wakil || "-",
        contents.ucapan || "",
        new Date().toLocaleString("ms-MY")
      ]);
      sheet.autoResizeColumns(1, 8);
      
      // Jika ada ucapan, auto masukkan ke tab Ucapan sekali
      if (contents.ucapan && contents.ucapan !== "Tiada ucapan" && contents.ucapan.trim() !== "") {
        var ucapanSheet = ss.getSheetByName("Ucapan");
        if (!ucapanSheet) {
          ucapanSheet = ss.insertSheet("Ucapan");
          ucapanSheet.appendRow(["Nama", "Ucapan", "Tarikh", "Likes"]);
        }
        ucapanSheet.appendRow([
          contents.nama || "",
          contents.ucapan || "",
          new Date().toLocaleDateString("ms-MY"),
          1
        ]);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "RSVP disimpan" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "ucapan") {
      var sheet = ss.getSheetByName("Ucapan");
      if (!sheet) {
        sheet = ss.insertSheet("Ucapan");
        sheet.appendRow(["Nama", "Ucapan", "Tarikh", "Likes"]);
      }
      sheet.appendRow([
        contents.nama || "",
        contents.teks || "",
        new Date().toLocaleDateString("ms-MY"),
        1
      ]);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Ucapan disimpan" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Sokongan kemaskini Like melalui POST
    if (action === "like_ucapan" || action === "like") {
      return handleLike(ss, contents);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Action tidak sah" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Fungsi utama kemaskini Like dalam tab 'Ucapan'
function handleLike(ss, params) {
  var sheet = ss.getSheetByName("Ucapan");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sheet Ucapan tidak dijumpai" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var row = params.row ? Number(params.row) : null;
  var targetNama = String(params.nama || "").replace(/\s+/g, " ").trim();
  var targetTeks = String(params.teks || "").replace(/\s+/g, " ").trim();
  var delta = Number(params.delta) || 1;
  var explicitLikes = (params.likes !== undefined && params.likes !== null && params.likes !== "") ? Number(params.likes) : null;

  // 1. Semak baris fizikal jika dibekalkan
  if (row && row >= 2 && row <= data.length) {
    var checkNama = String(data[row - 1][0] || "").replace(/\s+/g, " ").trim();
    if (!targetNama || checkNama === targetNama) {
      var curVal = Number(data[row - 1][3]) || 1;
      var updatedVal = explicitLikes !== null && !isNaN(explicitLikes) ? Math.max(0, explicitLikes) : Math.max(0, curVal + delta);
      sheet.getRange(row, 4).setValue(updatedVal);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", row: row, likes: updatedVal }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 2. Padankan berdasarkan Nama & teks ucapan
  for (var i = 1; i < data.length; i++) {
    var rNama = String(data[i][0] || "").replace(/\s+/g, " ").trim();
    var rTeks = String(data[i][1] || "").replace(/\s+/g, " ").trim();
    
    var namaSama = (rNama === targetNama);
    var teksSama = targetTeks && (rTeks.indexOf(targetTeks) !== -1 || targetTeks.indexOf(rTeks) !== -1);
    
    if (namaSama && (!targetTeks || teksSama)) {
      var cur = Number(data[i][3]) || 1;
      var updated = explicitLikes !== null && !isNaN(explicitLikes) ? Math.max(0, explicitLikes) : Math.max(0, cur + delta);
      sheet.getRange(i + 1, 4).setValue(updated);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", row: i + 1, likes: updated }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "not_found", message: "Ucapan tidak dijumpai" }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

---

## 🌐 Langkah 3: Deploy atau Kemaskini Web App

### Jika Anda Sudah Mempunyai Deployment Sedia Ada (Tak Perlu Tukar URL):
1. Dalam editor Apps Script, tampal kod penuh terkini di atas dan tekan **Ctrl + S** (Save).
2. Di penjuru atas kanan, klik **Deploy** > **Manage deployments**.
3. Klik ikon pensel **✏️ (Edit)** di sebelah deployment aktif anda.
4. Di bahagian **Version**, pilih **New version**.
5. Klik butang **Deploy**.
6. URL Web App anda kekal sama dan sistem terus berfungsi serta-merta!

### Jika Kali Pertama Deploy:
1. Di bahagian atas kanan editor Apps Script, klik butang biru **Deploy** > **New deployment**.
2. Klik ikon gear di sebelah "Select type" > pilih **Web app**.
3. Isikan tetapan:
   - **Description**: `RSVP & Ucapan API v2`
   - **Execute as**: `Me (emel-anda@gmail.com)`
   - **Who has access**: **`Anyone`** *(PENTING)*.
4. Klik **Deploy** dan benarkan akses (Authorize access).
5. Salin **Web app URL** yang terhasil.

---

## 🔗 Langkah 4: Masukkan URL ke dalam Kod Laman Web
Buka fail `index.html`, cari bahagian:
```javascript
const GOOGLE_SCRIPT_URL = "";
```
Tampal URL Web App anda di antara dua tanda petik tersebut:
```javascript
const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxNFd6eqBMXCRO4CNvshZgxbgOz3rG0aTW5tTiTilt1AEUdJKUSaF1CjDDs9zyxhzBt/exec";
```

Siap! Selepas itu:
- Semua RSVP tetamu akan automatik tersusun kemas dalam tab **RSVP** di Google Sheets anda.
- Semua ucapan tetamu akan disimpan secara langsung dalam tab **Ucapan**, dan semua pelawat dari mana-mana peranti akan dapat membaca ucapan yang dikirimkan!
