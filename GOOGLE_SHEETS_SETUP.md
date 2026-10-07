# Panduan Sambungkan Kad Kahwin ke Google Sheets (100% Percuma)

Anda **TIDAK PERLU MEMBAYAR APA-APA**. Google Sheets & Google Apps Script adalah perkhidmatan percuma seumur hidup yang disediakan oleh Google melalui akaun Gmail anda.

---

## 🚀 Langkah 1: Buka Google Sheets Baru
1. Buka pelayar dan pergi ke: [https://sheets.google.com](https://sheets.google.com)
2. Log masuk menggunakan **Akaun Google anda**.
3. Cipta fail Spreadsheet baharu dan namakannya: **RSVP & Ucapan Shaukie Tysha**.

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
      sheet.appendRow(["Pak Ngah & Mak Ngah", "Barakallahu lakuma wa baraka 'alaikuma. Semoga ikatan Shaukie dan Tysha diberkati hingga ke anak cucu dan kekal till Jannah!", "Semalam", 8]);
      sheet.appendRow(["Farhan & Keluarga", "Tahniah Shaukie & Tysha! Selamat melangkah ke fasa baharu kehidupan berumah tangga. Semoga sentiasa sakinah, mawaddah, warahmah.", "2 hari lepas", 5]);
      sheet.appendRow(["Alia Syahirah", "So happy for both of you! Cantik sangat pelamin rose merah ni. Cant wait to see you guys on 26 Disember nanti!", "3 hari lepas", 12]);
    }
    
    var data = sheet.getDataRange().getValues();
    var list = [];
    // Abaikan baris header pertama
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        list.push({
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
    
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Action tidak sah" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

---

## 🌐 Langkah 3: Deploy sebagai Web App
1. Di bahagian atas kanan editor Apps Script, klik butang biru **Deploy** > **New deployment**.
2. Klik ikon gear di sebelah "Select type" > pilih **Web app**.
3. Isikan tetapan berikut:
   - **Description**: `RSVP & Ucapan API`
   - **Execute as**: `Me (emel-anda@gmail.com)`
   - **Who has access**: **`Anyone`** *(PENTING: Pilih "Anyone" supaya tetamu boleh hantar RSVP dan ucapan tanpa perlu login akaun Google)*.
4. Klik **Deploy**.
5. Jika keluar tetingkap *"Authorization Required"*:
   - Klik **Authorize access** > pilih akaun Google anda > klik **Advanced** > klik **Go to Untitled project (unsafe)** > klik **Allow**. *(Ini normal untuk skrip peribadi anda sendiri)*.
6. Salin **Web app URL** yang terhasil (contoh format: `https://script.google.com/macros/s/AKfycb.../exec`).

---

## 🔗 Langkah 4: Masukkan URL ke dalam Kod Laman Web
Buka fail `index.html`, cari bahagian:
```javascript
const GOOGLE_SCRIPT_URL = "";
```
Tampal URL Web App anda di antara dua tanda petik tersebut:
```javascript
const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxxxxxxxxxxxxxxxxxxxxxx/exec";
```

Siap! Selepas itu:
- Semua RSVP tetamu akan automatik tersusun kemas dalam tab **RSVP** di Google Sheets anda.
- Semua ucapan tetamu akan disimpan secara langsung dalam tab **Ucapan**, dan semua pelawat dari mana-mana peranti akan dapat membaca ucapan yang dikirimkan!
