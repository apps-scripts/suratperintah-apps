// ==========================================
// KONFIGURASI
// ==========================================
const SPREADSHEET_ID = '1rUcZO70L_Di5lEUWC_k0M6jxeWReoAzW7TWtP0wVMTM';
const TEMPLATE_ID_KECIL = '1Z8jDSpoNBVguPCU5B2VoXa9ePnGC2gVMz6JIOwxtwyY'; 
const TEMPLATE_ID_BESAR = '1f9cQ1I2Jfgm9Bh6LC7MpXOU977go5JWgWXZOvCjvTgc';   

const SHEET_PEGAWAI = 'DataPegawai';  
const SHEET_KEGIATAN = 'DataKegiatan'; 
const SHEET_PEJABAT = 'DataPejabat';
const SHEET_DASAR_HUKUM = 'DataDasarHukum';
const SHEET_RIWAYAT = 'DataRiwayat';

// ==========================================
// 1. FUNGSI UTAMA
// ==========================================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Aplikasi Surat Perintah')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// 2. FUNGSI DATABASE (GET DATA MASTER)
// ==========================================
function getDataMaster() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    const getSheetData = (sheetName, numCols) => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 0) {
        const startRow = (sheetName === SHEET_DASAR_HUKUM) ? 1 : 2; 
        const lastRow = sheet.getLastRow();
        if (lastRow >= startRow) {
           return sheet.getRange(startRow, 1, lastRow - (startRow - 1), numCols).getDisplayValues();
        }
      }
      return[];
    };

    return { 
      pegawai: getSheetData(SHEET_PEGAWAI, 4),
      kegiatan: getSheetData(SHEET_KEGIATAN, 2),
      pejabat: getSheetData(SHEET_PEJABAT, 3),
      dasarHukum: getSheetData(SHEET_DASAR_HUKUM, 1),
      riwayat: getSheetData(SHEET_RIWAYAT, 6)
    };
  } catch (e) {
    throw new Error("Gagal Baca Database: " + e.message);
  }
}

function getRiwayatData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RIWAYAT);
  if (sheet && sheet.getLastRow() > 1) {
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
  }
  return[];
}

function hapusRiwayatById(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_RIWAYAT);
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return "Sukses";
      }
    }
  }
  throw new Error("Data Riwayat tidak ditemukan");
}

// --- CRUD PEGAWAI ---
function tambahPegawai(data) {
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PEGAWAI).appendRow(["'"+data.nama, "'"+data.nip, data.gol, data.jab]);
  return "Sukses";
}

// --- FUNGSI BULK UPLOAD PEGAWAI (EXCEL) ---
function tambahPegawaiBulk(dataArray) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PEGAWAI);
    const rows = dataArray.map(function(item) {
      return ["'" + item.nama, "'" + item.nip, item.gol, item.jab];
    });
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    }
    return "Sukses";
  } catch (e) {
    throw new Error("Gagal menyimpan data Excel: " + e.message);
  }
}

function editPegawai(data, index) {
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PEGAWAI).getRange(index + 2, 1, 1, 4).setValues([["'"+data.nama, "'"+data.nip, data.gol, data.jab]]);
  return "Sukses";
}
function hapusPegawaiByIndex(index) {
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PEGAWAI).deleteRow(index + 2);
  return "Sukses";
}

// --- CRUD KEGIATAN ---
function tambahKegiatan(data) {
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_KEGIATAN).appendRow([data.kode, data.nama]);
  return "Sukses";
}
function editKegiatan(data, index) {
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_KEGIATAN).getRange(index + 2, 1, 1, 2).setValues([[data.kode, data.nama]]);
  return "Sukses";
}
function hapusKegiatanByIndex(index) {
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_KEGIATAN).deleteRow(index + 2);
  return "Sukses";
}

// --- CRUD PEJABAT ---
function tambahPejabat(data) {
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PEJABAT).appendRow(["'"+data.nama, "'"+data.nip, data.jab]);
  return "Sukses";
}
function editPejabat(data, index) {
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PEJABAT).getRange(index + 2, 1, 1, 3).setValues([["'"+data.nama, "'"+data.nip, data.jab]]);
  return "Sukses";
}
function hapusPejabatByIndex(index) {
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PEJABAT).deleteRow(index + 2);
  return "Sukses";
}

// ==========================================
// 3. GENERATE PDF & SIMPAN RIWAYAT
// ==========================================
// ==========================================
// UPDATE FUNGSI GENERATE PDF (LOGIKA 2 TEMPLATE)
// ==========================================

function generatePDF(formData) {
  let tempFile;
  try {
    const timestamp = formData.idRiwayat ? formData.idRiwayat : new Date().getTime().toString();
    formData.idRiwayat = timestamp; 
    
    // --- 1. SIMPAN KE RIWAYAT ---
    try {
      const sheetRiwayat = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_RIWAYAT);
      if (sheetRiwayat) {
        const rowData = [
          timestamp,
          formData.nomorSP,
          formData.tanggalSP,
          formData.pejabat,
          formData.namaKegiatan,
          JSON.stringify(formData) 
        ];
        let isUpdated = false;
        if (sheetRiwayat.getLastRow() > 1) {
          const data = sheetRiwayat.getRange(2, 1, sheetRiwayat.getLastRow() - 1, 1).getValues();
          for (let i = 0; i < data.length; i++) {
            if (data[i][0] == timestamp) {
              sheetRiwayat.getRange(i + 2, 1, 1, 6).setValues([rowData]);
              isUpdated = true;
              break;
            }
          }
        }
        if (!isUpdated) sheetRiwayat.appendRow(rowData);
      }
    } catch(e) { console.log("Gagal menyimpan riwayat: " + e.message); }

    // --- 2. PILIH TEMPLATE BERDASARKAN JUMLAH PEGAWAI ---
    const jmlPegawai = formData.listPegawai ? formData.listPegawai.length : 0;
    // Gunakan ID Template yang sesuai (Pastikan variabel ini ada di bagian atas script Anda)
    const selectedTemplateId = (jmlPegawai <= 2) ? TEMPLATE_ID_KECIL : TEMPLATE_ID_BESAR;
    
    const templateFile = DriveApp.getFileById(selectedTemplateId);
    const tempFolder = DriveApp.getRootFolder();
    const nomorAman = formData.nomorSP.replace(/\//g, '-');
    const namaFileBaru = 'NOMOR_SP_' + formData.nomorSP; 
    tempFile = templateFile.makeCopy(namaFileBaru, tempFolder);
    const tempDoc = DocumentApp.openById(tempFile.getId());
    
    const body = tempDoc.getBody();
    const header = tempDoc.getHeader();
    const footer = tempDoc.getFooter();

    // Fungsi pembantu replace
    const universalReplace = (tag, value) => {
      const val = value || ""; 
      body.replaceText(tag, val);
      if (header) header.replaceText(tag, val);
      if (footer) footer.replaceText(tag, val);
    };

    const replaceOrCollapseLine = (tag, value) => {
      const found = body.findText(tag);
      if (found) {
        const element = found.getElement();
        const paragraph = element.getParent().asParagraph();
        if (value && value.trim() !== "") {
          paragraph.replaceText(tag, value);
        } else {
          paragraph.replaceText(tag, "");
          paragraph.setSpacingAfter(0);
          paragraph.setSpacingBefore(0);
          paragraph.setLineSpacing(1);
          paragraph.editAsText().setFontSize(1);
        }
      }
    };

    const tglIndo = formatTanggalIndonesia(formData.tanggalSP);
    
    // --- 3. LOGIKA REPLACE TAG GLOBAL ---
    // Header TTD (a.n. / Sekretaris)
    replaceOrCollapseLine('{{TTD_HEADER_1}}', formData.ttdHeader1);
    replaceOrCollapseLine('{{TTD_HEADER_2}}', formData.ttdHeader2);
    replaceOrCollapseLine('{{TTD_HEADER_3}}', formData.ttdHeader3); 

    // Info Pejabat & TTE
    universalReplace('{{JABATAN_TTE}}', formData.tteInfo);
    universalReplace('{{NAMA_PEJABAT}}', formData.pejabat);
    if (formData.pejabatData) {
       const nipText = formData.pejabatData.nip ? "" + formData.pejabatData.nip : "";
       universalReplace('{{NIP_PEJABAT}}', nipText);
    }
    universalReplace('{{JABATAN_PEJABAT}}', ""); // Tags lama

    // Nomor & Tanggal
    universalReplace('{{NOMOR_SP}}', formData.nomorSP);
    universalReplace('{{TANGGAL_SP}}', tglIndo); 

    // Kegiatan
    const kalimatKegiatan = `Sub Kegiatan ${formData.namaKegiatan || "..."} sesuai kodering : ${formData.kodeKegiatan || "..."}`;
    universalReplace('{{KODE_KEGIATAN}}', kalimatKegiatan);
    universalReplace('{{NAMA_KEGIATAN}}', formData.namaKegiatan);

    // List Dasar Hukum & Uraian Tugas
    if (formData.dasarHukum) processListInTable(body, '{{DASAR_HUKUM}}', formData.dasarHukum);
    if (formData.uraian) processListInTable(body, '{{URAIAN_TUGAS}}', formData.uraian);

    // --- 4. LOGIKA PENGISIAN PEGAWAI (DISTINCT PER TEMPLATE) ---
    const tables = body.getTables();
    
    if (jmlPegawai <= 2) {
      let tableKepada;
      for (let i = 0; i < tables.length; i++) {
        if (tables[i].getText().includes("{{DAFTAR_KEPADA}}")) {
          tableKepada = tables[i];
          break;
        }
      }

      if (tableKepada) {
        tableKepada.setBorderWidth(0);

        // --- FUNGSI PEMBANTU YANG SUDAH DIOPTIMALKAN ---
        const rapatkanSel = (cell) => {
          let p = cell.getChild(0).asParagraph();
          p.setSpacingAfter(0);   // Jarak bawah paragraf 0
          p.setSpacingBefore(0);  // Jarak atas paragraf 0
          p.setLineSpacing(1.0);  // Spasi baris tunggal
          
          // KUNCI UTAMA: Hilangkan Jarak (Padding) Sel
          cell.setPaddingTop(0);    // Jarak teks ke atas sel 0
          cell.setPaddingBottom(1); // Beri 1pt saja agar tidak terlalu nempel garis atas
          cell.setPaddingLeft(0);   
          cell.setPaddingRight(0);
          return cell;
        };

        formData.listPegawai.forEach((peg, index) => {
          // Baris Nama
          let r1 = tableKepada.appendTableRow();
          r1.setMinimumHeight(0); // Biarkan sistem mengikuti tinggi teks terkecil
          rapatkanSel(r1.appendTableCell((index + 1) + ". ")).setWidth(25);
          rapatkanSel(r1.appendTableCell("Nama")).setWidth(140);
          rapatkanSel(r1.appendTableCell(": " + (peg.nama || "")));
          
          // Baris Pangkat
          if (peg.gol && peg.gol !== "-") {
            let r2 = tableKepada.appendTableRow();
            r2.setMinimumHeight(0);
            rapatkanSel(r2.appendTableCell("")).setWidth(25);
            rapatkanSel(r2.appendTableCell("Pangkat / Gol. Ruang")).setWidth(140);
            rapatkanSel(r2.appendTableCell(": " + peg.gol));
          }

          // Baris NIP
          if (peg.nip && peg.nip !== "-" && peg.nip !== "") {
            let r3 = tableKepada.appendTableRow();
            r3.setMinimumHeight(0);
            rapatkanSel(r3.appendTableCell("")).setWidth(25);
            rapatkanSel(r3.appendTableCell("NIP.")).setWidth(140);
            rapatkanSel(r3.appendTableCell(": " + peg.nip));
          }

          // Baris Jabatan
          let r4 = tableKepada.appendTableRow();
          r4.setMinimumHeight(0);
          rapatkanSel(r4.appendTableCell("")).setWidth(25);
          rapatkanSel(r4.appendTableCell("Jabatan")).setWidth(140);
          rapatkanSel(r4.appendTableCell(": " + (peg.jab || "")));

          // Jarak antar pegawai (Spacer) - dibuat sangat tipis
          if (index < formData.listPegawai.length - 1) {
            let spacerRow = tableKepada.appendTableRow();
            spacerRow.setMinimumHeight(8); // Atur lebar jarak antar orang di sini (misal 8pt)
            rapatkanSel(spacerRow.appendTableCell("")).setWidth(25);
            rapatkanSel(spacerRow.appendTableCell("")).setWidth(140);
            rapatkanSel(spacerRow.appendTableCell(""));
          }
        });
        tableKepada.removeRow(0); 
      }

    } else {
      /**
       * TEMPLATE BESAR (Format Tabel Horizontal)
       * Mencari tabel yang memiliki header "NAMA DAN NIP"
       */
      let tablePegawai;
      for (let i = 0; i < tables.length; i++) {
        if (tables[i].getText().toUpperCase().includes("NAMA DAN NIP")) {
          tablePegawai = tables[i];
          break;
        }
      }
      if (!tablePegawai) tablePegawai = tables[tables.length - 1];

      if (formData.listPegawai) {
        formData.listPegawai.forEach((peg, index) => {
          const tr = tablePegawai.appendTableRow();
          tr.appendTableCell((index + 1).toString()).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          
          const stringNip = (peg.nip && peg.nip !== "-") ? "\nNIP. " + peg.nip : "";
          tr.appendTableCell(peg.nama + stringNip);
          
          const stringGol = (peg.gol && peg.gol !== "-") ? "\n" + peg.gol : "";
          tr.appendTableCell((peg.jab || "") + stringGol);
        });
      }
    }

    // --- 5. FINALIZE ---
    tempDoc.saveAndClose();
    const pdfBlob = tempFile.getAs(MimeType.PDF);
    const base64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    tempFile.setTrashed(true);
    
    return { base64: base64, idRiwayat: formData.idRiwayat };
    
  } catch (e) {
    if (tempFile) try { tempFile.setTrashed(true); } catch(x){}
    console.error("Error Detail: " + e.message);
    throw new Error("Gagal Generate PDF: " + e.message);
  }
}

function formatTanggalIndonesia(dateString) {
  if (!dateString) return "-";
  const parts = dateString.split('-');
  if (parts.length !== 3) return dateString;
  const namaBulan =["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return `${parts[2]} ${namaBulan[parseInt(parts[1]) - 1]} ${parts[0]}`;
}

function processListInTable(body, placeholder, dataArray) {
  const found = body.findText(placeholder);
  if (found) {
    const element = found.getElement();
    const paragraph = element.getParent();
    const cell = paragraph.getParent();
    
    let cleanData =[];
    if (dataArray && Array.isArray(dataArray)) {
      cleanData = dataArray.filter(item => item && item.toString().trim() !== "");
    }

    if (cleanData.length > 0) {
      cleanData.forEach(item => {
        const listItem = cell.appendListItem(item.toString())
            .setGlyphType(DocumentApp.GlyphType.NUMBER)
            .setNestingLevel(0)
            .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY); 
        
        listItem.setAttributes({[DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,[DocumentApp.Attribute.INDENT_START]: 18 
        });
      });
      try { paragraph.removeFromParent(); } catch (e) { element.setText(""); }
    } else {
      element.setText("-");
    }
  }
}
