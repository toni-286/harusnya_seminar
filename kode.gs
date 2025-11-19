/**
 * =================================================================
 * APLIKASI PENDAFTARAN SEMINAR + PRESENSI QCODE
 * =================================================================
 */

// GANTI DENGAN ID SPREADSHEET ANDA
const sheetId = "19SLtTVH4iuk3UtnOZshPTVhtl2Jd78gUS3JXfz14QBI";
const sheetName = "Lomba"; //sesuaikan dengan nama sheet nya

const TIMEZONE = Session.getScriptTimeZone();
const DATE_FORMAT = "dd MMMM yyyy, HH:mm:ss";

/**
 * Fungsi utama yang dimodifikasi untuk membaca tanggal dari sheet
 * dan mengirimkannya ke frontend.
 */
function doGet(e) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  // Baca tanggal dari H1 dan I1, konversi ke format ISO agar mudah dibaca JavaScript
  const startDate = new Date(sheet.getRange("H1").getValue()).toISOString();
  const endDate = new Date(sheet.getRange("I1").getValue());
  // Tambahkan satu hari ke tanggal akhir untuk membuat batas penutupan yang benar
  endDate.setDate(endDate.getDate() + 1);
  const deadlineDate = endDate.toISOString();

  // Kirim data tanggal ke template HTML
  const template = HtmlService.createTemplateFromFile('index');
  template.dates = {
    start: startDate,
    end: deadlineDate,
    displayEnd: sheet.getRange("I1").getDisplayValue() // Untuk ditampilkan ke pengguna
  };
  
  return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fungsi baru untuk input kehadiran manual.
 * Menggunakan logika yang sama dengan checkIn.
 * @param {string} phoneNumber - Nomor HP yang diinput manual.
 * @returns {object} - Objek JSON berisi status check-in.
 */
function manualCheckIn(phoneNumber) {
  let originalNoHp = phoneNumber.trim();
  let formattedNoHp = originalNoHp.startsWith('0') ? '62' + originalNoHp.substring(1) : originalNoHp;
  return checkIn(formattedNoHp); // Memanggil fungsi checkIn yang sudah ada
}


/**
 * Menyimpan data pendaftar baru, sekarang dengan tanggal dinamis.
 */
function saveData(data) {
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    
    // Baca tanggal dari H1 dan I1 setiap kali ada pendaftaran
    const startDate = new Date(sheet.getRange("H1").getValue());
    const endDate = new Date(sheet.getRange("I1").getValue());
    endDate.setDate(endDate.getDate() + 1); // Deadline adalah H+1 dari tanggal tutup
    const now = new Date();

    if (now < startDate) {
      return JSON.stringify({
        status: "CLOSED",
        message: "Pendaftaran belum dibuka."
      });
    }
    if (now >= endDate) {
      return JSON.stringify({
        status: "CLOSED",
        message: `Pendaftaran telah ditutup pada ${sheet.getRange("I1").getDisplayValue()}.`
      });
    }
    
    let originalNoHp = data.noHp.trim();
    let formattedNoHp = originalNoHp.startsWith('0') ? '62' + originalNoHp.substring(1) : originalNoHp;
    const dataRange = sheet.getDataRange();
    const values = dataRange.getDisplayValues();
    const noHpColumn = 3;

    for (let i = 1; i < values.length; i++) {
      if (values[i][noHpColumn] == formattedNoHp) {
        return JSON.stringify({ status: "EXISTS", message: "Nomor HP ini sudah terdaftar." });
      }
    }

    const timestamp = Utilities.formatDate(new Date(), TIMEZONE, DATE_FORMAT);
    const newRow = [timestamp, data.nama, data.alamat, formattedNoHp, data.instansi, ""];
    sheet.appendRow(newRow);

    data.noHp = formattedNoHp;
    return JSON.stringify({ status: "SUCCESS", data: data });

  } catch (error) {
    return JSON.stringify({ status: "ERROR", message: error.toString() });
  }
}

/**
 * Mencatat waktu kehadiran peserta dengan metode pencarian cepat menggunakan TextFinder.
 */
function checkIn(scannedPhoneNumber) {
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    
    const phoneColumnRange = sheet.getRange("D2:D" + sheet.getLastRow());
    
    const textFinder = phoneColumnRange.createTextFinder(scannedPhoneNumber).matchEntireCell(true);
    const foundCell = textFinder.findNext();

    if (foundCell) {
      const row = foundCell.getRow();
      const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
      const participantName = rowValues[1];
      const attendanceTime = rowValues[5];

      if (attendanceTime === "") {
        const checkInTime = Utilities.formatDate(new Date(), TIMEZONE, DATE_FORMAT);
        sheet.getRange(row, 6).setValue(checkInTime);
        return { status: "SUCCESS", name: participantName, time: checkInTime };
      } else {
        return { status: "INFO", name: participantName, time: attendanceTime };
      }
    } else {
      return { status: "GAGAL", message: `Nomor HP ${scannedPhoneNumber} tidak terdaftar.` };
    }
  } catch (error) {
    return { status: "ERROR", message: error.toString() };
  }
}

/**
 * Mengambil seluruh data peserta dan menghitung ringkasannya.
 */
function getParticipants() {
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    if (sheet.getLastRow() < 2) {
      return { data: [], summary: { total: 0, hadir: 0, belumHadir: 0 } };
    }
    
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const values = dataRange.getDisplayValues();

    let hadirCount = 0;
    const totalCount = values.length;
    const WaktuHadirColumn = 5;

    values.forEach(row => {
      if (row[WaktuHadirColumn] !== "") {
        hadirCount++;
      }
    });

    const belumHadirCount = totalCount - hadirCount;

    return {
      data: values,
      summary: { total: totalCount, hadir: hadirCount, belumHadir: belumHadirCount }
    };

  } catch (error) {
    throw new Error("Gagal mengambil data peserta: " + error.message);
  }
}

/**
 * Mencari data seorang peserta berdasarkan nomor HP untuk keperluan cetak ulang.
 */
function getParticipantByPhoneNumber(phoneNumber) {
  try {
    let originalPhoneNumber = phoneNumber.trim();
    let formattedPhoneNumber = originalPhoneNumber.startsWith('0') ? '62' + originalPhoneNumber.substring(1) : originalPhoneNumber;

    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getDisplayValues();

    for (let i = 1; i < values.length; i++) {
      if (values[i][3] == formattedPhoneNumber) {
        return {
          status: "FOUND",
          nama: values[i][1],
          noHp: values[i][3],
          instansi: values[i][4]
        };
      }
    }
    return { status: "NOT_FOUND" };
  } catch (error) {
    return { status: "ERROR", message: error.toString() };
  }
}
