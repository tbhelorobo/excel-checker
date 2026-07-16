/* ==========================================
   Telefon Düzenleyici
   Bölüm 1 - Başlangıç
========================================== */

const MAX_PREVIEW_ROWS = 20;
let workbook = null;
let outputWorkbook = null;
let originalRows = [];
let processedRows = [];
let phoneColumnIndex = -1;
let phoneColumnName = "";

const fileInput = document.getElementById("fileInput");
const convertBtn = document.getElementById("convertBtn");
const convertText = document.getElementById("convertText");
const downloadBtn = document.getElementById("downloadBtn");

const countrySelect = document.getElementById("countrySelect");
const countryCodeInput = document.getElementById("countryCode");

const previewBody = document.getElementById("previewBody");
const errorBody = document.getElementById("errorBody");

const progressFill = document.getElementById("progressFill");

const totalCount = document.getElementById("totalCount");
const validCount = document.getElementById("validCount");
const duplicateCount = document.getElementById("duplicateCount");
const invalidCount = document.getElementById("invalidCount");

const dropArea = document.getElementById("dropArea");
const uploadTitle = document.getElementById("uploadTitle");
const uploadText = document.getElementById("uploadText");

const phoneColumnSelect = document.getElementById("phoneColumnSelect");

const previewTitle = document.getElementById("previewTitle");
const previewInfo = document.getElementById("previewInfo");

const toast = document.getElementById("toast");
const toastMessage = document.getElementById("toastMessage");

let file = null;

/* ==========================================
   Ülke Kodu
========================================== */

countrySelect.addEventListener("change", () => {
  if (countrySelect.value === "manual") {
    countryCodeInput.disabled = false;

    countryCodeInput.value = "";

    countryCodeInput.focus();
  } else {
    countryCodeInput.disabled = true;

    countryCodeInput.value = countrySelect.value;
  }
});

countryCodeInput.value = countrySelect.value;
countryCodeInput.disabled = true;

/* ==========================================
   Drag Drop
========================================== */

dropArea.addEventListener("dragover", function (e) {
  e.preventDefault();

  dropArea.style.borderColor = "#22c55e";
  dropArea.style.background = "#162235";
});

dropArea.addEventListener("dragleave", function () {
  dropArea.style.borderColor = "#475569";
  dropArea.style.background = "";
});

dropArea.addEventListener("drop", function (e) {
  e.preventDefault();

  dropArea.style.borderColor = "#475569";
  dropArea.style.background = "";

  if (e.dataTransfer.files.length === 0) return;

  file = e.dataTransfer.files[0];

  fileInput.files = e.dataTransfer.files;

  loadExcel();
});

fileInput.addEventListener("change", function () {
  if (!this.files.length) return;

  file = this.files[0];

  loadExcel();
});

/* ==========================================
   Excel Oku
========================================== */

function loadExcel() {
  resetStats();
  clearTables();
  resetProgress();

  downloadBtn.disabled = true;

  processedRows = [];
  outputWorkbook = null;

  if (!file) return;

  uploadTitle.innerHTML = "📄 " + file.name;

  uploadText.innerHTML = Math.round(file.size / 1024) + " KB";

  dropArea.classList.add("success");

  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);

    workbook = XLSX.read(data, {
      type: "array",
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    originalRows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
    });

    findPhoneColumn();
    fillPhoneColumns();
  };

  reader.readAsArrayBuffer(file);
}

/* ==========================================
   Telefon Kolonu Doldur
========================================== */

function fillPhoneColumns() {
  phoneColumnSelect.innerHTML = "";

  const headers = originalRows[0];

  headers.forEach((header, index) => {
    const option = document.createElement("option");

    option.value = index;

    option.innerText = header;

    phoneColumnSelect.appendChild(option);
  });

  // Otomatik bulunan kolon varsa onu seç

  if (phoneColumnIndex != -1) {
    phoneColumnSelect.value = phoneColumnIndex;
  }
}

phoneColumnSelect.addEventListener("change", () => {
  phoneColumnIndex = parseInt(phoneColumnSelect.value);
});

/* ==========================================
   Telefon Kolonu Bul
========================================== */

function findPhoneColumn() {
  const headers = originalRows[0];

  const keywords = [
    "telefon",
    "telefon no",
    "tel",
    "gsm",
    "mobile",
    "phone",
    "cep",
    "müş",
    "müş. tel no",
    "customer phone",
  ];

  phoneColumnIndex = -1;

  headers.forEach((header, index) => {
    const value = String(header).toLowerCase();

    keywords.forEach((keyword) => {
      if (value.includes(keyword)) phoneColumnIndex = index;
    });
  });

  if (phoneColumnIndex === -1) {
    phoneColumnIndex = 0;
  }

  phoneColumnName = headers[phoneColumnIndex];

  console.log("Telefon Kolonu:", phoneColumnName);
}

/* ==========================================
   Progress
========================================== */

function setProgress(percent) {
  progressFill.style.width = percent + "%";
}

/* ==========================================
   Sayaçlar
========================================== */

function resetStats() {
  totalCount.innerText = "0";
  validCount.innerText = "0";
  duplicateCount.innerText = "0";
  invalidCount.innerText = "0";
}

/* ==========================================
   Önizleme Temizle
========================================== */

function clearTables() {
  previewBody.innerHTML = "";

  errorBody.innerHTML = "";

  previewTitle.innerText = "Önizleme";

  previewInfo.innerText = "Henüz işlem yapılmadı.";
}

/* ==========================================
   Dönüştür Butonu
========================================== */
function startConvert() {
  convertBtn.disabled = true;

  convertBtn.classList.remove("success");

  convertBtn.classList.add("loading");

  convertText.innerHTML = "⏳ İşleniyor...";
}

function finishConvert() {
  convertBtn.classList.remove("loading");

  convertBtn.classList.add("success");

  convertText.innerHTML = "✅ Tamamlandı";

  setTimeout(() => {
    convertBtn.disabled = false;

    convertBtn.classList.remove("success");

    convertText.innerHTML = "🚀 Dönüştür";
  }, 1200);
}

function convertError() {
  convertBtn.disabled = false;

  convertBtn.classList.remove("loading");

  convertText.innerHTML = "❌ Hata Oluştu";

  setTimeout(() => {
    convertText.innerHTML = "🚀 Dönüştür";
  }, 2000);
}

convertBtn.addEventListener("click", () => {
  if (!workbook) {
    alert("Lütfen önce Excel dosyası seçiniz.");

    return;
  }

  startConvert();

  resetStats();

  clearTables();

  processedRows = [];

  resetDuplicateList();

  try {
    processExcel();

    finishConvert();
  } catch (e) {
    console.error(e);

    convertError();
  }
});

/* ==========================================
   Telefon Düzenleyici
   Bölüm 2 - Telefon İşlemleri
========================================== */

/* ==========================================
   Telefonu Normalize Et
========================================== */

function normalizePhone(phone) {
  if (phone === undefined || phone === null) return "";

  let value = String(phone).trim();

  // Boş ise
  if (value === "") return "";

  // Tüm rakam olmayan karakterleri sil
  value = value.replace(/\D/g, "");

  // Baştaki 00 kaldır
  if (value.startsWith("00")) {
    value = value.substring(2);
  }

  const countryCode = countryCodeInput.value.trim();

  // Mevcut ülke kodunu koru seçili mi?
  const keepCountryCode = document.getElementById(
    "keepExistingCountryCode",
  ).checked;

  if (keepCountryCode) {
    // Zaten ülke koduyla başlıyorsa olduğu gibi bırak
    if (value.startsWith(countryCode)) {
      return value;
    }
  }

  // Baştaki sıfırları kaldır
  while (value.startsWith("0")) {
    value = value.substring(1);
  }

  // Ülke kodu ekle
  value = countryCode + value;

  return value;
}

/* ==========================================
   Telefon Geçerli mi?
========================================== */

function isValidPhone(phone) {
  const cc = countryCodeInput.value.trim();

  if (phone === "") return false;

  // Genel uzunluk kontrolü
  if (phone.length < cc.length + 8) return false;

  if (phone.length > 15) return false;

  // Türkiye için ekstra kontrol

  if (cc == "90") {
    // Mevcut ülke kodlarını koru açıksa
    // yabancı numaraları geçerli kabul et.

    if (
      document.getElementById("keepExistingCountryCode").checked &&
      isForeignNumber(phone)
    ) {
      return true;
    }

    return /^905[0-9]{9}$/.test(phone);
  }

  // Diğer ülkeler için
  return true;
}

function isForeignNumber(phone) {
  const cc = countryCodeInput.value.trim();

  if (phone.startsWith(cc)) return false;

  const knownCodes = ["90", "49", "44", "33", "39", "34", "1"];

  for (let code of knownCodes) {
    if (code == cc) continue;

    if (phone.startsWith(code)) return true;
  }

  return false;
}

/* ==========================================
   Sebep Döndür
========================================== */

function invalidReason(phone) {
  const cc = countryCodeInput.value.trim();

  if (phone === "") return "Telefon boş";

  if (phone.length < cc.length + 8) return "Numara çok kısa";

  if (phone.length > 15) return "Numara çok uzun";

  if (cc === "90") {
    if (!phone.startsWith("90")) return "Ülke kodu hatalı";

    if (phone.length !== 12) return "12 haneli olmalı";

    if (phone.charAt(2) !== "5") return "5 ile başlamalı";
  }

  return "Geçersiz";
}

/* ==========================================
   Duplicate Kontrol
========================================== */

const duplicateSet = new Set();

function isDuplicate(phone) {
  if (duplicateSet.has(phone)) return true;

  duplicateSet.add(phone);

  return false;
}

/* ==========================================
   Önizlemeye Satır Ekle
========================================== */

function addPreview(oldPhone, newPhone, status) {
  if (previewBody.rows.length >= MAX_PREVIEW_ROWS) return;

  const row = previewBody.insertRow();

  row.insertCell().innerText = oldPhone;

  row.insertCell().innerText = newPhone;

  const statusCell = row.insertCell();

  if (status === "Geçerli") {
    statusCell.innerHTML = '<span class="badge success">✅ Geçerli</span>';
  } else if (status === "Duplicate") {
    statusCell.innerHTML = '<span class="badge warning">🟡 Tekrarlı</span>';
  } else if (status === "Yabancı") {
    statusCell.innerHTML = '<span class="badge info">🌍 Yabancı Ülke</span>';
  } else {
    statusCell.innerHTML = '<span class="badge error">❌ Geçersiz</span>';
  }
}

/* ==========================================
   Hata Listesine Ekle
========================================== */

function addError(rowNumber, phone, reason) {
  const row = errorBody.insertRow();

  row.insertCell().innerText = rowNumber;

  row.insertCell().innerText = phone;

  row.insertCell().innerText = reason;
}

/* ==========================================
   Duplicate Hafızasını Temizle
========================================== */

function resetDuplicateList() {
  duplicateSet.clear();
}

/* ==========================================
   Telefon Düzenleyici
   Bölüm 3 - Excel İşleme
========================================== */

function processExcel() {
  setProgress(0);

  const removeDuplicates = document.getElementById("removeDuplicates").checked;

  const removeInvalid = document.getElementById("removeInvalid").checked;

  processedRows = [];

  // Başlık satırını ekle
  processedRows.push([...originalRows[0]]);

  let total = originalRows.length - 1;
  let valid = 0;
  let duplicate = 0;
  let invalid = 0;

  totalCount.innerText = total;

  for (let i = 1; i < originalRows.length; i++) {
    let row = [...originalRows[i]];

    let oldPhone = row[phoneColumnIndex];

    let newPhone = normalizePhone(oldPhone);

    row[phoneColumnIndex] = newPhone;

    /* ---------------------------
           Geçersiz Numara
        ----------------------------*/

    if (!isValidPhone(newPhone)) {
      invalid++;

      addError(i + 1, oldPhone, invalidReason(newPhone));

      if (removeInvalid) {
        addPreview(oldPhone, newPhone, "Geçersiz");

        continue;
      }
    }

    const removeForeign = document.getElementById(
      "removeForeignNumbers",
    ).checked;

    if (removeForeign) {
      if (isForeignNumber(newPhone)) {
        addPreview(oldPhone, newPhone, "Yabancı");

        addError(i + 1, oldPhone, "Farklı ülke kodu");

        continue;
      }
    }

    /* ---------------------------
           Duplicate
        ----------------------------*/

    if (removeDuplicates) {
      if (isDuplicate(newPhone)) {
        duplicate++;

        addPreview(oldPhone, newPhone, "Duplicate");

        continue;
      }
    }

    valid++;

    addPreview(oldPhone, newPhone, "Geçerli");

    processedRows.push(row);

    /* ---------------------------
           Progress
        ----------------------------*/

    if (i % 20 === 0) {
      let percent = Math.floor((i / total) * 100);

      setProgress(percent);
    }
  }

  setProgress(100);

  validCount.innerText = valid;

  duplicateCount.innerText = duplicate;

  invalidCount.innerText = invalid;

  const previewCount = Math.min(valid, 20);

  previewTitle.innerText = "Önizleme (" + previewCount + " Kayıt)";

  previewInfo.innerText =
    processedRows.length -
    1 +
    " kayıt Excel'e yazılacak. Önizlemede ilk " +
    previewCount +
    " kayıt gösteriliyor.";

  createWorkbook();
}

/* ==========================================
   Workbook Oluştur
========================================== */

function createWorkbook() {
  const sheet = XLSX.utils.aoa_to_sheet(processedRows);

  outputWorkbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    outputWorkbook,

    sheet,

    "Duzenlenmis",
  );

  downloadBtn.disabled = false;
}

/* ==========================================
   İndir
========================================== */

/* ==========================================
   Telefon Düzenleyici
   Bölüm 4 - Final
========================================== */

/* ==========================================
   Gelişmiş Normalize
========================================== */

function normalizePhone(phone) {
  if (phone === undefined || phone === null) return "";

  let value = String(phone).trim();

  if (value === "") return "";

  value = value.replace(/\D/g, "");

  const cc = countryCodeInput.value.trim();

  // 0090
  if (value.startsWith("00")) value = value.substring(2);

  // zaten seçilen ülke koduyla başlıyor
  if (value.startsWith(cc)) return value;

  // başka ülke kodu olabilir
  // örnek:
  // 90555...
  // 49555...
  // 44555...

  if (value.length >= 11 && value.length <= 15) {
    const knownCodes = ["90", "49", "44", "33", "39", "34", "1"];

    for (let code of knownCodes) {
      if (code !== cc && value.startsWith(code)) {
        // mevcut ülke kodunu koru seçiliyse
        if (document.getElementById("keepExistingCountryCode").checked) {
          return value;
        }

        value = value.substring(code.length);

        break;
      }
    }
  }

  while (value.startsWith("0")) value = value.substring(1);

  return cc + value;
}

/* ==========================================
   Telefon Kolonu Manuel Seç
========================================== */

function askPhoneColumn() {
  const headers = originalRows[0];

  let list = "";

  headers.forEach((h, i) => {
    list += i + " : " + h + "\n";
  });

  let answer = prompt(
    "Telefon kolonu bulunamadı.\n\n" + "Kolon numarasını giriniz.\n\n" + list,
  );

  if (answer === null) return false;

  answer = parseInt(answer);

  if (isNaN(answer)) return false;

  phoneColumnIndex = answer;

  phoneColumnName = headers[answer];

  return true;
}

/* ==========================================
   Telefon Kolonu Kontrol
========================================== */

const oldFind = findPhoneColumn;

findPhoneColumn = function () {
  oldFind();

  if (phoneColumnIndex === -1) {
    if (!askPhoneColumn()) return;
  }
};

/* ==========================================
   Dosya Adı
========================================== */

function getFileName() {
  if (!file) return "telefonlar.xlsx";

  const name = file.name;

  const dot = name.lastIndexOf(".");

  if (dot === -1) return name + "_duzenlenmis.xlsx";

  return name.substring(0, dot) + "_duzenlenmis.xlsx";
}

/* ==========================================
   Download
========================================== */

downloadBtn.addEventListener("click", () => {
  if (!outputWorkbook) return;

  XLSX.writeFile(
    outputWorkbook,

    getFileName(),
  );
  showToast(getFileName() + " başarıyla indirildi.");
});

/* ==========================================
   Progress Reset
========================================== */

function resetProgress() {
  progressFill.style.width = "0%";
}

/* ==========================================
   Başlangıç
========================================== */

resetProgress();

resetStats();

clearTables();

console.log(
  "%cTelefon Düzenleyici Hazır",

  "color:#22c55e;font-size:18px;font-weight:bold;",
);

/* ==========================================
   Toast
========================================== */

function showToast(message) {
  toastMessage.innerHTML = message;

  toast.classList.add("show");

  setTimeout(() => {
    toast.classList.remove("show");
  }, 3000);
}
