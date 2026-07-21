// Data mapping untuk gejala agar pengguna melihat nama gejala lengkap
const gejalaMap = {
  G01: "Terdapat bercak merah atau pendarahan pada tubuh ikan",
  G02: "Muncul benang putih seperti kapas pada kulit atau tubuh ikan",
  G03: "Terdapat bintik-bintik putih pada kulit, sirip, atau insang",
  G04: "Nafsu makan ikan menurun atau tidak mau makan",
  G05: "Ikan tampak lemas dan pergerakannya lambat",
  G06: "Insang berwarna merah terang atau mengalami peradangan",
  G07: "Terdapat luka, borok, atau lecet pada tubuh ikan",
  G08: "Sirip ikan rusak, robek, atau geripis"
};

const penyakitInfo = {
  Aeromonas: {
    deskripsi: "Penyakit Aeromonas menyerang jaringan kulit, sisik, dan insang, menyebabkan ikan tampak berlendir dan mudah terluka.",
    penyebab: "Bakteri Aeromonas hydrophila yang berkembang pada kondisi air kotor dan stres.",
    penanganan: "Segera bersihkan air, berikan aerasi, ganti sebagian air, dan gunakan obat bakteri sesuai petunjuk.",
    pencegahan: "Jaga kualitas air, bersihkan kolam secara berkala, dan hindari kepadatan ikan berlebihan."
  },
  Saprolegniasis: {
    deskripsi: "Saprolegniasis adalah infeksi jamur yang muncul sebagai benang putih kapas pada kulit ikan.",
    penyebab: "Lingkungan air dingin, kualitas air buruk, dan ikan yang terluka mudah terinfeksi.",
    penanganan: "Pisahkan ikan sakit, gunakan larutan antiseptik atau garam, dan tingkatkan kualitas air.",
    pencegahan: "Hindari luka pada ikan, pertahankan kebersihan kolam, dan lakukan karantina untuk ikan baru."
  },
  Dropsy: {
    deskripsi: "Dropsy menyebabkan perut ikan membengkak akibat cairan menumpuk di tubuh ikan.",
    penyebab: "Infeksi bakteri, masalah organ internal, atau kondisi stres kronis.",
    penanganan: "Perbaiki kualitas air, berikan pakan bergizi, dan konsultasi dengan ahli jika perlu.",
    pencegahan: "Pertahankan kualitas air stabil, hindari perubahan mendadak, dan jaga sanitasi tempat budidaya."
  },
  "White Spot": {
    deskripsi: "White Spot ditandai bintik putih kecil pada kulit ikan dan sirip.",
    penyebab: "Infeksi parasit Ichthyophthirius multifiliis akibat stres dan air kurang bersih.",
    penanganan: "Lakukan pengobatan dengan garam, perlakuan suhu, dan perbaiki sirkulasi air.",
    pencegahan: "Karantina ikan baru, jaga kebersihan air, dan hindari fluktuasi suhu."
  },
  "Insang Merah": {
    deskripsi: "Insang Merah menandakan peradangan pada insang sehingga berwarna merah terang.",
    penyebab: "Bakteri atau kondisi air buruk yang memicu iritasi insang.",
    penanganan: "Segera ganti air, tingkatkan aerasi, dan jaga pH agar stabil.",
    pencegahan: "Pertahankan kualitas air, bersihkan kotoran, dan jangan biarkan limbah menumpuk."
  },
  Kanibalisme: {
    deskripsi: "Kanibalisme terjadi ketika ikan memakan anggota kelompoknya sendiri karena stres atau kelaparan.",
    penyebab: "Kepadatan berlebih, pakan tidak cukup, dan kondisi kolam buruk.",
    penanganan: "Kurangi kepadatan, beri pakan bergizi, dan pisahkan ikan agresif.",
    pencegahan: "Atur kepadatan ideal, penuhi kebutuhan pakan, dan perbaiki kualitas air."
  }
};

let dataset = [];

function normalizeValue(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value === 0 ? 0 : 1;

  const normalized = String(value).trim().toLowerCase();
  if (["ya", "y", "iya", "1", "true", "yes"].includes(normalized)) return 1;
  if (["tidak", "t", "0", "false", "no", "-"].includes(normalized)) return 0;
  return 0;
}

function normalizeDataset(rawRows) {
  return rawRows
    .map(row => {
      const normalized = {};
      Object.entries(row).forEach(([key, value]) => {
        if (key === "Penyakit" || key === "penyakit") {
          normalized.Penyakit = value;
          return;
        }
        if (/^G\d+$/i.test(key)) {
          normalized[key.toUpperCase()] = normalizeValue(value);
        }
      });
      return normalized;
    })
    .filter(row => row.Penyakit);
}

// Muat dataset dari file Excel yang baru
async function loadDataset() {
  if (dataset.length) return dataset;

  const datasetUrl = encodeURI("gejala & penyakit ikan lele.xlsx");

  try {
    const response = await fetch(datasetUrl);
    if (!response.ok) throw new Error(`Excel fetch gagal: ${response.status}`);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: 0 });
    dataset = normalizeDataset(rawData);
    return dataset;
  } catch (error) {
    console.error("Gagal memuat dataset:", error);
    dataset = [];
    return [];
  }
}

// Merender daftar gejala sebagai checkbox pada halaman diagnosis
function renderGejala() {
  const gejalaList = document.getElementById("gejalaList");
  if (!gejalaList) return;

  gejalaList.innerHTML = Object.entries(gejalaMap)
    .map(([kode, nama]) => {
      return `
        <div class="gejala-item">
          <input type="checkbox" id="${kode}" name="gejala" value="${kode}" />
          <label for="${kode}">${nama}</label>
        </div>
      `;
    })
    .join("");
}

// Menghitung persentase kecocokan untuk setiap penyakit
function hitungPersentase(matchData, selectedCount) {
  return matchData.map(item => {
    const averageMatch = selectedCount > 0 ? (item.totalMatches / item.recordCount / selectedCount) * 100 : 0;
    const percent = Math.round(averageMatch * 10) / 10;
    return { ...item, percent };
  });
}

// Menampilkan hasil diagnosis dan informasi penyakit
function tampilkanHasil(sortedResults) {
  const resultArea = document.getElementById("resultArea");
  if (!resultArea) return;

  if (sortedResults.length === 0) {
    resultArea.innerHTML = `<p class="muted-text">Silakan pilih minimal satu gejala sebelum menekan tombol Diagnosa.</p>`;
    return;
  }

  const primary = sortedResults[0];
  const info = penyakitInfo[primary.name] || {
    deskripsi: "Informasi penyakit tidak tersedia.",
    penyebab: "-",
    penanganan: "-",
    pencegahan: "-"
  };

  resultArea.innerHTML = `
    <div class="result-summary">
      <h3>Hasil Utama: ${primary.name}</h3>
      <p>Keyakinan: <strong>${primary.percent}%</strong></p>
    </div>
    <div class="result-section">
      <h4>Perkiraan penyakit lain</h4>
      <ul class="result-list">
        ${sortedResults.map(item => `
          <li class="result-item">
            <div class="progress-row">
              <span>${item.name}</span>
              <strong>${item.percent}%</strong>
            </div>
            <div class="progress-bar"><div class="progress-fill" style="width: ${item.percent}%;"></div></div>
          </li>
        `).join("")}
      </ul>
    </div>
    <div class="detail-block">
      <h4>Informasi Penyakit: ${primary.name}</h4>
      <p><strong>Deskripsi:</strong> ${info.deskripsi}</p>
      <p><strong>Penyebab:</strong> ${info.penyebab}</p>
      <p><strong>Penanganan awal:</strong> ${info.penanganan}</p>
      <p><strong>Pencegahan:</strong> ${info.pencegahan}</p>
    </div>
  `;
}

// Menjalankan proses diagnosis berdasarkan gejala yang dipilih
async function diagnosa() {
  const checked = Array.from(document.querySelectorAll('input[name="gejala"]:checked'))
    .map(input => input.value);

  const resultArea = document.getElementById("resultArea");
  if (!resultArea) return;

  if (checked.length === 0) {
    resultArea.innerHTML = `<p class="muted-text">Pilih minimal satu gejala sebelum melakukan diagnosis.</p>`;
    return;
  }

  if (!dataset.length) {
    await loadDataset();
  }

  const diseaseMap = {};

  dataset.forEach(row => {
    const penyakit = row.Penyakit || row.penyakit || "Tidak Diketahui";
    const matchCount = checked.reduce((count, kode) => {
      return count + (Number(row[kode]) === 1 ? 1 : 0);
    }, 0);

    if (!diseaseMap[penyakit]) {
      diseaseMap[penyakit] = { totalMatches: 0, recordCount: 0, name: penyakit };
    }
    diseaseMap[penyakit].totalMatches += matchCount;
    diseaseMap[penyakit].recordCount += 1;
  });

  const matchData = Object.values(diseaseMap);
  const sortedResults = hitungPersentase(matchData, checked.length)
    .sort((a, b) => b.percent - a.percent);

  tampilkanHasil(sortedResults);
}

// Mengembalikan formulir ke kondisi awal
function resetForm() {
  const inputs = document.querySelectorAll('input[name="gejala"]');
  inputs.forEach(el => {
    el.checked = false;
  });
  const resultArea = document.getElementById("resultArea");
  if (resultArea) {
    resultArea.innerHTML = `<p class="muted-text">Hasil diagnosis akan muncul di sini setelah Anda memilih gejala.</p>`;
  }
}

// Inisialisasi halaman ketika dokumen siap
document.addEventListener("DOMContentLoaded", async () => {
  renderGejala();

  const diagnosaBtn = document.getElementById("diagnosaBtn");
  const resetBtn = document.getElementById("resetBtn");

  if (diagnosaBtn) {
    diagnosaBtn.addEventListener("click", diagnosa);
  }

  if (resetBtn) {
    resetBtn.addEventListener("click", resetForm);
  }

  // Muat dataset lebih awal agar proses diagnosis lebih cepat
  if (document.getElementById("gejalaList")) {
    await loadDataset();
  }
});
