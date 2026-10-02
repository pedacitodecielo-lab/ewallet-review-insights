const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, LevelFormat, Footer, PageNumber, TableOfContents, PageBreak,
} = require("docx");
const NOTES = require("./notes.js");

const INK = "0E1B2C", MINT = "17A589", GOLD = "B7860F", MUTED = "5B6B7C", SOFT = "EEF4F2", GOLDSOFT = "FCF1DA";
const FONT = "Calibri", HFONT = "Cambria";
const CONTENT_W = 9026; // A4 minus 1" margins, in DXA

const numbering = {
  config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] },
  ],
};
let numInstance = 0;

function runs(text, base = {}) {
  // **bold** markup inside strings
  return String(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((t) =>
    t.startsWith("**") ? new TextRun({ text: t.slice(2, -2), bold: true, ...base }) : new TextRun({ text: t, ...base }));
}
const P = (t, o = {}) => new Paragraph({ children: runs(t, o.run), spacing: { after: 140, line: 300 }, ...o.para });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)], spacing: { before: 360, after: 160 }, pageBreakBefore: false });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)], spacing: { before: 260, after: 120 } });
const BUL = (items) => items.map((t) => new Paragraph({ numbering: { reference: "bul", level: 0 }, children: runs(t), spacing: { after: 80, line: 290 } }));
const NUM = (items) => { const inst = numInstance++; return items.map((t) => new Paragraph({ numbering: { reference: "num", level: 0, instance: inst }, children: runs(t), spacing: { after: 80, line: 290 } })); };
const BOX = (title, lines, fill = SOFT) => {
  const border = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill },
      borders: { top: border, bottom: border, left: border, right: border },
      margins: { top: 140, bottom: 140, left: 220, right: 220 },
      children: [
        new Paragraph({ children: [new TextRun({ text: title, bold: true, color: fill === GOLDSOFT ? GOLD : MINT })], spacing: { after: 80 } }),
        ...lines.map((l) => new Paragraph({ children: runs(l), spacing: { after: 60, line: 290 } })),
      ],
    })] })],
  });
};
const GAP = () => new Paragraph({ children: [], spacing: { after: 120 } });
function TABLE(header, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: "D5E0DC" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cell = (t, i, head, zebra) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA }, borders,
    shading: { type: ShadingType.CLEAR, color: "auto", fill: head ? INK : (zebra ? "F6F9F8" : "FFFFFF") },
    margins: { top: 80, bottom: 80, left: 110, right: 110 },
    children: [new Paragraph({ children: runs(t, head ? { bold: true, color: "FFFFFF" } : (i === 0 ? { bold: true } : {})), spacing: { after: 0, line: 270 } })],
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true)) }),
      ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, i, false, ri % 2 === 1)) }))],
  });
}

const styles = {
  default: { document: { run: { font: FONT, size: 22, color: INK } } },
  paragraphStyles: [
    { id: "Title", name: "Title", basedOn: "Normal", run: { font: HFONT, size: 52, bold: true, color: INK }, paragraph: { spacing: { after: 120 } } },
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: HFONT, size: 32, bold: true, color: INK }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 25, bold: true, color: MINT }, paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
  ],
};
const footer = (label) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
  new TextRun({ text: label + "  |  Halaman ", color: MUTED, size: 18 }), new TextRun({ children: [PageNumber.CURRENT], color: MUTED, size: 18 }),
] })] });
const page = { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } };

function cover(title, sub, meta) {
  return [
    new Paragraph({ children: [new TextRun({ text: "MATERI REGENERASI DIVISI", bold: true, color: GOLD, size: 20, characterSpacing: 40 })], spacing: { before: 1800, after: 200 } }),
    new Paragraph({ style: "Title", children: [new TextRun(title)] }),
    new Paragraph({ children: [new TextRun({ text: sub, size: 28, color: MUTED })], spacing: { after: 600 } }),
    ...meta.map((m) => new Paragraph({ children: runs(m, { color: MUTED }), spacing: { after: 80 } })),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ================= MODUL =================
const modul = [
  ...cover("Modul Belajar Reksa Dana Pasar Uang", "Pegangan untuk penerus divisi: dibaca sebelum, selama, dan sesudah sesi materi.", [
    "**Data per:** Oktober 2026",
    "**Pendamping:** file presentasi Reksa_Dana_Pasar_Uang.pptx, Naskah_Presentasi_RDPU.docx, dan Simulator_RDPU.html",
    "**Waktu baca:** sekitar 45 menit, ditambah 20 menit latihan soal",
  ]),
  H1("Daftar isi"),
  new TableOfContents("Daftar isi", { hyperlink: true, headingStyleRange: "1-1" }),
  P("Jika daftar isi tampak kosong, klik kanan di area ini lalu pilih Update Field.", { run: { italics: true, color: MUTED, size: 18 } }),
  new Paragraph({ children: [new PageBreak()] }),

  H1("1. Cara memakai modul ini"),
  P("Modul ini ditulis untuk kamu yang akan meneruskan divisi dan suatu hari membawakan materi yang sama ke adik tingkat. Jadi targetnya ada dua: kamu paham, dan kamu bisa menjelaskan ulang."),
  P("Saran urutan belajarnya:"),
  ...NUM([
    "Baca bagian 2 sampai 6 untuk konsep dasar. Jangan lompati bagian NAB, karena hampir semua pertanyaan peserta nanti berujung ke sana.",
    "Baca bagian 7 sampai 11 untuk untung, pajak, risiko, dan praktik.",
    "Kerjakan latihan soal di bagian 13 tanpa melihat kunci jawaban.",
    "Coba Simulator_RDPU.html dan ubah angkanya sendiri. Ini cara paling cepat untuk merasakan efek pajak dan suku bunga.",
    "Sebelum membawakan materi, cek bagian 16 untuk daftar angka yang wajib diperbarui.",
  ]),
  H2("Tujuan belajar"),
  P("Setelah menyelesaikan modul ini, kamu bisa:"),
  ...BUL([
    "Menjelaskan apa itu RDPU dan bedanya dengan jenis reksa dana lain.",
    "Menghitung jumlah unit, nilai investasi, dan keuntungan dari NAB/UP.",
    "Menjelaskan perlakuan pajak dan biaya RDPU dengan benar.",
    "Menyebutkan risiko RDPU tanpa melebih-lebihkan atau meremehkannya.",
    "Memilih produk RDPU dengan kriteria yang masuk akal dan mengenali tawaran investasi bodong.",
  ]),

  H1("2. Kenapa kita perlu paham"),
  P("Uang yang disimpan di dompet atau di rekening dengan bunga sangat kecil memang tidak berkurang nominalnya. Masalahnya, harga barang terus naik. Bank Indonesia menargetkan inflasi 2,5% plus minus 1% untuk 2026 dan 2027. Kalau uangmu tumbuh di bawah angka itu, daya belinya turun pelan-pelan."),
  P("Contoh sederhana: seporsi nasi goreng Rp15.000 hari ini bisa jadi Rp15.500 tahun depan. Uang Rp15.000 yang hanya disimpan tidak lagi cukup untuk porsi yang sama."),
  H2("Angka acuan per Oktober 2026"),
  TABLE(["Indikator", "Angka", "Keterangan"], [
    ["BI Rate", "5,75%", "Ditahan pada RDG 22 sampai 23 September 2026, setelah naik total 100 bps di Mei sampai Juni 2026."],
    ["Tingkat bunga penjaminan LPS", "3,75%", "Simpanan rupiah di bank umum, periode 1 Juli sampai 30 September 2026. Cek periode terbaru di lps.go.id."],
    ["Target inflasi BI", "2,5% ± 1%", "Untuk 2026 dan 2027."],
    ["Investor pasar modal", "30,06 juta SID", "Per akhir Juli 2026 (KSEI)."],
    ["Investor reksa dana", "19,19 juta SID", "Per akhir 2025, naik 36,77% dari 2024 (KSEI)."],
  ], [2600, 1700, 4726]),
  GAP(),
  P("Data KSEI juga menunjukkan lebih dari separuh investor individu berusia 30 tahun ke bawah. Banyak dari mereka masuk lewat aplikasi, dan RDPU sering jadi produk pertama yang dibeli. Sayangnya, tidak sedikit yang membeli tanpa paham isinya."),

  H1("3. Mengenal reksa dana dan RDPU"),
  H2("Reksa dana dengan analogi patungan"),
  P("Bayangkan satu angkatan ingin makan enak tapi tidak ada yang bisa masak. Semua patungan, lalu menyewa koki profesional. Koki yang memilih bahan dan cara masak, dan hasilnya dibagi sesuai besar iuran masing-masing."),
  P("Reksa dana bekerja dengan cara yang sama. Investor adalah peserta patungan. Kokinya adalah **Manajer Investasi**, perusahaan berizin OJK. Bahannya adalah efek seperti deposito, surat utang negara, obligasi, atau saham. Secara hukum, reksa dana di Indonesia umumnya berbentuk Kontrak Investasi Kolektif (KIK) sesuai UU No. 8 Tahun 1995 tentang Pasar Modal."),
  H2("Definisi resmi RDPU"),
  BOX("Definisi", [
    "Reksa Dana Pasar Uang adalah reksa dana yang **100% asetnya** ditempatkan pada instrumen pasar uang dalam negeri dan/atau efek bersifat utang yang **jatuh temponya tidak lebih dari 1 tahun**, atau **sisa jatuh temponya tidak lebih dari 1 tahun**.",
    "Dasar aturan: POJK 23/POJK.04/2016 tentang Reksa Dana berbentuk KIK beserta perubahannya.",
  ]),
  GAP(),
  P("Batas satu tahun ini yang membuat RDPU stabil. Harga surat utang bertenor pendek hampir tidak terpengaruh perubahan suku bunga, berbeda dengan obligasi bertenor panjang."),
  H2("Isi portofolio RDPU"),
  ...BUL([
    "**Deposito bank.** Biasanya porsi terbesar. Bunganya dikenai PPh final 20%.",
    "**Surat utang negara jangka pendek.** Misalnya Surat Perbendaharaan Negara (SPN) dan sukuk negara bertenor pendek.",
    "**Obligasi atau sukuk korporasi** yang jatuh tempo, atau sisa jatuh temponya, maksimal satu tahun.",
  ]),
  P("Komposisi persisnya berbeda di setiap produk dan bisa dicek di fund fact sheet bulanan."),
  H2("RDPU syariah"),
  P("RDPU syariah menempatkan dana di deposito syariah dan sukuk, serta diawasi Dewan Pengawas Syariah. Cara kerja, cara menghitung NAB, dan jenis risikonya sama dengan RDPU konvensional."),

  H1("4. Posisi RDPU di antara jenis reksa dana"),
  TABLE(["Jenis", "Isi utama", "Tingkat risiko", "Horizon umum"], [
    ["Pasar Uang", "Deposito dan surat utang ≤ 1 tahun", "Paling rendah", "Di bawah 1 tahun"],
    ["Pendapatan Tetap", "Minimal 80% efek utang (obligasi)", "Rendah sampai sedang", "1 sampai 3 tahun"],
    ["Campuran", "Kombinasi saham, obligasi, pasar uang", "Sedang", "3 sampai 5 tahun"],
    ["Saham", "Minimal 80% saham", "Paling tinggi", "Di atas 5 tahun"],
  ], [1900, 3000, 2000, 2126]),
  GAP(),
  P("Horizon waktu di tabel adalah panduan umum yang lazim dipakai industri, bukan aturan baku. Semakin tinggi jenisnya, potensi hasil semakin besar tetapi naik turunnya juga semakin tajam."),
  H2("Seberapa besar RDPU di Indonesia"),
  P("Menurut KSEI (data S-INVEST, akhir Juli 2026), dana kelolaan RDPU sekitar **Rp133,58 triliun** dari **197 produk**. Sebagai perbandingan, reksa dana pendapatan tetap Rp217,02 triliun, terproteksi Rp144,84 triliun, saham Rp62,69 triliun, dan campuran Rp34,83 triliun."),

  H1("5. Pihak yang terlibat"),
  TABLE(["Pihak", "Peran"], [
    ["OJK", "Regulator. Memberi izin dan mengawasi seluruh pihak di industri reksa dana."],
    ["Manajer Investasi (MI)", "Mengelola portofolio sesuai kebijakan investasi di prospektus. Tidak memegang uang investor."],
    ["Bank Kustodian", "Menyimpan aset reksa dana, mengadministrasikan, dan menghitung NAB setiap hari bursa."],
    ["APERD", "Agen Penjual Efek Reksa Dana: bank atau aplikasi investasi tempat investor bertransaksi."],
    ["KSEI (S-INVEST)", "Mencatat kepemilikan unit atas nama SID investor dan mengirim konfirmasi."],
    ["Investor", "Pemilik unit penyertaan."],
  ], [2600, 6426]),
  GAP(),
  BOX("Kenapa ini penting", [
    "Uang investor tidak masuk ke rekening Manajer Investasi. Dana masuk ke rekening reksa dana di Bank Kustodian, dan asetnya dicatat terpisah dari harta MI. Kalau MI bermasalah, aset reksa dana bukan bagian dari harta perusahaan itu.",
    "Konsekuensinya: kalau ada yang meminta transfer ke rekening pribadi, itu pasti bukan transaksi reksa dana yang benar.",
  ]),

  H1("6. Cara kerja: NAB dan unit penyertaan"),
  H2("Dua istilah wajib"),
  ...BUL([
    "**NAB (Nilai Aktiva Bersih):** total nilai aset reksa dana dikurangi kewajibannya, misalnya biaya yang belum dibayar.",
    "**Unit Penyertaan (UP):** satuan kepemilikan investor. Harga per unitnya disebut **NAB/UP** dan diumumkan setiap hari bursa.",
  ]),
  BOX("Rumus", [
    "NAB/UP = NAB total ÷ jumlah unit beredar",
    "Jumlah unit yang didapat = nominal pembelian ÷ NAB/UP saat beli",
    "Nilai investasi = jumlah unit × NAB/UP terbaru",
  ]),
  GAP(),
  H2("Contoh hitung"),
  TABLE(["Langkah", "Perhitungan", "Hasil"], [
    ["Beli", "Rp1.000.000 ÷ Rp1.500", "666,67 unit"],
    ["Setahun kemudian", "NAB/UP naik 4,5%: 1.500 × 1,045", "Rp1.567,50"],
    ["Nilai investasi", "666,67 × Rp1.567,50", "Rp1.045.000"],
    ["Keuntungan", "Rp1.045.000 − Rp1.000.000", "Rp45.000"],
  ], [2200, 4226, 2600]),
  GAP(),
  P("Hal yang paling sering bikin bingung: jumlah unit tidak berubah selama kamu tidak membeli atau menjual. Yang bergerak adalah NAB/UP. Jadi saldo di aplikasi bisa naik walaupun kamu tidak menambah dana."),
  H2("Kenapa grafik RDPU naik seperti tangga"),
  P("Bunga deposito dan kupon surat utang diakui sedikit demi sedikit setiap hari, sehingga NAB/UP RDPU naik pelan dan mulus. Grafik reksa dana saham terlihat jauh lebih bergerigi karena harga saham berubah setiap hari."),
  P("Mulus bukan berarti tidak bisa turun. Kalau ada penerbit surat utang di dalam portofolio yang gagal bayar, NAB/UP bisa turun. Kejadiannya jarang, tetapi bukan mustahil."),
  H2("Alur transaksi"),
  ...NUM([
    "**Pilih produk.** Baca fund fact sheet dan prospektus.",
    "**Beli (subscription).** Ada batas waktu harian, umumnya pukul 13.00 WIB. Order sebelum batas memakai NAB/UP hari itu. Setelahnya memakai NAB/UP hari bursa berikutnya.",
    "**Unit tercatat** atas nama SID kamu di KSEI. Kamu menerima konfirmasi transaksi.",
    "**Pantau.** Nilai = jumlah unit × NAB/UP terbaru.",
    "**Jual (redemption).** Aturan OJK mewajibkan pembayaran paling lambat 7 hari bursa sejak permintaan jual yang lengkap diterima. Untuk RDPU, dalam praktiknya dana umumnya masuk dalam 1 sampai 2 hari bursa. Hari libur bursa tidak dihitung.",
  ]),

  H1("7. Untung, biaya, dan pajak"),
  H2("Sumber keuntungan"),
  P("Keuntungan RDPU berasal dari bunga deposito dan kupon surat utang di dalam portofolio. Investor tidak menerima transfer bunga bulanan seperti deposito. Keuntungan tercermin dari kenaikan NAB/UP. Per pertengahan 2026, return satu tahun RDPU umumnya berada di kisaran 4 sampai 5% per tahun, tergantung produk. Angka ini bisa berubah mengikuti suku bunga."),
  H2("Biaya"),
  ...BUL([
    "**Biaya pengelolaan MI dan biaya kustodian** dibebankan ke reksa dana dan sudah dipotong dari NAB. Return yang kamu lihat di aplikasi sudah bersih dari biaya ini.",
    "**Biaya pembelian dan penjualan** RDPU di banyak APERD umumnya Rp0. Tetap cek prospektus produknya.",
    "Perhatikan biaya transfer bank atau biaya platform jika ada.",
  ]),
  H2("Pajak"),
  BOX("Ringkasnya", [
    "Keuntungan yang diterima pemegang unit reksa dana **bukan objek pajak** (UU Pajak Penghasilan Pasal 4 ayat 3). Kamu tidak perlu menyetor pajak lagi atas keuntungan RDPU.",
    "Pajak sudah dipotong di dalam portofolio: bunga deposito dikenai PPh final 20% (PP 131/2000) dan bunga obligasi dikenai PPh final 10% (PP 91/2021).",
  ], GOLDSOFT),
  GAP(),
  H2("Simulasi Rp10 juta selama satu tahun"),
  TABLE(["Produk", "Asumsi", "Perhitungan", "Untung bersih"], [
    ["Tabungan", "0,5% bersih", "10.000.000 × 0,5%", "Rp50.000"],
    ["Deposito", "3,75% sebelum pajak", "10.000.000 × 3,75% = 375.000, dipotong 20%", "Rp300.000"],
    ["RDPU", "4,5% bersih", "10.000.000 × 4,5%", "Rp450.000"],
  ], [1700, 2000, 3326, 2000]),
  GAP(),
  P("Ini ilustrasi dengan asumsi tetap. Return RDPU tidak dijamin dan bisa lebih tinggi atau lebih rendah dari angka di atas.", { run: { italics: true, color: MUTED } }),

  H1("8. RDPU dibanding tabungan dan deposito"),
  TABLE(["Aspek", "Tabungan", "Deposito", "RDPU"], [
    ["Modal awal", "Kecil, sesuai syarat bank", "Umumnya jutaan rupiah", "Mulai sekitar Rp10.000 di banyak aplikasi"],
    ["Imbal hasil", "Rendah, sering di bawah 1% per tahun", "Bunga tetap selama tenor", "Mengikuti isi portofolio, tidak dijamin"],
    ["Pajak", "Bunga dikenai 20% final jika saldo di atas Rp7,5 juta", "Bunga dikenai 20% final", "Keuntungan bukan objek pajak bagi investor"],
    ["Pencairan", "Kapan saja", "Saat jatuh tempo, cair lebih awal biasanya kena penalti", "Kapan saja, dana umumnya masuk 1 sampai 2 hari bursa"],
    ["Penjaminan LPS", "Ya, sampai Rp2 miliar per nasabah per bank, selama bunga tidak melebihi batas LPS", "Ya, dengan syarat yang sama", "Tidak. Dilindungi lewat pengawasan OJK dan pemisahan aset di Bank Kustodian"],
    ["Cocok untuk", "Transaksi harian", "Dana yang pasti tidak dipakai selama tenor", "Dana darurat dan tujuan di bawah 1 tahun"],
  ], [1700, 2300, 2400, 2626]),
  GAP(),
  P("Kesimpulannya, RDPU bukan pengganti tabungan. Tabungan tetap dibutuhkan untuk transaksi harian, deposito cocok untuk dana yang pasti tidak dipakai, dan RDPU mengisi celah di antara keduanya."),

  H1("9. Risiko RDPU"),
  P("Risikonya rendah, tetapi bukan nol. Enam risiko yang perlu diketahui:"),
  ...NUM([
    "**Imbal hasil tidak pasti.** Saat suku bunga turun, return RDPU ikut turun secara bertahap.",
    "**Gagal bayar penerbit.** Jika penerbit obligasi di portofolio gagal bayar, NAB/UP bisa turun.",
    "**Risiko likuiditas.** Jika banyak investor menjual dalam jumlah besar secara bersamaan, MI harus menjual aset lebih cepat.",
    "**Tidak dijamin LPS.** Berbeda dengan tabungan dan deposito.",
    "**Pembubaran.** Reksa dana wajib dibubarkan jika NAB-nya di bawah Rp10 miliar selama 120 hari bursa berturut-turut (POJK 23/POJK.04/2016). Dana investor tidak hilang. Aset dilikuidasi dan hasilnya dibagikan sesuai porsi unit.",
    "**Perubahan peraturan,** misalnya perubahan tarif pajak yang memengaruhi hasil bersih.",
  ]),

  H1("10. Konteks suku bunga 2024 sampai 2026"),
  TABLE(["Periode", "BI Rate", "Peristiwa"], [
    ["Januari 2024", "6,00%", "Posisi awal."],
    ["April 2024", "6,25%", "Naik 25 bps."],
    ["September 2024", "6,00%", "Turun 25 bps."],
    ["Januari sampai September 2025", "5,75% ke 4,75%", "Turun lima kali (Januari, Mei, Juli, Agustus, September), masing-masing 25 bps."],
    ["Mei sampai Juni 2026", "4,75% ke 5,75%", "Naik total 100 bps dalam dua bulan untuk menjaga stabilitas rupiah."],
    ["Juli sampai September 2026", "5,75%", "Ditahan. Terakhir pada RDG 22 sampai 23 September 2026."],
  ], [2800, 1900, 4326]),
  GAP(),
  P("**Apa artinya buat RDPU?** Isi RDPU jatuh temponya pendek. Setiap kali deposito atau surat utangnya jatuh tempo, MI menempatkan ulang dana ke instrumen baru dengan bunga yang berlaku saat itu. Ketika suku bunga naik seperti di pertengahan 2026, return RDPU cenderung ikut naik dengan jeda waktu. Sebaliknya, sepanjang 2025 ketika suku bunga turun, return RDPU ikut melandai."),
  P("Reksa dana pendapatan tetap bereaksi berbeda. Saat suku bunga naik, harga obligasi bertenor panjang turun, sehingga NAB/UP-nya bisa tertekan. Pelajarannya: tidak ada instrumen yang selalu paling unggul. Kondisi suku bunga mengubah siapa yang diuntungkan."),

  H1("11. Praktik: memilih, waspada, dan memulai"),
  H2("Enam hal yang dicek sebelum memilih RDPU"),
  ...NUM([
    "**Legalitas.** MI dan APERD terdaftar serta diawasi OJK.",
    "**Dana kelolaan.** Pilih yang besar dan stabil, jauh di atas batas bubar Rp10 miliar.",
    "**Konsistensi kinerja.** Lihat return 1 tahun dan 3 tahun. Return yang stabil lebih penting daripada yang paling tinggi sesaat, karena return tinggi kadang datang dari surat utang yang lebih berisiko.",
    "**Isi portofolio.** Baca fund fact sheet: dananya ditempatkan di bank atau perusahaan mana saja.",
    "**Biaya dan pencairan.** Biaya pengelolaan, biaya beli dan jual, serta berapa hari dana cair.",
    "**Konvensional atau syariah,** sesuai preferensi.",
  ]),
  H2("Cara membaca fund fact sheet"),
  ...BUL([
    "**Tujuan dan kebijakan investasi:** memastikan produk benar-benar RDPU.",
    "**Kinerja:** return 1 bulan, 1 tahun, 3 tahun, dan sejak diluncurkan, beserta pembandingnya.",
    "**Efek terbesar dalam portofolio:** daftar bank atau penerbit tempat dana ditempatkan.",
    "**Biaya:** biaya pengelolaan dan kustodian per tahun.",
    "**Bank Kustodian dan total dana kelolaan.**",
  ]),
  H2("Waspada janji untung pasti"),
  BOX("Kasus nyata", [
    "Pada 21 November 2019, OJK melalui surat S-1422/PM.21/2019 memerintahkan PT Minna Padi Aset Manajemen membubarkan enam reksa dananya yang dipasarkan dengan janji imbal hasil pasti (fixed return). Total dana kelolaan keenam produk itu sekitar Rp6 triliun. OJK juga memberhentikan direktur utamanya.",
    "Pelajarannya: reksa dana tidak boleh menjanjikan imbal hasil pasti. Janji seperti itu adalah tanda bahaya.",
  ], GOLDSOFT),
  GAP(),
  P("Tanda bahaya lain: diminta transfer ke rekening pribadi, dan perusahaan atau aplikasinya tidak terdaftar di OJK. Cek legalitas di situs OJK atau Kontak OJK 157, dan laporkan tawaran mencurigakan ke Satgas PASTI."),
  H2("Cara mulai"),
  ...NUM([
    "Siapkan KTP, rekening bank atas nama sendiri, dan NPWP jika punya.",
    "Pilih APERD (bank atau aplikasi) yang terdaftar di OJK.",
    "Isi profil risiko dengan jujur.",
    "Lakukan pembelian pertama. Banyak aplikasi menerima mulai Rp10.000.",
    "Tambah secara rutin dan cek konfirmasi transaksi dari KSEI. Kepemilikan juga bisa dicek di akses.ksei.co.id.",
  ]),
  H2("Kapan RDPU tepat, kapan tidak"),
  TABLE(["Tepat untuk", "Kurang tepat untuk"], [
    ["Dana darurat", "Tujuan di atas 5 tahun seperti dana pensiun"],
    ["Tujuan di bawah 1 tahun (biaya KKN, laptop, UKT)", "Mengejar untung besar dalam waktu singkat"],
    ["Tempat parkir sementara sebelum dipindah ke instrumen lain", "Pengganti tabungan untuk transaksi harian"],
    ["Belajar berinvestasi dengan risiko kecil", "Investasi dengan uang pinjaman, termasuk paylater"],
  ], [4513, 4513]),

  H1("12. Mitos dan fakta"),
  TABLE(["Mitos", "Fakta"], [
    ["RDPU pasti untung dan tidak mungkin turun.", "NAB/UP RDPU bisa turun, misalnya saat penerbit surat utang gagal bayar. Jarang terjadi, tapi mungkin."],
    ["RDPU dijamin LPS karena belinya lewat bank.", "Penjaminan LPS berlaku untuk simpanan bank. RDPU tidak dijamin LPS, sekalipun dibeli lewat bank."],
    ["Kalau Manajer Investasi bangkrut, uang kita ikut hilang.", "Aset reksa dana disimpan di Bank Kustodian dan dicatat terpisah dari harta MI."],
    ["Untung RDPU harus dilaporkan dan dibayar pajaknya.", "Keuntungan bagi pemegang unit bukan objek pajak. Pajak sudah dipotong di dalam portofolio. Kepemilikannya tetap dicantumkan sebagai harta di SPT bagi yang wajib lapor."],
    ["Produk dengan return tertinggi pasti yang terbaik.", "Return tinggi bisa berasal dari surat utang yang lebih berisiko. Lihat konsistensi dan isi portofolio."],
    ["Harus punya modal besar untuk mulai.", "Banyak aplikasi menerima pembelian mulai Rp10.000."],
  ], [4000, 5026]),

  H1("13. Latihan soal"),
  P("Kerjakan dulu, baru cocokkan dengan kunci di bawah."),
  ...NUM([
    "Sebutkan batas jatuh tempo efek utang yang boleh dibeli RDPU.",
    "Kamu membeli RDPU Rp500.000 saat NAB/UP Rp1.250. Berapa unit yang kamu dapat?",
    "Lanjutan soal 2: sebulan kemudian NAB/UP menjadi Rp1.255. Berapa nilai investasimu?",
    "Siapa yang menyimpan aset reksa dana dan menghitung NAB?",
    "Kenapa keuntungan RDPU tidak dipotong pajak lagi saat dicairkan?",
    "Berapa batas minimum NAB sebelum reksa dana wajib dibubarkan, dan berapa lama?",
    "Kamu mengorder pembelian pukul 15.00 WIB pada hari bursa. NAB/UP hari apa yang dipakai?",
    "Deposito Rp5.000.000 berbunga 4% setahun. Berapa bunga bersih setelah pajak dalam setahun?",
    "Saat BI Rate naik, apa yang cenderung terjadi pada return RDPU? Jelaskan alasannya.",
    "Sebutkan dua tanda bahaya investasi bodong.",
  ]),
  H2("Kunci jawaban"),
  ...NUM([
    "Jatuh tempo, atau sisa jatuh tempo, tidak lebih dari 1 tahun.",
    "500.000 ÷ 1.250 = 400 unit.",
    "400 × 1.255 = Rp502.000.",
    "Bank Kustodian.",
    "Karena keuntungan yang diterima pemegang unit reksa dana bukan objek pajak (UU PPh Pasal 4 ayat 3). Pajak atas bunga deposito dan obligasi sudah dipotong di dalam portofolio.",
    "Di bawah Rp10 miliar selama 120 hari bursa berturut-turut.",
    "NAB/UP hari bursa berikutnya, karena melewati batas waktu (umumnya pukul 13.00 WIB).",
    "5.000.000 × 4% = Rp200.000, dipotong 20% = Rp160.000.",
    "Cenderung naik dengan jeda waktu, karena saat instrumen lama jatuh tempo, dananya ditempatkan ulang di deposito dan surat utang baru yang bunganya lebih tinggi.",
    "Contoh: janji untung pasti, diminta transfer ke rekening pribadi, perusahaan tidak terdaftar di OJK.",
  ]),

  H1("14. Kamus istilah"),
  TABLE(["Istilah", "Arti"], [
    ["APERD", "Agen Penjual Efek Reksa Dana: bank atau aplikasi tempat membeli reksa dana."],
    ["Bank Kustodian", "Bank yang menyimpan aset reksa dana dan menghitung NAB."],
    ["BI Rate", "Suku bunga acuan Bank Indonesia."],
    ["Fund fact sheet", "Laporan bulanan berisi kinerja, komposisi portofolio, dan biaya reksa dana."],
    ["KIK", "Kontrak Investasi Kolektif, bentuk hukum reksa dana di Indonesia."],
    ["LPS", "Lembaga Penjamin Simpanan, menjamin simpanan nasabah di bank."],
    ["Manajer Investasi", "Perusahaan berizin OJK yang mengelola portofolio reksa dana."],
    ["NAB", "Nilai Aktiva Bersih: total aset reksa dana dikurangi kewajiban."],
    ["NAB/UP", "Harga satu unit penyertaan."],
    ["Prospektus", "Dokumen resmi berisi kebijakan investasi, biaya, risiko, dan tata cara transaksi."],
    ["Redemption", "Penjualan kembali unit reksa dana."],
    ["SID", "Single Investor Identification, nomor identitas tunggal investor di KSEI."],
    ["Subscription", "Pembelian unit reksa dana."],
    ["Unit Penyertaan", "Satuan kepemilikan investor dalam reksa dana."],
  ], [2400, 6626]),

  H1("15. Sumber"),
  ...BUL([
    "UU No. 8 Tahun 1995 tentang Pasar Modal.",
    "POJK 23/POJK.04/2016 tentang Reksa Dana berbentuk Kontrak Investasi Kolektif, beserta perubahannya.",
    "UU Pajak Penghasilan Pasal 4 ayat (3); PP 131/2000 (bunga deposito); PP 91/2021 (bunga obligasi).",
    "KSEI, Statistik Publik Juli 2026.",
    "Bank Indonesia, hasil Rapat Dewan Gubernur 2024 sampai September 2026, dirangkum dari Bisnis.com, Kontan, dan Investor Trust.",
    "LPS, Tingkat Bunga Penjaminan periode 1 Juli sampai 30 September 2026.",
    "Kontan dan Bareksa, data kinerja reksa dana pasar uang 2026.",
    "OJK, surat S-1422/PM.21/2019 tentang pembubaran reksa dana Minna Padi, dirangkum dari Bareksa dan CNBC Indonesia.",
  ]),

  H1("16. Catatan untuk pembawa materi berikutnya"),
  P("Angka pasar berubah. Sebelum membawakan materi ini lagi, perbarui data berikut:"),
  ...BUL([
    "BI Rate terbaru (bi.go.id). Ada di slide 4 dan 20, serta bagian 2 dan 10 modul ini.",
    "Tingkat bunga penjaminan LPS periode berjalan (lps.go.id). Slide 4, 17, 18.",
    "Jumlah investor dan dana kelolaan per jenis reksa dana (Statistik Publik KSEI terbaru). Slide 5 dan 9.",
    "Kisaran return RDPU satu tahun terakhir (fund fact sheet atau aplikasi APERD). Slide 16.",
  ]),
  P("Hal yang tidak perlu diubah kecuali ada aturan baru: definisi RDPU, perlakuan pajak, batas pembubaran, dan batas waktu pembayaran redemption."),
];

// ================= NASKAH =================
const order = Object.keys(NOTES).map(Number).sort((a, b) => a - b);
const totalMin = order.reduce((a, k) => a + NOTES[k].min, 0);
const naskah = [
  ...cover("Naskah Presentasi Reksa Dana Pasar Uang", "Panduan bicara slide demi slide, sama dengan speaker notes di file PowerPoint.", [
    `**Durasi perkiraan:** sekitar ${Math.round(totalMin)} menit, belum termasuk tanya jawab`,
    "**Format:** 35 slide, enam bagian, ditutup kuis interaktif",
  ]),
  H1("Sebelum mulai"),
  ...BUL([
    "Buka presentasi dalam mode slideshow. Tombol \"Daftar isi\" dan pilihan jawaban kuis hanya bisa diklik di mode ini.",
    "Siapkan Simulator_RDPU.html di tab browser terpisah untuk slide 18.",
    "Cek ulang angka BI Rate dan LPS di hari presentasi.",
    "Naskah ini panduan, bukan teks hafalan. Pakai bahasamu sendiri selama poinnya tersampaikan.",
  ]),
  H2("Pembagian waktu"),
  TABLE(["Bagian", "Slide", "Perkiraan"], [
    ["Pembuka", "1 sampai 2", "2 menit"],
    ["1. Kenapa perlu paham", "3 sampai 5", "4 menit"],
    ["2. Mengenal RDPU", "6 sampai 10", "8,5 menit"],
    ["3. Cara kerjanya", "11 sampai 14", "7,5 menit"],
    ["4. Untung, biaya, pajak, risiko", "15 sampai 20", "13,5 menit"],
    ["5. Praktik", "21 sampai 25", "9 menit"],
    ["6. Kuis", "26 sampai 32", "6,5 menit"],
    ["Penutup", "33 sampai 35", "3 menit"],
  ], [4200, 2400, 2426]),
  new Paragraph({ children: [new PageBreak()] }),
  ...order.flatMap((k) => {
    const n = NOTES[k];
    return [
      H2(`Slide ${k}: ${n.title}`),
      new Paragraph({ children: [new TextRun({ text: `Sekitar ${String(n.min).replace(".", ",")} menit`, italics: true, color: MUTED, size: 18 })], spacing: { after: 100 } }),
      ...n.script.split("\n\n").map((para) => P(para)),
    ];
  }),
  H1("Pertanyaan yang sering muncul"),
  ...[
    ["Kalau aplikasinya tutup, uang saya bagaimana?", "Kepemilikan unit tercatat atas nama SID kamu di KSEI dan asetnya ada di Bank Kustodian. Penjualan tetap bisa diurus lewat Manajer Investasi produk tersebut."],
    ["Lebih baik RDPU atau deposito?", "Tergantung kebutuhan. Deposito dijamin LPS dan bunganya pasti selama tenor. RDPU lebih fleksibel dicairkan, mulai dari nominal kecil, dan tidak ada pajak tambahan bagi investor, tetapi tidak dijamin LPS dan return-nya tidak pasti."],
    ["Kenapa saldo saya turun sedikit di satu hari?", "Bisa karena koreksi harga surat utang di portofolio atau pencatatan biaya. Kalau turunnya besar atau berlanjut, cek fund fact sheet dan berita tentang penerbit di portofolio."],
    ["Apakah RDPU perlu dilaporkan di SPT?", "Keuntungannya bukan objek pajak, tetapi kepemilikan reksa dana tetap dicantumkan sebagai harta di SPT tahunan bagi yang sudah wajib lapor."],
    ["Berapa banyak produk RDPU yang sebaiknya dibeli?", "Untuk pemula, satu atau dua produk sudah cukup. Menambah produk tidak banyak mengurangi risiko karena isinya mirip."],
  ].flatMap(([q, a]) => [new Paragraph({ children: [new TextRun({ text: q, bold: true })], spacing: { before: 160, after: 60 } }), P(a)]),
];

(async () => {
  const mk = (children, label) => new Document({
    creator: "Divisi", title: label, styles, numbering, features: { updateFields: true },
    sections: [{ properties: { page }, footers: { default: footer(label) }, children }],
  });
  fs.writeFileSync("Modul_Belajar_RDPU.docx", await Packer.toBuffer(mk(modul, "Modul Belajar RDPU")));
  fs.writeFileSync("Naskah_Presentasi_RDPU.docx", await Packer.toBuffer(mk(naskah, "Naskah Presentasi RDPU")));
  console.log("ok", totalMin);
})();
