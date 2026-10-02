const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");
const NOTES = require("./notes.js");

const OUT = process.argv[2] || "Reksa_Dana_Pasar_Uang.pptx";
const SIM_URL = process.env.SIM_URL || "";

const THEME = {
  name: "RDPU Ink Mint",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "0E1B2C", lt1: "FFFFFF", dk2: "1D3A55", lt2: "EEF4F2",
    accent1: "17A589", accent2: "F2B33D", accent3: "4F7FD9",
    accent4: "E0625F", accent5: "6B7B8C", accent6: "9FD8CB",
    hlink: "17A589", folHlink: "1D3A55",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Reksa Dana Pasar Uang";
pres.subject = "Materi regenerasi divisi";
pres.author = "Divisi";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

const W = 13.333;
const FOOT = "Reksa Dana Pasar Uang  |  Materi Regenerasi Divisi";

// ---------- icons ----------
const iconCache = {};
async function icon(name, color) {
  const key = name + color;
  if (iconCache[key]) return iconCache[key];
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(fa[name], { color: "#" + color, size: 256 })
  );
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  iconCache[key] = "image/png;base64," + buf.toString("base64");
  return iconCache[key];
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.text1 },
  objects: [
    { ellipse: { x: 8.6, y: -2.2, w: 6.6, h: 6.6, fill: { color: C.accent1, transparency: 82 }, line: { color: C.accent1, transparency: 100 } } },
    { ellipse: { x: 10.4, y: 3.9, w: 4.2, h: 4.2, fill: { color: C.accent2, transparency: 86 }, line: { color: C.accent2, transparency: 100 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 9.2, h: 1.9, fontFace: THEME.headFontFace, fontSize: 48, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.35, w: 8.5, h: 1.0, fontSize: 20, color: C.accent6, valign: "top", margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "SECTION",
  background: { color: C.text1 },
  objects: [
    { ellipse: { x: 9.3, y: 1.2, w: 5.2, h: 5.2, fill: { color: C.accent1, transparency: 85 }, line: { color: C.accent1, transparency: 100 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 3.1, w: 9.0, h: 1.3, fontFace: THEME.headFontFace, fontSize: 36, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.45, w: 8.4, h: 1.0, fontSize: 18, color: C.accent6, valign: "top", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.8, y: 6.95, w: 7, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
});

pres.defineSlideMaster({
  title: "LIGHT",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 10.6, h: 0.8, fontFace: THEME.headFontFace, fontSize: 28, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.6, y: 6.95, w: 7, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
});

pres.defineSlideMaster({
  title: "DARK",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 10.6, h: 0.8, fontFace: THEME.headFontFace, fontSize: 28, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.6, y: 6.95, w: 7, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
});

// ---------- helpers ----------
let slideNo = 0;
function addSlide(master, section) {
  slideNo++;
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  s._no = slideNo;
  return s;
}
function notes(s) {
  const n = NOTES[s._no];
  if (!n) throw new Error("Missing notes for slide " + s._no);
  s.addNotes(n.script);
}
function homeLink(s, dark) {
  s.addText([{ text: "Daftar isi", options: { hyperlink: { slide: 2, tooltip: "Kembali ke daftar isi" } } }], {
    x: 11.45, y: 0.55, w: 1.3, h: 0.45, fontSize: 11, bold: true, align: "center", valign: "middle",
    color: dark ? C.accent6 : C.accent1, shape: pres.ShapeType.roundRect, rectRadius: 0.22,
    line: { color: dark ? C.accent6 : C.accent1, width: 1 }, isTextBox: true, objectName: "Link daftar isi",
  });
}
function card(s, x, y, w, h, fill, name) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.12, fill: { color: fill || C.background2 }, line: { color: fill || C.background2, width: 0 },
    objectName: name || "Kartu",
  });
}
function shadowCard(s, x, y, w, h, name) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.12, fill: { color: C.background1 }, line: { color: "DCE6E2", width: 0.75 },
    shadow: { type: "outer", color: "0E1B2C", opacity: 0.12, blur: 10, offset: 3, angle: 90 },
    objectName: name || "Kartu",
  });
}
async function iconCircle(s, x, y, d, iconName, fill, iconColorHex, name) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: (name || iconName) + " lingkaran" });
  const p = d * 0.25;
  s.addImage({ data: await icon(iconName, iconColorHex), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: (name || iconName) + " ikon" });
}
function numCircle(s, x, y, d, label, fill, color, size) {
  s.addText(label, {
    x, y, w: d, h: d, shape: pres.ShapeType.ellipse, fill: { color: fill }, line: { color: fill, width: 0 },
    fontSize: size || 16, bold: true, color, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Nomor " + label,
  });
}
function txt(s, text, opts) {
  s.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", color: C.text1, fontSize: 15 }, opts));
}
function source(s, text, dark) {
  txt(s, text, { x: 0.6, y: 6.55, w: 11.5, h: 0.3, fontSize: 10, italic: true, color: dark ? C.accent6 : C.accent5 });
}
function sectionSlide(sec, num, title, sub) {
  const s = addSlide("SECTION", sec);
  txt(s, "BAGIAN " + num, { x: 0.8, y: 2.45, w: 4, h: 0.45, fontSize: 14, bold: true, color: C.accent2, charSpacing: 4 });
  s.addText(title, { placeholder: "title" });
  s.addText(sub, { placeholder: "body" });
  txt(s, num, { x: 9.3, y: 2.3, w: 5.2, h: 3.0, fontSize: 150, bold: true, fontFace: THEME.headFontFace, color: C.accent1, align: "center", valign: "middle" });
  homeLink(s, true);
  notes(s);
  return s;
}

// slide numbers used for links
const SL = { agenda: 2, s1: 3, s2: 6, s3: 11, s4: 15, s5: 21, s6: 26, a1: 28, a2: 30, a3: 32, q1: 27, q2: 29, q3: 31 };

(async () => {
  // 1. COVER
  pres.addSection({ title: "Pembuka" });
  let s = addSlide("COVER", "Pembuka");
  txt(s, "MATERI REGENERASI DIVISI  |  OKTOBER 2026", { x: 0.8, y: 1.6, w: 8, h: 0.4, fontSize: 13, bold: true, color: C.accent2, charSpacing: 3 });
  s.addText("Reksa Dana Pasar Uang", { placeholder: "title" });
  s.addText("Cara paling tenang untuk mulai berinvestasi: apa isinya, bagaimana cara kerjanya, dan apa risikonya.", { placeholder: "body" });
  await iconCircle(s, 10.55, 0.75, 1.6, "FaCoins", C.accent2, HEX.dk1, "Koin");
  txt(s, "Klik \"Daftar isi\" di pojok kanan atas tiap slide untuk berpindah bagian.", { x: 0.8, y: 6.6, w: 8, h: 0.35, fontSize: 11, color: C.accent5 });
  notes(s);

  // 2. AGENDA (interactive)
  s = addSlide("LIGHT", "Pembuka");
  s.addText("Daftar isi", { placeholder: "title" });
  txt(s, "Klik salah satu kartu untuk langsung lompat ke bagian itu.", { x: 0.6, y: 1.2, w: 10, h: 0.4, fontSize: 14, color: C.accent5 });
  const agenda = [
    ["01", "Kenapa kita perlu paham", "Uang diam vs uang bekerja, dan data investor muda", SL.s1, "FaLightbulb"],
    ["02", "Mengenal RDPU", "Definisi resmi, isi portofolio, pihak yang terlibat", SL.s2, "FaBookOpen"],
    ["03", "Cara kerjanya", "NAB, unit penyertaan, alur beli dan jual", SL.s3, "FaRoute"],
    ["04", "Untung, biaya, pajak, risiko", "Perbandingan dengan tabungan dan deposito", SL.s4, "FaScaleBalanced"],
    ["05", "Praktik", "Cara memilih produk, waspada bodong, cara mulai", SL.s5, "FaListCheck"],
    ["06", "Kuis interaktif", "Tiga soal untuk mengecek pemahaman", SL.s6, "FaCircleQuestion"],
  ];
  for (let i = 0; i < agenda.length; i++) {
    const [n, t, d, target, ic] = agenda[i];
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.1, y = 1.85 + row * 2.45, w = 3.8, h = 2.15;
    shadowCard(s, x, y, w, h, "Kartu agenda " + n);
    await iconCircle(s, x + 0.3, y + 0.3, 0.75, ic, C.accent1, "FFFFFF", "Agenda " + n);
    txt(s, n, { x: x + 2.6, y: y + 0.3, w: 0.95, h: 0.6, fontSize: 28, bold: true, fontFace: THEME.headFontFace, color: C.accent6, align: "right" });
    s.addText([{ text: t, options: { hyperlink: { slide: target, tooltip: "Buka bagian " + n } } }], {
      x: x + 0.3, y: y + 1.15, w: w - 0.5, h: 0.42, fontSize: 15, bold: true, color: C.text1, margin: 0, isTextBox: true, valign: "top", objectName: "Link agenda " + n,
    });
    txt(s, d, { x: x + 0.3, y: y + 1.6, w: w - 0.6, h: 0.5, fontSize: 12, color: C.accent5 });
  }
  notes(s);

  // 3. SECTION 1
  pres.addSection({ title: "1. Kenapa perlu paham" });
  sectionSlide("1. Kenapa perlu paham", "01", "Kenapa kita perlu paham?", "Uang yang cuma diam pelan-pelan kalah oleh kenaikan harga.");

  // 4. Hook stats
  s = addSlide("DARK", "1. Kenapa perlu paham");
  s.addText("Uang diam vs uang yang bekerja", { placeholder: "title" });
  homeLink(s, true);
  const stats = [
    ["5,75%", "BI Rate", "Ditahan BI pada RDG 22 sampai 23 September 2026, setelah naik total 100 bps di Mei sampai Juni 2026.", C.accent1],
    ["3,75%", "Batas bunga LPS", "Batas bunga simpanan rupiah di bank umum yang dijamin LPS, periode 1 Juli sampai 30 September 2026.", C.accent2],
    ["2,5±1%", "Target inflasi BI", "Kalau uangmu tumbuh di bawah angka inflasi, daya belinya turun walau nominalnya tidak berkurang.", C.accent3],
  ];
  stats.forEach(([big, lab, desc, col], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.6, 3.8, 3.9, C.text2, "Kartu statistik " + lab);
    txt(s, big, { x: x + 0.35, y: 1.9, w: 3.2, h: 1.1, fontSize: 52, bold: true, fontFace: THEME.headFontFace, color: col });
    txt(s, lab, { x: x + 0.35, y: 3.05, w: 3.2, h: 0.45, fontSize: 18, bold: true, color: C.background1 });
    txt(s, desc, { x: x + 0.35, y: 3.55, w: 3.1, h: 1.8, fontSize: 13, color: C.accent6 });
  });
  txt(s, [
    { text: "Intinya: ", options: { bold: true, color: C.accent2 } },
    { text: "uang di dompet nilainya tidak tumbuh. RDPU adalah langkah pertama paling tenang untuk membuat uang ikut bekerja.", options: { color: C.background1 } },
  ], { x: 0.6, y: 5.75, w: 12, h: 0.6, fontSize: 16 });
  source(s, "Sumber: Bank Indonesia (RDG September 2026), LPS (TBP Juli sampai September 2026).", true);
  notes(s);

  // 5. Investor data chart
  s = addSlide("LIGHT", "1. Kenapa perlu paham");
  s.addText("Investor reksa dana tumbuh cepat", { placeholder: "title" });
  homeLink(s, false);
  s.addChart(pres.charts.BAR, [{ name: "Investor reksa dana (juta SID)", labels: ["2023", "2024", "2025"], values: [11.42, 14.03, 19.19] }], {
    x: 0.6, y: 1.45, w: 7.2, h: 4.95, barDir: "col", chartColors: [HEX.accent1],
    showTitle: true, title: "Jumlah investor reksa dana (juta SID, akhir tahun)", titleFontSize: 14, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelColor: HEX.dk1, dataLabelFormatCode: "0.00", dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.accent5, valAxisLabelColor: HEX.accent5, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 13,
    valGridLine: { color: "E3EAE7", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 25, barGapWidthPct: 70,
    valAxisHidden: true,
  });
  const st2 = [
    ["30,06 juta", "total investor pasar modal (SID) per akhir Juli 2026"],
    ["+36,77%", "pertumbuhan investor reksa dana sepanjang 2025"],
    ["> 50%", "investor individu berusia 30 tahun ke bawah"],
  ];
  st2.forEach(([b, d], i) => {
    const y = 1.5 + i * 1.65;
    card(s, 8.2, y, 4.55, 1.45, C.background2, "Kartu data " + i);
    txt(s, b, { x: 8.5, y: y + 0.18, w: 4, h: 0.65, fontSize: 30, bold: true, fontFace: THEME.headFontFace, color: C.accent1 });
    txt(s, d, { x: 8.5, y: y + 0.85, w: 4.0, h: 0.5, fontSize: 13, color: C.text1 });
  });
  source(s, "Sumber: KSEI, Statistik Publik Juli 2026. Satu SID dapat memiliki lebih dari satu jenis aset.");
  notes(s);

  // 6. SECTION 2
  pres.addSection({ title: "2. Mengenal RDPU" });
  sectionSlide("2. Mengenal RDPU", "02", "Mengenal RDPU", "Apa itu reksa dana, apa yang membuat versi \"pasar uang\" berbeda, dan siapa saja yang terlibat.");

  // 7. Analogi patungan
  s = addSlide("LIGHT", "2. Mengenal RDPU");
  s.addText("Reksa dana: patungan yang dikelola profesional", { placeholder: "title" });
  homeLink(s, false);
  // three columns diagram: investors -> MI -> instruments
  const colX = [0.6, 4.95, 9.3];
  const colT = ["Banyak investor", "Manajer Investasi", "Portofolio efek"];
  const colD = [
    "Kamu dan ribuan orang lain menyetor uang, mulai dari nominal kecil.",
    "Perusahaan berizin OJK yang memilih dan mengatur isi portofolio sesuai kebijakan investasi.",
    "Uang dibelikan deposito, surat utang negara, atau obligasi. Hasilnya dibagi sesuai porsi unit masing-masing.",
  ];
  const colI = ["FaUsers", "FaUserTie", "FaWallet"];
  for (let i = 0; i < 3; i++) {
    shadowCard(s, colX[i], 1.6, 3.45, 3.6, "Kartu alur " + colT[i]);
    await iconCircle(s, colX[i] + 1.27, 1.85, 0.9, colI[i], i === 1 ? C.accent2 : C.accent1, i === 1 ? HEX.dk1 : "FFFFFF", colT[i]);
    txt(s, colT[i], { x: colX[i] + 0.25, y: 2.9, w: 2.95, h: 0.45, fontSize: 18, bold: true, align: "center" });
    txt(s, colD[i], { x: colX[i] + 0.3, y: 3.4, w: 2.85, h: 1.7, fontSize: 13, color: C.accent5, align: "center" });
    if (i < 2) s.addShape(pres.ShapeType.rightArrow, { x: colX[i] + 3.55, y: 3.15, w: 0.7, h: 0.5, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, objectName: "Panah alur " + i });
  }
  card(s, 0.6, 5.45, 12.15, 0.95, C.background2, "Kartu analogi");
  txt(s, [
    { text: "Analogi: ", options: { bold: true, color: C.accent1 } },
    { text: "seperti satu angkatan patungan menyewa koki. Kamu tidak perlu jago masak, cukup pilih koki yang jelas izinnya dan ikut menikmati hasilnya sesuai porsi iuranmu.", options: {} },
  ], { x: 0.9, y: 5.6, w: 11.6, h: 0.7, fontSize: 15, valign: "middle" });
  notes(s);

  // 8. Definisi + isi
  s = addSlide("LIGHT", "2. Mengenal RDPU");
  s.addText("Definisi resmi: apa yang membuatnya \"pasar uang\"", { placeholder: "title" });
  homeLink(s, false);
  card(s, 0.6, 1.5, 6.3, 2.35, C.text1, "Kartu definisi");
  txt(s, "Reksa Dana Pasar Uang (RDPU)", { x: 0.95, y: 1.7, w: 5.7, h: 0.45, fontSize: 18, bold: true, color: C.accent2 });
  txt(s, "Reksa dana yang 100% asetnya ditempatkan pada instrumen pasar uang dalam negeri dan/atau efek bersifat utang yang jatuh temponya tidak lebih dari 1 tahun, atau sisa jatuh temponya tidak lebih dari 1 tahun.", { x: 0.95, y: 2.2, w: 5.7, h: 1.55, fontSize: 14, color: C.background1 });
  txt(s, "Kata kuncinya: jangka pendek (maksimal 1 tahun). Karena tenornya pendek, harga isinya jarang bergejolak.", { x: 0.6, y: 4.05, w: 6.3, h: 0.8, fontSize: 14, color: C.text1 });
  const items = [
    ["FaLandmark", "Deposito bank", "Biasanya porsi terbesar. Bunganya dipotong pajak final 20%."],
    ["FaFileInvoiceDollar", "Surat utang negara jangka pendek", "Contoh: Surat Perbendaharaan Negara (SPN) dan sukuk negara tenor pendek."],
    ["FaBuildingColumns", "Obligasi atau sukuk korporasi", "Yang jatuh temponya, atau sisa jatuh temponya, maksimal 1 tahun."],
  ];
  for (let i = 0; i < items.length; i++) {
    const y = 1.5 + i * 1.6;
    shadowCard(s, 7.25, y, 5.5, 1.38, "Kartu isi " + i);
    await iconCircle(s, 7.5, y + 0.3, 0.78, items[i][0], C.accent1, "FFFFFF", items[i][1]);
    txt(s, items[i][1], { x: 8.5, y: y + 0.2, w: 4.1, h: 0.42, fontSize: 15, bold: true });
    txt(s, items[i][2], { x: 8.5, y: y + 0.62, w: 4.1, h: 0.7, fontSize: 12, color: C.accent5 });
  }
  card(s, 0.6, 5.0, 6.3, 1.35, C.background2, "Kartu syariah");
  txt(s, [
    { text: "Ada versi syariah. ", options: { bold: true, color: C.accent1 } },
    { text: "RDPU Syariah menempatkan dana di deposito syariah dan sukuk, diawasi Dewan Pengawas Syariah. Cara kerjanya sama.", options: {} },
  ], { x: 0.9, y: 5.15, w: 5.8, h: 1.1, fontSize: 13, valign: "middle" });
  source(s, "Sumber: POJK 23/POJK.04/2016 tentang Reksa Dana berbentuk KIK dan perubahannya; prospektus RDPU.");
  notes(s);

  // 9. Peta jenis reksa dana + AUM chart
  s = addSlide("LIGHT", "2. Mengenal RDPU");
  s.addText("Posisi RDPU di antara jenis reksa dana lain", { placeholder: "title" });
  homeLink(s, false);
  const ladder = [
    ["Pasar Uang", "Risiko paling rendah", "Di bawah 1 tahun", HEX.accent1],
    ["Pendapatan Tetap", "Rendah sampai sedang", "1 sampai 3 tahun", HEX.accent3],
    ["Campuran", "Sedang", "3 sampai 5 tahun", HEX.accent2],
    ["Saham", "Paling tinggi", "Di atas 5 tahun", HEX.accent4],
  ];
  txt(s, "Tangga risiko dan horizon investasi yang umum dipakai", { x: 0.6, y: 1.4, w: 6, h: 0.4, fontSize: 14, bold: true, color: C.accent5 });
  ladder.forEach(([n, r, h, col], i) => {
    const y = 5.15 - i * 1.0, x = 0.6 + i * 0.45, w = 5.9 - i * 0.45;
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.85, rectRadius: 0.1, fill: { color: col, transparency: i === 0 ? 0 : 15 }, line: { color: col, width: 0 }, objectName: "Tangga " + n });
    txt(s, n, { x: x + 0.25, y: y + 0.08, w: 2.6, h: 0.4, fontSize: 15, bold: true, color: i === 2 ? C.text1 : C.background1 });
    txt(s, r + "  |  " + h, { x: x + 0.25, y: y + 0.45, w: w - 0.4, h: 0.35, fontSize: 11.5, color: i === 2 ? C.text1 : C.background1 });
  });
  txt(s, "Makin ke atas: potensi hasil makin tinggi, naik turunnya juga makin besar.", { x: 0.6, y: 6.08, w: 6.3, h: 0.3, fontSize: 11, italic: true, color: C.accent5 });
  s.addChart(pres.charts.BAR, [{
    name: "AUM (Rp triliun)",
    labels: ["Pendapatan Tetap", "Terproteksi", "Pasar Uang", "Saham", "Campuran"],
    values: [217.02, 144.84, 133.58, 62.69, 34.83],
  }], {
    x: 7.1, y: 1.4, w: 5.65, h: 4.95, barDir: "bar",
    chartColors: [HEX.accent5, HEX.accent5, HEX.accent1, HEX.accent5, HEX.accent5],
    showTitle: true, title: "Dana kelolaan per jenis, Juli 2026 (Rp triliun)", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFormatCode: "#,##0.00", dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, catAxisOrientation: "maxMin",
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, valAxisMaxVal: 270, barGapWidthPct: 55,
  });
  source(s, "Sumber AUM: KSEI Statistik Publik Juli 2026 (data S-INVEST). RDPU: 197 produk, Rp133,58 triliun.");
  notes(s);

  // 10. Pihak-pihak
  s = addSlide("LIGHT", "2. Mengenal RDPU");
  s.addText("Siapa saja yang terlibat", { placeholder: "title" });
  homeLink(s, false);
  const parties = [
    ["FaGavel", "OJK", "Regulator. Memberi izin dan mengawasi seluruh pihak."],
    ["FaUserTie", "Manajer Investasi", "Mengelola portofolio. Tidak memegang uang investor."],
    ["FaVault", "Bank Kustodian", "Menyimpan aset reksa dana dan menghitung NAB setiap hari bursa."],
    ["FaMobileScreen", "APERD", "Agen penjual: bank atau aplikasi investasi tempat kamu membeli."],
    ["FaShieldHalved", "KSEI (S-INVEST)", "Mencatat kepemilikan unit atas nama SID kamu."],
    ["FaUsers", "Investor", "Pemilik unit penyertaan. Bisa cek saldo dan konfirmasi transaksi."],
  ];
  for (let i = 0; i < parties.length; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 3.85, y = 1.5 + row * 1.6;
    shadowCard(s, x, y, 3.6, 1.4, "Kartu pihak " + parties[i][1]);
    await iconCircle(s, x + 0.22, y + 0.3, 0.75, parties[i][0], C.accent1, "FFFFFF", parties[i][1]);
    txt(s, parties[i][1], { x: x + 1.15, y: y + 0.18, w: 2.35, h: 0.4, fontSize: 14, bold: true });
    txt(s, parties[i][2], { x: x + 1.15, y: y + 0.58, w: 2.35, h: 0.78, fontSize: 11, color: C.accent5 });
  }
  card(s, 8.4, 1.5, 4.35, 4.6, C.text1, "Kartu poin penting");
  await iconCircle(s, 8.75, 1.8, 0.8, "FaShieldHalved", C.accent2, HEX.dk1, "Perlindungan");
  txt(s, "Poin penting", { x: 8.75, y: 2.75, w: 3.7, h: 0.45, fontSize: 18, bold: true, color: C.accent2 });
  txt(s, "Uangmu tidak disimpan di rekening Manajer Investasi. Asetnya dititipkan di Bank Kustodian dan dicatat terpisah dari harta MI. Kalau MI bermasalah, aset reksa dana bukan bagian dari harta MI.", { x: 8.75, y: 3.25, w: 3.7, h: 2.7, fontSize: 13.5, color: C.background1 });
  source(s, "Sumber: UU No. 8 Tahun 1995 tentang Pasar Modal; POJK 23/POJK.04/2016; KSEI.");
  notes(s);

  // 11. SECTION 3
  pres.addSection({ title: "3. Cara kerjanya" });
  sectionSlide("3. Cara kerjanya", "03", "Cara kerjanya", "NAB, unit penyertaan, dan kenapa grafik RDPU naik seperti tangga.");

  // 12. NAB & unit worked example
  s = addSlide("LIGHT", "3. Cara kerjanya");
  s.addText("NAB dan unit penyertaan", { placeholder: "title" });
  homeLink(s, false);
  card(s, 0.6, 1.5, 5.6, 4.85, C.background2, "Kartu konsep");
  txt(s, "Dua istilah wajib", { x: 0.9, y: 1.7, w: 5, h: 0.45, fontSize: 18, bold: true, color: C.accent1 });
  txt(s, [
    { text: "NAB (Nilai Aktiva Bersih)", options: { bold: true, breakLine: true } },
    { text: "Total nilai seluruh aset reksa dana dikurangi kewajibannya (misalnya biaya yang belum dibayar).", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "Unit Penyertaan (UP)", options: { bold: true, breakLine: true } },
    { text: "\"Potongan kue\" yang kamu miliki. Harga per potongnya disebut NAB/UP dan diumumkan setiap hari bursa.", options: { breakLine: true } },
  ], { x: 0.9, y: 2.2, w: 5.0, h: 2.6, fontSize: 13.5 });
  s.addShape(pres.ShapeType.roundRect, { x: 0.9, y: 4.95, w: 5.0, h: 1.15, rectRadius: 0.1, fill: { color: C.text1 }, line: { color: C.text1, width: 0 }, objectName: "Kotak rumus" });
  txt(s, "NAB/UP = NAB total ÷ jumlah unit", { x: 1.1, y: 5.05, w: 4.6, h: 0.95, fontSize: 15, bold: true, color: C.accent2, align: "center", valign: "middle" });
  // steps
  const steps = [
    ["1", "Beli", "Kamu beli Rp1.000.000 saat NAB/UP Rp1.500", "Unit = 1.000.000 ÷ 1.500 = 666,67 unit"],
    ["2", "Tunggu", "Setahun lagi NAB/UP naik 4,5% jadi Rp1.567,50", "Jumlah unitmu tetap 666,67 unit"],
    ["3", "Jual", "Nilai investasi = 666,67 × 1.567,50", "= Rp1.045.000 (untung Rp45.000)"],
  ];
  steps.forEach(([n, t, a, b], i) => {
    const y = 1.5 + i * 1.65;
    shadowCard(s, 6.55, y, 6.2, 1.45, "Langkah hitung " + n);
    numCircle(s, 6.8, y + 0.4, 0.65, n, C.accent1, C.background1, 18);
    txt(s, t, { x: 7.7, y: y + 0.15, w: 4.8, h: 0.4, fontSize: 15, bold: true });
    txt(s, a, { x: 7.7, y: y + 0.52, w: 4.85, h: 0.4, fontSize: 12.5, color: C.accent5 });
    txt(s, b, { x: 7.7, y: y + 0.9, w: 4.85, h: 0.4, fontSize: 13, bold: true, color: C.accent1 });
  });
  source(s, "Angka di atas ilustrasi untuk latihan hitung, bukan kinerja produk tertentu.");
  notes(s);

  // 13. Grafik tangga
  s = addSlide("LIGHT", "3. Cara kerjanya");
  s.addText("Kenapa grafik RDPU naik pelan seperti tangga?", { placeholder: "title" });
  homeLink(s, false);
  const months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
  // illustrative: RDPU grows ~0.37%/month; equity fund volatile
  const rdpu = [], eq = [];
  const eqPath = [1000, 1032, 985, 1010, 1061, 1012, 968, 1023, 1078, 1040, 1092, 1069, 1110];
  for (let i = 0; i <= 12; i++) rdpu.push(+(1000 * Math.pow(1 + 0.045 / 12, i)).toFixed(1));
  const lbl = ["Awal"].concat(months);
  s.addChart(pres.charts.LINE, [
    { name: "RDPU (ilustrasi)", labels: lbl, values: rdpu },
    { name: "Reksa dana saham (ilustrasi)", labels: lbl, values: eqPath },
  ], {
    x: 0.6, y: 1.45, w: 7.6, h: 4.95, chartColors: [HEX.accent1, HEX.accent4], lineSize: 3, lineDataSymbol: "none",
    showTitle: true, title: "Ilustrasi pergerakan NAB/UP dalam 12 bulan (awal = 1.000)", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.dk1,
    catAxisLabelColor: HEX.accent5, valAxisLabelColor: HEX.accent5, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    valAxisMinVal: 940, valAxisMaxVal: 1140, valGridLine: { color: "E3EAE7", size: 0.5 }, catGridLine: { style: "none" }, valAxisLabelFormatCode: "#,##0",
  });
  const why = [
    ["FaClock", "Bunga dihitung harian", "Bunga deposito dan kupon surat utang diakui sedikit demi sedikit setiap hari, jadi NAB/UP ikut naik pelan."],
    ["FaChartLine", "Tenor pendek, harga stabil", "Surat utang yang jatuh tempo kurang dari setahun harganya nyaris tidak terpengaruh perubahan suku bunga."],
    ["FaTriangleExclamation", "Tetap bisa turun", "Kalau ada penerbit surat utang yang gagal bayar, NAB/UP bisa turun. Jarang, tapi pernah terjadi."],
  ];
  for (let i = 0; i < why.length; i++) {
    const y = 1.5 + i * 1.65;
    shadowCard(s, 8.5, y, 4.25, 1.45, "Alasan " + i);
    await iconCircle(s, 8.7, y + 0.33, 0.7, why[i][0], i === 2 ? C.accent4 : C.accent1, "FFFFFF", why[i][1]);
    txt(s, why[i][1], { x: 9.6, y: y + 0.13, w: 3.0, h: 0.38, fontSize: 13.5, bold: true });
    txt(s, why[i][2], { x: 9.6, y: y + 0.5, w: 3.05, h: 0.92, fontSize: 10.5, color: C.accent5 });
  }
  source(s, "Grafik adalah ilustrasi pola, bukan data produk. RDPU memakai asumsi 4,5% per tahun.");
  notes(s);

  // 14. Alur transaksi
  s = addSlide("LIGHT", "3. Cara kerjanya");
  s.addText("Alur beli dan jual", { placeholder: "title" });
  homeLink(s, false);
  const flow = [
    ["FaMagnifyingGlass", "Pilih produk", "Baca fund fact sheet dan prospektus."],
    ["FaCoins", "Beli (subscription)", "Order sebelum batas waktu, biasanya pukul 13.00 WIB, memakai NAB/UP hari itu."],
    ["FaIdCard", "Unit tercatat", "Unit masuk atas nama SID kamu di KSEI. Ada konfirmasi transaksi."],
    ["FaChartLine", "Pantau", "Nilai = jumlah unit × NAB/UP terbaru."],
    ["FaWallet", "Jual (redemption)", "Aturan OJK: dana dibayar paling lambat 7 hari bursa. RDPU umumnya 1 sampai 2 hari bursa."],
  ];
  s.addShape(pres.ShapeType.line, { x: 1.2, y: 2.35, w: 10.9, h: 0, line: { color: C.accent6, width: 3 }, objectName: "Garis alur" });
  for (let i = 0; i < flow.length; i++) {
    const cx = 0.6 + i * 2.48;
    await iconCircle(s, cx + 0.62, 1.75, 1.15, flow[i][0], C.accent1, "FFFFFF", flow[i][1]);
    numCircle(s, cx + 1.5, 1.6, 0.42, String(i + 1), C.accent2, C.text1, 12);
    shadowCard(s, cx, 3.2, 2.3, 2.15, "Langkah " + (i + 1));
    txt(s, flow[i][1], { x: cx + 0.15, y: 3.35, w: 2.0, h: 0.6, fontSize: 14, bold: true, align: "center" });
    txt(s, flow[i][2], { x: cx + 0.15, y: 3.95, w: 2.0, h: 1.35, fontSize: 12, color: C.accent5, align: "center" });
  }
  card(s, 0.6, 5.65, 12.15, 0.6, C.background2, "Catatan cut off");
  txt(s, "Order setelah batas waktu diproses memakai NAB/UP hari bursa berikutnya. Hari libur bursa tidak dihitung.", { x: 0.85, y: 5.7, w: 11.7, h: 0.5, fontSize: 13, valign: "middle" });
  notes(s);

  // 15. SECTION 4
  pres.addSection({ title: "4. Untung, biaya, pajak, risiko" });
  sectionSlide("4. Untung, biaya, pajak, risiko", "04", "Untung, biaya, pajak, dan risiko", "Bagian yang paling sering disalahpahami. Pelan-pelan saja.");

  // 16. Hasil biaya pajak
  s = addSlide("LIGHT", "4. Untung, biaya, pajak, risiko");
  s.addText("Sumber untung, biaya, dan pajak", { placeholder: "title" });
  homeLink(s, false);
  const tri = [
    ["FaArrowTrendUp", "Sumber keuntungan", [
      "Bunga deposito dan kupon surat utang yang masuk ke portofolio setiap hari.",
      "Keuntungan kamu terlihat dari kenaikan NAB/UP, bukan dari transfer bunga bulanan.",
      "Return 1 tahun RDPU per pertengahan 2026 umumnya di kisaran 4 sampai 5% per tahun, tergantung produk.",
    ]],
    ["FaPercent", "Biaya", [
      "Biaya pengelolaan MI dan biaya kustodian dipotong dari NAB, jadi return yang kamu lihat sudah bersih dari biaya ini.",
      "Biaya beli dan jual RDPU di banyak APERD umumnya Rp0. Tetap cek di prospektus.",
      "Hati-hati biaya transfer bank atau biaya platform, kalau ada.",
    ]],
    ["FaFileInvoiceDollar", "Pajak", [
      "Keuntungan yang diterima pemegang unit reksa dana bukan objek pajak (UU PPh Pasal 4 ayat 3).",
      "Pajak sudah dipotong di dalam portofolio: bunga deposito 20% final, bunga obligasi 10% final (PP 91/2021).",
      "Jadi kamu tidak perlu menyetor pajak lagi atas untung RDPU.",
    ]],
  ];
  for (let i = 0; i < tri.length; i++) {
    const x = 0.6 + i * 4.1;
    shadowCard(s, x, 1.5, 3.85, 4.9, "Kolom " + tri[i][1]);
    await iconCircle(s, x + 0.3, 1.75, 0.8, tri[i][0], i === 2 ? C.accent2 : C.accent1, i === 2 ? HEX.dk1 : "FFFFFF", tri[i][1]);
    txt(s, tri[i][1], { x: x + 1.25, y: 1.95, w: 2.5, h: 0.45, fontSize: 18, bold: true });
    s.addText(tri[i][2].map((t, j) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: j < tri[i][2].length - 1 } })), {
      x: x + 0.3, y: 2.8, w: 3.3, h: 3.5, fontSize: 13.5, color: C.text1, paraSpaceAfter: 8, valign: "top", margin: 0, isTextBox: true,
    });
  }
  source(s, "Sumber: UU PPh Pasal 4 ayat (3); PP 131/2000 (bunga deposito); PP 91/2021 (bunga obligasi); data return RDPU dari Kontan dan Bareksa, 2026.");
  notes(s);

  // 17. Comparison table
  s = addSlide("LIGHT", "4. Untung, biaya, pajak, risiko");
  s.addText("RDPU dibanding tabungan dan deposito", { placeholder: "title" });
  homeLink(s, false);
  const H = (t, col) => ({ text: t, options: { bold: true, color: col || HEX.lt1, fill: { color: HEX.dk1 }, fontSize: 14, align: "center", valign: "middle" } });
  const R = (t, opt) => ({ text: t, options: Object.assign({ fontSize: 12, color: HEX.dk1, valign: "middle" }, opt || {}) });
  const Lb = (t) => R(t, { bold: true, fill: { color: HEX.lt2 } });
  const G = (t) => R(t, { fill: { color: "E3F4EF" } });
  const rows = [
    [H("Aspek"), H("Tabungan"), H("Deposito"), H("RDPU", HEX.accent2)],
    [Lb("Modal awal"), R("Kecil, sesuai syarat bank"), R("Umumnya jutaan rupiah"), G("Mulai sekitar Rp10.000 di banyak aplikasi")],
    [Lb("Imbal hasil"), R("Rendah, sering di bawah 1% per tahun"), R("Bunga tetap selama tenor"), G("Mengikuti isi portofolio, tidak dijamin")],
    [Lb("Pajak"), R("Bunga kena 20% final (saldo di atas Rp7,5 juta)"), R("Bunga kena 20% final"), G("Untung bukan objek pajak bagi investor")],
    [Lb("Pencairan"), R("Kapan saja"), R("Saat jatuh tempo, cair lebih awal biasanya kena penalti"), G("Kapan saja, dana masuk umumnya 1 sampai 2 hari bursa")],
    [Lb("Penjaminan LPS"), R("Ya, sampai Rp2 miliar per nasabah per bank, jika bunga sesuai batas LPS"), R("Ya, dengan syarat yang sama"), R("Tidak. Dilindungi lewat pengawasan OJK dan Bank Kustodian", { fill: { color: "FBE6E5" } })],
    [Lb("Cocok untuk"), R("Transaksi harian"), R("Dana yang pasti tidak dipakai selama tenor"), G("Dana darurat dan tujuan di bawah 1 tahun")],
  ];
  s.addTable(rows, {
    x: 0.6, y: 1.45, w: 12.15, colW: [2.15, 3.2, 3.3, 3.5], rowH: [0.5, 0.72, 0.72, 0.72, 0.8, 0.88, 0.72],
    border: { type: "solid", pt: 1, color: "FFFFFF" }, fontFace: "Calibri", margin: 0.08, objectName: "Tabel perbandingan",
  });
  source(s, "Sumber: LPS (ketentuan penjaminan simpanan), PP 131/2000, prospektus RDPU. Syarat tiap bank dan produk bisa berbeda.");
  notes(s);

  // 18. Simulasi
  s = addSlide("LIGHT", "4. Untung, biaya, pajak, risiko");
  s.addText("Simulasi: Rp10 juta disimpan selama 1 tahun", { placeholder: "title" });
  homeLink(s, false);
  s.addChart(pres.charts.BAR, [{ name: "Keuntungan bersih (Rp ribu)", labels: ["Tabungan (0,5% bersih)", "Deposito (3,75% sebelum pajak)", "RDPU (4,5% bersih)"], values: [50, 300, 450] }], {
    x: 0.6, y: 1.45, w: 7.3, h: 4.95, barDir: "col", chartColors: [HEX.accent5, HEX.accent3, HEX.accent1],
    showTitle: true, title: "Perkiraan keuntungan bersih setelah 1 tahun (Rp ribu)", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelColor: HEX.dk1, dataLabelFormatCode: "#,##0", dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 520, barGapWidthPct: 60,
  });
  card(s, 8.2, 1.5, 4.55, 3.15, C.background2, "Kartu cara hitung");
  txt(s, "Cara menghitungnya", { x: 8.5, y: 1.7, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
  txt(s, [
    { text: "Tabungan: 10 juta × 0,5% = Rp50.000", options: { breakLine: true } },
    { text: "Deposito: 10 juta × 3,75% = Rp375.000, dipotong pajak 20% = Rp300.000", options: { breakLine: true } },
    { text: "RDPU: 10 juta × 4,5% = Rp450.000, tanpa pajak tambahan", options: {} },
  ], { x: 8.5, y: 2.2, w: 4.05, h: 2.4, fontSize: 14.5, paraSpaceAfter: 10 });
  card(s, 8.2, 4.85, 4.55, 1.55, C.text1, "Kartu peringatan");
  txt(s, "Ini ilustrasi dengan asumsi tetap. Return RDPU bisa lebih tinggi atau lebih rendah, dan tidak dijamin.", { x: 8.45, y: 4.95, w: 4.1, h: 0.85, fontSize: 12, color: C.background1 });
  if (SIM_URL) {
    s.addText([{ text: "Coba simulator interaktif", options: { hyperlink: { url: SIM_URL, tooltip: "Buka simulator di browser" } } }], {
      x: 8.45, y: 5.65, w: 4.1, h: 0.35, fontSize: 13, bold: true, color: C.accent2, margin: 0, isTextBox: true, objectName: "Link simulator",
    });
    txt(s, "Versi offline: buka file Simulator_RDPU.html di browser.", { x: 8.45, y: 6.0, w: 4.1, h: 0.3, fontSize: 10.5, color: C.accent6 });
  } else {
    txt(s, "Coba sendiri di file Simulator_RDPU.html", { x: 8.45, y: 5.8, w: 4.1, h: 0.45, fontSize: 13, bold: true, color: C.accent2 });
  }
  source(s, "Asumsi ilustrasi: tabungan 0,5% bersih, deposito 3,75% (setara batas penjaminan LPS Juli sampai September 2026), RDPU 4,5% bersih.");
  notes(s);

  // 19. Risiko
  s = addSlide("LIGHT", "4. Untung, biaya, pajak, risiko");
  s.addText("Risikonya rendah, tapi bukan nol", { placeholder: "title" });
  homeLink(s, false);
  const risks = [
    ["FaArrowTrendUp", "Imbal hasil tidak pasti", "Saat suku bunga turun, return RDPU ikut turun pelan-pelan."],
    ["FaBan", "Gagal bayar penerbit", "Jika penerbit obligasi di portofolio gagal bayar, NAB/UP bisa turun."],
    ["FaDroplet", "Risiko likuiditas", "Kalau penarikan dana besar terjadi serentak, MI harus menjual aset lebih cepat."],
    ["FaUmbrella", "Tidak dijamin LPS", "Berbeda dari deposito. Perlindungannya lewat aturan OJK dan pemisahan aset di kustodian."],
    ["FaGavel", "Pembubaran", "Reksa dana wajib bubar jika NAB di bawah Rp10 miliar selama 120 hari bursa berturut-turut. Dana dikembalikan sesuai hasil likuidasi."],
    ["FaLandmark", "Perubahan aturan", "Misalnya aturan pajak berubah, sehingga hasil bersih ikut berubah."],
  ];
  for (let i = 0; i < risks.length; i++) {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.1, y = 1.5 + row * 2.45;
    shadowCard(s, x, y, 3.85, 2.2, "Risiko " + risks[i][1]);
    await iconCircle(s, x + 0.25, y + 0.25, 0.7, risks[i][0], C.accent4, "FFFFFF", risks[i][1]);
    txt(s, risks[i][1], { x: x + 1.1, y: y + 0.35, w: 2.6, h: 0.5, fontSize: 15, bold: true, valign: "middle" });
    txt(s, risks[i][2], { x: x + 0.25, y: y + 1.1, w: 3.4, h: 1.05, fontSize: 12, color: C.accent5 });
  }
  source(s, "Sumber: POJK 23/POJK.04/2016 Pasal 45; bagian faktor risiko pada prospektus RDPU.");
  notes(s);

  // 20. BI Rate context
  s = addSlide("LIGHT", "4. Untung, biaya, pajak, risiko");
  s.addText("2026: suku bunga naik, apa efeknya ke RDPU?", { placeholder: "title" });
  homeLink(s, false);
  const biLabels = [], biVals = [];
  const mn = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
  const bi = {
    2024: [6.00, 6.00, 6.00, 6.25, 6.25, 6.25, 6.25, 6.25, 6.00, 6.00, 6.00, 6.00],
    2025: [5.75, 5.75, 5.75, 5.75, 5.50, 5.50, 5.25, 5.00, 4.75, 4.75, 4.75, 4.75],
    2026: [4.75, 4.75, 4.75, 4.75, 5.25, 5.75, 5.75, 5.75, 5.75],
  };
  for (const y of [2024, 2025, 2026]) bi[y].forEach((v, i) => { biLabels.push(mn[i] + " " + String(y).slice(2)); biVals.push(v); });
  s.addChart(pres.charts.LINE, [{ name: "BI Rate (%)", labels: biLabels, values: biVals }], {
    x: 0.6, y: 1.45, w: 7.6, h: 4.95, chartColors: [HEX.accent1], lineSize: 3, lineDataSymbol: "none",
    showTitle: true, title: "BI Rate akhir bulan, Januari 2024 sampai September 2026 (%)", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    showLegend: false, catAxisLabelColor: HEX.accent5, valAxisLabelColor: HEX.accent5, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    catAxisLabelFontSize: 10, catAxisLabelFrequency: 3, valAxisMinVal: 4.5, valAxisMaxVal: 6.5, valAxisMajorUnit: 0.5, valAxisLabelFormatCode: "0.00",
    valGridLine: { color: "E3EAE7", size: 0.5 }, catGridLine: { style: "none" },
  });
  const ctx = [
    ["2025", "BI memangkas suku bunga 5 kali, dari 6,00% ke 4,75%. Return RDPU ikut melandai."],
    ["Mei sampai Juni 2026", "BI menaikkan BI Rate total 100 bps ke 5,75% untuk menjaga rupiah. Ditahan sampai September 2026."],
    ["Artinya", "RDPU menempatkan ulang dana ke deposito dan surat utang baru dengan bunga lebih tinggi, jadi return-nya cenderung ikut naik dengan jeda. Reksa dana pendapatan tetap justru bisa tertekan karena harga obligasi turun saat bunga naik."],
  ];
  let cy = 1.5;
  const ch = [1.05, 1.25, 2.3];
  ctx.forEach(([t, d], i) => {
    card(s, 8.5, cy, 4.25, ch[i], i === 2 ? C.text1 : C.background2, "Konteks " + t);
    txt(s, t, { x: 8.75, y: cy + 0.12, w: 3.8, h: 0.35, fontSize: 13.5, bold: true, color: i === 2 ? C.accent2 : C.accent1 });
    txt(s, d, { x: 8.75, y: cy + 0.47, w: 3.8, h: ch[i] - 0.52, fontSize: 11.5, color: i === 2 ? C.background1 : C.text1 });
    cy += ch[i] + 0.15;
  });
  source(s, "Sumber: Bank Indonesia, hasil RDG 2024 sampai September 2026 (dirangkum dari Bisnis.com, Kontan, Investor Trust).");
  notes(s);

  // 21. SECTION 5
  pres.addSection({ title: "5. Praktik" });
  sectionSlide("5. Praktik", "05", "Dari paham ke praktik", "Cara memilih produk, mengenali investasi bodong, dan memulai dengan aman.");

  // 22. Checklist memilih
  s = addSlide("LIGHT", "5. Praktik");
  s.addText("Enam hal yang dicek sebelum memilih RDPU", { placeholder: "title" });
  homeLink(s, false);
  const checks = [
    ["Legal", "Manajer Investasi dan APERD terdaftar dan diawasi OJK. Cek di situs OJK."],
    ["Dana kelolaan", "Pilih yang AUM-nya besar dan stabil, jauh di atas batas bubar Rp10 miliar."],
    ["Konsistensi kinerja", "Lihat return 1 tahun dan 3 tahun. Yang stabil lebih penting dari yang paling tinggi sesaat."],
    ["Isi portofolio", "Baca fund fact sheet: ke bank atau penerbit mana saja dananya ditempatkan."],
    ["Biaya dan pencairan", "Cek biaya pengelolaan, biaya beli dan jual, serta berapa hari dana cair."],
    ["Konvensional atau syariah", "Sesuaikan dengan preferensi kamu. Prinsip kerja dan risikonya mirip."],
  ];
  for (let i = 0; i < checks.length; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.15, y = 1.5 + row * 1.65;
    shadowCard(s, x, y, 5.9, 1.45, "Cek " + checks[i][0]);
    await iconCircle(s, x + 0.3, y + 0.35, 0.75, "FaCheck", C.accent1, "FFFFFF", "Centang " + i);
    txt(s, checks[i][0], { x: x + 1.3, y: y + 0.18, w: 4.4, h: 0.4, fontSize: 15, bold: true });
    txt(s, checks[i][1], { x: x + 1.3, y: y + 0.58, w: 4.4, h: 0.8, fontSize: 12.5, color: C.accent5 });
  }
  source(s, "Fund fact sheet terbit tiap bulan di situs Manajer Investasi atau di aplikasi APERD.");
  notes(s);

  // 23. Waspada bodong
  s = addSlide("DARK", "5. Praktik");
  s.addText("Waspada janji untung pasti", { placeholder: "title" });
  homeLink(s, true);
  card(s, 0.6, 1.5, 6.0, 4.85, C.text2, "Kartu kasus");
  await iconCircle(s, 0.9, 1.8, 0.85, "FaTriangleExclamation", C.accent4, "FFFFFF", "Peringatan");
  txt(s, "Pelajaran dari kasus nyata", { x: 1.95, y: 1.95, w: 4.5, h: 0.5, fontSize: 18, bold: true, color: C.accent2 });
  txt(s, "Pada November 2019, OJK memerintahkan PT Minna Padi Aset Manajemen membubarkan enam reksa dananya. Salah satu alasannya: produk tersebut dipasarkan dengan janji imbal hasil pasti, padahal reksa dana tidak boleh menjanjikan hasil tetap.", { x: 0.9, y: 2.85, w: 5.4, h: 2.2, fontSize: 14, color: C.background1 });
  txt(s, "Kalau ada yang bilang \"RDPU pasti 10% setahun\", itu tanda bahaya.", { x: 0.9, y: 5.2, w: 5.4, h: 0.9, fontSize: 14, bold: true, color: C.accent6 });
  const flags = [
    ["FaXmark", "Janji untung tetap dan tinggi", C.accent4],
    ["FaXmark", "Transfer ke rekening pribadi, bukan ke rekening reksa dana di bank kustodian", C.accent4],
    ["FaXmark", "Perusahaan atau aplikasi tidak terdaftar di OJK", C.accent4],
    ["FaCheck", "Cek legalitas di situs OJK atau Kontak OJK 157", C.accent1],
    ["FaCheck", "Laporkan tawaran mencurigakan ke Satgas PASTI (Pemberantasan Aktivitas Keuangan Ilegal)", C.accent1],
  ];
  for (let i = 0; i < flags.length; i++) {
    const y = 1.5 + i * 0.99;
    card(s, 6.95, y, 5.8, 0.82, C.text2, "Tanda " + i);
    await iconCircle(s, 7.15, y + 0.16, 0.5, flags[i][0], flags[i][2], "FFFFFF", "Tanda ikon " + i);
    txt(s, flags[i][1], { x: 7.85, y: y + 0.06, w: 4.75, h: 0.7, fontSize: 12.5, color: C.background1, valign: "middle" });
  }
  source(s, "Sumber: siaran pers OJK tentang pembubaran reksa dana Minna Padi (2019); OJK.", true);
  notes(s);

  // 24. Cara mulai
  s = addSlide("LIGHT", "5. Praktik");
  s.addText("Cara mulai dalam lima langkah", { placeholder: "title" });
  homeLink(s, false);
  const start = [
    ["FaIdCard", "Siapkan data", "KTP, rekening bank atas nama sendiri, dan NPWP jika punya."],
    ["FaMobileScreen", "Pilih APERD", "Bank atau aplikasi investasi yang terdaftar di OJK."],
    ["FaListCheck", "Isi profil risiko", "Jawab kuesioner dengan jujur. RDPU cocok untuk profil konservatif."],
    ["FaCoins", "Beli pertama", "Mulai dari nominal kecil, banyak aplikasi menerima mulai Rp10.000."],
    ["FaSeedling", "Rutin dan sabar", "Tambah berkala (misalnya tiap bulan) dan cek konfirmasi transaksi dari KSEI."],
  ];
  for (let i = 0; i < start.length; i++) {
    const y = 1.45 + i * 1.03;
    numCircle(s, 0.6, y + 0.12, 0.65, String(i + 1), C.accent1, C.background1, 18);
    shadowCard(s, 1.45, y, 6.4, 0.88, "Langkah mulai " + (i + 1));
    await iconCircle(s, 1.6, y + 0.14, 0.6, start[i][0], C.background2, HEX.accent1, start[i][1]);
    txt(s, start[i][1], { x: 2.35, y: y + 0.08, w: 5.3, h: 0.36, fontSize: 14, bold: true });
    txt(s, start[i][2], { x: 2.35, y: y + 0.44, w: 5.4, h: 0.4, fontSize: 11.5, color: C.accent5 });
  }
  card(s, 8.2, 1.45, 4.55, 4.95, C.text1, "Kartu tips");
  await iconCircle(s, 8.5, 1.75, 0.8, "FaGraduationCap", C.accent2, HEX.dk1, "Tips mahasiswa");
  txt(s, "Tips untuk mahasiswa", { x: 8.5, y: 2.7, w: 4, h: 0.45, fontSize: 18, bold: true, color: C.accent2 });
  s.addText([
    "Sisihkan di awal, bukan dari sisa uang bulanan.",
    "Mulai dari Rp50.000 sampai Rp100.000 per bulan sudah cukup untuk membangun kebiasaan.",
    "Pisahkan RDPU untuk dana darurat dari uang jajan.",
    "Jangan pakai uang pinjaman untuk berinvestasi.",
  ].map((t, j, a) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: j < a.length - 1 } })), {
    x: 8.5, y: 3.25, w: 3.95, h: 3.0, fontSize: 13, color: C.background1, paraSpaceAfter: 8, valign: "top", margin: 0, isTextBox: true,
  });
  notes(s);

  // 25. Use cases
  s = addSlide("LIGHT", "5. Praktik");
  s.addText("Kapan RDPU tepat dipakai, dan kapan tidak", { placeholder: "title" });
  homeLink(s, false);
  const good = [
    ["FaUmbrella", "Dana darurat", "Bisa dicairkan cepat dan nilainya stabil."],
    ["FaGraduationCap", "Tujuan di bawah 1 tahun", "Contoh: biaya KKN, laptop, atau UKT semester depan."],
    ["FaWallet", "Tempat parkir sementara", "Menampung dana sebelum dipindah ke instrumen lain."],
    ["FaSeedling", "Belajar berinvestasi", "Risiko kecil, cocok untuk membiasakan diri membaca NAB dan laporan."],
  ];
  txt(s, "Tepat dipakai untuk", { x: 0.6, y: 1.4, w: 7, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
  for (let i = 0; i < good.length; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 3.95, y = 1.9 + row * 2.2;
    shadowCard(s, x, y, 3.7, 2.0, "Kegunaan " + good[i][1]);
    await iconCircle(s, x + 0.25, y + 0.25, 0.7, good[i][0], C.accent1, "FFFFFF", good[i][1]);
    txt(s, good[i][1], { x: x + 1.1, y: y + 0.3, w: 2.5, h: 0.6, fontSize: 14.5, bold: true, valign: "middle" });
    txt(s, good[i][2], { x: x + 0.25, y: y + 1.05, w: 3.25, h: 0.9, fontSize: 12, color: C.accent5 });
  }
  card(s, 8.65, 1.4, 4.1, 5.0, C.background2, "Kartu kurang tepat");
  txt(s, "Kurang tepat untuk", { x: 8.95, y: 1.6, w: 3.6, h: 0.4, fontSize: 16, bold: true, color: C.accent4 });
  s.addText([
    "Tujuan jangka panjang di atas 5 tahun, seperti dana pensiun. Return-nya bisa kalah dari instrumen yang lebih agresif.",
    "Mengejar untung besar dalam waktu singkat.",
    "Pengganti tabungan untuk transaksi harian, karena pencairannya butuh 1 sampai 2 hari bursa.",
  ].map((t, j, a) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: j < a.length - 1 } })), {
    x: 8.95, y: 2.15, w: 3.55, h: 4.1, fontSize: 13, color: C.text1, paraSpaceAfter: 10, valign: "top", margin: 0, isTextBox: true,
  });
  notes(s);

  // 26. SECTION 6
  pres.addSection({ title: "6. Kuis" });
  sectionSlide("6. Kuis", "06", "Kuis interaktif", "Tiga soal. Klik jawaban pilihanmu untuk melihat pembahasannya.");

  // quiz
  const quiz = [
    {
      q: "RDPU hanya boleh membeli efek utang yang jatuh tempo atau sisa jatuh temponya maksimal...",
      opts: ["A.  3 tahun", "B.  1 tahun", "C.  5 tahun"], correct: 1,
      exp: "Jawaban B, 1 tahun. Batas tenor pendek inilah yang membuat NAB/UP RDPU stabil dan risikonya paling rendah di antara jenis reksa dana lain.",
    },
    {
      q: "Kamu membeli RDPU senilai Rp1.000.000 saat NAB/UP Rp1.250. Berapa unit yang kamu dapat?",
      opts: ["A.  800 unit", "B.  1.250 unit", "C.  125 unit"], correct: 0,
      exp: "Jawaban A, 800 unit. Rumusnya: nominal beli dibagi NAB/UP, yaitu 1.000.000 ÷ 1.250 = 800 unit.",
    },
    {
      q: "Mana pernyataan yang BENAR tentang RDPU?",
      opts: ["A.  Dijamin LPS seperti deposito", "B.  Untung yang diterima investor bukan objek pajak", "C.  Menjamin return tetap 6% per tahun"], correct: 1,
      exp: "Jawaban B. RDPU tidak dijamin LPS (A salah) dan tidak boleh menjanjikan return tetap (C salah). Pajak sudah dipotong di dalam portofolio.",
    },
  ];
  const qTargets = [[SL.a1, SL.q2], [SL.a2, SL.q3], [SL.a3, 33]];
  for (let k = 0; k < quiz.length; k++) {
    const qz = quiz[k];
    s = addSlide("DARK", "6. Kuis");
    s.addText("Soal " + (k + 1) + " dari 3", { placeholder: "title" });
    homeLink(s, true);
    txt(s, qz.q, { x: 0.6, y: 1.5, w: 12.1, h: 1.2, fontSize: 24, bold: true, color: C.background1, fontFace: THEME.headFontFace });
    qz.opts.forEach((o, i) => {
      const y = 2.95 + i * 1.1;
      s.addText([{ text: o, options: { hyperlink: { slide: qTargets[k][0], tooltip: "Lihat pembahasan" } } }], {
        x: 0.6, y, w: 8.5, h: 0.85, shape: pres.ShapeType.roundRect, rectRadius: 0.12, fill: { color: C.text2 }, line: { color: C.accent1, width: 1.25 },
        fontSize: 18, bold: true, color: C.background1, margin: [20, 10, 0, 0], valign: "middle", isTextBox: true, objectName: "Pilihan " + (k + 1) + "." + (i + 1),
      });
    });
    await iconCircle(s, 10.0, 3.0, 2.2, "FaCircleQuestion", C.accent2, HEX.dk1, "Ikon soal");
    txt(s, "Diskusikan dulu, lalu klik jawaban.", { x: 9.4, y: 5.4, w: 3.4, h: 0.5, fontSize: 12, color: C.accent6, align: "center" });
    notes(s);

    s = addSlide("LIGHT", "6. Kuis");
    s.addText("Pembahasan soal " + (k + 1), { placeholder: "title" });
    homeLink(s, false);
    qz.opts.forEach((o, i) => {
      const y = 1.5 + i * 0.95, ok = i === qz.correct;
      s.addShape(pres.ShapeType.roundRect, { x: 0.6, y, w: 7.4, h: 0.78, rectRadius: 0.12, fill: { color: ok ? C.accent1 : C.background2 }, line: { color: ok ? C.accent1 : C.background2, width: 0 }, objectName: "Opsi " + i });
      txt(s, o, { x: 0.85, y, w: 6.45, h: 0.78, fontSize: 14, bold: ok, color: ok ? C.background1 : C.accent5, valign: "middle" });
    });
    const okIdx = qz.correct;
    s.addImage({ data: await icon("FaCheck", "FFFFFF"), x: 7.4, y: 1.5 + okIdx * 0.95 + 0.2, w: 0.38, h: 0.38, objectName: "Tanda benar" });
    card(s, 8.35, 1.5, 4.4, 2.75, C.text1, "Kartu pembahasan");
    txt(s, qz.exp, { x: 8.6, y: 1.7, w: 3.95, h: 2.4, fontSize: 15, color: C.background1 });
    const nextLabel = k < 2 ? "Lanjut ke soal " + (k + 2) : "Lanjut ke rangkuman";
    s.addText([{ text: nextLabel, options: { hyperlink: { slide: qTargets[k][1], tooltip: nextLabel } } }], {
      x: 0.6, y: 4.65, w: 3.6, h: 0.6, shape: pres.ShapeType.roundRect, rectRadius: 0.3, fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 },
      fontSize: 15, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Tombol lanjut",
    });
    s.addText([{ text: "Ulangi soal", options: { hyperlink: { slide: [SL.q1, SL.q2, SL.q3][k], tooltip: "Kembali ke soal" } } }], {
      x: 4.45, y: 4.65, w: 2.4, h: 0.6, shape: pres.ShapeType.roundRect, rectRadius: 0.3, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.25 },
      fontSize: 15, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Tombol ulangi",
    });
    notes(s);
  }

  // 33. Rangkuman
  pres.addSection({ title: "Penutup" });
  s = addSlide("LIGHT", "Penutup");
  s.addText("Rangkuman: enam hal yang perlu diingat", { placeholder: "title" });
  homeLink(s, false);
  const sum = [
    "RDPU menaruh 100% dana di instrumen pasar uang dan surat utang bertenor maksimal 1 tahun.",
    "Nilai investasimu = jumlah unit × NAB/UP. Unitnya tetap, NAB/UP yang bergerak.",
    "Untung bagi investor bukan objek pajak. Biaya pengelolaan sudah dipotong dari NAB.",
    "Risikonya rendah tapi bukan nol, dan RDPU tidak dijamin LPS.",
    "Cocok untuk dana darurat dan tujuan di bawah 1 tahun.",
    "Hanya beli lewat pihak yang terdaftar di OJK. Janji untung pasti adalah tanda bahaya.",
  ];
  sum.forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.15, y = 1.5 + row * 1.65;
    shadowCard(s, x, y, 5.9, 1.45, "Rangkuman " + (i + 1));
    numCircle(s, x + 0.3, y + 0.38, 0.7, String(i + 1), C.accent1, C.background1, 18);
    txt(s, t, { x: x + 1.25, y: y + 0.15, w: 4.45, h: 1.15, fontSize: 14, valign: "middle" });
  });
  notes(s);

  // 34. Sumber
  s = addSlide("LIGHT", "Penutup");
  s.addText("Sumber dan bacaan lanjutan", { placeholder: "title" });
  homeLink(s, false);
  const src = [
    ["Regulasi", [
      "UU No. 8 Tahun 1995 tentang Pasar Modal",
      "POJK 23/POJK.04/2016 tentang Reksa Dana berbentuk KIK, beserta perubahannya",
      "UU Pajak Penghasilan Pasal 4 ayat (3); PP 131/2000; PP 91/2021",
    ]],
    ["Data", [
      "KSEI, Statistik Publik Juli 2026 (investor dan AUM per jenis reksa dana)",
      "Bank Indonesia, hasil Rapat Dewan Gubernur 2024 sampai September 2026",
      "LPS, Tingkat Bunga Penjaminan periode 1 Juli sampai 30 September 2026",
    ]],
    ["Belajar mandiri", [
      "ojk.go.id dan sikapiuangmu.ojk.go.id (edukasi resmi OJK)",
      "Fund fact sheet dan prospektus di situs Manajer Investasi",
      "akses.ksei.co.id untuk cek kepemilikan unit atas nama sendiri",
    ]],
  ];
  for (let i = 0; i < src.length; i++) {
    const x = 0.6 + i * 4.1;
    shadowCard(s, x, 1.5, 3.85, 4.3, "Sumber " + src[i][0]);
    txt(s, src[i][0], { x: x + 0.3, y: 1.75, w: 3.3, h: 0.45, fontSize: 18, bold: true, color: C.accent1 });
    s.addText(src[i][1].map((t, j, a) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: j < a.length - 1 } })), {
      x: x + 0.3, y: 2.35, w: 3.3, h: 3.3, fontSize: 14, color: C.text1, paraSpaceAfter: 10, valign: "top", margin: 0, isTextBox: true,
    });
  }
  txt(s, "Angka suku bunga dan data pasar berubah. Sebelum presentasi berikutnya, perbarui slide 4, 5, 9, dan 20 dengan data terbaru.", { x: 0.6, y: 6.1, w: 12, h: 0.4, fontSize: 12, italic: true, color: C.accent5 });
  notes(s);

  // 35. Closing
  s = addSlide("COVER", "Penutup");
  txt(s, "TERIMA KASIH", { x: 0.8, y: 1.6, w: 8, h: 0.4, fontSize: 13, bold: true, color: C.accent2, charSpacing: 3 });
  s.addText("Mulai kecil, yang penting paham.", { placeholder: "title" });
  s.addText("Investasi terbaik untuk pemula adalah yang kamu mengerti cara kerjanya. Pertanyaan dan diskusi dipersilakan.", { placeholder: "body" });
  await iconCircle(s, 10.55, 0.75, 1.6, "FaSeedling", C.accent2, HEX.dk1, "Tunas");
  notes(s);

  if (slideNo !== 35) throw new Error("Expected 35 slides, got " + slideNo);
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT, slideNo);
})().catch((e) => { console.error(e); process.exit(1); });
