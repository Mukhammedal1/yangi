// Excel -> etiketka -> PDF
// Mahalliy:  npm start   ->  http://localhost:3000
// Vercel:    api/index.js orqali chaqiriladi

const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");
const bwipjs = require("bwip-js");
const { Writable } = require("stream");
const fs = require("fs");
const path = require("path");

// ------------------------------------------------------------- SOZLAMALAR

const PORT = process.env.PORT || 3000;
const MM = 2.834645; // 1 mm = 2.834645 pt

const EN = 40; // qog'oz eni, mm (oynadan o'zgartiriladi)
const BOYI = 38; // qog'oz bo'yi, mm
const ORALIQ = 1; // qatorlar orasidagi masofa, mm

// Barcode bloki — frontenddagi qiymatlar bilan BIR XIL bo'lishi shart
const RAQAM_H = 3.5; // barcode tagidagi raqam balandligi, mm
const BAR_ORALIQ = 1.5; // barcode bilan raqam orasi, mm
const PAST_CHET = 1.5; // blokdan pastdagi bo'sh joy, mm

// Excel sarlavhalarida qidiriladigan so'zlar (faqat taxmin uchun)
const NOM_KEYS = ["nomi", "tovar", "mahsulot", "наименование", "название", "товар", "name"];
const NARX_KEYS = ["sotuv", "narx", "розничная", "цена", "price"];
const KOD_KEYS = ["shtrix", "штрих", "barcode", "kod", "код"];

// ----------------------------------------------------------------- FONT
//
// MUHIM: font repo ichidagi fonts/ papkadan olinadi.
//
// npm paketidan require.resolve() bilan olish Vercel da ISHLAMAYDI:
// bundler faqat statik require() larni kuzatadi, dinamik topilgan .ttf
// fayl bundle ga tushmaydi va fs.existsSync false qaytaradi. Keyin
// PDFKit Helvetica ga tushadi, uning .cjs fayli ham bundle da yo'q va
// "Cannot find module .../standard-fonts/Helvetica.cjs" xatosi chiqadi.
//
// vercel.json dagi includeFiles ham fonts/** ni o'z ichiga olishi shart.

const FONT_DIR = path.join(__dirname, "public", "fonts");

const topFont = (yollar, zahira) =>
  yollar.filter(Boolean).find(fs.existsSync) || zahira;

const FONT = topFont(
  [
    path.join(FONT_DIR, "DejaVuSans.ttf"),
    "C:\\Windows\\Fonts\\arial.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
  ],
  "Helvetica"
);

const FONT_B = topFont(
  [
    path.join(FONT_DIR, "DejaVuSans-Bold.ttf"),
    "C:\\Windows\\Fonts\\arialbd.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
  ],
  "Helvetica-Bold"
);

// ------------------------------------------------------------------ EXCEL

// Kalit so'zlarni berilgan tartibda qidiradi: "sotuv" "narx" dan ustun turadi
function ustunTopish(sarlavha, kalitlar) {
  for (const k of kalitlar) {
    for (let i = 0; i < sarlavha.length; i++) {
      const s = String(sarlavha[i] || "").toLowerCase();
      if (s && s.includes(k)) return i;
    }
  }
  return -1;
}

function excelOqish(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const hammasi = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  if (!hammasi.length) return { ustunlar: [], qatorlar: [], taxmin: {} };

  // sarlavha qatorini topamiz
  let bosh = 0;
  for (let i = 0; i < Math.min(10, hammasi.length); i++) {
    if (ustunTopish(hammasi[i] || [], NOM_KEYS) !== -1) {
      bosh = i;
      break;
    }
  }

  const sarlavha = hammasi[bosh] || [];
  const ustunlar = sarlavha.map((u, i) => String(u || "").trim() || `Ustun ${i + 1}`);

  const qatorlar = [];
  for (const q of hammasi.slice(bosh + 1)) {
    const qator = ustunlar.map((_, i) =>
      q[i] === undefined || q[i] === null ? "" : String(q[i]).trim()
    );
    if (qator.some((x) => x !== "")) qatorlar.push(qator);
  }

  let taxmin = {
    nom: ustunTopish(sarlavha, NOM_KEYS),
    narx: ustunTopish(sarlavha, NARX_KEYS),
    kod: ustunTopish(sarlavha, KOD_KEYS),
  };
  if (taxmin.nom === -1) taxmin = { nom: 0, narx: 1, kod: 2 };

  return { ustunlar, qatorlar, taxmin };
}

// --------------------------------------------------------------- ETIKETKA

// Matnni berilgan enga sig'adigan satrlarga bo'ladi
function satrlarga(doc, matn, kenglik, olcham, qalin) {
  doc.font(qalin ? FONT_B : FONT).fontSize(olcham);
  const satrlar = [];
  let joriy = "";
  for (const s of String(matn).split(/\s+/)) {
    const sinov = (joriy + " " + s).trim();
    if (!joriy || doc.widthOfString(sinov) <= kenglik) joriy = sinov;
    else {
      satrlar.push(joriy);
      joriy = s;
    }
  }
  if (joriy) satrlar.push(joriy);
  return satrlar;
}

async function shtrixRasm(kod) {
  const toza = String(kod || "").replace(/[^a-zA-Z0-9]/g, "");
  if (!toza) return null;

  let bcid = "code128";
  if (/^\d{13}$/.test(toza)) bcid = "ean13";
  else if (/^\d{8}$/.test(toza)) bcid = "ean8";

  const chiz = (turi) =>
    bwipjs.toBuffer({
      bcid: turi,
      text: toza,
      scale: 4,
      height: 10,
      includetext: false,
      paddingwidth: 0,
      paddingheight: 0,
    });

  try {
    return await chiz(bcid);
  } catch (e) {
    // EAN nazorat raqami noto'g'ri bo'lsa - Code128 qilib chizamiz
    if (bcid === "code128") return null;
    try {
      return await chiz("code128");
    } catch (e2) {
      return null;
    }
  }
}

// Bitta satrni chizadi. qoshimcha (masalan "so'm") kichikroq shriftda yoziladi.
function satrChizish(doc, matn, qoshimcha, x, y, kenglik, olcham, qalin) {
  doc.font(qalin ? FONT_B : FONT).fontSize(olcham);

  if (!qoshimcha) {
    doc.text(matn, x, y, { width: kenglik, align: "center", lineBreak: false });
    return;
  }

  const kichik = olcham * 0.55;
  const w1 = doc.widthOfString(matn);
  doc.fontSize(kichik);
  const w2 = doc.widthOfString(" " + qoshimcha);

  const bosh = x + (kenglik - w1 - w2) / 2;
  doc.fontSize(olcham).text(matn, bosh, y, { lineBreak: false });
  doc
    .fontSize(kichik)
    .text(" " + qoshimcha, bosh + w1, y + (olcham - kichik) * 0.85, {
      lineBreak: false,
    });
}

// ------------------------------------------------------------------- PDF
//
// Joylashuv frontenddagi geometriya() bilan bir xil:
//
//   [ matn bloki  ]  <- qolgan joy, vertikal o'rtada, sig'masa kichrayadi
//   [ BARCODE     ]
//   [ kod raqami  ]
//   [ pastki chet ]

async function pdfYasash(mahsulotlar, stream, sozlama = {}) {
  const enMM = Number(sozlama.en) || EN;
  const boyiMM = Number(sozlama.boyi) || BOYI;
  const oraliq = (sozlama.oraliq == null ? ORALIQ : Number(sozlama.oraliq)) * MM;

  const W = enMM * MM;
  const H = boyiMM * MM;
  const chetMM = Math.min(2, boyiMM * 0.06);
  const chet = chetMM * MM;
  const ich = W - 2 * chet;

  const barKor = !!sozlama.bar;
  const barHMM = barKor ? Math.max(5, boyiMM * 0.28) : 0;
  const barH = barHMM * MM;

  const blokH = barKor ? (barHMM + BAR_ORALIQ + RAQAM_H + PAST_CHET) * MM : 0;
  const matnH = Math.max(0, H - 2 * chet - blokH);

  const doc = new PDFDocument({ size: [W, H], margin: 0 });
  doc.pipe(stream);

  for (let i = 0; i < mahsulotlar.length; i++) {
    if (i > 0) doc.addPage({ size: [W, H], margin: 0 });

    const m = mahsulotlar[i] || {};
    const kiruvchi = (m.qatorlar || []).filter((q) => String(q.matn).trim() !== "");

    // hamma satr sig'maguncha shriftlarni birdek kichraytiramiz
    let tayyor = [];
    let balandlik = 0;
    for (let k = 1; k >= 0.4; k -= 0.05) {
      tayyor = [];
      balandlik = 0;
      for (const q of kiruvchi) {
        const olcham = Math.max(4, q.olcham * k);
        const satrlar = satrlarga(doc, q.matn, ich, olcham, q.qalin);
        tayyor.push({ satrlar, olcham, qoshimcha: q.qoshimcha, qalin: q.qalin });
        balandlik += satrlar.length * olcham * 1.2;
      }
      balandlik += Math.max(0, tayyor.length - 1) * oraliq;
      if (balandlik <= matnH) break;
    }

    // matn blokini matn joyi ichida vertikal o'rtaga qo'yamiz
    let y = chet + Math.max(0, (matnH - balandlik) / 2);
    for (const q of tayyor) {
      q.satrlar.forEach((satr, n) => {
        const oxirgi = n === q.satrlar.length - 1;
        satrChizish(doc, satr, oxirgi ? q.qoshimcha : "", chet, y, ich, q.olcham, q.qalin);
        y += q.olcham * 1.2;
      });
      y += oraliq;
    }

    // barcode + tagidagi raqam
    if (barKor) {
      const rasm = await shtrixRasm(m.kod);
      if (rasm) {
        const raqamY = H - PAST_CHET * MM - RAQAM_H * MM;
        const barY = raqamY - BAR_ORALIQ * MM - barH;

        doc.image(rasm, chet, barY, { width: ich, height: barH });

        doc.font(FONT).fontSize(RAQAM_H * MM * 0.8);
        doc.text(String(m.kod || ""), chet, raqamY, {
          width: ich,
          align: "center",
          lineBreak: false,
        });
      }
    }
  }

  doc.end();
  return new Promise((res) => stream.on("finish", res));
}

// Lambda javobni stream qilmaydi — PDF ni xotirada yig'amiz
function pdfBuffer(mahsulotlar, sozlama) {
  return new Promise((res, rej) => {
    const bolaklar = [];
    const yigib = new Writable({
      write(b, e, cb) {
        bolaklar.push(b);
        cb();
      },
    });
    yigib.on("finish", () => res(Buffer.concat(bolaklar)));
    pdfYasash(mahsulotlar, yigib, sozlama).catch(rej);
  });
}

// -------------------------------------------------------------- SERVER

const app = express();
app.use(express.json({ limit: "20mb" }));
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({ storage: multer.memoryStorage() });

app.post("/api/excel", upload.single("fayl"), (req, res) => {
  try {
    res.json(excelOqish(req.file.buffer));
  } catch (e) {
    console.error("Excel xato:", e);
    res.status(400).json({ xato: e.message });
  }
});

// Ro'yxatni PDF qilib qaytaradi
app.post("/api/pdf", async (req, res) => {
  try {
    const buf = await pdfBuffer(req.body.mahsulotlar || [], req.body.sozlama);
    res.setHeader("Content-Type", "application/pdf");
    res.end(buf);
  } catch (e) {
    console.error("PDF xato:", e);
    res.status(500).json({ xato: e.message });
  }
});

// Bitta shtrix kodni PNG qilib qaytaradi (preview uchun)
app.get("/api/shtrix", async (req, res) => {
  try {
    const rasm = await shtrixRasm(String(req.query.kod || ""));
    if (!rasm) return res.status(404).end();
    res.setHeader("Content-Type", "image/png");
    res.setHeader("Cache-Control", "max-age=3600");
    res.end(rasm);
  } catch (e) {
    console.error("Shtrix xato:", e);
    res.status(500).end();
  }
});

if (require.main === module) {
  app.listen(PORT, () => console.log(`http://localhost:${PORT}`));
}

module.exports = app;
