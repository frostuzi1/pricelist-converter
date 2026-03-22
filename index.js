require("dotenv").config();

const path = require("path");
const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const { GoogleGenerativeAI } = require("@google/generative-ai");

const PORT = Number(process.env.PORT) || 3000;
const MODEL_ID = "gemini-3.1-flash-lite-preview";

const SYSTEM_INSTRUCTION = `You are an OCR expert specializing in extraction from structured tables (invoices). Look at the provided image. Ignore all handwritten marks, ink circles, or arrows. Your task is to generate a clean, JSON-formatted list. Every item in the list must represent a product row from the table and contain EXACTLY these four fields (use these exact key names): "DESCRIPTION", "PRICE", "DEAL", and "DISC %". Do not infer values; only extract what is written in the respective columns. Set any missing cell to null. For PRICE, preserve the invoice formatting when possible (numbers with two decimal places; thousands separators like commas are fine). For "DISC %", copy the discount as shown (e.g. 5%, 15%, 10%) or null if blank. For DEAL, copy text such as promotional deals (e.g. 5+1) or null if blank. If the source table has blank separator rows between groups, include them as an object with all four fields null or empty strings.`;

const USER_PROMPT = `Return ONLY a valid JSON array of objects. Each object must have exactly these keys: "DESCRIPTION", "PRICE", "DEAL", "DISC %". No markdown, no explanation, no code fences.`;

const app = express();
app.use(express.json({ limit: "2mb" }));

const EXPORT_COLUMNS = ["DESCRIPTION", "PRICE", "DEAL", "DISC %"];

function strOrEmpty(val) {
  if (val === null || val === undefined) return "";
  return String(val);
}

function formatDiscForXlsx(val) {
  if (val === null || val === undefined || val === "") return "";
  const s = String(val).trim();
  if (s === "") return "";
  const num = parseFloat(s.replace(/%/g, "").replace(/,/g, "").trim());
  if (!Number.isNaN(num)) return `${num}%`;
  return s;
}

function parsePriceForXlsx(val) {
  if (val === null || val === undefined || val === "") return { kind: "empty" };
  const raw = String(val)
    .replace(/,/g, "")
    .replace(/^\s*\$\s*/, "")
    .trim();
  const n = parseFloat(raw);
  if (!Number.isNaN(n)) return { kind: "number", value: n };
  return { kind: "text", value: String(val).trim() };
}

const BORDER_THIN_BLACK = { style: "thin", color: { argb: "FF000000" } };

function applyAllBorders(cell) {
  cell.border = {
    top: BORDER_THIN_BLACK,
    left: BORDER_THIN_BLACK,
    bottom: BORDER_THIN_BLACK,
    right: BORDER_THIN_BLACK,
  };
}

async function buildInvoiceWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Invoice", {
    views: [{ state: "frozen", ySplit: 1 }],
  });

  sheet.columns = [{ width: 48 }, { width: 14 }, { width: 12 }, { width: 10 }];

  const header = sheet.getRow(1);
  EXPORT_COLUMNS.forEach((title, i) => {
    const cell = header.getCell(i + 1);
    cell.value = title;
    cell.font = { name: "Calibri", bold: true, size: 14 };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    applyAllBorders(cell);
  });
  header.height = 22;

  rows.forEach((row) => {
    const price = parsePriceForXlsx(row.PRICE);
    const priceCell =
      price.kind === "number"
        ? price.value
        : price.kind === "text"
          ? price.value
          : "";

    const excelRow = sheet.addRow([
      strOrEmpty(row.DESCRIPTION),
      priceCell,
      strOrEmpty(row.DEAL),
      formatDiscForXlsx(row["DISC %"]),
    ]);

    excelRow.font = { name: "Calibri", size: 11 };
    if (price.kind === "number") {
      excelRow.getCell(2).numFmt = "#,##0.00";
    }
    excelRow.getCell(1).alignment = { horizontal: "left", vertical: "top", wrapText: true };
    excelRow.getCell(2).alignment = { horizontal: "right", vertical: "top" };
    excelRow.getCell(3).alignment = { horizontal: "center", vertical: "top" };
    excelRow.getCell(4).alignment = { horizontal: "center", vertical: "top" };
    for (let c = 1; c <= 4; c++) {
      applyAllBorders(excelRow.getCell(c));
    }
  });

  return workbook;
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 },
});

function parseJsonFromModelText(text) {
  const trimmed = text.trim();
  const fence = trimmed.match(/^```(?:json)?\s*([\s\S]*?)```$/m);
  const raw = fence ? fence[1].trim() : trimmed;
  return JSON.parse(raw);
}

app.post("/export-xlsx", async (req, res) => {
  const rows = req.body && req.body.rows;
  if (!Array.isArray(rows) || rows.length === 0) {
    return res.status(400).json({ error: "Request body must include a non-empty rows array." });
  }

  try {
    const workbook = await buildInvoiceWorkbook(rows);
    const buffer = await workbook.xlsx.writeBuffer();
    const rawName = (req.body.baseName && String(req.body.baseName)) || "invoice";
    const safeBase =
      rawName
        .replace(/\.[^.]+$/, "")
        .replace(/[^\w\-]+/g, "_")
        .slice(0, 80) || "invoice";
    const filename = `${safeBase}-extracted.xlsx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Failed to build spreadsheet." });
  }
});

function getGeminiApiKey() {
  return (
    process.env.GEMINI_API_KEY ||
    process.env.GOOGLE_API_KEY ||
    ""
  ).trim();
}

app.post("/analyze-invoice", upload.single("image"), async (req, res) => {
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    return res.status(500).json({
      error:
        "No API key in .env. Set GEMINI_API_KEY or GOOGLE_API_KEY (Google AI Studio).",
    });
  }

  if (!req.file) {
    return res.status(400).json({ error: "No image file provided. Use field name 'image'." });
  }

  const mimeType = req.file.mimetype || "image/jpeg";
  if (!mimeType.startsWith("image/")) {
    return res.status(400).json({ error: "Uploaded file must be an image." });
  }

  const base64 = req.file.buffer.toString("base64");

  try {
    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({
      model: MODEL_ID,
      systemInstruction: SYSTEM_INSTRUCTION,
    });

    const result = await model.generateContent([
      {
        inlineData: {
          mimeType,
          data: base64,
        },
      },
      { text: USER_PROMPT },
    ]);

    const response = result.response;
    const text = response.text();
    const rows = parseJsonFromModelText(text);

    if (!Array.isArray(rows)) {
      return res.status(502).json({
        error: "Model did not return a JSON array.",
        raw: text,
      });
    }

    res.json({ rows });
  } catch (err) {
    console.error(err);
    const message = err.message || "Analysis failed";
    res.status(502).json({ error: message });
  }
});

app.use(express.static(path.join(__dirname, "public")));

module.exports = app;

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Invoice scanner running at http://localhost:${PORT}`);
  });
}
