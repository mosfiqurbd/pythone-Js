const fs = require("fs");
const path = require("path");
const pdf = require("pdf-poppler");
const Tesseract = require("tesseract.js");
const { Document, Packer, Paragraph, HeadingLevel } = require("docx");
const ExcelJS = require("exceljs");

//SETTINGS
const PDF_PATH = "C:\\pdf_project\\book.pdf";
const OUTPUT_FOLDER = "C:\\pdf_project\\images_js";
const TOTAL_PAGES = 20;

async function main() {
  // Step 1: PDF to Image & Rename
  console.log("Step 1: PDF to Image Cronverting...");
  if (!fs.existsSync(OUTPUT_FOLDER)) fs.mkdirSync(OUTPUT_FOLDER);

  const opts = {
    format: "png",
    out_dir: OUTPUT_FOLDER,
    out_prefix: "page",
    page: null,
    poppler_path: "C:\\poppler\\Library\\bin\\poppler-25.12.0\\Library\\bin",
  };

  for (let i = 1; i <= TOTAL_PAGES; i++) {
    opts.page = i;
    opts.out_prefix = `page_${String(i).padStart(2, "0")}`;
    await pdf.convert(PDF_PATH, opts);
    console.log(` page_${String(i).padStart(2, "0")}.png saved`);
  }

  // Step 2: OCR
  console.log("\nStep 2:OCR [Image to text]...");
  const allText = [];

  const files = fs
    .readdirSync(OUTPUT_FOLDER)
    .filter((f) => f.endsWith(".png"))
    .sort();

  for (const file of files) {
    const imgPath = path.join(OUTPUT_FOLDER, file);
    const {
      data: { text },
    } = await Tesseract.recognize(imgPath, "ben", {
      logger: () => {},
    });
    allText.push({ page: file, text: text.trim() });
    console.log(`   ${file} OCR done`);
  }

  // Step 3: Word File
  console.log("\nStep 3: Word file Creating...");
  const children = [
    new Paragraph({ text: "আল-আদাবুল মুফরাদ", heading: HeadingLevel.TITLE }),
  ];
  for (const item of allText) {
    children.push(
      new Paragraph({ text: item.page, heading: HeadingLevel.HEADING_2 }),
    );
    children.push(new Paragraph({ text: item.text }));
  }
  const doc = new Document({ sections: [{ children }] });
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync("C:\\pdf_project\\output_js.docx", buffer);
  console.log("output_js.docx saved");

  // Step 4: Excel File
  console.log("\nStep 4: Excel file Creating...");
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("আদাবুল মুফরাদ");
  ws.columns = [
    { header: "Page Number", key: "page", width: 15 },
    { header: "Text", key: "text", width: 100 },
  ];
  // Header formatting
  ws.getRow(1).font = { bold: true, size: 12 };
  ws.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4472C4" },
  };
  ws.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };

  for (const item of allText) {
    ws.addRow({ page: item.page, text: item.text });
  }
  await wb.xlsx.writeFile("C:\\pdf_project\\output_js.xlsx");
  console.log("   output_js.xlsx saved");

  console.log("\n All work is done!");
  console.log(" C:\\pdf_project\\ Look in the folder");
}

main().catch(console.error);
