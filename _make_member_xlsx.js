// Build รายชื่อสมาชิก.xlsx from transcribed names.
// Columns: ลำดับ | ชื่อสมาชิก (no prefix) | เบอร์โทร (sequential, 10-digit text starting 0000000001)
const fs = require("fs");
const zlib = require("zlib");

// --- Full names (with prefix) ---
const fromImages = [
  "นางสาวมลทิพย์ พิมพ์เสน","นายทองดี พลสวัสดิ์","นายจำนงค์ หวังสุดดี","นางรัชฎาวรรณ เจือจันทร์",
  "นางสุภานัน จงอุตส่าห์","นางวรลักษณ์ เพชรหนองชุม","นางสมเดช ไพสเป็น","นางสาวรุ่งภา เอิบยิ้ม",
  "นายสุระ อุณชาติ","นางสาวบุญตา บรรหารทุกข์","นางสาวศณินันท์ ทำทอง","นายวิศรุต ค้ำคูณ",
  "นางสาวธัญญธร หมายเจริญ","นางสาวรัญยา สุระพล","นายอริยา เชิดสุข","นางสาวธนิษฐา วุฒิยา",
  "นางสาวอรินทร์ ลออลม","นางสาวนิชชาวัลย์ บุญสุข","นายวีรวัฒน์ ศรีไสว","นางสาวศุภลักษณ์ สุขเต็ม",
  "นายธีระพล แสนแก้ว","นายพัฒนพงษ์ อรรคสาร","นางสาวสุภาภรณ์ จันทบุตร","นางสาวหนูโกร โกรงาม",
  "นายสุริยา นาลา","นายพีรพัฒน์ กรยศทอง","นางสาวกาญจนา ดาทอง","นางสาวอนัญญา วันนภัส",
  "นายสุรชัย บุญหรัง","นางสาวพรทิพย์ เศษคงดา","นายคุณเกียรติ เหล่าธีรกาญจนา","นางสาวขนัฏฐา บัวลา",
  "นายสันทนา พิเศษ","นางจันทร์เพ็ญ พยุงเกษม","นางสาวนวพร พงษ์นรรัมย์","นายเพชร อุดธูร",
  "นายทนงศักดิ์ บุญสอน","นายวุฒิชัย พ่อค้า","นายสิทธิชัย เสกกล้า","นางสาวสุภาวดี ศรีจันทร์",
  "นายสมนึก อยู่นาน","นายคงฤทธิ์ จิรันคร","นายปิยะ วรรณทอง","นายศรราม แสนกล้า",
  "นายนิธิรร วิทยาสิงห์","นายธนารักษ์ สุดยอดสุข","นางสาวพัชรีภรณ์ ทองคำสาร","นายดุลยกิตติ์ คิดดี",
  "นายเชิด อยู่ดี","นายสัมฤทธิ์ ทองอ้ม","นางสาวชุติมา พลบูรณ์ศรี","นางสาวลดาพรรณ ชมภูนุช",
  "นายธีรวัฒน์ อสิพงษ์","นางสาวรัญฤทัย จันทร์ฝอย","นางสาววิจิตรา สันงวง","นายพงศธร ทวีศรี",
  "นางสาวพรพิมล โพธิ์แก้ว","นายกิตติพงษ์ อุปไมเลิศ","นางสาวฐิติมา อุดมศักดิ์","นายวิรัตน์ รัตนวิวัฒน์",
  "นายณัฐพล คำเสมอ","นายกำชัย บุญนับ","นางสาวมาลิน คุมสุข","นายสุพิศ ภักดีไสย์",
  "นางนิภากร อุดหนุน","นายธีรวัฒน์ จันทะสิงห์","นายคมสัน จันใจมั่น","นางสาวดุจดาว สุขอยู่",
  "นางพัน ภักดีไสย์","นายสุรชาติ เซี่ยวชาญ","นายกัณฑภณ จารัตน์","นางสาวสุมาลี เชื่อมแสง",
  "นายสุรศักดิ์ แจ่มแจ้ง","นายภักรา ประดับสุข","นายวินัย สำรวย","นางสาวนัญธิดา ศรีสุข",
  "นางสาวจุฑิกานต์ ขุลีรัมย์","นางสาวภูภา สาแก้ว","นายวีรชัย ประจวบเตา","นายสมภพ หวังผล",
  "นายอภิรีรัตน์ บุญสอน","นางสาวเอื้อนณะภา บุญแย้ม","นางสาวณิชชาริรัตน์ เส้นศูนย์","นายปุณชัย ถาวร",
  "นายชานนท์ ใยแดง","นายพรศักดิ์ หวังสุดดี","นางสาวพักตรา บุตรลักษณ์","นางสาวเหทัย ในพลทอง",
  "นายโกรพม ศรีพรหม","นายศิวัช กะทอง","นายวรรุฬย์ โพธิ์หิรัญ","นางสาวปาราจนา แน่นหนา",
  "นายธีรพงษ์ ศรีสุข","นายณัฐดนัย อุปนันต์","นางสาวกากนกพรรณ ชุดดอน","นายพงษ์พรรณ เหลื่อมล้ำ",
  "นายสมบัติ วาพัดไทย","นายรัฐกานต์ สุนทรารักษ์","นายบัวหอม เกษมสุข","นายบุญเชิด อุไร",
  "นายแสงอาทิตย์ ประคองจิตร","นายสมชาย แสนเจตนา","นายวิไลรัตน์ ยวนยี","นายทศพร เป็นสุข",
  "นายณัฐธยาน์ คณาดี","นายภูษิต ขวัญยืน","นายอนุศาสตร์ เรือนริน","นางสาวสุมาลี สร้อยจิตร",
  "นายศราวุธ แสนรู้","นายจิรเดช จงกลาง","นางสาวเสาวลักษณ์ อุไร","นางวรรณภา ผดุงดี",
  "นางสาวนุสิตา สุดหนองบัว","นางสาวปัทมา สวนงาม","นางสุราลัย อุตธูร","นางสาวชิสา สมานโสร์",
  "นายธนากร จะลงประโคน"
];

const fromDocx = [
  "นางณีระนุช ใจองอาจ","นางกันทิยา อุไร","นางศรีจันทร์ อุไร","นางสมพงษ์ แจ่มแจ้ง",
  "นางละเอียด สาคร","นายวันลบ สุดแสนฉุน","นางพิสมัย สาคร","นางสายทิพย์ วันนิยม",
  "นายวิสุทธิ์ สระทองอ่อน","นางพิมพา โด่งดัง","นายสมคิด ประจันทร์","นายคาน ดำเนินงาม",
  "นายกิตติพงษ์ ประดับ"
];

const all = [...fromImages, ...fromDocx];

function stripPrefix(name) {
  let n = name.trim();
  for (const p of ["นางสาว", "นาง", "นาย"]) {
    if (n.startsWith(p)) { n = n.slice(p.length); break; }
  }
  return n.trim().replace(/\s+/g, " ");
}

const rows = all.map((full, i) => ({
  no: i + 1,
  name: stripPrefix(full),
  phone: String(i + 1).padStart(10, "0"),
}));

console.log("Total members:", rows.length, "(images:", fromImages.length, "+ docx:", fromDocx.length + ")");

// --- XLSX builders ---
function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function colRef(c, r) { return String.fromCharCode(65 + c) + r; }

function inlineStrCell(ref, val) {
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(val)}</t></is></c>`;
}
function numCell(ref, val, style) {
  return `<c r="${ref}"${style ? ` s="${style}"` : ""}><v>${val}</v></c>`;
}

const header = ["ลำดับ", "ชื่อสมาชิก", "เบอร์โทร"];
let sheetRows = "";
// header row (style 1 = bold)
sheetRows += `<row r="1">` +
  `<c r="A1" t="inlineStr" s="1"><is><t>${esc(header[0])}</t></is></c>` +
  `<c r="B1" t="inlineStr" s="1"><is><t>${esc(header[1])}</t></is></c>` +
  `<c r="C1" t="inlineStr" s="1"><is><t>${esc(header[2])}</t></is></c>` +
  `</row>`;
rows.forEach((row, idx) => {
  const r = idx + 2;
  sheetRows += `<row r="${r}">` +
    numCell(`A${r}`, row.no) +
    inlineStrCell(`B${r}`, row.name) +
    inlineStrCell(`C${r}`, row.phone) +
    `</row>`;
});

const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<cols><col min="1" max="1" width="8" customWidth="1"/><col min="2" max="2" width="30" customWidth="1"/><col min="3" max="3" width="16" customWidth="1"/></cols>
<sheetData>${sheetRows}</sheetData>
</worksheet>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="TH Sarabun New"/></font><font><b/><sz val="11"/><name val="TH Sarabun New"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
</styleSheet>`;

const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="สมาชิก" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

const files = [
  { name: "[Content_Types].xml", data: contentTypes },
  { name: "_rels/.rels", data: rootRels },
  { name: "xl/workbook.xml", data: workbookXml },
  { name: "xl/_rels/workbook.xml.rels", data: workbookRels },
  { name: "xl/worksheets/sheet1.xml", data: sheetXml },
  { name: "xl/styles.xml", data: stylesXml },
];

// --- CRC32 ---
const crcTable = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// --- ZIP ---
const localParts = [];
const central = [];
let offset = 0;
for (const f of files) {
  const nameBuf = Buffer.from(f.name, "utf8");
  const dataBuf = Buffer.from(f.data, "utf8");
  const comp = zlib.deflateRawSync(dataBuf);
  const crc = crc32(dataBuf);

  const lh = Buffer.alloc(30);
  lh.writeUInt32LE(0x04034b50, 0);
  lh.writeUInt16LE(20, 4);
  lh.writeUInt16LE(0x0800, 6); // UTF-8 flag
  lh.writeUInt16LE(8, 8);      // deflate
  lh.writeUInt16LE(0, 10);
  lh.writeUInt16LE(0, 12);
  lh.writeUInt32LE(crc, 14);
  lh.writeUInt32LE(comp.length, 18);
  lh.writeUInt32LE(dataBuf.length, 22);
  lh.writeUInt16LE(nameBuf.length, 26);
  lh.writeUInt16LE(0, 28);
  localParts.push(lh, nameBuf, comp);

  const ch = Buffer.alloc(46);
  ch.writeUInt32LE(0x02014b50, 0);
  ch.writeUInt16LE(20, 4);
  ch.writeUInt16LE(20, 6);
  ch.writeUInt16LE(0x0800, 8);
  ch.writeUInt16LE(8, 10);
  ch.writeUInt16LE(0, 12);
  ch.writeUInt16LE(0, 14);
  ch.writeUInt32LE(crc, 16);
  ch.writeUInt32LE(comp.length, 20);
  ch.writeUInt32LE(dataBuf.length, 24);
  ch.writeUInt16LE(nameBuf.length, 28);
  ch.writeUInt16LE(0, 30);
  ch.writeUInt16LE(0, 32);
  ch.writeUInt16LE(0, 34);
  ch.writeUInt16LE(0, 36);
  ch.writeUInt32LE(0, 38);
  ch.writeUInt32LE(offset, 42);
  central.push(ch, nameBuf);

  offset += lh.length + nameBuf.length + comp.length;
}
const centralBuf = Buffer.concat(central);
const localBuf = Buffer.concat(localParts);
const eocd = Buffer.alloc(22);
eocd.writeUInt32LE(0x06054b50, 0);
eocd.writeUInt16LE(0, 4);
eocd.writeUInt16LE(0, 6);
eocd.writeUInt16LE(files.length, 8);
eocd.writeUInt16LE(files.length, 10);
eocd.writeUInt32LE(centralBuf.length, 12);
eocd.writeUInt32LE(localBuf.length, 16);
eocd.writeUInt16LE(0, 20);

const zip = Buffer.concat([localBuf, centralBuf, eocd]);
const outPath = "สมาชิก/รายชื่อสมาชิก.xlsx";
fs.writeFileSync(outPath, zip);
console.log("Wrote:", outPath, "(" + zip.length + " bytes)");
console.log("First 3:", rows.slice(0,3).map(r => `${r.no}|${r.name}|${r.phone}`).join("  ::  "));
console.log("Last 3:", rows.slice(-3).map(r => `${r.no}|${r.name}|${r.phone}`).join("  ::  "));
