// Builds the Windows download: dist/Clause-Craft-for-Windows.zip, plus download/ for the GitHub Pages download page
//
//   Clause Craft/
//     START HERE.html             how to use it (browser, desktop Word, Word on the web)
//     Clause Craft.html           the browser version - double-click to open, works offline
//     Install in Word.cmd         registers the add-in with desktop Word (current user, no admin), then
//                                 opens Clause Craft Demo Lease.docx so Word starts with the panel showing
//     Clause Craft Demo Lease.docx  the demo lease, tagged to open the Clause Craft panel
//     Uninstall from Word.cmd
//     Clause Craft manifest.xml   for Word on the web: Home > Add-ins > More Settings > Upload My Add-in
//     files/                      app.js, styles.css, engine/
//
// Usage: node tools/package.mjs
import { readFileSync, writeFileSync, mkdirSync, rmSync, existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import zlib from "node:zlib";
import { createRequire } from "node:module";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const OUT = join(ROOT, "dist");
const TOP = "Clause Craft/";

const page = readFileSync(join(ROOT, "index.html"), "utf8")
  .replace('href="styles.css"', 'href="files/styles.css"')
  .replace('src="engine/clause-engine.js"', 'src="files/engine/clause-engine.js"')
  .replace('src="engine/demo-docs.js"', 'src="files/engine/demo-docs.js"')
  .replace('src="app.js"', 'src="files/app.js"');
for (const ref of ["files/styles.css", "files/engine/clause-engine.js", "files/engine/demo-docs.js", "files/app.js"]) {
  if (!page.includes(ref)) throw new Error("index.html no longer references " + ref);
}

const manifest = readFileSync(join(ROOT, "word", "manifest.xml"), "utf8");
if (!/https:\/\/chadnasir\.github\.io\/clause-craft-cre\/word\/taskpane\.html/.test(manifest)) throw new Error("manifest is not pointing at GitHub Pages");
const ADDIN_ID = /<Id>([0-9A-Fa-f-]{36})<\/Id>/.exec(manifest)[1];
const ADDIN_VERSION = /<Version>([\d.]+)<\/Version>/.exec(manifest)[1];
const Demo = createRequire(import.meta.url)(join(ROOT, "engine", "demo-docs.js"));

/**
 * The demo lease as a Word document tagged to open the Clause Craft panel. The web extension parts
 * mirror the template Microsoft's office-addin-dev-settings uses to open a sideloaded add-in
 * (reference store="developer" storeType="Registry", task pane visible).
 */
function startDocx(paragraphs) {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
  const esc = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const run = (t, bold) => "<w:r>" + (bold ? "<w:rPr><w:b/></w:rPr>" : "") + '<w:t xml:space="preserve">' + esc(t) + "</w:t></w:r>";
  const body = paragraphs.map((text, i) => {
    const m = /^(\d+\.\s+[^.]{2,60}\.)\s+([\s\S]+)$/.exec(text);
    const heading = !m && /^[^a-z]{3,90}$/.test(text) && !/^By:/.test(text);
    const pPr = i === 0 ? '<w:pPr><w:spacing w:after="360"/><w:jc w:val="center"/></w:pPr>' : '<w:pPr><w:jc w:val="both"/></w:pPr>';
    if (!text.trim()) return "<w:p/>";
    if (m) return "<w:p>" + pPr + run(m[1], true) + run(" " + m[2], false) + "</w:p>";
    return "<w:p>" + pPr + run(text, heading) + "</w:p>";
  }).join("");
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  return zip([
    ["[Content_Types].xml", Buffer.from(xml + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
      '<Override PartName="/word/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>' +
      '<Override PartName="/word/webextensions/webextension.xml" ContentType="application/vnd.ms-office.webextension+xml"/></Types>')],
    ["_rels/.rels", Buffer.from(xml + '<Relationships xmlns="' + REL + '">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="/word/webextensions/taskpanes.xml"/></Relationships>')],
    ["word/_rels/document.xml.rels", Buffer.from(xml + '<Relationships xmlns="' + REL + '">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')],
    ["word/document.xml", Buffer.from(xml + '<w:document xmlns:w="' + W + '"><w:body>' + body +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')],
    ["word/styles.xml", Buffer.from(xml + '<w:styles xmlns:w="' + W + '"><w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>')],
    ["word/settings.xml", Buffer.from(xml + '<w:settings xmlns:w="' + W + '"><w:defaultTabStop w:val="720"/>' +
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')],
    ["word/webextensions/taskpanes.xml", Buffer.from(xml + '<wetp:taskpanes xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">' +
      '<wetp:taskpane dockstate="" visibility="1" width="350" row="1"><wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></wetp:taskpane></wetp:taskpanes>')],
    ["word/webextensions/_rels/taskpanes.xml.rels", Buffer.from(xml + '<Relationships xmlns="' + REL + '">' +
      '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="/word/webextensions/webextension.xml"/></Relationships>')],
    ["word/webextensions/webextension.xml", Buffer.from(xml + '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{' + ADDIN_ID + '}">' +
      '<we:reference id="' + ADDIN_ID + '" version="' + ADDIN_VERSION + '" store="developer" storeType="Registry"/><we:alternateReferences/><we:properties></we:properties><we:bindings/></we:webextension>')]
  ]);
}

const entries = [
  ["START HERE.html", readFileSync(join(ROOT, "windows", "START HERE.html"))],
  ["Clause Craft.html", Buffer.from(page, "utf8")],
  ["Install in Word.cmd", readFileSync(join(ROOT, "windows", "Install in Word.cmd"))],
  ["Clause Craft Demo Lease.docx", null],
  ["Uninstall from Word.cmd", readFileSync(join(ROOT, "windows", "Uninstall from Word.cmd"))],
  ["Clause Craft manifest.xml", Buffer.from(manifest, "utf8")],
  ["files/app.js", readFileSync(join(ROOT, "app.js"))],
  ["files/styles.css", readFileSync(join(ROOT, "styles.css"))],
  ["files/engine/clause-engine.js", readFileSync(join(ROOT, "engine", "clause-engine.js"))],
  ["files/engine/demo-docs.js", readFileSync(join(ROOT, "engine", "demo-docs.js"))]
];
for (const [name, data] of entries) {
  if (name.endsWith(".cmd") && data.includes("\n") && !data.includes("\r\n")) throw new Error(name + " must have Windows (CRLF) line endings");
}

function zip(files) {
  const now = new Date();
  const time = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
  const date = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  const locals = [], centrals = [];
  let offset = 0;
  for (const [name, data] of files) {
    const nameBuf = Buffer.from(name, "utf8");
    const packed = zlib.deflateRawSync(data, { level: 9 });
    const crc = zlib.crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x0800, 6); local.writeUInt16LE(8, 8);
    local.writeUInt16LE(time, 10); local.writeUInt16LE(date, 12); local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(packed.length, 18); local.writeUInt32LE(data.length, 22); local.writeUInt16LE(nameBuf.length, 26); local.writeUInt16LE(0, 28);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt16LE(0x0800, 8); central.writeUInt16LE(8, 10);
    central.writeUInt16LE(time, 12); central.writeUInt16LE(date, 14); central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(packed.length, 20); central.writeUInt32LE(data.length, 24); central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    locals.push(local, nameBuf, packed);
    centrals.push(central, nameBuf);
    offset += 30 + nameBuf.length + packed.length;
  }
  const cd = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, end]);
}

entries.find((e) => e[0] === "Clause Craft Demo Lease.docx")[1] = startDocx(Demo.lease.paragraphs);

if (existsSync(OUT)) rmSync(OUT, { recursive: true, force: true });
mkdirSync(join(OUT, "Clause Craft", "files", "engine"), { recursive: true });
for (const [name, data] of entries) writeFileSync(join(OUT, "Clause Craft", name), data);
const zipped = zip(entries.map(([name, data]) => [TOP + name, data]));
writeFileSync(join(OUT, "Clause-Craft-for-Windows.zip"), zipped);
// The download page (download/index.html, served by GitHub Pages) links to these two files.
writeFileSync(join(ROOT, "download", "Clause-Craft-for-Windows.zip"), zipped);
writeFileSync(join(ROOT, "download", "Clause-Craft-manifest.xml"), manifest);
console.log("dist/Clause-Craft-for-Windows.zip and download/  " + Math.round(zipped.length / 1024) + " KB, " + entries.length + " files");
