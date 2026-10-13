// Builds the Windows download: dist/Clause-Craft-for-Windows.zip
//
//   Clause Craft/
//     START HERE.html             how to use it (browser, desktop Word, Word on the web)
//     Clause Craft.html           the browser version - double-click to open, works offline
//     Install in Word.cmd         registers the add-in with desktop Word (current user, no admin)
//     Uninstall from Word.cmd
//     Clause Craft manifest.xml   for Word on the web: Home > Add-ins > More Settings > Upload My Add-in
//     files/                      app.js, styles.css, engine/
//
// Usage: node tools/package.mjs
import { readFileSync, writeFileSync, mkdirSync, rmSync, existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import zlib from "node:zlib";

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

const entries = [
  ["START HERE.html", readFileSync(join(ROOT, "windows", "START HERE.html"))],
  ["Clause Craft.html", Buffer.from(page, "utf8")],
  ["Install in Word.cmd", readFileSync(join(ROOT, "windows", "Install in Word.cmd"))],
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

if (existsSync(OUT)) rmSync(OUT, { recursive: true, force: true });
mkdirSync(join(OUT, "Clause Craft", "files", "engine"), { recursive: true });
for (const [name, data] of entries) writeFileSync(join(OUT, "Clause Craft", name), data);
const zipped = zip(entries.map(([name, data]) => [TOP + name, data]));
writeFileSync(join(OUT, "Clause-Craft-for-Windows.zip"), zipped);
console.log("dist/Clause-Craft-for-Windows.zip  " + Math.round(zipped.length / 1024) + " KB, " + entries.length + " files");
