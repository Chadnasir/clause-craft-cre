# Clause Craft

Spellbook-style contract review that sits beside your document, built for CRE leases and PSAs.

Two ways to use it, same review engine (`engine/clause-engine.js`):

1. **Inside Microsoft Word** (`word/`): a **Review Contract** button on the Home tab opens the Clause Craft pane next to your lease or PSA.
2. **In the browser** (`index.html`): a Word-style page with the same review panel. Nothing to install.

Everything runs on your computer. No AI service, no API keys, no account, and the contract text is never uploaded.

- Live (browser): https://raw.githack.com/Chadnasir/clause-craft-cre/main/index.html
- Repo: https://github.com/Chadnasir/clause-craft-cre

## What it does

- **Review**: flags one-sided terms for the side you represent (tenant, landlord, buyer, seller, or both sides), with the reason, the exact words, and suggested language. Severity filters: High / Medium / Low / Missing.
- **Redline**: one click applies the fix. In Word it is a real tracked change, or a Word comment if you'd rather flag it. In the browser it shows as red strike-through and blue double underline. "Redline all" and "Comment all" in Word.
- **Missing clauses**: CASp disclosure (Cal. Civ. Code §1938), waiver of subrogation, quiet enjoyment, renewal option, 1031 cooperation, seller operating covenants and more, with starting-point language to insert.
- **Clauses**: what the document contains, section by section. Click to jump. Word also has a clause library to insert at the cursor.
- **Ask**: plain-English questions ("Who can terminate?", "What's the NNN?", "Is the deposit refundable?") answered from the document's own words, with the section cited.
- **Summary**: parties, premises or property, term, rent or price, deposit, key risks and missing items. Insert into the document or copy.

Covers office and retail leases and purchase and sale agreements / escrow instructions, with California checks (CASp, pre-dispute jury waivers, Civil Code waivers, §1542 releases). Demo lease (Irvine office) and demo PSA (Perris net-lease retail) are included.

## Install the Word add-in

The add-in loads its pages from this GitHub repo through raw.githack.com, so the `word/` and `engine/` folders need to be on the `main` branch first.

**Word on the web (quickest)**
1. Open a document in Word at office.com.
2. Home > Add-ins > More Settings > **Upload My Add-in**.
3. Browse to `word/manifest.xml` and choose **Upload**.
4. Home tab > **Review Contract**.

**Word for Windows (desktop)**
1. Make a folder, for example `C:\ClauseCraft`, and copy `word/manifest.xml` into it.
2. Right-click the folder > Properties > Sharing > **Share** > Share. Note the network path it shows (like `\\YOUR-PC\ClauseCraft`).
3. In Word: File > Options > Trust Center > Trust Center Settings > **Trusted Add-in Catalogs**. Paste the network path into Catalog Url, choose **Add catalog**, tick **Show in Menu**, then OK and OK. Close and reopen Word.
4. Home > Add-ins > **Advanced** > **SHARED FOLDER** > Clause Craft > **Add**.
5. Home tab > **Review Contract**.

Word versions: tracked changes and comments need a current Word (Microsoft 365, Word 2024, or Word on the web). Older versions such as Word 2019 still work, but redlines are applied as formatting (red strike-through, blue double underline) and comments are turned off. The pane tells you which mode it is in.

To check the manifest with Microsoft's validator (needs internet): `npx office-addin-manifest validate word/manifest.xml`

## Use it in the browser

Open `index.html` in Chrome or Edge (double-clicking the file works) or use the live link above.

- **Load Demo Lease**, or **Upload** a .docx, .txt or .md contract (or drag the file onto the page). Tracked deletions in a .docx are left out; insertions are kept.
- **Run Review**, then **Accept** or **Reject** each item. Change "Reviewing for" to switch sides.
- **Export HTML** saves a redlined copy. **Print** prints only the document, redlines in color.
- Your working copy is saved in that browser only, and survives a page reload.

## Privacy

Contract text stays on the computer. The only files loaded from the internet are Microsoft's Office.js (every Word add-in needs it) and the Google Fonts typefaces. The browser tests confirm no other network requests during a full review session.

## Limits

- Rule-based checks written for California CRE forms, not AI. It catches the landlord- and seller-favorable language it was built for and will miss issues worded differently. Treat it as a first pass, not a substitute for counsel.
- Tested against an Office.js simulator and in Chromium (119 automated checks). Not yet run inside a real copy of Word: try it on the demo lease first.
- The Word add-in needs an internet connection to load.
- .docx import in the browser reads the text only: no formatting, headers, footers or footnotes.

## Files

| Path | What it is |
|---|---|
| `engine/clause-engine.js` | Review rules, missing-clause checks, Ask, Summary |
| `engine/demo-docs.js` | Demo office lease and PSA |
| `word/manifest.xml` | Word add-in manifest (what you upload or put in the shared folder) |
| `word/taskpane.*`, `word/commands.html`, `word/assets/` | Word task pane and ribbon icons |
| `index.html`, `app.js`, `styles.css` | Browser version |
| `tests/` | Engine tests, Word add-in tests (Office.js simulator), browser tests, .docx fixtures |

## Run the tests

```
npm install
npm test           # engine: 15 tests
npm run test:ui    # Word add-in against the Office.js simulator: 45 checks
npm run test:web   # browser version: 59 checks
```

The UI and browser tests drive Chromium. Point `CHROMIUM_PATH` at Chrome or Edge (for example `C:\Program Files\Google\Chrome\Application\chrome.exe`), or run `npx playwright-core install chromium` once.

## vs Spellbook

Spellbook is a Microsoft Word add-in with cloud AI, Playbooks, market benchmarks, and Associate for multi-doc matters (custom per-seat pricing, 7-day trial). Clause Craft is narrower: CRE-tuned rules, a Word add-in plus a browser version, runs offline, free to use. Not a law-firm platform and **not legal advice**.

## vs typical Grok Build

Broker brand, real redlining in the document, CRE clause library, Ask + Summary, persistence. Not a generic chat toy.

Demo only. Not legal advice. Have counsel review before signing.
