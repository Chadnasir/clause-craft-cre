# Clause Craft

Contract review that sits beside your document, built for CRE leases and PSAs.

Two ways to use it, same review engine (`engine/clause-engine.js`):

1. **Inside Microsoft Word** (`word/`): a **Review Contract** button on the Home tab opens the Clause Craft pane next to your lease or PSA.
2. **In the browser** (`index.html`): a Word-style page with the same panel. Nothing to install, and it works offline when opened straight from the folder.

Everything runs on your computer. No AI service, no API keys, no account, and the contract text is never uploaded.

- Browser version: https://chadnasir.github.io/clause-craft-cre/
- Download for Windows (browser version + Word installer, with steps): https://chadnasir.github.io/clause-craft-cre/download/
- Repo: https://github.com/Chadnasir/clause-craft-cre

## What it does

- **Review setup**: review type (comprehensive, quick = high risks only, drafting check), your party (with the party names read from the document; **Refresh parties** re-reads them and **Add party** lets you name a side yourself), jurisdiction (detect, California, another state), which playbook to use (or none), and deal notes that stay with the document.
- **Review**: flags one-sided terms for the side you represent (tenant, landlord, buyer, seller, or both sides), with the reason, the exact words and suggested language. Filter by High / Medium / Low / Missing, or **Review selection** to see only what's in the paragraphs you selected.
- **Document checks**: term dates add up, rent matches rate x square feet, party names match the signature blocks, defined terms are used and capitalized terms are defined, section cross-references and exhibits exist, written numbers match the figures.
- **Redline**: one click applies the fix. In Word it is a real tracked change, or a Word comment if you'd rather flag it; **Redline all** / **Comment all**. In the browser it shows as red strike-through and blue double underline; **Accept all**; **Download .docx** opens in Word with every change as a tracked change.
- **Playbook**: your own standards (escalation cap, late charge, holdover, cure periods, estoppel turnaround, deposit, opex cap, due diligence and closing periods, rep survival, deposit size and more), checked against the document and redlined to your number. Edit them once; they are saved on that computer. **Save as new playbook** keeps a separate set (say, for credit-tenant deals) that you can pick in the setup.
- **Missing clauses**: CASp disclosure (Cal. Civ. Code 1938), waiver of subrogation, quiet enjoyment, renewal option, 1031 cooperation, seller operating covenants and more, with starting-point language to insert at the cursor or at the end.
- **Draft**: option to extend, early termination, free rent, TI allowance, opex cap, guaranty cap, exclusive use, ROFO, due diligence, DD extension, closing and rep survival, written in contract style ("sixty (60) days") from the numbers you type. Clause library, plus **My clauses** to save language you like from any document.
- **Ask / Explain**: plain-English questions ("Who can terminate?", "Is the deposit refundable?") answered from the document's own words with the section cited; **Explain selected text** describes any clause and lists what's flagged in it.
- **Summary and issues list**: parties, premises or property, term, rent or price, deposit, key risks and missing items; and a numbered issues list for the other side, with your deal notes at the top. Insert either into the document or copy it. Clause Craft ignores its own summaries and issues lists when it reviews the document again.

Covers office and retail leases and purchase and sale agreements / escrow instructions, with California checks (CASp, pre-dispute jury waivers, Civil Code waivers, 1542 releases). Demo lease (Irvine office) and demo PSA (Perris net-lease retail) are included.

## Install the Word add-in

The add-in's pages are served by GitHub Pages from the `gh-pages` branch (https://chadnasir.github.io/clause-craft-cre/word/taskpane.html). Word needs an internet connection to open the pane; the review itself runs inside Word.

**Word for Windows (desktop), one line**: press Windows + R, paste the line below, press Enter. It saves Clause Craft to Documents\Clause Craft, adds it to Word and opens Word with the panel showing (if Word is open it asks, then restarts it).

```
powershell -c "irm https://chadnasir.github.io/clause-craft-cre/install.ps1|iex"
```

**Or from the download**: double-click `Install in Word.cmd`. It registers the add-in for your Windows user, the same way Microsoft's own `office-addin-dev-settings` tool does (`HKCU\Software\Microsoft\Office\16.0\WEF\Developer`), with no admin rights needed. Then it opens `Clause Craft Demo Lease.docx`, which is tagged to open the Clause Craft panel. After that, open any document and click Home > **Review Contract**. `Uninstall from Word.cmd` removes it.

**Word on the web**: open a document, then Home > Add-ins > More Settings > **Upload My Add-in** > Browse to `Clause Craft manifest.xml` (or `word/manifest.xml`) > **Upload**. Word on the web keeps it until that browser's cache is cleared.

**If IT blocks the installer**: put `word/manifest.xml` in a shared folder, add the folder's network path under File > Options > Trust Center > Trust Center Settings > **Trusted Add-in Catalogs** (tick Show in Menu), restart Word, then Home > Add-ins > Advanced > **SHARED FOLDER**.

Word versions: tracked changes and comments need Microsoft 365, Word 2024 or Word on the web. Older versions such as Word 2019 still work, but redlines are applied as formatting (red strike-through, blue double underline) and comments are turned off. The pane tells you which mode it is in.

## Use it in the browser

Open `index.html` (or `Clause Craft.html` in the Windows download) in Chrome or Edge; double-clicking the file works. Or use the link above.

- **Load Demo Lease**, a demo PSA, or **Upload** a .docx, .txt or .md contract (or drag the file onto the page). Tracked deletions in a .docx are left out; insertions are kept.
- **Start review**, then **Accept** or **Reject** each item. **Change** reopens the setup.
- **Download .docx** saves a Word file with your redlines as tracked changes (track changes stays on for edits you make in Word). **Export HTML** and **Print** are there too.
- Your working copy, setup, standards and saved clauses stay in that browser only, and survive closing the page.

## Privacy

Contract text stays on the computer. The only files loaded from the internet are Microsoft's Office.js (every Word add-in needs it), the Clause Craft pages themselves from GitHub Pages, and the Google Fonts typefaces. The browser tests confirm no other network requests during a full session.

## Limits

- Rule-based checks written for California CRE forms, not AI. It catches the landlord- and seller-favorable language it was built for and will miss issues worded differently. Treat it as a first pass, not a substitute for counsel.
- No free-form chat or "edit the document for me", no market benchmarks, no web research and no multi-document matters: those need an AI service, which this deliberately doesn't use.
- Tested against an Office.js simulator, in Chromium (including opened straight from the folder) and the Windows installer under Wine. Not yet run inside a real copy of Word: try it on the demo lease first.
- .docx import in the browser reads the text only: no formatting, headers, footers or footnotes. The .docx download is plain formatting (Times New Roman, bold captions) with your redlines as tracked changes.

## Files

| Path | What it is |
|---|---|
| `engine/clause-engine.js` | Review rules, checks, playbook, drafting, Ask, Explain, Summary, issues list |
| `engine/demo-docs.js` | Demo office lease and PSA |
| `word/manifest.xml` | Word add-in manifest (points at GitHub Pages) |
| `word/taskpane.*`, `word/commands.html`, `word/assets/` | Word task pane and ribbon icons |
| `index.html`, `app.js`, `styles.css` | Browser version |
| `windows/` | `Install in Word.cmd`, `Uninstall from Word.cmd`, `START HERE.html` for the Windows download |
| `install.ps1` | One-line setup for Windows (downloads the zip, then runs `Install in Word.cmd`) |
| `tools/package.mjs` | Builds `dist/Clause-Craft-for-Windows.zip` and the files behind the download page |
| `download/` | Download page with install steps, the Windows zip and the add-in manifest (served by GitHub Pages) |
| `tests/` | Engine tests, Word add-in tests (Office.js simulator), browser tests, .docx fixtures |

## Publish an update

```
git push origin main            # the code
git push origin main:gh-pages   # what Word and the browser link load
node tools/package.mjs          # new Windows download in dist/ and download/ (commit download/ before pushing gh-pages)
```

## Run the tests

```
npm install
npm test           # engine: 25 tests
npm run test:ui    # Word add-in against the Office.js simulator: 84 checks
npm run test:web   # browser version: 120 checks (incl. file:// and the .docx download)
```

The UI and browser tests drive Chromium. Point `CHROMIUM_PATH` at Chrome or Edge (for example `C:\Program Files\Google\Chrome\Application\chrome.exe`), or run `npx playwright-core install chromium` once. The .docx check also opens the download with python-docx and LibreOffice when they are installed.

## vs Spellbook

Spellbook is a Microsoft Word add-in built on cloud AI: chat and edit, Playbooks, market benchmarks and Associate for multi-document matters (custom per-seat pricing). Clause Craft covers the review, playbook, drafting, issues-list and setup workflow with CRE-tuned rules instead, runs offline, and is free to use. Not a law-firm platform and **not legal advice**.

## vs typical Grok Build

Broker brand, real redlining in the document, CRE clause library, Ask + Summary, persistence. Not a generic chat toy.

Demo only. Not legal advice. Have counsel review before signing.
