# PSWP Register Schema and Build Conventions (LEAN core, v51 workbook, 18-Aug-2026)

Project knowledge file - the ONLY schema file that belongs in project knowledge. Build history (v10-v31 positions and per-build notes) lives in `PSWP_Register_Schema_HISTORY.md`, shipped in the session zip, read only when a historical question arises.
POSITIONS ARE READ FROM THE WORKBOOK CONFIG SHEET (rule 19.6). Any position quoted in this file is illustrative; Config is authoritative. Companions: `pswp_build_batch.py` (rule 20 brief-driven capture build), `pswp_ident_batch.py` (rule 20 identification build), `pswp_json_repair.py` (F1-F11 + rule 16 gate), `pswp_line_restructure.py` (F9), `pswp_build_lib.py` (LO profile), `pswp_parsers.py` (raw-text route).

## 1. File identity

- Current: `PS_WP_Transaction_Register_3FY_v39_HANDOVER.xlsx` (~11MB). Register: 21,469 data lines (rows 5:21473), 127 columns, header row 4, control total row 21475 col T = $11,377,312.51 (must never change on a capture/enrichment build).
- Fonts Cambria; money `$#,##0.00;($#,##0.00);"-"`; dates `d-mmm-yyyy`; live formulas only (SUMIFS/COUNTIF/INDEX/MATCH; never XLOOKUP/FILTER/SORT).
- Recalc: LibreOffice convert route only (`soffice --headless --convert-to xlsx` with an isolated profile, OOXMLRecalcMode=0), ~90-120 s. `recalc.py` (macro route) hangs in the container and is banned. Reads via python-calamine (data_only equivalent), writes via openpyxl.
- Sheets (21): Handover | Method | Theme_Map | Register | Themes | PK000022 | Vendor_Series | Summary | Sites | PK_Listing | Coverage | Open_Items | Creditor_Lines | SE2_Budget | Evidence_Invoices | Evidence_Invoice_Lines | Data_Acquisition | Config | Project_Instructions | Vendor_Boilerplate | Freeze_Log.

## 2. Register column map (1-indexed)

Analysis block:
1 LineKey | 2 Section code | 3 Section name | 4 Doc Date | 5 FY loaded | 6 FY posted | 7 FY service | 8 Reference | 9 Doc Type | 10 Source | 11 Creditor Code | 12 Contractor | 13 ABN | 14 Natural Account | 15 Nat Acct Name | 16 PK Header | 17 PK Charged | 18 PK basis | 19 (id) | 20 Amount ex GST | 21 Amount incl GST (est, `=ROUND(T{r}*1.1,2)`) | 22 Narration (GL line) | 23 Enquiry Narration | 24 Nature Category | 25 Nature Detail | 26 Theme (formula lookup on Theme_Map A5:A31) | 27 Nature Basis | 28 Evidence | 29 Evidence Tier | 30 Attachment in TechOne | 31 Coding Verdict | 32 Coding Note | 33 Status | 34 Follow-up

Parsed references: 35 Contract Ref | 36 DM Ref | 37 Quote Ref | 38 Work Ref | 39 REGO | 40 Toll TAG | 41 Retailer Acct | 42 Site/supply point | 43 Person Named | 44 Standing Order flag

Enquiry block: 45-53 (Enq Amount incl GST, GST Date, Due Date, Outstanding, Applied, Ageing, On Hold, Payment Details, Billing System)

Grey source block (verbatim, never touch): 54-87 (54 Src Short Description, 55 Src Account, 56 Period, 57 Reference, 58 Date, 59 Doc Type Desc, 60 Details, 61 Transaction Amount, 62 Has Attachment, 63 Source, 64 Has Note, 65 Account Chain Description, 66 Work Order, 67 Work System, 68 Document Unique ID, 69 Src Document File, 70 Contract, 71 Contract Schedule, 72 Attachment, 73 Asset, 74 Asset Register, 75 WO Transaction Number, 76 Natural Account, 77 Service No, 78 Doc Type Code, 79-80 User Fld 1-2, 81 Section, 82 Program, 83 Directorate, 84 WO Task, 85 Posted Date, 86 Posted Time, 87 Ledger). At v13 the 26SLACT layout mapped by NAME, not position (CurrencyAmount -> 61, ATBatchName -> 69, ATDImageFilename -> 72).

## 3. Rule 17 green evidence block (cols 88-127, CJ to DW)

88 Ev Invoice ID | 89 Vendor (printed) | 90 ABN (printed, verbatim incl. mis-groupings) | 91 ACN | 92 Vendor address | 93 Vendor phone | 94 Invoice Date (datetime) | 95 Due Date | 96 Purchase Order | 97 Contract # | 98 Bill To | 99 Ship To | 100 Requesting Officer | 101 Request Date | 102 Request Via | 103 CR/WO # | 104 PK # (as printed) | 105 PK normalised | 106 Site Details | 107 Site Contact | 108 Work Description | 109 Technician | 110 Line count | 111 Line items (verbatim, "1. Desc | qty @ $unit = $amt. 2. ...") | 112 Works carried out / attendance narrative (verbatim; condense whitespace runs before write - see §11 BLC trap) | 113 Printed Sub Total ex GST | 114 Printed GST | 115 Printed Total incl GST | 116 Payments Made | 117 Balance Due | 118 Payment details | 119 Terms & notices | 120 Pages | 121 Lines sum ex GST (live) | 122 Chk lines = printed subtotal | 123 Chk register = printed subtotal | 124 Chk printed GST = 10% | 125 Source file | 126 Captured stamp | 127 Anomalies & capture notes

Conventions:
- Fields printed blank → the literal string `(blank as printed)`. Fields the layout does not carry → `(not printed)`. Never fill from another source.
- Col 121 (new rows): `=ROUND(SUMIF(Evidence_Invoice_Lines!$A$5:$A${EIL_END},CJ{r},Evidence_Invoice_Lines!$G$5:$G${EIL_END}),2)` — v14 EIL_END = 1436 (v12 rows use $A$5:$A$1197, v11 rows $A$5:$A$613, pre-v11 exemplar rows 13365/13369 $A$5:$A$56; all remain valid since their invoices' lines sit inside those ranges; leave them).
- Anomalies (127): any check-formula variant line MUST contain the literal tag `[check variant]` — a control counts these (>= 46 at v14).

## 4. Check formulas — standard and ratified variants (Method 11.0)

Standard:
- 122: `=IF(DQ{r}=DI{r},"TRUE","FALSE")`
- 123: `=IF(ROUND(T{r},2)=DI{r},"TRUE","FALSE")`
- 124: `=IF(DJ{r}=ROUND(DI{r}*0.1,2),"TRUE","FALSE")`

Variants (each requires the `[check variant]` anomalies tag; ratified v6):
- 122 inclB (Harpley template B, lines print GST-inclusive): `=IF(OR(DQ{r}=DI{r},DQ{r}=DK{r}),"TRUE","FALSE")`
- 123 tol1c (TechOne 1c ex-GST derivation, rule 7): `=IF(ROUND(ABS(ROUND(T{r},2)-DI{r}),2)<=0.01,"TRUE","FALSE")`
- 123 deriv (register = printed incl / 1.1; Hilltops, Bethania): `=IF(OR(ROUND(T{r},2)=DI{r},ROUND(T{r},2)=ROUND(DK{r}/1.1,2)),"TRUE","FALSE")`
- 123 split (one invoice, several register lines; ties by LineKey): `=IF(ROUND(T{r},2)=ROUND(SUMIFS(Evidence_Invoice_Lines!$G$5:$G${END},Evidence_Invoice_Lines!$A$5:$A${END},CJ{r},Evidence_Invoice_Lines!$J$5:$J${END},A{r}),2),"TRUE","FALSE")`
- 124 gstfree (FRE supplies, printed Tax $0.00): `=IF(OR(DJ{r}=ROUND(DI{r}*0.1,2),DJ{r}=0),"TRUE","FALSE")`
- 124 tol1cgst / tol2c (vendor GST rounding): `=IF(ROUND(ABS(DJ{r}-ROUND(DI{r}*0.1,2)),2)<=0.01,"TRUE","FALSE")` (0.02 for per-line-rounding baskets, e.g. Bethania). v14 example: Harpley 00014014 prints GST $121.25, 1c under 10%, while its own printed total implies $121.26.

TRAP: never write a bare `ABS(x)<=0.01` tolerance — binary floats make 0.01 store as 0.01000000002 and the check fails. Always wrap: `ROUND(ABS(...),2)<=`.
TRAP: Python `round()` is banker's rounding; Excel ROUND is half-away-from-zero. Pre-compute check-3 expectations with `Decimal(...).quantize(..., ROUND_HALF_UP)`, or a printed GST of e.g. $150.54 on a $1,505.35 subtotal will falsely flag.

## 5. Evidence sheets (layout; positions from Config)

Evidence_Invoices (32 cols, header row 4):
- Cols: 1 Invoice | 2 Vendor | 3 ABN | 4 Invoice Date | 5 Due Date | 6 LCC Reference | 7 Contract/Quote | 8 Site(s) | 9 Lines | 10 Ex GST | 11 GST | 12 Incl GST | 13 Verification (cites the invoice's recon row) | 14-31 staged register-block fields | 32 Staged (source/capture stamp).
- Data rows 5-379 (5-13 pre-v11; 14-157 v11; 158-292 v12; 293-379 the v14 batch: 42 Batch 6 + 45 Batch 7). Register-captured invoices carry the col-14 marker `Complete on the register line (green block, columns CJ to DW)`. MJR-3/495 (no register match, Open Items #49) holds its staged fields in cols 14-31.
- Totals row 381 (`=SUM(J5:J379)` etc). Correspondence rows 383-384 (verbatim, rule 16d). Rule 17 controls: title 386, rows 387-398 (twelve).
- Rows 5-11: the seven pre-v4 FY2025/26 invoices completed rule 17 at v13 (26SLACT); rows carry completion markers.
- Ev Invoice ID prefixes: DRL- (Robinson), INV-000xxxxx (Harpley), KC- (Kachel), QP- (Q Power), EDSS- (EDSS/CINC EDSS series), HT- (Hilltops), LBL- (Lockwise), ELM-, T2-, LED- (Shailer claims), TCS-, MJR-, HIG-, BW-, UFF-, CALDME-, BLC- (NEW v14), CLM- (NEW v14, the 0042232-02 claim); C# and printed INV-# ids stay as printed.

Evidence_Invoice_Lines (11 cols, header row 4):
- Cols: A Invoice | B Line | C Description (verbatim) | D Qty | E Unit Price | F GST | G Amount ex GST | H PK as printed | I PK normalised | J Register LineKey | K Notes.
- Data rows 5-1436 (5-56 pre-v11; 57-613 v11; 614-1197 v12; 1198-1436 the v14 batch, 239 lines). Reconciliation block: header 1438, rows 1439-1813 (375 invoices: col C invoice id, col D full-range SUMIF over $A$5:$A$1436, col E TRUE/FALSE against the printed target).
- Template-B Harpley amounts in col G are AS PRINTED (GST-inclusive); their recon rows test against the printed Total Inc GST; every such line's Notes says so.
- Any recon-block move must restate the EI Verification (col 13) row citations in the same build — the v14 build fixed 279 stale citations left by earlier moves.

Controls (Evidence_Invoices rows 387-398, all must be TRUE to ship):
sighted-line count (`COUNTIF(Register!$AA$5:$AA$19232,"Sighted invoice line")` = 376 at v14) | Ev-ID coverage | count tie | check 1/2/3 counts = sighted count | register total = 10,677,585.33 | header ex-GST totals = SUM(EIL G$5:G$1436) minus $25,579.06 GST-inclusive allowance ($25,383.26 v13 + $195.80 INV-00013880 template B) | staging count = 7 | recon TRUE count = 375 (E1439:E1813) | `[check variant]` count >= 46 | Coverage grand total (Coverage!C130) = Register!$T$19234.

## 6. Other sheets

- Theme_Map: categories A5:A31 -> themes B5:B31. Nature Category values written to Register col 24 MUST exist in A5:A31 (27 categories; list unchanged since v12).
- Vendor_Series: header row 4; label | ABN | COUNTIF | SUMIF over Register!$L/$T$5:$19232. Labels written to Register col 12 must match EXACTLY. Data blocks rows 5-66, 71-92, 96-111, 115-117 (v14: 115 BLC Queensland Pty Ltd, 116 CINC EDSS Pty Ltd, 117 Unidentified (EDSS-format progress claim, licence 53209)); totals row 68 sums all four blocks; footnotes rows 70, 94, 113, 119.
- Open_Items: header row 4; # | FY | Status | Contractor/series | Lines | $ ex GST (net) | Action. Last # = 54 at v14 (rows to 58).
- Data_Acquisition: single-column text. Section F logs transcription batches with md5 hashes: F1-F3 (v11 raw text), F5 Batch 4, F6 Batch 5, F7 Batch 6 + F8 Batch 7 (v14, structured JSON). F4 (extraction tool identity) OPEN for batches 1-4, 6, 7; stated for Batch 5 only (M365 Copilot, GPT-5).
- Coverage: Panel A per FY x natural account (live COUNTIFS/SUMIFS); FY totals; grand total row C130 (tied by EI control 398); Panel C working order (ranks static as at v12).
- Config: key/value block rows 3-19 (v14 values: REGISTER_DATA 5:19232, REGISTER_TOTAL_ROW 19234, CONTROL_TOTAL 10677585.33, EIL_DATA 5:1436, EIL_END 1436, RECON_ROWS 1439:1813, RECON_COUNT 375, EI_DATA 5:379, EI_TOTALS_ROW 381, EI_CONTROL_ROWS 387:398, SIGHTED_COUNT 376, VARIANT_TAG_MIN 46, EXTRACT_26SLACT 3429658.75). The toolkit reads these instead of hardcoding (rule 19.6).
- Handover row 20 = v14 change note; Method rows 31-34 = 14.0 (v14 batch capture, structural moves, identities).

## 7. Build mechanics (hard-won; see pswp_build_lib.py)

- Read/edit with openpyxl (not read_only for writes); pandas reads need `engine='calamine'`. Register reads: `header=3`, data row = df index + 5.
- MERGED-CELL TRAP: unmerge any range intersecting a write zone first (v13/v14 evidence sheets currently carry none, but check).
- Styles: copy `cell._style` (via `copy.copy`) from template rows. Green block template = Register row 13365 cols 88-127; EI template row 158; EIL template row 6; recon template row (capture BEFORE clearing the old block). Overwritten regions keep stale styles unless every written cell gets the template style.
- Setting `.value` on an existing cell preserves its style — analysis-column upgrades need no style work.
- One build script per session (rule 18/20), one save, ONE recalc via the LibreOffice convert route (isolated profile, OOXMLRecalcMode=0; recalc.py banned), one verify pass reading the RECALCULATED file with python-calamine and requiring exact "TRUE" on every control.
- Duplicate-upload check before any parse: md5 the uploads against Data_Acquisition.
- Cell text cap 32,767 chars: condense whitespace runs (`[ \t]{3,}` -> two spaces, `\n{3,}` -> two) on narrative/anomalies writes; the printed text is preserved, only layout padding drops (note it in the anomalies where length forced it).

## 8. Register-matching logic

- The Reference column (col 8) carries invoice numbers directly. Match on exact string; Harpley refs keep leading zeros. References can TRUNCATE long invoice numbers (TCS 17118059 posts as "171180") — check truncation before declaring a non-match.
- Green-block only AP lines (`Doc Type = PUR Cred Invoice`); journals rest on pairing (rule 6/T9) and are never green-blocked.
- Target row selection: the AP row whose amount equals the printed subtotal (±1c) or the incl/1.1 derivation. Multi-line postings all get the block with the split variant.
- Zero-amount companion AP rows: no green block; Coding Note cross-references the sibling LineKey — ONLY when the reference has exactly one non-zero AP row.
- Status/Basis on companions untouched (control: every "Sighted invoice line" row must carry an Ev Invoice ID).

## 9. Standing traps and facts

- PoolShop bills one monthly invoice split across five PKs but the GL charges PK000022 in full (Open Items #41, both prior FYs).
- Hilltops invoices are GST-FREE: TechOne posts incl/1.1 — phantom GST credit; whole series suspect (Open Items #42).
- Harpley template B ("Your Order No:" header) prints line amounts GST-INCLUSIVE.
- PK typo normalisation: PK00022 = PK000022 (record printed form alongside).
- Multi-page invoices may print totals on the last page only — capture note, not an extraction loss (rule 13).
- Binders carry duplicate copy pages: whitespace-normalise and hash to dedupe; capture once, note the duplicate.
- NA-safe everything: `fillna('')` before comparing PK/WO columns (rule 1).
- On JSON batches run repair_and_gate() (families F1-F9) before anything else; F9 (aggregate-line capture) rebuilds items from the retained page text via `pswp_line_restructure.py`, replacing only when the rebuilt sum ties exactly.
- LCC's own ABN (21 627 796 435) prints as the bill-to; extractors grab it as the vendor ABN (F8 repair). The vendor ABN is in the supplier letterhead/footer.
- PLAY FORCE PTY LTD (69 106 457 176) ≠ Play Force Australia Pty Ltd (89 677 476 541): different entities, separate series/labels. Same for the two BrizSouth locksmith entities.
- Xero INV-# references are shared across vendors AND years — target selection by amount, extra rows untouched.


## 10. Ratified variants and standing traps lifted from the build history

### Check-2 sum-tie variant (ratified v23, Method 23.0; fold into instructions v10)

Where the GL posts one printed invoice line as SEVERAL non-zero AP register lines with NO printed counterpart to the split (distinct from the LineKey split-posting variant, where printed lines map to register lines), each row's check 2 sums the sibling register amounts against the row's printed target:
`=IF(ROUND(T{r1}+T{r2},2)=DI{r},"TRUE","FALSE")` on each sibling, tagged [check variant] in the anomalies note.
Precedent: INV-38915 ($6,695.00, one printed line) posting as $6,694.55 "Quote 657743" + $0.45 "variance" (rows 14732-14733). First occurrence in the workbook; VARIANT_TAG_MIN 863 -> 865.

### Recon E-formula forms (build-trap note, v23)

Reconciliation rows carry TWO ratified E-column forms: exact `=IF(D{r}={target},...)` and tolerance `=IF(ROUND(ABS(D{r}-{target}),2)=0,...)` (incl-GST-basis and rounding-tolerance invoices). Any recon regeneration must parse and preserve the row's existing form; new ex-GST-tied rows take the exact form. (v23 build trap: the tolerance form first appears at v22 recon row 3583, INV-7187.)

### EI control formula fixed references (build-trap note, v23)

EI control row "header totals tie to captured lines" cites the totals row by absolute cell (v22 J1464 -> v23 J1506) and carries a fixed delta (-39,979.12, the GST content of incl-GST-basis captures; unchanged at v23, all Levai lines print ex-GST). Any EI append must rewrite this reference alongside the range transforms, and any incl-GST-basis capture must restate the delta.
### Citation-rewrite build trap (v25 patch cycle)

Old-format EI col-13 texts (v12-era) carry TWO "row N" tokens: the recon row AND a tail phrase "rule 17 checks live on register row N" whose number was ALREADY overwritten with the recon row by earlier blanket rewrites. The citation rewrite must therefore replace ALL "row \d+" occurrences in col 13 (blanket re.sub, no count limit), as v23/v24 did; a first-occurrence-only rewrite leaves the tail token stale against the citation audit (which matches the dotted "row N." form). The tail phrase does not hold a real register row anywhere in the current workbook.

- v29/v30 tail-shift traps (Method 30.2): (a) A blind `1589 -> 1591`-family string shift can write a SELF-INCLUSIVE `=SUM(J5:J1591)` into the totals row itself - circular, computes blank after the convert-route recalc. Always rewrite the totals-row sums explicitly after any tail shift. (b) Inserting rows above a panel whose SUMIFS cite absolute `$A` criteria rows does NOT shift those criteria (openpyxl writes raw strings): the v29 FY27 Coverage insert left Panel B reading `$A126:$A129`, so all four FY rows computed FY2026/27. Repoint criteria to each row's own label on any panel move. (c) A verify pass that tests controls for `== "FALSE"` lets blank/circular results through. STANDING RULE from v30: the verify pass asserts EI control cells are exactly "TRUE" (counts rows exactly 1,496) and audits the Coverage all-years tie (Coverage!$C$142 = Register!$T$21475) directly.
- PK_Listing sheet: NEW SHEET PK_Listing (before Coverage): section / cost centre / verbatim D1Description1 segment / PK charged, live COUNTIFS-SUMIFS over Register FY2026/27 (cols B section code, S cost centre, Q PK charged) + tie-check row (Config PKL_DATA, PKL_CHECK_ROW). Coverage gained 12 FY2026/27 account rows (panels B/C relocated +12 with captured styles/formulas).
- APLEDGER creditor-history layout (v17): 26 cols, system names row 4, display names row 5, data row 6+; TRAP: trailing blank-reference row carries the account balance (grand-total trap) - exclude on parse; Transaction Amounts are INCL GST. Creditor label format "CODE (Name)".
- Register total re-bound at v29 (19232->21473 / 19234->21475): any fixed-range formula rewrite uses digit-boundary-guarded regex, never bare substring replace.

## 32. v32 positions (Batch 29, Binder3.pdf, 11-Aug-2026)

Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged; capture build).
Evidence_Invoices: data 5:1677 (+50), totals 1679, correspondence 1681-1682, controls title 1684, controls 1685:1698 (fourteen). Sighted count 1,555 (v31 1,505 + 50).
Evidence_Invoice_Lines: data 5:3915 (+289), reconciliation header 3918, reconciliation rows 3919:5591 (1,673 rows).
Config (v32): WORKBOOK_VERSION v32; EIL_DATA 5:3915; EIL_END 3915; RECON_ROWS 3919:5591; RECON_COUNT 1673; EI_DATA 5:1677; EI_TOTALS_ROW 1679; EI_CONTROL_ROWS 1685:1698; SIGHTED_COUNT 1555; VARIANT_TAG_MIN 872 (count 872); GSTINC_ALLOWANCE 40,685.82 (unchanged - no incl-GST-basis captures in the batch); BOILERPLATE_KEYS 71.
New boilerplate keys: BP:BSP-P1, BP:ELM-P1, BP:ELM-T1, BP:CMBM-P1, BP:CMBM-T1. Reused with bumps: BP:FD-P1 +8, BP:INV-P4 +6, BP:INV-P21 +19, BP:INV-T3 +19, BP:INV-P6 +7.

### Duplicate screening is a GATE, not a nicety (v32 precedent)

Kachel 7546 arrived in the Batch 29 corpus as a re-sighting of the v30 capture KC-7546. Its two register rows (205, 241) are the v30 HELD rows - the GL posts $43,409.64 PK000025 + $10,776.36 PK000030 against five printed whole-dollar zone lines, unallocatable to the printed lines. The parse-session match table listed both rows as ordinary `exact` targets. Green-blocking them would have silently reversed a documented held decision. Screen EvID (WITH the vendor prefix convention - the corpus id was the bare `7546`, the existing capture `KC-7546`) against Evidence_Invoices column A BEFORE the match is trusted, and re-assert that held rows still carry no EvID after the build.

### Check-2 partial-scope variant (ratified v32, Method 32.1; fold into instructions v10)

Where a sighted invoice is Council-wide and the loaded PS/WP scope carries only a SUBSET of its printed lines - the balance posting outside 4090240/4090260 and therefore invisible to this register - check 2 ties the register amount to that row's own captured line(s) by LineKey, using the ratified LineKey split-posting formula; checks 1 and 3 continue to tie the whole printed invoice as printed.
`=IF(ROUND(T{r},2)=ROUND(SUMIFS(Evidence_Invoice_Lines!$G$5:$G${END},Evidence_Invoice_Lines!$A$5:$A${END},CJ{r},Evidence_Invoice_Lines!$J$5:$J${END},A{r}),2),"TRUE","FALSE")`
The in-scope line(s) carry the register LineKey in EIL column J; every out-of-scope line carries a blank J and a Notes entry saying so. The association must be FIXED by evidence, never inferred - precedent CMBM SI-047499, where the GL narration "Berrinba Wetlands" plus the $58.38 amount identify exactly one of 159 printed lines (the only Berrinba line). Where no evidence fixes the line, the row is HELD, not green-blocked. Tagged `[check variant]`; floor 869 -> 872.

### Defect family F11: column-header token bleed (new v32)

On column-ruled templates the wrapped tail of a column header can attach to the FIRST captured description ("Line Amount Excl. GST" -> "GST Animal Management Centre - 8 x Auto Femcare Bin - Monthly Service"). Detect by a leading header token on line 1 only; repair from the retained page text with the basis in the anomalies field. Observed once (CMBM SI-047499).

### Citation drift: BOTH limbs (build trap, repaired v32)

Evidence_Invoices col 13 carries TWO citations - the reconciliation row on Evidence_Invoice_Lines, and the register row(s) on which the rule 17 checks live. Successive builds rewrote the recon limb with a blanket `re.sub(r'row \d+', ...)`, which also overwrote the REGISTER limb with the recon-row number: 141 of the 285 register-row clauses were wrong at v31 (all of the lower-case `register row` phrasing; the capitalised `Register row` phrasing survived). v32 rewrites both limbs from authoritative maps - EvID -> recon row, and EvID -> the register rows carrying that green block, read from column CJ - and the verify pass now audits BOTH permanently. Invoices with no register match cite `(no register match)`. Never rewrite either limb by positional shift or blanket pattern.

### Header restatement from retained page text (v32 method)

Where a corpus carries amounts and lines but only a partial header set, the rule 17 header fields are restated per vendor template from the retained verbatim page text, and EVERY literal token is asserted present in that invoice's own page text before writing (v32: 51 invoices, 0 failures - see b29_headers.py). Two letterheads print address blocks in side-by-side columns that the text layer interleaves (CMBM, Elemental); those blocks are de-interleaved and the reconstruction is stated in the anomalies field with its basis. Fields the layout does not carry stay `(not printed)`.


## 33. v33 positions (Batch 30, Binder1.pdf re-cut, 17-Aug-2026)

Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged; capture build). One green block added: row 17071 (Harpley 00014715, INV-00014715, $870.00, PK000415).
Evidence_Invoices: data 5:1678 (+1), totals 1680, correspondence 1682-1683, controls title 1685, controls 1686:1699 (fourteen). Sighted count 1,556 (v32 1,555 + 1).
Evidence_Invoice_Lines: data 5:3918 (+3), reconciliation header 3921, reconciliation rows 3922:5595 (1,674 rows).
Config (v33): WORKBOOK_VERSION v33; EIL_DATA 5:3918; EIL_END 3918; RECON_ROWS 3922:5595; RECON_COUNT 1674; EI_DATA 5:1678; EI_TOTALS_ROW 1680; EI_CONTROL_ROWS 1686:1699; SIGHTED_COUNT 1556; VARIANT_TAG_MIN 872 (count 872, no new variants); GSTINC_ALLOWANCE 40,685.82 (unchanged); BOILERPLATE_KEYS 71 (BP:INV-P1 68, BP:INV-T1 97).
Data_Acquisition F34 (Batch 30). Method 33.0-33.2 (Panel C repair recorded in the driver and this file; add Method 33.3 at v34). Open Items unchanged (last #87).

### A whole binder can be a re-cut (rule 12 precedent, v33)

Binder1.pdf (fourth distinct file of that name; 69 pages, 34 invoices, v3 corpus md5 7a7a75f62894132ae7faa679d1ca2afd) was 33/34 re-sightings of captures already on the register (Batch 24/v28 Harpley HAR073 creditor batch and Batch 28/v31). Screen EVERY corpus id, prefix-aware, before trusting the match table - here the existing EvIDs used the INV- prefix on CINC EDSS 0058241/0058368 and Traffex 180246 as well as on Harpley (Batch 28 convention), so the screen must try the bare id AND every prefix on the sheet, not only the family's own prefix. Re-sightings are audited against the existing Evidence_Invoices row (subtotal, GST, total, priced-line count - `duplicate_expect` in the brief) and never re-captured; 33/33 agreed to the cent.

### Coverage Panel C criteria were 12 rows stale (found and repaired v33, rule-18 patch cycle)

The v29 12-row FY2026/27 insert left Panel C (working order, rows 148-169) citing `$B{r-12}` in its COUNTIFS/SUMIFS, so every listing displayed the counts and dollars of the listing twelve rows below it (Levai's 2,131 Partial lines showed against rank 13, the FY2023/24 5-digit series against rank 14, and ranks 1-12 read blank cells as 0). Panel B was repointed at v30; Panel C was missed because the verify audited Coverage only for FALSE cells and the all-years tie. `pswp_build_batch.py` now (a) repoints any Coverage formula that tests `Register!$L` against a `$B{n}` label to its own row (44 cells at v33), and (b) the verify recounts every Coverage row whose col B is a Register col-L label against the register itself (lines and $ not Confirmed) - 19 label rows at v33. Panel C ranks remain static as at v12; the true clearance order at v33 is Levai ($1.94M Partial), FY2023/24 7-digit ($130,716.19), FY2024/25 5-digit ($82,053.52), Q Power Partial, then the rest.

### F1 recurs on v3 corpora for Harpley template A (v33)

Harpley template A prints the GST amount as a bare `$` line under `Sub Total:` (the label is in the image layer), so `printed_gst` in the v3 corpus held the Total Inc GST value on all 30 Harpley invoices (self-tie still passed because the extractor's tie is on the subtotal). Restate from the bare-amount line, assert ROUND_HALF_UP(subtotal/10) and subtotal + GST = printed total; multi-page copies print the totals block without amounts on the first page (rule 13, last-page totals). Template B (`Your Order No` header) prints `GST:` labelled and its lines GST-inclusive (00014639: captured $2,997.50 = printed Total Inc GST, not a loss). Queue: tighten extraction prompt v3 for this line.

### v3 corpus adaptation notes (v33)

`pswp_corpus_from_v3.py` maps `unit` from `l['unit']` (always null) - the unit price lives in `unit_price_ex_gst`; and the printed description lives in the PRICED record's `note` field, not `line_text` (which is the padded layout row). The v33 session script (b30_prepare.py in the session zip) applies both; fold into the adapter before its next use. Per-line GST is `(not printed)` on Harpley template A (no GST column on the line table).


## 34. v34 positions (Batch 31, playforce.pdf, 17-Aug-2026)

Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged; capture build). 34 green blocks added across both Play Force entities.
Evidence_Invoices: data 5:1712 (+34), totals 1714, correspondence 1716-1717, controls title 1719, controls 1720:1733 (fourteen). Sighted count 1,590 (v33 1,556 + 34).
Evidence_Invoice_Lines: data 5:3986 (+68), reconciliation header 3989, reconciliation rows 3990:5697 (1,708 rows).
Config (v34): WORKBOOK_VERSION v34; EIL_DATA 5:3986; EIL_END 3986; RECON_ROWS 3990:5697; RECON_COUNT 1708; EI_DATA 5:1712; EI_TOTALS_ROW 1714; EI_CONTROL_ROWS 1720:1733; SIGHTED_COUNT 1590; VARIANT_TAG_MIN 872 (count 872, no new variants); GSTINC_ALLOWANCE 40,685.82 (unchanged); BOILERPLATE_KEYS 72 (new BP:INV-P22; bumps BP:INV-P3 +10, BP:INV-P9 +17, BP:INV-T3 +20, BP:INV-T2 +17).
Data_Acquisition F35 (Batch 31). Method 34.0-34.3. Open Items to #88.

### Defect family F12: wrapped description tail (new v34)

A printed description wider than its column prints the tail on the NEXT layout row, which the extractor classifies as NARRATIVE and drops, truncating the captured description mid-phrase ("Routine Inspection and removal of debris from" losing "skating area"). Repair ONLY where the tail is the immediately adjacent narrative line, carries no amount, carries no label prefix and runs to six words or fewer; restate by joining and state the basis in the anomalies. Observed on 8 of 68 lines (Play Force classic layout); the same description prints un-wrapped on the 6 Xero-layout lines, which cross-checks the repair to the word. Distinct from F11 (column-header token bleed), which attaches a header tail to the FIRST description.

### A shared Xero reference can be a COLLISION, not a re-sighting (rule 12 precedent, v34)

Xero INV-# references collide across vendors and years (section 9). The rule 12 EvID screen must therefore separate a re-sighting from a collision - the same id belonging to a different supplier. INV-3412 is both a Pool Shop QLD invoice ($31,142.16, 26-Jul-2023, register row 210, captured earlier) and a Play Force Australia invoice ($1,142.28, 19-May-2025, register row 1563). Resolved by an EvID override in the brief (PF-3412); the register Reference is left as posted on both rows and target selection is by printed subtotal, never by reference alone.

### One boilerplate key per cell (build trap, v34)

The Amendment 2 control does an exact whole-cell `COUNTIF(Vendor_Boilerplate!$A$5:$A$300, ...)` over Register DN and DO, so a "; "-joined list of BP keys in either cell reads as an unknown key and the control returns FALSE (caught on the v34 first pass, fixed in a rule-18 patch cycle). Where an invoice prints more than one payment or terms block, the cell carries the FIRST key and the driver writes the remainder into the anomalies with their role.

### Driver extension: evid_overrides and bp_overrides (rule 20, v34)

`pswp_build_batch.py` gains two brief-driven maps. `evid_overrides` (invoice id -> EvID) resolves shared-reference collisions. `bp_overrides` (invoice id -> {payment, terms}, single key or list) carries per-invoice boilerplate, because one vendor can print several blocks across a series or an entity change - Play Force prints three payment blocks (BP:INV-P3 old entity; BP:INV-P9 new entity classic layout; BP:INV-P22 new entity Xero layout, minted at v34) and two terms notices (BP:INV-T3 terms pages; BP:INV-T2 the 1-Jul-2024 name-change notice). Both maps are validated against Vendor_Boilerplate and recorded in the build summary.

### F1 recurs on the v3 corpus in three fields (v34)

On the Play Force corpus `printed_gst` was null on the 7 Xero-layout invoices, `printed_total_incl_gst` null on the 7 classic-layout invoices, and `invoice_date` null on the 7 Xero-layout invoices (the Xero letterhead prints the label and the value on separate layout rows, interleaved with the supplier address column, so a same-line regex misses it). All 21 restated from the invoice's own retained page text and re-tied arithmetically. Queue with the v33 Harpley finding for extraction prompt v3.

### v3 corpus adaptation notes (v34, extends the v33 block)

`unit` on a PRICED record can hold the printed ITEM CODE (L1, MISC1, PG236, L67), not a unit of measure - map the unit price from `unit_price_ex_gst` and push the item code to the EIL Notes column. The printed per-line GST is a RATE column ("10%") on the Xero layout and absent on the classic layout, so per-line GST is "10%" or "(not printed)", never a derived dollar figure.


## 35. v35 positions (Batch 32, five APLEDGER creditor histories, 17-Aug-2026)

IDENTIFICATION build, not capture. Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged - column T is never written on an identification build). 290 lines upgraded to Tier 1, $116,322.34 ex GST. Sighted count unchanged at 1,590; EI/EIL/recon positions all unchanged from v34.
Creditor_Lines: data 5:5592 (+254). Vendor_Series: 2 new rows (Farmcraft Pty Ltd, Parklands Distributors Pty Ltd t/a AMS Eco Products) plus footnote. Open Items to #89. Data_Acquisition F36. Method 35.0-35.3.
Config (v35): WORKBOOK_VERSION v35; CL_DATA 5:5592; ENQ_MATCHED_ROWS 4418; TIER1_UPGRADED 2619. All capture-side keys unchanged.

### Identification builds run through pswp_ident_batch.py (NEW driver, rule 20)

Sibling of `pswp_build_batch.py`: same Config-driven positions, same brief-driven facts, same one-call build/recalc/verify/ship chain. It writes WHO was paid and nothing else - Register cols 11/12/13, 23, 24/25 (only where the brief authorises a category), 27/28/29/30, 33, 34, 45-53; Creditor_Lines append; Vendor_Series, Open_Items, Data_Acquisition, Handover, Method, Config. It NEVER writes Confirmed, never touches a green block and never writes column T, so the control total cannot move. Verify re-runs the whole capture-side control set unchanged (it must still pass) and adds: every target row carries contractor, ABN, creditor code, Tier 1 and an enquiry block; every target reference exists on Creditor_Lines; no target row carries an Ev Invoice ID.

### Amount corroboration is the gate, not the reference (standing rule, v35)

Short numeric creditor series collide across vendors exactly as Xero INV-# references do (Method 34.2). A creditor-history reference match is accepted ONLY where the creditor transaction amount - which is INCL GST - equals the register line ex GST at 1.1 within 2c. Batch 32 held 43 non-zero matches back on this test; Open Item 89 proves them a second vendor (the FY2024/25 5-digit series is shared between Logan Plumbing Service Pty Ltd and Q Power (Qld) Pty Ltd, ABN 82 067 507 591, whose invoice 14150 was sighted in the same upload for the exact register amount $3,252.40 against Logan Plumbing's own 14150 at $244.65 incl GST).

### Zero-amount companion rows need a corroborated sibling (standing rule, v35)

A $0.00 register row cannot corroborate on amount, so reference alone would identify it - and on a generic reference form such as INV-#### that is not evidence. A zero-amount row is identified ONLY where a NON-ZERO sibling row carrying the same reference is itself amount-corroborated against the same creditor account (schema section 8 companion convention). Batch 32: 113 zero rows matched, 36 kept, 77 dropped.

### A Confirmed row is never an identification target (standing rule, v35)

A row can be Confirmed by journal pairing or the system record without carrying an Ev Invoice ID, so screening on the EvID alone is not enough. An identification build skips any target whose Status is already Confirmed rather than rewriting its basis, status or identity. Batch 32 skipped 12, including the five narration-only "Logan Plumbing" rows.

### Config CL_DATA drift (found and repaired v35)

Config carried CL_DATA 5:3353 while Creditor_Lines held contiguous data to row 5338 - 1,985 rows of creditor evidence outside the declared range, so any toolkit reading positions from Config (rule 19.6) would have silently truncated the sheet. Repaired at v35 and now rewritten by the driver on every append. Audit the other Config ranges against their sheets when a sheet is next touched.

## 36. v36 positions (Batch 33, kachel.pdf / robinson.pdf / brizsouth.pdf v3 corpora, 17-Aug-2026)

Capture build. Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged). 58 green blocks added (Kachel Cleaning 4, David Robinson Landscaping Pty Ltd 42, Lockwise BrizSouth Locksmiths 12); no check variants. Sighted count 1,648 (v35 1,590 + 58).
Evidence_Invoices: data 5:1771 (+59, including DRL-11151 with no register match), totals 1773, correspondence 1775-1776, controls title 1778, controls 1779:1792 (fourteen).
Evidence_Invoice_Lines: data 5:4210 (+224), reconciliation header 4213, reconciliation rows 4214:5980 (1,767 rows).
Config (v36): WORKBOOK_VERSION v36; EIL_DATA 5:4210; EIL_END 4210; RECON_ROWS 4214:5980; RECON_COUNT 1767; EI_DATA 5:1771; EI_TOTALS_ROW 1773; EI_CONTROL_ROWS 1779:1792; SIGHTED_COUNT 1648; VARIANT_TAG_MIN 872 (count 872); GSTINC_ALLOWANCE 40,685.82 (unchanged); BOILERPLATE_KEYS 72 (no new keys; bumps BP:VENDOR-P1 +43, BP:VENDOR-T1 +43, BP:LBL-P1 +12, BP:LBL-T1 +12).
Data_Acquisition F37 (Batch 33). Method 36.0-36.5. Open Items to #94. Vendor_Series: no new labels, one footnote. Register row 17777 ($0.00 companion, previously Unidentified) identified from its sighted sibling.

### Rule 12 held-row re-sighting, second occurrence (v36)

Kachel 7546 arrived again (kachel.pdf, page 5) as a bare `7546`; screened against KC-7546 (v30 HELD, Open Items #82), audited to the cent (`duplicate_expect`), not re-captured; rows 205/241 re-asserted with no EvID after the build. The screen must be prefix-aware AND vendor-aware: three families in one build, three prefixes (KC-, DRL-, LBL-).

### Rule 16 capture with NO register line: `unmatched` (driver extension, v36)

A sighted invoice with no PS/WP register line (Robinson 11151, $520.00, Beenleigh Aquatic Centre - an MV asset cost code, out of scope) is captured under rule 16 only: EIL lines with blank LineKey and a Notes marker, a reconciliation row, an EI row whose LCC Reference reads `(no register match; Open Items #n)`, whose col 13 register limb reads `Register row (no register match)`, and whose cols 14-31 carry the staged header fields (the MJR-3/495 precedent, section 5). No green block, no Confirmed; the verify asserts the EvID appears on no register row and the EI row is staged, not marked complete.

### F12 across three templates; `desc_rows` row-level verbatim gate (v36)

Kachel prints one claim over several layout rows with the amount in split dollar/cent cells (`180 | 00`, `$ 18   00`); Robinson prints each charge under a `SITE - CR n - OFFICER` header row with the description wrapped over 1-5 rows (once, 11302, the header row itself carries the amount and the description follows on the next row); Lockwise (Jobber) prints a two-column line table (Product/Service | Description, joined with one space) whose column tails wrap onto the following row ("Call out fee - Logan City" + "Council"). Every joined description carries `desc_rows` on the corpus line and the driver's rule 19.2 spot-check asserts EVERY row present in the page text (the 5-word shingle cannot pass a 5-word join such as "Toilet indicator bolt Internal Furniture"). Postscript rows (`NB: ...`) go to the EIL Notes, never into the description. Site/CR header rows are prefixed to each charge under them and named in the Notes.

### F1 recurs on the v3 corpus (Robinson) and split-cell amounts (Kachel)

`printed_gst` held the SUBTOTAL on all 43 Robinson invoices (the `GST:` line interleaves with the payment block in the text layer); restated from the printed `GST: $x` line and re-tied. Kachel header amounts print in split cells and are asserted by a split-cell regex (`18\s*[.\s]\s*00`). Queue with the v33/v34 findings for extraction prompt v3.

### Cost code in the PK position (build trap, v36 patch cycle)

Robinson 11151 prints the cost code `MR4863` where the other invoices print the PK; the adapter carried it into the EIL PK columns as PK004863 on the first pass. Line PK is written only where the printed token is a PK (`^PK\s?\d{6,7}$` gate in the adapter); `pk_typo_map` (PK4000022 -> PK000022, Kachel 7055) handles genuine typos with the printed form retained in col DB.

### Stale pre-capture notes on Confirmed rows (found v36, Open Items #94)

179 rows green-blocked Confirmed by earlier builds still carry the Coding Note "Supplier unidentified; sight attachment before confirming nature and coding" (330 the follow-up "Sight attachment in TechOne; record supplier name and ABN"). The driver now writes a per-family `coding_note_default` and clears the stale follow-up on every row it green-blocks, and the verify rejects a stale note on any target row; the earlier rows are queued for one housekeeping rewrite. Companion ($0.00) rows keep their pre-capture follow-up text until that pass.

### Driver extensions summary (rule 20, v36)

`unmatched`, `companion_rows` (coding note cross-referencing the sighted sibling LineKey; identity only where Unidentified; Status/Basis/green block untouched; verified), `verdict_overrides` (Correct | Review | Confirm), `nature_label` per family and `nature_label_overrides` per invoice, `pk_typo_map`, `desc_rows`, `coding_note_default` with `stale_note_prefixes` / `stale_followup_prefixes`. Corpus adaptation from three v3 corpora is in `b33_prepare.py` (session zip); the v3 adapter notes of v33/v34 still apply (unit price in `unit_price_ex_gst`, description in the PRICED record's `note`; Robinson/Kachel PRICED `note` is null, so the description is rebuilt from the layout rows).

## 37. v37 positions (Batch 34, T & H Levai Pty Ltd.pdf v3 corpus, 17-Aug-2026)

Capture build. Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged). 48 green blocks added (T & H Levai Pty Ltd), no check variants; sighted count 1,696 (v36 1,648 + 48).
Evidence_Invoices: data 5:1820 (+49, including INV-38706 with no register match), totals 1822, correspondence 1824-1825, controls title 1827, controls 1828:1841 (fourteen).
Evidence_Invoice_Lines: data 5:4261 (+51), reconciliation header 4264, reconciliation rows 4265:6080 (1,816 rows).
Config (v37): WORKBOOK_VERSION v37; EIL_DATA 5:4261; EIL_END 4261; RECON_ROWS 4265:6080; RECON_COUNT 1816; EI_DATA 5:1820; EI_TOTALS_ROW 1822; EI_CONTROL_ROWS 1828:1841; SIGHTED_COUNT 1696; VARIANT_TAG_MIN 872 (count 872); GSTINC_ALLOWANCE 40,685.82 (unchanged); BOILERPLATE_KEYS 72 (bump BP:INV-P4 +49).
Data_Acquisition F38 (Batch 34). Method 36.6-36.9. Open Items to #97.

### Legacy Confirmed without rule 17 (found v37, Open Items #95)

202 AP rows carry Status Confirmed on Nature Basis "Line narration" with no Ev Invoice ID and no green block - set by an early build from the LEV002 creditor cross-match, before rule 17. The sighted-line controls never counted them, so every control still ties. A CAPTURE build may green-block such a row when its invoice is sighted (five done at v37: rows 6968/6922/6587/6220/6229); the v35 "a Confirmed row is never an identification target" rule is for identification builds only. Target selection in `b34_prepare.py` therefore does NOT exclude Confirmed rows; `pswp_ident_batch.py` still does. Remaining 197 queued for re-status to Partial.

### Levai (Xero) template (v37 adapter notes)

Officer / contract / PK / "CR - Site (Ref: n)" / "Completed - date" print as the first rows under DESCRIPTION, so header fields are read from fixed row positions and every token asserted present; the site line is parsed (CR = first 7-digit number, `Quote No. n`, site = the segment carrying `(Ref: n)`, trailing segment = work summary). Descriptions wrap onto rows above the first priced row and below a priced row (F12; `desc_rows`). "GST (10%)" prints on the totals block only (per-line GST 10%). Printed contract-reference typos (PAR/3356L/2023, PSR/335L/2023) are captured verbatim; the register Contract Ref carries the printed form.

### Stale follow-up survives a per-invoice note (build trap, v37 patch cycle)

The v36 stale-note replacement ran only where the brief carried no per-invoice coding note, so a row with a per-invoice note kept the pre-capture follow-up "Sight attachment; one invoice can cover ..." (row 9991, INV-36224) - caught by the v36 verify check. Driver fixed: the stale follow-up is cleared on every green-blocked row unless the brief supplies one; the per-invoice note wins over the family default.

## 38. v38 positions (Batch 35, C00052769.zip Weis Contractors PDFs, raw-text route, 17-Aug-2026)

Capture build. Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged). 32 green blocks added (Weis Contractors), no check variants; sighted count 1,728 (v37 1,696 + 32).
Evidence_Invoices: data 5:1852 (+32), totals 1854, correspondence 1856-1857, controls title 1859, controls 1860:1873 (fourteen).
Evidence_Invoice_Lines: data 5:4558 (+297), reconciliation header 4561, reconciliation rows 4562:6409 (1,848 rows).
Config (v38): WORKBOOK_VERSION v38; EIL_DATA 5:4558; EIL_END 4558; RECON_ROWS 4562:6409; RECON_COUNT 1848; EI_DATA 5:1852; EI_TOTALS_ROW 1854; EI_CONTROL_ROWS 1860:1873; SIGHTED_COUNT 1728; VARIANT_TAG_MIN 872; GSTINC_ALLOWANCE 40,685.82 (unchanged); BOILERPLATE_KEYS 72 (bump BP:INV-P14 +32).
Data_Acquisition F39 (Batch 35). Method 37.0-37.3. Open Items to #100.

### Raw-text route gates (v38, first use since v14)

Where no extraction corpus exists, the in-session parser (`b35_prepare.py`) reproduces every corpus gate: per-line qty x unit = amount (297/297), captured lines = printed subtotal, printed GST = ROUND_HALF_UP(subtotal/10), subtotal + GST = printed total, every header token and every description layout row asserted verbatim in the retained page text (`desc_rows`), and normalised-hash dedupe of duplicate copies (2 of 34 files, also identical by file md5). `pdftotext -layout` is the reader; the retained page text ships in `pages_b35.json`.

### Weis (Xero) template

The line table prints the works date and each site as ZERO-amount rows (qty 1.00, unit 0.00) above the priced work lines - 161 of the 297 captured lines - and Contract / P/O # / PK as three unpriced head rows. On INV-2112 and INV-2125 that head block is itself a printed zero-amount table row and is captured as one line with three joined layout rows; elsewhere it is header text captured in the green-block Contract and Purchase Order fields. GST prints as a rate column (10%), so per-line GST is the printed rate, never a derived dollar figure. Page 2 is the PAYMENT ADVICE remittance slip. Header labels and values print on separate rows interleaved with the supplier address column and are de-interleaved on parse.

### Findings (Open Items #98-#100)

The PK000417 high-pressure cleaning round is proven across fifteen sighted invoices (10-Jul-2023 to 24-Jul-2024) billing seven named parks at fixed per-site rates ($245/$325/$385/$245/$180/$195/$180) roughly fortnightly under PO 494375, ceasing after INV-1918; later cleaning moves to PK000022 under PO 703472/709937 (#98). INV-2526 ($13,571.00, Springwood Park plinth recoating) is project-scale work billed on the cleaning contract (#99). INV-2182 prints PK000022 but posts to PK000417 (#100). INV-1582 is a Play Force reference collision resolved by the EvID override WC-1582 (Method 34.2).

## 39. v39 positions (Batch 36, HIG010 Higgins Coatings + Harpley Services creditor-attachment zips, raw-text route, 17-Aug-2026)

Capture build. Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged). 40 green blocks added (Higgins Coatings Pty Ltd 26 / $37,611.00; Harpley Services Pty Ltd 14 / $31,765.13), one check variant (INV-00013128 template B, check-1 inclB); sighted count 1,768 (v38 1,728 + 40). 25 Higgins $0.00 companion rows cross-referenced (schema section 8), not green-blocked. Rows 1848/2036 (legacy Confirmed on the narration, Open Items #95 pattern) green-blocked and re-identified HAR073 / 22 162 601 694.
Evidence_Invoices: data 5:1892 (+40), totals 1894, correspondence 1896-1897, controls title 1899, controls 1900:1913 (fourteen).
Evidence_Invoice_Lines: data 5:4626 (+68), reconciliation header 4629, reconciliation rows 4630:6517 (1,888 rows).
Config (v39): WORKBOOK_VERSION v39; EIL_DATA 5:4626; EIL_END 4626; RECON_ROWS 4630:6517; RECON_COUNT 1888; EI_DATA 5:1892; EI_TOTALS_ROW 1894; EI_CONTROL_ROWS 1900:1913; SIGHTED_COUNT 1768; VARIANT_TAG_MIN 873 (count 873); GSTINC_ALLOWANCE 40,890.42 (+204.60, the template B capture); BOILERPLATE_KEYS 75 (new BP:HIG-P2, BP:HIG-T2, BP:HIG-T3; bumps BP:INV-P1 +14, BP:INV-T1 +14).
Data_Acquisition F40 (Batch 36). Method 38.0-38.4. Open Items to #105. Project Instructions v10 EMBEDDED on the Project_Instructions sheet at this build (§6.3 / §7.3 same-step rule); the sheet no longer carries v9.

### Driver absent from the uploads: self-contained one-call script (v39)

`pswp_build_batch.py` and the rest of the session-zip toolkit were not uploaded with the batch, so the build ran as `b36_build.py` (session zip): the same write set (EI/EIL append, recon relocation with per-row E-form preservation, green blocks from the row-14937 style template, companion notes, both citation limbs rewritten from the EvID maps, EI tail relocation with explicit totals-row SUMs and regex-retargeted control formulas, Vendor_Boilerplate keys and bumps, Handover/Method/Data_Acquisition/Open_Items/Vendor_Series/Config), the convert-route recalc under an isolated profile (`registrymodifications.xcu` with OOXMLRecalcMode=0 written in-session) and the exact-TRUE verify (EI controls exact "TRUE" with the two count rows equal to SIGHTED_COUNT, every sighted line's three checks, every recon row, both citation limbs, Coverage tie, register total, held rows and companions carry no EvID, stale notes absent, cited BP keys exist, no formula errors). One rule-18 patch cycle (verify logic only: the two EI count rows and a shadowed variable) - the workbook was correct on the first pass. Fold `b36_build.py`'s Higgins / Harpley template handling into `pswp_build_batch.py` and `b36_prepare.py`'s parsers into `pswp_parsers.py` when the toolkit next travels.

### Higgins (JD Edwards FORMAR001) template (v39; extends the v11 HIG precedent)

One DESCRIPTION OF SUPPLY line per invoice printed over 2-4 layout rows with the amount on one of them: rows joined with " | " (v11 form, F12), every row asserted verbatim in the retained page text, `amount_row` retained on the corpus line. The Order Ref (`PO495370/PK000022/CR3508822`, PO702589 from FY2024/25, PO709746 from FY2025/26; three invoices print a PO-level reference with no PK) is captured verbatim in the Purchase Order field and PK / CR/WO are parsed from it. Requesting officer, ship-to, technician and payments/balance are not printed. Two terms wordings across the series (2004 Act to Jun-2024, 2017 Act from Sep-2024) and one payment block, minted verbatim as BP:HIG-P2 / T2 / T3; the v11 keys BP:HIG-P1 / T1 hold condensed wording (Open Items #104). Every Higgins invoice but one posts with a $0.00 companion AP row (26 invoices, 25 companions).

### Harpley template A/B and an image-only PDF (v39)

Template A as v33 (bare GST amount under Sub Total, wrapped descriptions per F12, `desc_rows` asserted); 00013391 prints the totals with amounts on page 2 only (rule 13); two files carry a "How to pay" (Pay online / QR) slip as page 2 - not invoice content, noted in Pages. Template B (00013128, "Your Order No"): lines print GST-INCLUSIVE and sum to the printed Total Inc GST - check 1 inclB, recon row on the tolerance form against $2,250.60, GSTINC_ALLOWANCE +$204.60, EIL Notes on every line. IMAGE-ONLY copy: C00041822 (4).pdf (00012692) carries no usable text layer (pdftotext returns "489578"); read visually from a 300 dpi render, transcribed into the template A layout so the same parser and gates ran, and cross-checked token-by-token against a tesseract OCR layer (identifiers, dates, all seven line descriptions and amounts, subtotal, total present; the OCR misreads the printed GST $1,258.78 as $1,298.78 - printed value confirmed visually and arithmetically). Both layers ship in `pages_b36.json`; the anomalies field states the basis. Where a raw-text batch carries an image-only file, this is the route: visual transcription + OCR cross-check + full arithmetic gates, never a summary row.

### Findings (Open Items #101-#105)

Seven Higgins printed-vs-charged PK divergences (PK000023 / PK000508 printed on PO702589 work orders, PK000022 / PK000417 charged; $7,044.00) (#101); Harpley 00012990 prints PK#000418 and posts to PK000415 (#102); Harpley 00012692 ($12,587.76, January-2023 pump outs, LCC-03-2018) is dated 1-Mar-2023 (FY2022/23) and posts in the FY2023/24 scope at row 227 (#103); Vendor_Boilerplate HIG-P1/T1 wording (#104); the same Higgins scope coded 73212 in FY2023/24 and 73123 from FY2024/25, with Towns Park NRL line marking recurring at a fixed $1,989.00 (#105).

## 41. v40 positions (Batch 37, binder_batch_20260818 corpus over Binder1.pdf + Binder111.pdf, 18-Aug-2026)

Capture build. Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged). 63 green blocks added (Play Force Australia Pty Ltd 43, BLC Queensland Pty Ltd 11, Weis Contractors 7, Flavell-Dau 1, Grange 1951 1); one invoice captured under rule 16 only (Ozifresh 0012910). Three check variants (BLC-001270 / BLC-001310 check 3 tol1cgst; INV-5142 check 2 tol1c + check 3 tol1cgst). Sighted count 1,831 (v39 1,768 + 63). 36 zero-amount companion rows cross-referenced. Four Play Force invoices HELD, uncaptured (Open Items #108).
Evidence_Invoices: data 5:1956 (+64), totals 1958, controls 1964:1977 (fourteen).
Evidence_Invoice_Lines: data 5:4785 (+159), reconciliation rows 4789:6740 (1,952 rows).
Config (v40): WORKBOOK_VERSION v40; EIL_DATA 5:4785; EIL_END 4785; RECON_ROWS 4789:6740; RECON_COUNT 1952; EI_DATA 5:1956; EI_TOTALS_ROW 1958; EI_CONTROL_ROWS 1964:1977; SIGHTED_COUNT 1831; VARIANT_TAG_MIN 876; GSTINC_ALLOWANCE 40,890.42 (unchanged); BOILERPLATE_KEYS 80 (new BP:BLC-P1, BP:BLC-T1, BP:GRA-P1, BP:OZI-P1, BP:INV-P23).
Data_Acquisition F41. Method 39.0-39.4. Open Items to #112.

### Green blocks are AP lines, not one doc-type string (v40)

Register row 228 (Grange 1951, reference 380CFR1) is the first green block on Doc Type `Creditors invoices`. Section 8 said "AP lines (Doc Type = PUR Cred Invoice)"; the workbook in fact green-blocks `PUR Cred Invoice` (961 at v39) and `DIR Cred Invoice` (807), and `Creditors invoices` is a source-export vintage of the same class carried by only seven rows (228, 358, 710, 1535, 4477, 8383, 12261). Read section 8 as naming all three. Journals and PCard rows are still never green-blocked (Open Items #111).

### A purchase-card row is not an AP line (v40, extends the v36 `unmatched` pathway)

Ozifresh 0012910 ($3,706.56) has no AP creditor line: the spend posts at row 256 as a `PCard Expense` (reference TE003933, narration `SQ *OZIFRESH`), already Confirmed on the system record, with the contractor recorded as the CARD PROVIDER (Corporate Travel Management) rather than the supplier. Captured under rule 16 only - EIL lines with blank LineKey, a recon row, an EI row whose LCC Reference reads `(no register match; Open Items #109)` and whose col 13 register limb reads `Register row (no register match)`, cols 14-31 staged. The verify asserts the EvID appears on no register row. Purchase-card postings are a standing identification gap: the register names the card, not the vendor.

### Extractor read the DUE date as the invoice date on every Xero layout (F1, v40)

The Xero letterhead prints the `Invoice Date` label and its value on SEPARATE interleaved layout rows, so a same-line regex returns the Due Date. Wrong on all 15 Xero-layout invoices in the batch (7 Play Force, 7 Weis, 1 Flavell-Dau) and it moves the FY on some (INV-2575/INV-2578 read FY2026/27, are FY2025/26). Restate by scanning the two rows below the label and cross-check every restated date against the register Doc Date - 65 of 66 matched rows then agree to the day. The one legitimate gap is a posting on the printed `Commencement Date` (INV-3895, posts 16-Jun-2025 against a 25-Jun-2025 invoice).

### Head-block scanning must stop at the first priced row (build trap, v40)

Play Force and Weis print `Account: PK000022` as a head row AND `Account: 10367833` in the payment block. An unbounded scan for `Account:` reads the bank account number as the PK. Bound the head-block scan to the rows between the column-header row and the first priced row.

### The duplicate screen runs on the corpus too, not only against the workbook (v40)

Invoice 001311 appears in BOTH source binders with identical normalised text while the extractor's own report declared "Duplicate copies: 0". Screen doc_ref collisions and normalised-text hashes WITHIN the corpus before the workbook screen. Separately, document segmentation merged two BLC invoices into one record (001290 cut as pages 13-16; pages 15-16 are 001291), so a whole invoice was absent from a corpus that reported full page coverage: scan every document's page text for invoice-number tokens that are not its own, whitespace-tolerant, because a degraded text layer prints `Invoice n u m b e r : 001 291`.

### Amount-column-only line tables (F9, v40)

Where a line table prints an amount column with no qty and no unit price (Grange 1951 misc template; Flavell-Dau single charge over seven dated rows), the extractor classifies every line NARRATIVE and the document captures $0.00 - the whole invoice is lost, not a line. Detect by a self-tie of exactly the printed subtotal below zero capture, and rebuild from the amount column with the basis in the anomalies.

### BLC template (v40)

BLC prints NO line table: a Description block and one printed Sub total, captured as a single-scope line tied to the printed subtotal. Amounts print with VARIABLE decimals ($3,660, $366, $65.1, $908.74), so a `\d+\.\d{2}` amount regex misses whole-dollar and one-decimal figures. `printed_total_incl_gst` held the SUBTOTAL on all 11 invoices (F1). The vendor ABN prints mis-grouped in the footer as `166 392 737 51` (= 16 639 273 751) while LCC's own ABN prints under `Client Details`; take the vendor ABN from the footer, assert digits-only, capture the printed grouping verbatim.

## 42. v41 positions (Batch 38, levi_bind_3, plus Open Items #107 and #108 closed, 18-Aug-2026)

Capture build. Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged). 46 green blocks added across three cohorts: Batch 38 proper (T & H Levai 39, 4Park 1, TradeTools 1), the Open Item #107 recovery (BLC 001291) and the Open Item #108 release (Play Force 4). Four check variants (check-1 tol1c). Sighted count 1,877 (v40 1,831 + 46). Two new Tier 1 Vendor_Series labels (4Park Pty Ltd, TradeTools Pty Ltd).
Evidence_Invoices: data 5:2002 (+46), totals 2004, controls 2010:2023 (fourteen).
Evidence_Invoice_Lines: data 5:4839 (+54), reconciliation rows 4843:6840 (1,998 rows).
Config (v41): WORKBOOK_VERSION v41; EIL_DATA 5:4839; EIL_END 4839; RECON_ROWS 4843:6840; RECON_COUNT 1998; EI_DATA 5:2002; EI_TOTALS_ROW 2004; EI_CONTROL_ROWS 2010:2023; SIGHTED_COUNT 1877; VARIANT_TAG_MIN 880; GSTINC_ALLOWANCE 40,890.38 (-0.04); BOILERPLATE_KEYS 83 (new BP:LEV-P1, BP:4PK-P1, BP:4PK-T1).
Data_Acquisition F42. Method 40.0-40.4. Open Items to #116 (#107, #108, #115, #116 closed).

### RATIFIED: check-1 per-line rounding tolerance (v41, Method 40.1; fold into instructions rule 17)

Where a vendor prints each line amount rounded to two decimals but computes the printed subtotal from the UNROUNDED values, captured lines cannot equal the printed subtotal and the printed document does not self-tie. Check 1 takes the standard tolerance form `=IF(ROUND(ABS(DQ{r}-DI{r}),2)<=0.01,"TRUE","FALSE")`, tagged `[check variant]`, with the delta and a worked example in the anomalies. This completes the tolerance set across all three checks, alongside the already-ratified check-2 one-cent (TechOne ex-GST derivation) and check-3 rounding tolerances. Precedent: Play Force Australia INV-4412 / INV-6739 / INV-7609 / INV-7759 (5.50 x $105.41 = $579.755, printed $579.75), held at v40 and released here.

TWO CONSEQUENCES that must be carried in the same build, both caught by the v41 verify:
- the reconciliation row for such an invoice takes the tolerance E-form `=IF(ROUND(ABS(D{r}-{target}),2)<=0.01,...)`, not the exact form (a third ratified recon form alongside exact and `=0`);
- the Evidence_Invoices header ex-GST totals then stand ABOVE the sum of captured lines, so GSTINC_ALLOWANCE moves by the negative of the total shortfall (-$0.04 for four invoices at 1c).

### A defective text layer is de-spaced and sibling-checked, not discarded (v41, Open Item #107 closed)

BLC 001291 was absent from the Batch 37 corpus because segmentation merged it into record 001290; it was recovered from that record's retained page text. The page-15 text layer inserts spurious spaces INSIDE words ("L o g a n City Council", "Invoice n u m b e r : 001 291", "9thM a y 2025"). Route: collapse single-character runs, then cross-check every reconstructed field against a CLEAN SIBLING invoice of the same template captured in the same batch (here eleven v40 BLC captures sharing vendor, address, phone and site-contact wording verbatim); printed FIGURES carry no spacing artefact and are read directly. State the basis in the anomalies field per rule 16a - reconstruction of a defective text layer is not paraphrase. This sits alongside the v39 image-only route (visual transcription plus OCR cross-check) and the v32 de-interleaving method.

### Defect family F9 is the dominant extraction failure (v41)

An amount-column-only line table - a DESCRIPTION / AMOUNT table with no qty and no unit-price column - makes the extractor classify every line NARRATIVE, so the document captures $0.00 and the WHOLE INVOICE is lost, not a line. Now observed on five documents across five vendors in two consecutive batches (Grange 1951, Flavell-Dau, and four Levai). Detect by a zero capture against a present printed subtotal, then rebuild from the amount column and re-tie. Two related rule 16 failures in the same batch: 4Park lumped both printed lines into a single summary record (a summary row in place of line capture is a critical failure), and TradeTools attached the line amount to a footer warranty clause row.

### Levai template and letterhead drift (v41)

Levai (Xero) prints DESCRIPTION / AMOUNT with head rows for the requesting officer, contract, PK, a `CR - Park (Ref: n)` job line and a completion date; per-line qty, unit price and GST are not printed. The letterhead address changes from `Lot 2 Industrial Ave` to `41 Industrial Ave`, Logan Village between May-2024 and Jun-2024 under the same ABN (and from `34-36 Calcium Court, Crestmead` before that), so the address is parsed per invoice and never carried across the series.

### Two vendors, one address (finding, Open Items #113)

T & H Levai Pty Ltd (65 100 395 480) and Play Force Australia Pty Ltd (89 677 476 541) print the SAME premises, 41 Industrial Ave Logan Village QLD 4207, with adjacent landlines ((07) 3803 0032 and (07) 3803 1788), across 123 and 118 sighted green blocks respectively. Separate legal entities on the printed ABNs (rule 8), but a related-party and probity question for Finance where both hold large concurrent Parks contracts against the same PKs.

## 43. v42 positions (Batch 39, site dimension + waste/energy analysis, 18-Aug-2026)

Enrichment build, not capture. Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged - column T is never written on an enrichment build). Sighted count unchanged at 1,877; EI/EIL/recon positions all unchanged from v41.
Register gains TWO COLUMNS: **129 (DY) Park / Site (normalised)** and **130 (DZ) Site basis**. The register is now 130 columns; the green block stays at 88:127 (CJ:DW) and column 128 (DX) remains the empty `Src __md5Row (26SLACT)` source field.
Sites sheet REBUILT: Panel A cost per park (483 parks, live COUNTIF/SUMIF/SUMIFS on col DY), two group rows keyed on col DZ, totals row and a three-way tie check; Panel B is the v41 electricity-by-retailer-account panel relocated unchanged (394 rows).
Config (v42): WORKBOOK_VERSION v42; SITE_COL 129 (DY); SITE_BASIS_COL 130 (DZ); SITES_PANELA 5:487; SITES_GROUP_ROWS 488:489; SITES_TOTAL_ROW 490; SITES_CHECK_ROW 491; SITES_PANELB 495:888; PARK_COUNT 483; SITE_ATTRIBUTED_LINES 5452; SITE_MULTI_LINES 319. All capture-side keys unchanged.
Data_Acquisition G1a/G2/G3. Method 41.0-41.4. Open Items to #128.

### The site dimension (NEW v42)

Cost per park is a first-class register dimension. Column DY carries the normalised park or site name, column DZ the basis. A line is attributed ONLY where a source field carries a FULL PARK NAME matching the gazetteer (user directive, 18-Aug-2026); nothing is inferred from a bare street address. Sources are read in evidence order and the first hit wins, with the source recorded as the basis: Rates parcel narration (1,208) > Sighted invoice site printed (1,016) > Sighted invoice work description (48) > Electricity supply point (3,062) > GL narration (402) > Enquiry narration (35).
Coverage: 5,452 lines / $2,407,766.77 (21.2%) attributed to 483 parks; 319 lines / $904,116.12 (7.9%) held multi-site; 15,698 lines / $8,065,429.62 (70.9%) carry no park name and read `(no site named)` in DZ with DY blank. Every row carries a basis - there are no blanks in DZ, so every group is SUMIF-able and the tie check is exhaustive.

### The gazetteer, and why generic names are blocklisted (v42)

646 names. Primary source: the internal-rates S07 narration series, which prints parcel id, street address and park name (1,139 parcels, 583 named) - authoritative, properly cased Logan park names with addresses. Secondary harvest: any capitalised multi-word name ending Park / Reserve / Gardens / Wetlands / Forest / Parklands / Oval / Bushland / Green / Square / Common recurring three or more times across site details, work descriptions and supply points.
BLOCKLIST is mandatory: General Park, Skate Park, Dog Park, Water Park, Car Park, Memorial Park, Sports Park, District Park, Adventure Park and the like name a facility TYPE, not a site. Unblocked, "General Park maintenance" attributed 152 lines and $72,856 to a park that does not exist.

### Multi-site lines are HELD, never split (v42, extends the v32 partial-scope principle)

319 lines name several parks on one line (a grounds contract billing a schedule of sites, a standing order across a round). DY carries the semicolon-joined list of every park named; DZ reads `Multiple sites named (not allocated)`. No allocation is made, because no evidence fixes the split. Precedent: row 14039 ($33,335.10, PAR/340E/2026) prints "Logan Gardens; Flagstone; Darlington Parklands; ..." and a single-park pass had attributed the whole amount to one of them.

### COUNTIF IS CASE-INSENSITIVE (build trap, caught by the v42 verify)

Panel A over-counted by 799 lines and $619,244 on the first pass because the rates-narration path wrote the RAW printed park name (`STURDEE PARK`, `LOGAN GARDENS`) while every other path wrote the gazetteer label (`Sturdee Park`, `Logan Gardens`). Both appear as distinct Panel A labels, but COUNTIF/SUMIF match case-insensitively, so each row counted twice. STANDING RULE: any label written to a register column that a COUNTIF panel keys on must come from ONE canonical map, and the build asserts no two labels share a case-folded key before writing. Fixed in a rule-18 patch cycle; the tie check now proves lines = 21,469 and $ = Register!$T$21475 exactly.

### Batch 39 reconciliations (analysis, shipped alongside the workbook)

FY2025/26 public place bins: all twelve monthly Parks postings equal the Waste branch driver (LCC DOCS-19541843-v1, row Parks, coded `(PK) PK000022-7B115`) to the cent, $1,123,128.88. FY2025/26 electricity: 746 register lines on the five Origin statements (A-FE35E476-021/-028/-029/-030/-031) group to 234 sites by NMI and every one ties register ex GST to the Origin statement site total to the cent, $27,621.35. Worked in `PSWP_Batch39_Waste_and_Energy_Analysis.xlsx`.
Waste account break: the Parks share of Cleanaway public place bin lifts is coded 73122 in FY2023/24 and FY2024/25 and 7B115 from FY2025/26 - trending either account alone shows a false 97.5% saving (Open Items #117).

## 44. v43 positions (combined enrichment build: Themes restated + exhaustive site basis, 18-Aug-2026)

Enrichment build, not capture. Register: data 5:21473, total row 21475, control total $11,377,312.51 (unchanged; column T never written; green-block count unchanged at 1,877 sighted). Only columns DY (129) and DZ (130) written, on the 15,698 rows that read `(no site named)` at v42 plus the 75 rows carrying the leaked facility-type label `Undertake General Park`.
Sites: Panel A 5:486 (482 parks), group rows 487:495 (nine closed-list classes), total 496, tie check 497, Panel B 501:894. Themes: rows 5:17, total 18, controls 21:26 (six, all live or frozen-correct), data-quality block 30:34.
Config (v43): WORKBOOK_VERSION v43; SITES_PANELA 5:486; SITES_GROUP_ROWS 487:495; SITES_TOTAL_ROW 496; SITES_CHECK_ROW 497; SITES_PANELB 501:894; PARK_COUNT 482; SITE_ATTRIBUTED_LINES 5482; SITE_MULTI_LINES 408; NEW SITE_CLASS_LABELS 9; SITE_NONE_LINES 250; SITE_SOURCE_COLS 111,106,112,108,127,99,42,22,23,80; SITE_BANNED_COLS 12,25,32,92,98,105; THEMES_TOTAL_ROW 18; THEMES_CONTROL_ROWS 21:26. All capture-side keys unchanged.
Method 42.0-42.5. Data_Acquisition G4-G7. Open Items to #134. Session script b43_build.py + b43_verify.py (session zip); the writes are listed in b43_site_writes_for_verification.json.

### Column DZ is exhaustive (standing rule from v43)

No register row may read `(no site named)`. Every row carries either a park-attribution source label, `Multiple sites named (not allocated)`, or one of the closed-list classes: `Non-park Council facility (ship-to = Parks Depot)`, `City-wide network service (no site by design)`, `Accounting event (no site by construction)`, `Internal charge (not site-specific)`, `Invoice sighted, printed site not in gazetteer (candidate harvest)`, `Supply-point address only (address route pending)`, `No invoice sighted (awaiting capture)`, `No site determinable`. The Sites sheet carries one COUNTIF/SUMIF group row per class and the tie check remains exhaustive. Position at v43: 5482 single-park lines / $2,853,325 (25.1%) on 482 parks; 408 multi-site holds / $1,742,439; network service 108 / $3.07M; awaiting capture 8,627 / $2.67M; candidate harvest 606; address only 3,075; accounting event 2,566; internal 326; depot 21; residual 250.

### The site source hierarchy reads the WHOLE ROW, with a ban list (v43)

v42 read six fields and missed the two richest. Order: 111 Ev Line items (printed) > 106 Ev Site Details > 112 Ev Works carried out > 108 Ev Work Description > 127 Ev Anomalies > 99 Ev Ship To (depot only) > 42 Site/supply point > 22 GL narration > 23 Enquiry narration > 80 Src User Fld 2. First hit wins; the basis records the source. Cols 111 and 112 alone carried $1.26M of park names never read before. BANNED as sources, permanently: 12 Contractor and 32 Coding Note (Claude's own labels), 25 Nature Detail (derived from 111), 92 Ev Vendor address (supplier premises - "15 Helios St, Shailer Park" matched the suburb on 20 lines), 98 Ev Bill To (Council address), 105 Ev PK normalised. Standing rule: a supplier's premises, Council's own address and any analysis-authored label are never site evidence.

### openpyxl `ws.cell(r, c, None)` does NOT clear a cell (build trap, v43 patch cycle)

Passing `value=None` to `ws.cell()` leaves the existing value in place. To blank a cell set `.value = None` explicitly. Caught by the verify (`leaked label absent from register`) on the second pass; the 75 `Undertake General Park` rows had kept their DY value while DZ moved to a class.

### Themes controls read FALSE from v29 to v42 (found and repaired v43, Open Items #134)

The grand-total control tested $10,677,585.33 (v14) and the FY2026/27 control a P1 subtotal; the sheet header still said three years. Neither the EI control set nor the Coverage tie covers Themes, so both were FALSE for thirteen builds. Restated four-year with six controls (three frozen prior-year extracts, FY2026/27 live as register minus those three, all-years = Register!$T$21475, PK000022 = SUMIFS). STANDING RULE: the verify asserts every Themes control exact TRUE, and any part-year scope ties live, never to a frozen figure.

### Address route: street-name join REJECTED, property-level route pending (v43, Open Items #131)

26 Parks Directory streets serve two or more parks (Underwood Rd four, Paradise Rd four); OSM/Nominatim geocodes to the road centre with no house-number precision, and a 200 m point-in-polygon against layer 121 disagreed with the directory on the three largest supply points (99 Ewing Rd -> Timms Park vs Rainbow Park). Reopen only on G-NAF QLD_ADDRESS_DEFAULT_GEOCODE points or Council cadastre keyed on Property_Key, zero-buffer point-in-polygon, every hit flagged for user check.

## 45. v44 positions (site read at printed-line level; electricity address crosswalk, 18-Aug-2026)

Enrichment build. Register: data 5:21473, total 21475, control total $11,377,312.51 (unchanged; column T never written; 1,877 green blocks unchanged). Only DY/DZ written (1,130 rows).
NEW Evidence_Invoice_Lines cols L Site (printed line) / M Site basis, rows 5:4839 (1,180 lines name one park, 42 several, 3,613 none). NEW sheet Site_Allocation (after Sites): data 5:632 (628 rows over 128 register LineKeys, $1,187,581.57), total row 634, tie check 635. Sites: Panel A 5:494 (490 parks; $ = SUMIF(Register DY) + SUMIF(Site_Allocation G:H)), group rows 495:504 (residue + nine classes), total 505, check 506, amber-electricity count row 507, Panel B 511:904.
Config (v44): WORKBOOK_VERSION v44; SITES_* repositioned; PARK_COUNT 490; SITE_ATTRIBUTED_LINES 6484; SITE_MULTI_LINES 280; NEW SITE_ALLOC_DATA 5:632, SITE_ALLOC_TOTAL_ROW 634, SITE_ALLOC_LINEKEYS 128, SITE_ALLOC_ROWS 628, EIL_SITE_COLS 12:13 (L:M), ELEC_ADDR_GREEN 420, ELEC_ADDR_AMBER 582, ELEC_ADDR_UNRESOLVED 2073. Method 43.0-43.3. Open Items to #139. Scripts b44_build.py / b44_verify.py / elec_xwalk.py; writes listed in b44_electricity_writes.xlsx.

### Two-level site model (standing from v44)

Level 1 = Evidence_Invoice_Lines L:M, site per PRINTED LINE. Level 2 = Register DY/DZ. A register row whose tied printed lines name several parks reads `Multiple sites named (allocated per printed line, see Site_Allocation)` and its dollars are exploded on Site_Allocation using the PRINTED line amounts, with a residue row `(residue: printed lines naming no park or several)` forcing each LineKey to tie to Register!T to the cent. This is rule 16 evidence read at its own granularity; nothing is apportioned. A multi-site row with NO tied printed lines stays `Multiple sites named (not allocated)`. Sites Panel A sums both levels; the tie check stays exhaustive. Standing rule: never hold on a register-level field what the printed lines beneath it already allocate (v43 held $1.19M this way; corrected v44). Precedent Pool Shop INV-5278 (one register line, fifteen printed lines across five water-park sites and four PKs).

### Electricity supply points: Council address records, three flags (v44)

Indexes: S07 internal-rates parcel narrations (Council's own parcel -> street -> park record), Parks Strategy Layer House_Add, Parks Directory address. number+street exact and unique -> `Electricity supply point (address = Council parcel record, number and street exact)` (420 lines); street unique, no number match -> `Electricity supply point (address = Council parcel record, street unique) [USER CHECK]` (582 lines, own group row, listed for the user); street ambiguous or unknown -> class unchanged (2073 lines; 99 Ewing Rd Woodridge 217 lines -> Booran / Rainbow / Willow). External geocoding remains rejected; unresolved rows await G-NAF property points or a Property_Key join.

### Gazetteer: water-park printed sites (v44)

Flagstone Water Play, Beenleigh Water Play, Forest Glen Water Feature added AS PRINTED (absent from all three park datasets); parent parks to be confirmed with Water Parks (Open Items #139). Underwood Water Feature -> Underwood Park; Logan Gardens Water Play -> Logan Gardens. None of the three new names has a coordinate yet.

## 46. v45 positions (Batch 40 capture: Q Power + BrizSouth Locksmiths, 18-Aug-2026)

Capture build. Register: data 5:21473, total 21475, control total $11,377,312.51 (unchanged). 44 green blocks added (Q Power (Qld) Pty Ltd 13 / QPO001, BrizSouth Locksmiths Pty Ltd 31 / BRI115), 121 printed lines, $17,677.49 ex GST; no check variants. Sighted count 1,877 -> 1921. 28 $0.00 companion rows cross-referenced, not green-blocked. 41 of 44 rows gain a Tier 1 printed site (and their companions with them).
Evidence_Invoices: data 5:2046 (+44), totals 2048, controls 2054:2067 (fourteen).
Evidence_Invoice_Lines: data 5:4960 (+121), reconciliation rows 4964:7005 (2042).
Config (v45): WORKBOOK_VERSION v45; EI_DATA 5:2046; EI_TOTALS_ROW 2048; EI_CONTROL_ROWS 2054:2067; EIL_DATA 5:4960; EIL_END 4960; RECON_ROWS 4964:7005; RECON_COUNT 2042; SIGHTED_COUNT 1921; BOILERPLATE_KEYS 83 (new BP:QP-P1, BP:QP-T1, BP:C-T1); VARIANT_TAG_MIN 880 and GSTINC_ALLOWANCE 40,890.38 unchanged. Data_Acquisition F43. Method 44.0-44.3. Open Items to #143.

### F1, third form: printed_total_incl_gst holds the SUBTOTAL (v45)

On all 31 BrizSouth Locksmiths invoices the corpus `printed_total_incl_gst` equalled the subtotal, because the printed TOTAL sits below the `PLUS 10% GST` line and the same-line regex returned the subtotal. Restated as subtotal + printed GST and re-tied on all 31, basis stated per invoice in the anomalies. This joins the v33 (Harpley bare-$ GST line), v34 (three null Xero fields) and v36 (Robinson GST = subtotal) findings; extraction prompt v3 needs one totals-block rule covering all four.

### Recon relocation must rewrite BOTH the range AND the row self-reference (build trap, v45 patch cycle 1)

Relocating the reconciliation block copies each row's D and E formulas verbatim. Two things break: the D-formula SUMIF still cites the OLD `$A$5:$A${old EIL_END}`, and the E-formula still cites the row's OLD row number (`=IF(D4843=...)` sitting on row 4964). Rewrite the EIL data range AND repoint `D{n}`/`C{n}` to the row's own new row, while PRESERVING the row's ratified E-form (exact, `=0` tolerance, or `<=0.01` tolerance - schema §10/§42). Verify: every recon row exact "TRUE".

### EI control formulas carry self-references into the control block (build trap, v45 patch cycle 1)

Controls cite each other (`=IF(AND(C2010=1921,C2010=C2011),...)`) and cite the EI data tail (`$AF$5:$AF$2003`). A tail shift must remap every `C{old control row}` to its new row and the `$AF` staging range to the new EI end, alongside the EIL/recon range transforms and the literal sighted/recon counts. Four controls read FALSE on the first pass from this alone.

### Nature Category must exist in Theme_Map or the Themes controls go FALSE (rule 4, enforced v45)

The build first wrote `Electrical, lighting and data` and `Security and access control` - neither in Theme_Map A5:A31 - so the Theme lookup returned blank on 44 rows and all six Themes controls read FALSE. Correct values are `Electrical & data services` (T3) and `Security & access` (T12). The verify now asserts every newly written Nature Category exists in Theme_Map and every new row resolves a Theme. Restating Nature Category on a capture build is a THEME-AFFECTING write.

### Site is printed on both templates (v45)

Q Power (simPRO) prints `Site:` and `Site Address:` as header rows, plus `Ordered By:`, `Technician:` and a `Works Completed:` block; BrizSouth prints `Job Location: <park>, <SUBURB>` in the description block and the PK inline. 41 of 44 invoices name a gazetteer park on their own face, so this batch is unusually strong for the site dimension. Two print no site (C13293, C13403); seven print no PK (Open Items #140).

## 47. v47 placeholder / v46 positions (Batch 41 identification + site corrections, 18-Aug-2026)

Identification build plus site corrections. Register: data 5:21473, total 21475, control total $11,377,312.51 (unchanged; column T never written; green blocks unchanged at 1,921). 485 rows identified, $330,915.43 ex GST.
Creditor_Lines: data 5:11194 (+5,602). NEW sheet Site_Crosswalk (after Sites), data 5:11. Sites: Panel A 5:498 (494 sites), groups 499:508, total 509, check 510, Panel B 514:907.
Config (v46): WORKBOOK_VERSION v46; CL_DATA 5:11194; SITES_* repositioned; PARK_COUNT 494; NEW SITE_CROSSWALK_DATA 5:11, IDENT_ROWS_V46 485, IDENT_HELD_V46 20. Method 45.0-45.3. Data_Acquisition F44. Open Items to #149 (#144, #145 closed).

### Five creditor histories close four unidentified series (v46)

ELE032 Electrical Data & Security Services Pty Ltd (4,444 invoices, ABN 17 099 805 277) resolves the FY2023/24 and FY2024/25 7-digit series, 325 rows / $160,485. QPO001 Q Power (Qld) Pty Ltd (1,131, ABN 82 067 507 591) resolves the FY2024/25, FY2025/26 and FY2026/27 5-digit series, 141 rows / $123,110, and closes Open Item #89 - the incl/1.1 amount test separates Q Power from Logan Plumbing on the shared series, 171 of 187 references corroborating. SEC022 (ABN 97 613 325 769) 17 rows; SEA029 Sea-Crete (ABN 84 127 054 291) the $34,760 invoice 7558; HIB003 (ABN 30 306 518 188) 1 row. 20 references matched on reference but FAILED the amount test and are held (Open Items #147). The 7-digit series is MULTI-VENDOR (ELE032 and SEC022 both post 7-digit references) - short numeric series never identify on their own.

### A published address is a starting point, not an authority (standing rule, v46)

The Parks Directory lists Ewing Park at Netball Drive and Rainbow Park at Ewing Road, so an address match sent all 237 lines at `99 Ewing Rd Woodridge` to Rainbow Park; the correct site is Ewing Park. STREET-ONLY matching is banned outright: `293 Logan St Eagleby` resolved to Doug Larsen Park, which is 41 Logan Street BEENLEIGH - a cross-suburb error on 42 lines whose true site is Olivers Sports Complex. Only street PLUS suburb may be proposed, and only as a flagged proposal. 121 street-only segments / 885 lines / $112,153 are held for user adjudication (Open Items #148).

### NEW sheet Site_Crosswalk: manual decisions outrank the matcher (v46)

One row per adjudicated source text: decided site, type, basis, decided by, date. Applied on every build BEFORE the matcher, so local knowledge survives a rebuild. Seeded with seven decisions. The removal of `Shailer Park` is the exemplar: it is a SUBURB and David Robinson Landscaping's business address (15 Helios St, Shailer Park, per PO 802122), it had taken $50,140 across 16 lines, and it beat `Homestead Park (Shailer Park)` on an exact match. Banning a column (v43) was not enough - a suburb can enter through any field.

### Non-park Council facilities are sites in their own right (v46)

`Non-park Council facility (named site, not a park)` extends the Parks Depot precedent from one facility to a class: Olivers Sports Complex, West Logan Netball Courts, Historical Society Amenities, Logan and District Services Club. Real Council sites carrying real spend, neither forced onto a park nor left unresolved.

## 48. v47 positions (Customer Request register join, 18-Aug-2026)

Enrichment build. Register: data 5:21473, total 21475, control total $11,377,312.51 (unchanged; column T never written; green blocks unchanged). 1290 previously unattributed lines gain a site from the Parks Services Customer Request register; 14 disagreements flagged, not overwritten.
Sites: Panel A 5:526 (522 sites), groups 527:536, total 537, check 538, CR-flag count row 539, Panel B 543:936. Site_Crosswalk to 25.
Config (v47): WORKBOOK_VERSION v47; SITES_* repositioned; PARK_COUNT 522; SITE_CROSSWALK_DATA 5:25; NEW CR_JOIN_ROWS_V47 1290, CR_DISAGREE_FLAGS_V47 14, GAZ_FROM_LISTS_V47 146. Method 46.0-46.4. Data_Acquisition G8. Open Items to #152.

### THE INVOICE IS AUTHORITATIVE (standing rule, v47)

The Parks Services Customer Request register (`_AGR_Data`, 12,190 CRs, 3,024 carrying an invoice number) is joined on the invoice number, normalised by stripping every non-alphanumeric character (`INV 2428` and `INV-2428` are one key). It writes a site ONLY where the register line has none. Where a site already exists from invoice evidence and the CR Location disagrees, the register site STANDS and the CR reading is appended to column DZ as `[CR DISAGREES: ...]` and written to Site_Crosswalk. The CR register is an operational tracker: its Actual Cost disagrees with the register on roughly half the matched lines, so it is used for the SITE only, never the amount, the PK or the vendor.

### Never overwrite an allocated-multi row (build trap, v47 patch cycle)

A row whose basis is `Multiple sites named (allocated per printed line, see Site_Allocation)` already carries its dollars on Site_Allocation. Rewriting its basis to anything else double counts - the register column T then lands in Panel A alongside the row's own allocation rows. The first pass overstated the Sites total by $32,270.27 and broke both the Sites tie and the Site_Allocation tie. Allocated-multi rows are now protected explicitly: printed-line allocation is invoice evidence and outranks every later source.

### Gazetteer harvested from a Council-maintained list (v47)

The Parks Services workbook maintains a curated `Lists` Park_Name column, 275 names, of which 146 were absent from the register gazetteer. Harvesting a Council-maintained operational list resolves most outstanding candidate names without individual adjudication and widens the match surface for every later build. Panel A grows from 494 to 522 sites.

### CR register data-quality gates (v47)

Location values `Various`, `0`, blank and `Scheduled` are junk and excluded; 52 invoice keys carry more than one distinct Location and are skipped rather than allocated; one Location field holds a product description (`Abus 83/45 Titanium Padlocks...`), so free-text Location is always resolved against the gazetteer and never taken verbatim.

## 49. v48 positions (full site sweep + Marsden Stores catalogue, 18-Aug-2026)

Enrichment build. Register: data 5:21473, control total $11,377,312.51 (unchanged; column T never written; green blocks unchanged). Sites Panel A 5:533 (529 sites, $4,624,523.59 = 40.6%), groups 534:544 (eleven classes), total 545, check 546, Panel B 550:943.
Config (v48): WORKBOOK_VERSION v48; SITES_* repositioned; PARK_COUNT 529; NEW SUBURB_BLOCKLIST 70, GAZ_UNIVERSE 534, SITE_WRITES_V48 75, LISTS_COLLAPSED_V48 193, STORES_CATALOGUE_MATCHES 109, EIL_SITE_LINES 1106. Method 47.0-47.5. Data_Acquisition G9. Open Items to #154. Four rule-18 patch cycles, all on the suburb blocklist and the allocated-multi guard.

### COLLECT-ALL-CANDIDATES (standing rule, v48)

The site read no longer stops at the first non-empty field. Every register column outside the ban list (12, 25, 32, 92, 98, 105) is read, every candidate collected, and the decision taken afterwards on the priority order 111 > 106 > 112 > 108 > 127 > 99 > 42 > 22 > 23 > 80. `(not printed)` and `(blank as printed)` are EMPTY, never an answer: at v47 they stopped the scan on 80 sighted rows carrying $340,399. Where candidates agree, the highest-priority source wins; where they name several parks the row becomes multi-site.

### The suburb blocklist is a first-class object, and the exception list with it (v48)

All 70 Logan suburbs are loaded and three assertions run: no gazetteer label is a suburb, no suburb is written as a site, no joined list contains one. Banning column 92 at v43 was not enough - Shailer Park re-entered through other fields and took $50,140. TRAP: `Regents Park` is BOTH a suburb and a genuine park, so an ALLOW set is required or the blocklist deletes a real site; the exception is recorded on Site_Crosswalk.

### A gazetteer change must trigger a joined-list re-screen (v48)

193 joined multi-site lists reduced to a SINGLE park once blocked terms were removed - `Haldham Park; Shailer Park`, `Lansdown Park; Shailer Park`, `Kimberley Forest Park; Shailer Park`, `Larry Storey Park; Shailer Park`, `Shailer Park; Shailer Pioneer Park` and others. Multi-site held falls from $546,259 to $389,636. The v46 removal fixed single-site rows only; the lists carried the error for two builds.

### An allocated-multi row is untouchable (build trap, v48 patch cycle)

Two later loops (the suburb clear and the joined-list strip) changed the BASIS of rows whose dollars sit on Site_Allocation, breaking both the Sites tie (over by $103.23) and the Site_Allocation tie. Any loop that rewrites DY or DZ must skip `Multiple sites named (allocated per printed line, see Site_Allocation)` first. Their split is fixed by printed invoice lines and outranks every later source.

### Marsden Stores catalogue identifies the ITEM, never the site (NEW v48)

Internal stores issues narrate `Despatch Stock Requisition 'nnnnnn'-ALLSTORE` and carry a six or seven digit PRODUCT number in the GL narration (col 22) and Src Details (col 60). Matched against the Marsden Stores 2026 Inventory Catalogue (293 product numbers parsed from the PDF), 109 of the 326 internal stores rows resolve to a named item: paver sand 20kg (11 rows, $4,193), galvanised post 50mm x 4.4m (9, $3,776), 1-ply toilet paper (2, $3,579), 450mm ground anchor spike (10, $1,871), star pickets, first aid kits, sign brackets. Recorded in the Coding Note. This identifies WHAT was issued and NOT where: a stores issue records goods leaving the store, not where they were installed. The site would have to come from the requisition's own work order (Open Items #153).

## 50. v49 positions (user web-verified address decisions applied, 18-Aug-2026)

Enrichment build. Control total $11,377,312.51 unchanged. 578 user-adjudicated address segments applied through Site_Crosswalk: 110 resolved (VERIFIED 5, SUPPORTED 104, PROBABLE 1) writing 1028 register lines; 1 NOT A PARK (90-110 Evans Road Meadowbrook, 50 lines, new class `Non-park Council parcel (user-adjudicated: not a park)`); 467 UNRESOLVED (439 no unique public match, 28 ambiguous streets) recorded so they are never re-asked. Sites Panel A 5:555 (551 sites), groups 556:567, total 568, check 569. Config: WORKBOOK_VERSION v49; SITE_CROSSWALK_DATA extended; USER_ADDR_DECISIONS_V49 578; USER_ADDR_WRITES_V49 1028. Method 48.0. Open Items to #155.

### A class row carries NO DY label (build trap, v49 patch cycle)

Writing a descriptive label into DY on a row whose DZ is a closed-list class double counts: Panel A picks the DY label up as a site while the group row picks the DZ class up. The first pass overstated Sites by 50 lines / $25,981.73. Class rows leave DY blank; the class in DZ is the whole record.

## 51. v50 positions (user override: 90-110 Evans Road is Riverdale Park, 18-Aug-2026)

Control total $11,377,312.51 unchanged. 50 lines moved from the one-row not-a-park class to Riverdale Park; the class removed; Site_Crosswalk row corrected with the superseded status retained. Sites Panel A 5:555 (551 sites), groups 556:566, total 567, check 568. Config WORKBOOK_VERSION v50. Method 48.1.

### Local knowledge outranks a public listing (standing rule, v50)

A property-portal search describes tenure, not use. It can neither confirm nor deny that a Council-rated parcel is a park. The web-verification pass marked 90-110 Evans Road Meadowbrook NOT A PARK; the user knows it as Riverdale Park, and the parcel (S07 9714148, Open Item #123) is the park's own. Order of authority on a site question: printed invoice > user local knowledge > Council operational record (CR register, rates parcel) > Council directory > any public listing.

## 40. Current positions = Config (v51). Update this section and Config in the same build; append one short block per build; do NOT accumulate per-build history here (it goes to the HISTORY file).

## 52. v51 positions (Batch 42 Q Power historical capture; Batch 43 Origin -022..-027 reconciliation; v45 date repair; site decisions, 18-Aug-2026)

Capture build. Register: data 5:21473, total 21475, control total $11,377,312.51 (unchanged). 48 green blocks added (Q Power (Qld) Pty Ltd, QPO001, $58,935.38 ex GST across the corpus, $52,851.22 green-blocked); 3 invoices with no register line captured under rule 16 only (14406, 14865, 15222; Open Items #158); no check variants; no companions. Sighted count 1,921 -> 1969. 43 rows gain a Tier 1 printed site.
Evidence_Invoices: data 5:2097 (+51), totals 2099, controls 2105:2118 (fourteen; the two count labels now say "must be 1,969").
Evidence_Invoice_Lines: data 5:5056 (+96: 64 priced lines + 32 single-scope), reconciliation rows 5060:7152 (2093).
Sites: Panel A 5:555 (551 sites), groups 556:566, total 567, check 568, Panel B 572:965. Site_Crosswalk data 5:639.
Config (v51): WORKBOOK_VERSION v51; EI_DATA 5:2097; EI_TOTALS_ROW 2099; EI_CONTROL_ROWS 2105:2118; EIL_DATA 5:5056; EIL_END 5056; RECON_ROWS 5060:7152; RECON_COUNT 2093; SIGHTED_COUNT 1969; BOILERPLATE_KEYS 83 (no new keys; bumps BP:QP-P1 +51, BP:QP-T1 +45, BP:QP-T2 +6); VARIANT_TAG_MIN 880 and GSTINC_ALLOWANCE 40,890.38 unchanged; SITES_* repositioned; PARK_COUNT 551; NEW DATE_REPAIRS_V51 13, SITE_ALIAS_WRITES_V51 92, ELEC_STMTS_RECONCILED_V51, ELEC_GROUPS_TIED_V51 375/375, ELEC_LINES_TIED_V51 1191. Data_Acquisition F45 (Batch 42), F46 (Batch 43). Method 49.0-49.5. Open Items to #164. Session scripts b51_prepare.py / b51_elec.py / b51_build.py / b51_verify.py / b51_elec_xlsx.py; writes listed in b51_site_writes_for_verification.xlsx.

### The Q Power work-note template prints NO line table (v51, extends the v40 BLC single-scope precedent)

32 of 51 invoices (Aug-2024 to Jun-2025) print a Description / Work Note narrative and a totals block only, no priced grid on the text layer or the OCR layer. Each is captured as ONE single-scope line = the printed Description block at the printed Sub-Total ex GST, qty / unit '(not printed)', basis stated in the EIL Notes and the anomalies; checks 2 and 3 tie on the printed figures. Not a summary row: the invoice has one scope. From Jun-2025 the simPRO template prints 'Part # | Item | Quantity | Unit Price | GST | Total'; wrapped items print over THREE rows (description above, part number / item on the priced row, continuation below) - the prefix row is always attached, the suffix row only when a prefix exists, every narrative row inside the block must be consumed and every row asserted verbatim.

### F1, fourth form: invoice_date = work-note date or PLEASE PAY BY date (v51)

Q Power prints 'PLEASE PAY BY | AMOUNT | INVOICE DATE' as three values on one header row ('Invoicing 05/09/2024 $712.80 22/08/2024'). The extractor's invoice_date disagreed with the printed INVOICE DATE on 36 of 51 documents (work-note date) and on all 13 v45 Q Power captures (pay-by date). Restate from the header row (due first, invoice date third) and cross-check every matched row against the register Doc Date. The 13 v45 green blocks and their EI rows were repaired from the retained Batch 40 text; seven v28 captures (QP-15124 family) show the same 14-day pattern but their page text was not retained (Open Items #159). Queue with the v33/v34/v36/v45 findings for extraction prompt v3: one header-row rule for simPRO.

### Citation drift was live again at v50 (repaired v51)

Every EI col-13 recon citation and every register col-28 recon citation was stale at v50 (the v45 relocation shifted the block by 12 rows without rewriting either limb, and the pre-v45 texts cited only 'Evidence_Invoice_Lines reconciliation block'). v51 rewrites both limbs from the maps for ALL rows (2,042 EI, 1,921 register) and the verify now audits both limbs on every row, every build - the b45-b50 verifies did not.

### Origin collective statement summaries reconcile the register to the sub-account (Batch 43, v51)

The 'Your collective account summary' PDF prints sub account | invoice number | full supply address | GST | charges incl GST for every Council sub-account. Register (statement, NMI) groups match by supply-address prefix (the register label is the Origin address cut at 40 characters: drop a truncated tail token before comparison) and tie register ex GST to (incl - GST) within 2c: 375/375 on -022..-027, 1,191 lines, $42,265.95. Sub-account and invoice number are now on Recon; the full address adds no park name the register label lacked.

### Site decisions and the alias map (v51)

Aliases are applied three ways: tokens in DY canonicalised (lists re-screened, MA lists cosmetically), Site_Allocation G and EIL L canonicalised, and rows whose deciding source column names an alias re-decided (a manual, CR, rates-parcel or user-adjudicated basis is never overwritten). Guards: 'Logan Reserve' is a SUBURB - the alias fires only where the field itself reads Logan Reserve ('X Park, Logan Reserve' never attributes); 'Sport Complex' only with 'Oliver' in the field; depot labels never enter a multi list; a substring inside a longer name never matches (v50 had 'Darlington Park' -> Arlington Park and 'Flindersia Riverside Park' -> Riverside Park). 'Mount Warren Park' is a suburb name and cannot be a gazetteer label without an ALLOW exception - 'Mt Warren Oval Park' added as printed pending the user. 'Logan River Park' (Origin label, Lot 2 Peacock Ave) merged into Logan River Parklands on the user's Council-directory decision. 'Flindersia Riverside Park' added as printed; possibly the rates label 'Logan Reserve Riverside Park' (Open Items #162).
