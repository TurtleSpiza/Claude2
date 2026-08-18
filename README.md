# PS/WP Transaction Register: session handover, v51 (18-Aug-2026)

## Verdict
v51 shipped on a clean exact-TRUE verify (one rule-18 patch cycle: citation rewrite, alias guards, Logan River Park label). Control total $11,377,312.51 unchanged since v29. Sighted invoices 1,921 -> 1,969.

## This kit (merged v50 HANDOVER kit + v51 session)
- `PS_WP_Transaction_Register_3FY_v51_HANDOVER.xlsx` (v50 alongside; v43-v49 in `prior_session/`).
- `PSWP_Register_Schema.md` - section 52 (v51). Replace the project-knowledge copy.
- `LCC_Parks_Expenditure_Map_v51.html` - rebuilt on v51: 366 parks plotted, $3,847,140; named-park share 42.2%. v1-v50 maps in `prior_session/`.
- `PSWP_Invoices_Left_To_Capture_v51.xlsx` - By creditor (66) and Invoices left to capture (8,015 references, 8,750 lines, $2,664,144); definition on the Method sheet. v45-v48 in `prior_session/`.
- `PSWP_Batch43_Electricity_Reconciliation_022_027.xlsx` - Origin -022..-027, 375/375 groups tie, $42,265.95.
- `PSWP_Site_Worklist_v51_unallocated_classes.xlsx` - the six unallocated classes (1,933 lines + 104 residue rows, $1,559,597) with a decision column per row and per summary sheet.
- `toolkit/` - b51_prepare / b51_elec / b51_build / b51_verify / b51_elec_xlsx and their JSONs, corpus_qpower_historical_20260818.json + capture report; all prior b37-b50 scripts, pswp_build_dashboard.py.
- `verification/` - b51_site_writes_for_verification.xlsx (92 DY/DZ writes) plus every prior verification file.
- `creditor_histories/`, `logan_geo/` unchanged from v50.

## v51 in one paragraph
Batch 42: Q Power historical corpus, 51 invoices, $58,935.38 - 48 green blocks ($52,851.22), 3 held rule-16-only (14406, 14865, 15222; #158); 32 invoices print no line table and are captured as single-scope lines (Method 49.1, #160); F1 fourth form (invoice_date = work-note / pay-by date) restated on 36 and REPAIRED on the 13 v45 Q Power blocks, 7 v28 blocks flagged (#159). Citation drift live at v50 on every EI col-13 and register col-28 recon citation: both limbs rewritten from the maps (2,042 + 1,921) and audited every build from now on. Batch 43: Origin -022..-027, 375/375 tie. Sites: 30 user candidate names + 17 Peacock Avenue = Logan River Parklands applied (92 writes; Darlington Park had been Arlington Park, Logan Gardens Park had mapped to a non-existent label). Open Items to #164.

## Position at v51
| | |
|---|---|
| Control total | $11,377,312.51 |
| Sighted invoices | 1,969 (45.7% of dollars) |
| Confirmed | 79.3% of dollars |
| Named park / site (incl. printed-line allocation) | $4,799,926 (42.2%), 551 sites |
| Multi-site held | $407,515 |
| Plotted on map | $3,847,140, 366 parks |
| AP invoices left to capture | 8,015 references, $2,664,144, 66 creditors (Levai $918,792, Harpley $365,330) |

## Waiting on you
1. `PSWP_Site_Worklist_v51_unallocated_classes.xlsx` decisions (five entry forms: park name / MAP: / NOT A PARK: / MULTI / LEAVE).
2. Flindersia Riverside Park vs Logan Reserve Riverside Park (#162); Mt Warren Oval Park stands - it IS the Council directory name (30 Yvonne Crescent) so no ALLOW exception is needed.
3. Origin -020 and -033 summaries (#163); re-sight QP-15124 family (#159).
4. Everything still open from the v50 list below.

---
# (v50 handover text retained below)

# PS/WP Transaction Register: session handover, v45 (18-Aug-2026)

## Verdict
v45 shipped on a clean exact-TRUE verify (24 checks). Control total $11,377,312.51 unchanged across v43, v44 and v45. Three builds this session, six rule-18 patch cycles, all caught by the verify, none reaching the shipped file.

## Session arc
- **v43** Themes restated four-year (three controls had read FALSE since v29); column DZ made exhaustive; source hierarchy widened to all 130 register columns with six columns banned.
- **v44** Site read at PRINTED-LINE level: Evidence_Invoice_Lines L/M, new Site_Allocation sheet (628 rows, $1,187,582) releasing the multi-site invoices on printed line amounts; electricity matched to Council parcel address records (420 exact, 582 flagged, 2,073 unresolved).
- **v45** Batch 40 rule 17 capture: 44 invoices, 121 printed lines, $17,677.49. Sighted 1,877 -> 1,921.

## Batch 40 (this build)
- Corpus `qpower_brizsouth_20260818`, sha256 `e3a4a0fb...9416`, Microsoft 365 Copilot / GPT-5, 2 PDFs, 64 pages, 46 records (44 primary + 2 duplicate copies).
- Gates: 44/44 self-tie; 121/121 priced records from TEXT lines; 64/64 pages; rule 12 prefix-aware screen over 28 EvID prefixes = 0 re-sightings; corpus-internal hash screen = 0 collisions; every invoice amount-corroborated to exactly one non-zero register row.
- **Q Power (Qld) Pty Ltd** QPO001, ABN 82 067 507 591, 13 invoices. **BrizSouth Locksmiths Pty Ltd** BRI115, ABN 78 126 351 184, 31 invoices.
- 28 $0.00 companions cross-referenced. 41 of 44 gain a Tier 1 printed site.
- New Open Items #140-#143: seven invoices print no PK; 15134 prints PK000442 and posts PK000412; three Q Power rows carry no creditor code; six referenced quotes/photos not supplied.

## Position at v45
| | |
|---|---|
| Control total | $11,377,312.51 |
| Sighted invoices | 1,921 |
| Named park (incl. printed-line allocation) | $3,856,260 (33.9%), 490 parks |
| Multi-site held | $554,857 |
| Plotted on map | $2,989,542, 328 parks |
| **AP invoices left to capture** | **6,781 references, 12,503 lines, $3,040,982, 61 creditors** |

## Files
- `PS_WP_Transaction_Register_3FY_v45_HANDOVER.xlsx` (v43, v44 also included).
- `PSWP_Register_Schema.md` - sections 44 (v43), 45 (v44), 46 (v45). Replace the project-knowledge copy.
- `PSWP_Invoices_Left_To_Capture_v45.xlsx` - **By creditor** (61 rows, creditor code + name + invoices + lines + dollars + date span) and **Invoices left to capture** (6,781 rows: creditor code, creditor, reference, lines, dollars, dates, FY, PK, status, tier, account, first narration).
- `toolkit/` - b43/b44/b45 build + verify scripts, b45_prepare.py (gates), elec_xwalk.py, payload builders, dashboard shell.
- `verification/` - b43 site writes (194), b44 electricity writes (1,002), fullscan worklist (45 gazetteer candidates).
- `LCC_Parks_Expenditure_Map_v45.html`, `logan_geo/`.

## Waiting on you
1. `b44_electricity_writes.xlsx` - 582 AMBER rows, Y/N.
2. `b43_site_writes_for_verification.xlsx` - 194 rows, Y/N.
3. 45 gazetteer candidates.
4. G-NAF: QLD_ADDRESS_DETAIL, QLD_ADDRESS_DEFAULT_GEOCODE, QLD_STREET_LOCALITY, QLD_LOCALITY, QLD_STREET_LOCALITY_ALIAS + Authority Code.
5. Waste branch: Cleanaway bin register by location ($3.07M, the largest single site lever).
6. Contract site schedules for the standing-order rounds ($554,857 of holds).
7. Authoritative area column; Strategy Layer re-exported with geometry; FY2026/27 extract completeness.

## Capture priority (from the listing)
1. **T & H Levai LEV002** - 2,139 references, $980,438, 32% of the remaining capture. Prints `CR - Park (Ref: n)` on every invoice, so capture yields a site immediately.
2. **Harpley HAR073** - 1,397 references, $365,330.
3. Robinson ROB004 $255,022, Weis WEI013 $244,321, Origin (no code) $203,323, FY2023/24 unidentified 7-digit series $130,716, STAR CARPENTRY STA051 $109,465.
Levai and Harpley together are 44% of what is left.

## Standing traps added this session
- Themes controls sat FALSE for thirteen builds because no verify covered them.
- openpyxl `ws.cell(r,c,None)` does not clear; set `.value=None`.
- Supplier premises, Council's own address and analysis-authored labels are never site evidence.
- Never hold on a register-level field what the printed lines beneath it already allocate.
- Recon relocation must rewrite the EIL range AND the row self-reference, preserving each row's ratified E-form.
- EI control formulas cite each other; a tail shift must remap every control self-reference.
- A Nature Category not in Theme_Map A5:A31 silently breaks the Themes controls.
