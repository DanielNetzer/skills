---
name: cloud-cost-projection
description: Build a Finance-grade, formula-driven cloud cost projection (GCP/AWS via a reseller) with three scenarios, per-project and per-account breakdowns, commitment flags, vendor-level cash flow with payment timing, a KPI cap comparison and a discretionary innovation envelope, from invoices, the cloud cost module, contracts and the cloud console. Use this whenever someone asks to project, forecast, budget, model or reconcile cloud spend, compare cloud cost to a revenue KPI or cap, size FinOps savings initiatives, model cloud vendor commitments (Datadog, Confluent, Databricks, CUDs, marketplace prepayments, credits milestones), or produce a cloud cash-flow view for a CFO. Trigger even if they only say "cloud budget", "cloud cost model", "FinOps plan", "how much will GCP cost next year" or "when do we pay the cloud vendors".
---

# Cloud cost projection

Produce one workbook that Finance and the CTO can both sign: amortized projection by product line and by project, three scenarios, commitments flagged, cash by vendor and by project on real payment timing, tie-outs that prove the numbers. `scripts/build_model.py` generates it from a JSON config; the skill's job is filling that config with reconciled facts and defensible assumptions.

## 1. Scope in three questions

Ask before pulling data (defaults in brackets):
1. Which projects / entities are in scope? [all; exclude spin-offs, subsidiaries, other departments by project-name suffix]
2. Account-level vendors (observability, support, marketplace contracts): allocate to projects by usage share, or keep as separate rows? [allocate by baseline usage share; excluded projects take their share out]
3. Cash basis: whole account (what the company pays) or in-scope share? [whole account, with a switch]
Confirm the KPI (e.g. cloud at 8% of revenue) and whether the revenue basis matches the cost scope (same entities in and out). This mismatch flatters or punishes the ratio more than any assumption.

## 2. Gather sources (read `references/sources.md` for exact queries)

Priority order; stop when the matrix reconciles:
1. **Reseller tax invoices** (PDFs). Read the latest full month line by line: it is the only source with project × product, support, marketplace postings, credits and VAT. Earlier months: page 1 gives the account-level block (support, marketplace postings).
2. **Cloud cost module** (Datadog Cloud Cost Management or native billing export): baseline month by product, by project, credit lines by type, monthly history. Cross-check to the invoice within ~1–4%.
3. **Contracts** from email: order forms, renewal threads, marketplace private offers. Extract minimum, term, cadence, uplift (premium support %), renewal date.
4. **Cloud console**: committed-use discounts (rates, terms, end dates), incentives / credit milestones (earned, remaining, thresholds), month-to-date forecast.
5. **Finance**: plan revenue, expense targets, prior budget by group, plan-vs-actual trackers.

Never estimate what an invoice can tell you. If a PDF sits in email and the connector cannot read attachments, save it to Drive through the browser and read it there.

## 3. Reconcile before modeling

Transcribe the baseline invoice into a project × product matrix (9 product lines is enough: compute incl. Kubernetes nodes, networking, cache, AI/ML, SQL, storage, device/emulator fleets, logging+monitoring, other). Check every project total against the invoice's "Total for source", then: project usage + account-level block + marketplace projects + IT memo items = invoice sub-total before VAT, to the cent. Only then fill `reconciliation` in the config. Excluded-but-paid items on the same invoice (Workspace seats) are memo lines, never totals.

## 4. Fill the config (schema in `references/config.md`, worked example in `examples/example_config.json`)

Rules that keep the model honest:
- **Run-rates** come from the baseline invoice/cost module; **elasticities** (0 flat … 1 tracks volume) per product line; growth per scenario compounds monthly.
- **Flags**: Variable (usage), Committed (contract, cannot shrink within term), Floor (contract minimum with usage above it), Discretionary (envelope). Committed and floor lines get their own MoM-spread subtotal.
- **Account-level vendors** are allocated by baseline usage share; a contract's saving inside its term is amortized only, never cash.
- **Initiatives** carry P/R/O saving share, added cost share, start month and ramp, plus a project scope (`all`, `qa+dev`) so savings land where they happen. Gate start months on CUD end dates (paying a commitment fee without usage is a cost, not a saving). A vendor the company has ruled out for savings gets no initiative.
- **Credits**: milestone credits post when cumulative spend crosses a threshold; model them as single-month tranches per scenario (spending less delays them), never as an even monthly spread.
- **Cash timing**: reseller-invoiced items for month M are invoiced early M+1 at Net 30, so cash is M + lag (default 2). Marketplace prepayments follow the same lag from their posting month. Direct-billed SaaS: next month. First lag months use the last actual invoices as the proxy, including that month's vendor postings.
- **Scenarios**: P = growth stays hot, initiatives deliver ~half and start late, no or late credits; R = plan; O = growth decouples, initiatives over-deliver, credits early. A budget envelope is a cap: vary utilization (110 / 100 / 85%), not the amount.
- **Terminating commitments** (e.g. an unused AWS savings plan billed by the reseller) get a last-month input, not a flat line.

## 5. Build and verify

```bash
python3 scripts/build_model.py <config.json> <output_dir>
```
The script writes the xlsx and a long-format CSV, recalculates with LibreOffice (via the xlsx skill's `recalc.py` when present) and prints the tie-outs. Do not deliver unless all four hold: zero formula errors; invoice reconciliation difference ≈ 0; every Shares row = 100%; Breakdown grand total = Projection net and Cash by project total = vendor cash total every month (within cents). Then open the Scenario summary and sanity-check the baseline month against the actual bill.

## 6. Deliver

Present the xlsx and CSV, then a short note that leads with the plan-year net in the report currency and the ratio to revenue per scenario, the gap to the cap, and the peak cash month with its cause. Always state the three things that move the answer most (usually: which shared contracts the excluded entities actually pay, credit milestone timing, and one unsigned contract), and list the open gaps with owners. Use `references/checklist.md` for the delivery checklist and the framing that has worked with CFOs.

## Files

- `scripts/build_model.py` — config → workbook + CSV + tie-outs (formula-driven, 11k+ formulas).
- `references/sources.md` — where each number lives and the exact pulls (cost-module queries, console pages, email search patterns, reseller invoice anatomy).
- `references/method.md` — formulas, allocation, timing, credits, CUD gates, scenario logic.
- `references/config.md` — every config field, with defaults.
- `references/checklist.md` — reconciliation, tie-outs, gaps register, CFO framing.
- `examples/example_config.json` — scrubbed full example (33 projects, 7 allocated vendors, 8 initiatives, envelope, cash events).
