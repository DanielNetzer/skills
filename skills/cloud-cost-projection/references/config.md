# Config reference (JSON)

Every top-level key used by `examples/example_config.json` and `scripts/build_model.py`. Optional keys noted. Do not invent fields the script does not read.

| Field | Type | Meaning / default |
|---|---|---|
| `company` | str | Workbook title label (default `"Company"`) |
| `report_currency` | str | Report currency code (default `"EUR"`) |
| `fx_usd_per_report_ccy` | num | USD per one unit of report currency (planning FX) |
| `months` | `{start: "YYYY-MM", count, plan_year}` | Horizon; plan-year months drive totals, cap and envelope |
| `revenue` | `{current_year, plan_year, note}` | Report currency, phased flat |
| `kpi_cap_share` | num | e.g. `0.08` |
| `payment_lag_months` | int | Usage month → cash month (`2` for Net 30 after an M+1 invoice) |
| `vat_rate` | num | Shown on the vendor cash sheet only |
| `baseline_label` | str | Provenance text for the baseline invoice matrix |
| `excluded_memo` | str[] | IT / other items on the invoice excluded from totals |
| `reconciliation` | `{project_usage, account_block, marketplace, excluded_items, invoice_subtotal}` | Invoice sub-total tie-out |
| `products` | `{name, elasticity}[]` | Product lines, in the order of `projects[].values` |
| `projects` | `{id, group, env, in_scope, values[]}[]` | Baseline matrix; `env` in `prod`/`qa`/`dev`/`other`; `in_scope` `0` excludes but keeps the project for allocation and whole-account cash |
| `group_order` | str[] | Breakdown group order (optional; else sorted unique groups) |
| `items` | object[] | Account-level vendors allocated by usage share. Fields: `name`, `aug_invoice`, `run_rate`, `elasticity?`, `flag`, `cash`, `note?`, `usage_billed?`, `support_ratio?`, `floor?`. `support_ratio: true` → ratio-driven (`run_rate` null). `floor: {min_monthly, uplift_share, note}` → floor vendor. `usage_billed: true` → own whole-account cash row |
| `aws_accounts` | `{name, baseline}[]` | Linked accounts from the AWS invoice |
| `aws_elasticity` | num | AWS usage elasticity (default `0.3`) |
| `terminating_commitments` | `{name, monthly, last_month_index, category, note}[]` | e.g. unused savings plan; first entry bound to Inputs B15 (optional) |
| `initiatives` | object[] | `{name, target, scope, saving{P,R,O}, added{P,R,O}?, start{P,R,O}, ramp, owner, note}`. `target` = product line or item name; `scope` `all` or env list like `qa+dev` |
| `envelope` | object | Optional innovation envelope: `{amount_report_ccy, reseller_billed_share, inside_kpi, utilization{P,R,O}, weights[]|null, governance[]}`; weights default flat |
| `scenarios` | `{growth_yoy{P,R,O}, credit_month_index{P,R,O}, credit_amount, credit_note}` | Credit month index `1` = first month of the horizon |
| `switches` | `{floor_year_fee_single}` | Bound to Inputs B13; cash events with `switch_equals` follow it (optional) |
| `cash_events` | `{name, amount, months[], switch_equals?, fallback_in_lag_months?}[]` | Whole-account postings by month label; fallback = that vendor’s posting on the last actual invoice (first lag months) |
| `cash_placeholders` | str[] | Vendors with unknown amounts, shown at zero (optional) |
| `cash_fallback` | `{gcp_usage, aws_usage}` | Last actual invoice usage for the first lag months (optional) |
| `commitments_register` | `{vendor, term, total, cadence, mom_spread, next_event, decision_point, note}[]` | Register rows (optional) |
| `cash_notes` | str[] | Free text on the cash sheet (optional) |
| `summary_notes` | str[] | Free text on the scenario summary (optional) |
| `sources_and_gaps` | str[] | Free text on the sources sheet (optional) |
| `recalc_script` | path | Optional. Defaults to the xlsx skill’s `scripts/recalc.py`. Script-only; omitted from the example. |

Input cell map (fixed): B4 FX, B5/B6 revenue, B7 cap, B8 support ratio, B9 floor minimum, B10 uplift, B11 lag, B13 floor-fee switch, B15 terminating commitment last month, B17 cash basis, B18 envelope amount, B19 envelope reseller share, B20 envelope in KPI; scenario table from row 23.
