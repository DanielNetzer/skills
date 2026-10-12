# Method: formulas and conventions the builder implements

## Time and scenarios
Months from `months.start` for `months.count`; plan year identified by `months.plan_year`. Monthly growth g = (1+YoY)^(1/12) − 1 per scenario. Baseline line value at month t: run_rate × (1 + e × ((1+g)^t − 1)), e = elasticity (0 flat, 1 tracks volume).

## Lines
- Product lines (GCP native): run-rate = Σ in-scope projects for that product (from the Projects matrix).
- Allocated account-level items: run-rate = item run-rate × in-scope usage share (Σ in-scope usage / Σ all usage). Support = ratio × GCP product spend (ratio = baseline support / baseline project usage).
- Floor vendor: MAX(minimum × (1 + uplift) × in-scope share, usage line). Savings on it are amortized only; cash follows the fee schedule.
- Terminating commitments: run-rate while t ≤ last month index, else 0.
- Envelope: amount × FX × utilization(scenario) × monthly weight, plan-year months only; flagged Discretionary; shown in both "with" and "without" KPI ratios.

## Initiatives
saving_t = target_line_t × (saving share − added cost share) × ramp_t, ramp_t = min(1, max(0, (t − start + 1)/ramp months)). Scope restricts which projects receive the saving in the breakdown (`qa+dev` = env in {qa, dev}); the amount is still sized on the whole in-scope line. Gate start months on CUD terms.

## Credits
Milestone programs pay a fixed tranche when cumulative eligible spend crosses a threshold. Estimate the crossing month from remaining spend ÷ monthly eligible spend per scenario (higher spend → earlier). In-scope credit = tranche × in-scope share; cash-basis credit reduces the invoice it posts on, shifted by the lag.

## Allocation and breakdown
Shares sheet: product share = project's product spend / in-scope product total; item and credit shares = project usage / in-scope usage; initiative share = project's target-product spend within scope / Σ within scope. Breakdown row = SUMPRODUCT(components column, shares column) where components = [products, items, −savings, −credit]; the AWS line is split by account baseline weights. Grand total must equal the Projection net.

## Cash
- Reseller-invoiced usage and marketplace postings: cash at t = value at t − lag (lag default 2 = invoice early M+1, Net 30). Direct-billed SaaS: t − 1.
- Fallback for the first lag months: the last actual invoice (usage, support, and that month's vendor postings).
- Whole-account basis adds excluded projects back at baseline spend × growth with no savings (the company pays for them); switch to in-scope share with one input.
- Vendor rows: one per payee with its timing; commitments carry orange flags. Cash by project: project usage net of its initiatives × (1 + support ratio) + project usage share × (vendor events + usage-billed marketplace + credit). Cash by project total must equal the vendor total.

## Scenario defaults
| | Pessimistic | Realistic | Optimistic |
|---|---|---|---|
| Growth | prior-year pace | plan | plan minus efficiency |
| Initiative saving share | ~60% of R | as proposed | ~120% of R |
| Start month | R + 2 | proposed | R or earlier |
| Credit tranche month | early (spend hot) | mid | late (spend cut) |
| Envelope utilization | 110% | 100% | 85% |

## Conventions
USD for all cost lines; report currency only for the KPI ratio, cap and headroom. No em dashes in sheet text. Blue cells are inputs. Every note names its source and date.
