# Cloud cost projection

Finance-grade, formula-driven cloud cost projection (GCP/AWS via a reseller). Three scenarios, per-project and per-account breakdowns, commitment flags, vendor cash on real payment timing, KPI cap comparison, optional innovation envelope.

- Canonical page: https://www.danielnetzer.com/skills/cloud-cost-projection
- GitHub (skill repo): https://github.com/DanielNetzer/skills/tree/main/skills/cloud-cost-projection
- Zip: https://www.danielnetzer.com/skills/files/cloud-cost-projection.zip

## Install

**npx (preferred):**

```bash
npx skills add DanielNetzer/skills --skill cloud-cost-projection
# optional Cursor target:
npx skills add DanielNetzer/skills --skill cloud-cost-projection --agent cursor
```

**Zip:** download from the canonical page, extract so `SKILL.md` is at the skill root, drop into your agent skills folder (Cursor: `.agents/skills/cloud-cost-projection` or `~/.cursor/skills/cloud-cost-projection`).

**Zip via npx:**

```bash
npx skills add https://www.danielnetzer.com/skills/files/cloud-cost-projection.zip
```

## Run

1. Follow `SKILL.md` — scope, gather sources, reconcile, fill config.
2. Schema: `references/config.md`. Scrubbed shape: `examples/example_config.json` (not your numbers).
3. Build:

```bash
python3 scripts/build_model.py <config.json> <output_dir>
```

Deliver only if the four tie-outs in `SKILL.md` §5 hold. Framing checklist: `references/checklist.md`.

## What’s in the box

| Path | Role |
|------|------|
| `SKILL.md` | Agent operating contract |
| `scripts/build_model.py` | Config → xlsx + CSV + tie-outs |
| `references/sources.md` | Where numbers live |
| `references/method.md` | Formulas and conventions |
| `references/config.md` | Every config field |
| `references/checklist.md` | Pre/post checks + CFO framing |
| `examples/example_config.json` | Scrubbed full example |

## Paste prompt

```
Read the cloud-cost-projection skill (SKILL.md + references/). Help me build a Finance-grade cloud cost projection for [company].

Ask the three scope questions first (projects in/out, account-level allocation, cash basis). Confirm the KPI (e.g. cloud as % of revenue) and that the revenue basis matches the cost scope.

Gather sources in order: reseller invoices → cloud cost module → contracts → console CUDs/credits → Finance plan. Reconcile the baseline invoice to the cent before modeling. Fill config.json from references/config.md; use examples/example_config.json only as a scrubbed shape, never as my numbers.

Then run: python3 scripts/build_model.py <config.json> <output_dir>
Do not deliver unless the four tie-outs in the skill pass. Lead the note with plan-year net, % of revenue per scenario, gap to the cap, peak cash month, the three assumptions that move the answer most, and a gaps register with owners.
```

## License / attribution

MIT © Daniel Netzer. Shared from danielnetzer.com skills + https://github.com/DanielNetzer/skills/tree/main/skills/cloud-cost-projection. Do not treat the example config as a real company bill.
