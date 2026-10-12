# Sources: where each number lives and how to pull it

## Reseller tax invoices (the ground truth)
Cloud resellers issue one invoice per cloud per month, dated ~3rd of M+1, Net 30, before VAT (add local VAT for the local-currency payment).
Anatomy of the GCP invoice: an account-level block ("charges not specific to a project": observability marketplace installments, support, CUD offsets, promotional credits), marketplace pseudo-projects (data-platform vendors billed through the marketplace, shown under opaque project ids), then one section per project with product totals ("Total for source name"), then IT items on the same account (Workspace seats). The AWS invoice lists linked accounts; look for "Unused Commitments" lines: savings plans the company pays for without using.
Pull: search email `from:<reseller billing address> subject:"Digital Invoice" newer_than:400d`; PDFs cannot be read through the mail connector, so open the message in the browser and use "Add all to Drive", then read from Drive (`read_file_content` extracts PDF text). Finance often keeps a folder of them already: search Drive `title contains '<invoice number prefix>' and mimeType = 'application/pdf'`.
Read the baseline month completely and transcribe every project × product cell; for earlier months page 1 is enough (support and marketplace postings).

## Cloud cost module
Datadog Cloud Cost Management (`get_datadog_metric` with `use_cloud_cost=true`):
- by product: `gcp.cost.amortized`, group_by `google_product`, scalar sum over the calendar month;
- by project: group_by `project_id`; by product within projects: filter `project_id IN (a,b,c)`;
- credits: filter `cost_type:credit`, group_by `credits_type` (promotion, fee_utilization_offset, subscription_benefit, committed_usage_discount…), 30-day rollups, to see how promotional tranches actually posted;
- history: `gcp.cost.amortized`, `aws.cost.amortized`, `datadog.cost.amortized` with 30-day rollups (buckets start on the from-date, not calendar months).
Marketplace one-day postings (streaming, database, search and observability vendors billed through the marketplace) appear inside `gcp.cost.amortized`; the observability vendor's own installment on the GCP bill double-counts against `datadog.cost.amortized`.
Native alternatives: GCP billing export to BigQuery; AWS Cost Explorer / CUR.

## Contracts (email)
Search patterns: `<vendor> (order form OR contract OR renewal OR invoice OR commitment) newer_than:400d`. Read the negotiation thread, not just the attachment: the buyer's reconciliation email usually states minimum, term, uplift and installment cadence in plain text. Marketplace private offers (`Q-…`) are accepted when a matching posting appears on the reseller invoice. Order-form PDFs need the browser-to-Drive route.

## Cloud console (GCP)
`console.cloud.google.com/billing/<account>/…`:
- Committed use discounts: each CUD's hourly rate, term and start date (end date = start + term); the "expiring soon" insight;
- Incentives: spend-based milestone programs with spend target, progress %, reward and end date (this is the credits schedule);
- Overview: month-to-date cost, savings and forecast.
AWS: Savings Plans inventory for commitment end dates (needs a signed-in browser).

## Finance
CFO plan deck (revenue, expense targets, the cloud line as % of revenue), prior-year budget by group (allocation keys for shared contracts), plan-vs-actual trackers, FinOps weekly notes. Board decks for credit programs and migration ROI (total credits, completion dates).

## What each source is good for
| Need | Source |
|---|---|
| Project × product baseline | Invoice (exact) → cost module (cross-check) |
| Monthly history | Cost module 30-day buckets; invoice page 1 for postings |
| Contract floor, cadence, uplift | Email threads |
| CUD end dates (initiative gates) | Console CUD page |
| Credit milestones | Console Incentives; cost-module credit lines for what posted |
| Excluded entities' share | Invoice project totals (usage share) |
| Payment timing | Invoice date and terms (Net 30) |
