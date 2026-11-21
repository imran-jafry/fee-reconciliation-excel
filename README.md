# Exchange Fee Calculation & Reconciliation Model (Excel)

An Excel model that recalculates exchange fees on futures trades, compares them to invoice lines, and flags billing exceptions. I built it to practise the fee calculation and reconciliation cycle used in fees & commissions operations.

> **Data disclaimer:** All data in this project is **simulated**. Trades, clients and invoice lines are randomly generated. Exchange names are used only as labels, and the fee rates and client discounts are illustrative, not actual exchange or broker fees. About 19 billing errors were **deliberately planted** in the invoice data so the model had something to detect. This is a demonstration of method, not a real-world finding.

## What it does

1. Looks up the fee per contract for each trade from a rate card (by exchange and product).
2. Applies a client discount tier.
3. Calculates the expected fee: `ROUND(Quantity x Rate x (1 - Discount), 2)`.
4. Matches each trade to its invoice lines and compares calculated vs billed amounts.
5. Assigns a status to every trade and lists all exceptions with a suggested action.
6. Summarises results by exchange, status and client, with a tie-out check.

## Workbook structure

| Sheet | Purpose |
|---|---|
| **README** | Overview, data disclaimer, limitations |
| **Summary** | KPIs, variance by exchange, exceptions by status and client, tie-out check, charts |
| **Assumptions** | Editable tolerance settings (amount and rate) |
| **Rate_Card** | Fee per contract by exchange and product (input) |
| **Clients** | Client tiers and discounts (input) |
| **Trades** | 300 simulated futures trades with rate lookup and calculated fee |
| **Billed** | 301 simulated invoice lines, with a check for lines that have no matching trade |
| **Recon** | One row per trade: calculated vs billed, variance, status |
| **Exceptions** | Live list of non-OK trades with a suggested action |

## Reconciliation logic

Each trade gets one status, tested in this order:

1. **Missing billing:** no invoice line found
2. **Duplicate billing:** more than one invoice line
3. **Quantity mismatch:** billed quantity differs from trade quantity
4. **OK:** variance within the tolerance (default $0.05)
5. **Rate mismatch:** billed effective rate differs from the calculated rate
6. **Amount mismatch:** any other difference above tolerance

Checks run in this order so a trade is labelled by its most basic problem first. A missing or duplicate line makes quantity and rate comparisons meaningless.

## Results on the simulated data

| Measure | Result |
|---|---|
| Trades processed | 300 |
| Calculated fees | $78,351.02 |
| Billed on matched trades | $77,635.62 |
| Trades with exceptions | 16 (5.3%) |
| Missing billing | 4 |
| Duplicate billing | 3 |
| Quantity mismatch | 3 |
| Rate mismatch | 6 |
| Orphan invoice lines (no trade) | 2 |
| Tie-out difference | $0.00 |

Three lines differing by $0.01 were correctly treated as rounding and not flagged.

## Excel skills used

- `INDEX` / `MATCH` lookups on a combined Exchange|Product key
- `SUMIF`, `COUNTIF`, `COUNTIFS`, `ROUND`, nested `IF`
- Data validation (Buy/Sell list, whole-number quantity, tolerance inputs)
- Conditional formatting and filters
- Summary tables and charts
- Input vs formula colour coding (blue = input, black = formula, yellow = editable assumption)

## Limitations

- Flat per-contract fees only. No tiered or volume pricing, fee caps, rebates or multiple currencies.
- No effective-dated rates or amendment history.
- Simple tolerance logic. Real processes involve more systems, approvals and an audit trail.

## Possible extensions

- Effective-dated rate card
- Tiered fees by monthly volume
- Currency conversion
- Power Query for loading real invoice files
- A Python version of the same checks

## Author

Mohd Imran, Hyderabad, India
