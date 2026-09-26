# Memo: What DANA should fix first, based on 3,157 App Store reviews

**To:** Product and Customer Experience leads (illustrative)
**Data:**
- App Store reviews of DANA, GoPay, OVO, ShopeePay and LinkAja, collected 26 Sep 2026.
- Bank Indonesia e-money statistics for 2021–2025.

## Bottom line

DANA is in the "good" half of the market:

- It averages 3.3 stars, against 1.3–1.6 for OVO and LinkAja.
- It collects the most reviews per day (20).
- It already beats the peer average on the three biggest industry complaints: failed transfers, customer service and login/OTP.

DANA's own gap is narrower. **Paylater (DANA Cicil)** and **money leaving the balance without clear consent** (auto-debit and bills) come up more often for DANA than for its peers.

## Evidence

| Theme | DANA | Peer average | Gap | Share of those reviews at 1 star (market) |
|---|---|---|---|---|
| Paylater & bills | 12.8% | 5.6% | +7.3 pp | 58% |
| Balance lost / auto-debit | 8.6% | 6.6% | +2.0 pp | 90% |
| Failed transfer / top-up | 9.2% | 19.0% | −9.8 pp | 86% |
| Customer service | 7.6% | 15.4% | −7.9 pp | 85% |

The share for each app is the share of its latest 500 reviews that mention the theme.

Reading of the DANA Paylater reviews: the most common story is "I have a DANA Cicil limit but cannot use it", often with no explanation. Some of these reviewers still give 5 stars. This means the users want the product, so the product has potential; the issue is friction, not rejection.

## Recommendations

1. **DANA Cicil transparency.** Show in the app why a limit is locked and when it will unlock. Measure it by the share of reviews that mention Cicil, and by the number of support tickets.
2. **Consent before each debit.** Ask for a confirmation or send a pre-notification before each auto-debit or bill payment, and send an instant receipt. This targets the most severe theme: 90% of these reviews are 1 star.
3. **Release guardrail.** Version 2.145.0 averages 3.05 stars, against about 3.4 for earlier September builds, with more customer-service and fee complaints. This is not yet significant (p = 0.10). Track 1-star share per release as an early-warning KPI.

## Caveats

- The data covers iOS only.
- Reviewers skew towards strong opinions.
- Themes come from a keyword classifier. On unseen reviews it has precision 0.80 and recall 0.64, so it undercounts some complaints. It does this equally across apps, so the comparisons still hold.
- Snapshot windows differ by app, so the findings are cross-sectional comparisons, not trends.
