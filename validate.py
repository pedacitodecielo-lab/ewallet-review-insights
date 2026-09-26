"""Compare rule-based themes with blind hand labels on 100 unseen negative reviews."""
from pathlib import Path

import pandas as pd

from classify import tag

HERE = Path(__file__).resolve().parent
CODES = {"B": "Balance lost / unauthorised deduction", "F": "Failed / pending transfer or top-up",
         "L": "Login, OTP & verification", "S": "Account security & fraud", "C": "Customer service",
         "P": "Fees & charges", "U": "Hard to use", "A": "App performance & bugs",
         "D": "Paylater, loans & debt collection", "X": "Ads, notifications & promos"}

sample = pd.read_csv(HERE / "validation_sample.csv")
gold = {}
for line in (HERE / "validation_labels.txt").read_text(encoding="utf-8").splitlines():
    if line.startswith("#") or "|" not in line:
        continue
    i, codes = line.split("|")
    gold[int(i)] = {CODES[c] for c in codes.split(",") if c}
assert len(gold) == len(sample) == 100

pred = {i: set(tag(t)) for i, t in enumerate(sample["review"])}
rows = []
for theme in CODES.values():
    tp = sum(theme in pred[i] and theme in gold[i] for i in gold)
    fp = sum(theme in pred[i] and theme not in gold[i] for i in gold)
    fn = sum(theme not in pred[i] and theme in gold[i] for i in gold)
    rows.append(dict(theme=theme, support=tp + fn, tp=tp, fp=fp, fn=fn,
                     precision=tp / (tp + fp) if tp + fp else float("nan"),
                     recall=tp / (tp + fn) if tp + fn else float("nan")))
res = pd.DataFrame(rows)
res["f1"] = 2 * res.precision * res.recall / (res.precision + res.recall)
tp, fp, fn = res.tp.sum(), res.fp.sum(), res.fn.sum()
micro = dict(theme="ALL (micro)", support=tp + fn, tp=tp, fp=fp, fn=fn,
             precision=tp / (tp + fp), recall=tp / (tp + fn))
micro["f1"] = 2 * micro["precision"] * micro["recall"] / (micro["precision"] + micro["recall"])
res = pd.concat([res, pd.DataFrame([micro])], ignore_index=True)
res.round(2).to_csv(HERE / "validation_results.csv", index=False)
print(res.round(2).to_string(index=False))

print("\nMisses and false alarms:")
for i in gold:
    if pred[i] != gold[i]:
        miss = {k for k, v in CODES.items() if v in gold[i] - pred[i]}
        extra = {k for k, v in CODES.items() if v in pred[i] - gold[i]}
        print(f"{i:>3} missed={''.join(sorted(miss)) or '-'} extra={''.join(sorted(extra)) or '-'} | {sample.review[i][:110]!r}")
