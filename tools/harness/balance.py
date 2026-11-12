"""Summarise CPU-vs-CPU balance runs (tools/harness/out/bal_*.log, written with CMD norender).
usage: python3 tools/harness/balance.py"""
import collections, glob, os
HERE = os.path.dirname(os.path.abspath(__file__))
pair = collections.defaultdict(lambda: [0, 0, 0])   # (a, b) sorted -> [a wins, b wins, draws]
char = collections.defaultdict(lambda: [0, 0])      # name -> [wins, matches]
for p in glob.glob(os.path.join(HERE, "out", "bal_*.log")):
    for line in open(p):
        if not line.startswith("R "):
            continue
        _, _, a, wa, b, wb = line.split()
        wa, wb = int(wa), int(wb)
        key = tuple(sorted((a, b)))
        if wa == wb:
            pair[key][2] += 1
        else:
            winner = a if wa > wb else b
            pair[key][0 if winner == key[0] else 1] += 1
            char[winner][0] += 1
        char[a][1] += 1
        char[b][1] += 1
print("matchup                     wins          win rate")
for (a, b), (x, y, d) in sorted(pair.items()):
    n = x + y + d
    print(f"{a:9s} vs {b:9s}  {x:3d} - {y:3d} (draw {d})   {100 * x / max(1, n):5.1f}% / {100 * y / max(1, n):5.1f}%   n={n}")
print()
for c, (w, n) in sorted(char.items(), key=lambda kv: -kv[1][0] / max(1, kv[1][1])):
    print(f"{c:9s} overall {100 * w / max(1, n):5.1f}%  ({w}/{n})")
