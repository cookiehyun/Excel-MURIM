# Balance (CPU HARD vs CPU HARD)

Measured headless with `tools/harness` (CMD norender, 60 s rounds, best of 3, both sides,
varied seeds, ~25 matches per pairing). CPU-vs-CPU numbers are a sanity check, not proof of
human-level balance.

| fighter | overall | notes |
|---|---|---|
| Cheolsan | 58% | high HP, command grab, slow (strong vs Hongryeon 79%) |
| Muyeong | 53% | longest reach (HP 980 -> 940 after naegong) |
| Heukrim | 50% | zoner: cheap needles (naegong 250), HP 1070 |
| Seolhwa | 48% | fastest, lowest HP |
| Hongryeon | 45% | mid-range fans, cheaper boomerang (naegong 300) |
| Danwoo | 45% | all-rounder, fireball naegong 350 |

Naegong (inner energy, added after players found fireball spam unanswerable):
specials 1 / 2 cost naegong (250-400 of 1000), refill 7 per tick, no refill while the special is out;
fireball recovery raised so a jump over it can be punished. Supers still use only the super gauge.

Pass history
1. 3 fighters: Cheolsan 64% -> command grab / throw / HP nerf; Muyeong 38% -> buffs
2. 4 fighters: Seolhwa 81% -> speed, damage, recovery nerfs
3. 6 fighters: Heukrim 15% -> HP, speed, startup, needles buffs; Muyeong slight nerf
4. CPU style per fighter (zoners keep range), varied seeds; Cheolsan small nerf
5. Heukrim 38% -> HP 1040, needles 65, sweep recovery 10
6. Naegong gauge: zoners collapsed (Heukrim 24%, Hongryeon 35%), Muyeong 67%
   -> cheaper zoner specials, Muyeong HP 940
7. Cheolsan 59% -> HP 1030; Heukrim HP 1070; Danwoo fireball 350 -> 45-58%

Re-run: generate `tools/harness/scripts/bal_<a>_<b>.txt`, run them (optionally in parallel with
`HARNESS_PIPE=<name>`), then `python3 tools/harness/balance.py`.
