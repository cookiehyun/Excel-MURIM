# Balance (CPU HARD vs CPU HARD)

Measured headless with `tools/harness` (CMD norender, 60 s rounds, best of 3, both sides,
varied seeds, ~25 matches per pairing). CPU-vs-CPU numbers are a sanity check, not proof of
human-level balance.

| fighter | overall | notes |
|---|---|---|
| Muyeong | 54% | longest reach, slower and weaker swings |
| Cheolsan | 54% | high HP, command grab, slow |
| Heukrim | 51% | zoner: fast needles, teleport (was 15% before buffs) |
| Seolhwa | 49% | fastest, lowest HP (was 81% before nerfs) |
| Danwoo | 48% | all-rounder |
| Hongryeon | 44% | mid-range fans, returning fan, fire-pillar super |

Pass history
1. 3 fighters: Cheolsan 64% -> command grab / throw / HP nerf; Muyeong 38% -> buffs
2. 4 fighters: Seolhwa 81% -> speed, damage, recovery nerfs
3. 6 fighters: Heukrim 15% -> HP, speed, startup, needles buffs; Muyeong slight nerf
4. CPU style per fighter (zoners keep range), varied seeds; Cheolsan small nerf
5. Heukrim 38% -> HP 1040, needles 65, sweep recovery 10

Re-run: generate `tools/harness/scripts/bal_<a>_<b>.txt`, run them (optionally in parallel with
`HARNESS_PIPE=<name>`), then `python3 tools/harness/balance.py`.
