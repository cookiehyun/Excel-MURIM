# Balance (CPU HARD vs CPU HARD)

Measured headless with `tools/harness` (CMD norender, 60 s rounds, best of 3, both sides,
varied seeds, ~25 matches per pairing). CPU-vs-CPU numbers are a sanity check, not proof of
human-level balance.

| fighter | overall | notes |
|---|---|---|
| Muyeong | 56% | longest reach |
| Hongryeon | 54% | HP 1000, uppercut invincible 4 ticks |
| Heukrim | 52% | zoner, cheap needles |
| Seolhwa | 48% | fastest, lowest HP |
| Cheolsan | 47% | command grab reach 24 -> 20, damage 160 -> 145 |
| Danwoo | 44% | fireball naegong 300 |

Lopsided pairings left (n ~ 20-30 each, so +-10%p noise): Heukrim > Cheolsan 82:18,
Hongryeon > Seolhwa 71:29, Cheolsan > Danwoo 65:35.

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
8. Matchups Cheolsan-Hongryeon 79:21 and Heukrim-Danwoo 73:27: Cheolsan's command grab dealt 63% of his damage
   -> grab reach 20 / damage 145; Hongryeon HP 1000 + longer uppercut invincibility; Danwoo fireball 300
   -> those two now 58:42 and 45:55, overall 44-56%

Re-run: generate `tools/harness/scripts/bal_<a>_<b>.txt`, run them (optionally in parallel with
`HARNESS_PIPE=<name>`), then `python3 tools/harness/balance.py`.
