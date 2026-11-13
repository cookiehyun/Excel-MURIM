# Excel MURIM

엑셀 시트 위에서 돌아가는 2D 무협 격투 게임입니다. 셀 하나를 픽셀 하나로 써서 화면을 그리고, 게임은 전부 VBA로 만들었습니다. 캐릭터 6명과 최종 보스, 아케이드·태그·2인 대전 모드가 있습니다.

A real-time 2D martial-arts (wuxia) fighting game that runs inside an Excel sheet. Every cell is one pixel, and the whole game is written in VBA. Six fighters plus a final boss, with arcade, tag-team and two-player versus modes.

Windows 엑셀(.xlsm) 전용입니다. 맥 엑셀에서는 돌아가지 않습니다. / Windows Excel (.xlsm) only. It does not run on Excel for Mac.

## 시작하기 / Getting started

| 한국어 | English |
|---|---|
| 처음 한 번만 하면 됩니다. 자세한 안내는 [docs/SETUP_KO.md](docs/SETUP_KO.md) | One-time setup. Step-by-step guide (Korean): [docs/SETUP_KO.md](docs/SETUP_KO.md) |
| 1. 엑셀에서 새 통합 문서를 만들어 `excel-murim` 폴더 안에 `ExcelMurim.xlsm`으로 저장. OneDrive 동기화 폴더는 피하세요. | 1. Create a new workbook and save it as `ExcelMurim.xlsm` inside the `excel-murim` folder. Avoid OneDrive-synced folders. |
| 2. 파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정에서 "VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음" 켜기 | 2. File > Options > Trust Center > Trust Center Settings > Macro Settings: turn on "Trust access to the VBA project object model". |
| 3. Alt+F11 → 파일 > 파일 가져오기 → `src/DevTools.bas` | 3. Alt+F11 → File > Import File → `src/DevTools.bas` |
| 4. Alt+F8 → `DevTools.ImportSrc` 실행 (나머지 모듈을 불러옴) | 4. Alt+F8 → run `DevTools.ImportSrc` (imports the other modules) |
| 업데이트: `git pull` → `ImportSrc` → VBE에서 디버그 > VBAProject 컴파일 | Updating: `git pull` → `ImportSrc` → in the VBE, Debug > Compile VBAProject |
| 실행: Alt+F8 → `StartGame` | Play: Alt+F8 → `StartGame` |

<details>
<summary><b>한국어 (게임 설명)</b></summary>

## 게임 모드

시작 화면에서 SINGLE PLAYER 또는 MULTIPLAYER를 고릅니다. 메뉴는 ↑↓로 움직이고 A로 결정, S로 뒤로 갑니다.

| 메뉴 | 모드 | 내용 |
|---|---|---|
| SINGLE PLAYER | ARCADE | 캐릭터와 난이도를 고르고 나머지 5명과 차례로 싸운 뒤 최종 보스 천마와 대결. 이기면 캐릭터별 엔딩 |
| SINGLE PLAYER | TAG BATTLE | 2명씩 팀을 짜서 CPU 팀과 싸웁니다. C 키로 교체 |
| SINGLE PLAYER | TRAINING | 시간 제한 없음. 콤보가 끝나면 체력이 다시 차고, 내공과 게이지는 항상 가득 |
| MULTIPLAYER | VERSUS | 키보드 하나로 1P와 2P가 1대1로 대결 |
| MULTIPLAYER | KEY CONFIG | 1P·2P 키 바꾸기. `data/keys.txt`에 저장됩니다 |

### 아케이드
- 상대 5명은 매번 순서가 섞입니다. 4번째 상대부터 CPU가 한 단계 강해지고, 보스는 항상 HARD입니다.
- 지면 CONTINUE? 화면이 뜹니다. 9초 안에 A를 누르면 같은 상대와 다시 싸우고, S를 누르거나 시간이 지나면 게임 오버입니다.
- 천마는 체력 1400에 흑림 계열 기술을 더 강하게 씁니다(마공탄, 순간이동, 궁극기 천마무황).

### TAG BATTLE
- 내 팀 2명을 고른 다음 CPU 팀 2명을 고릅니다.
- 서 있거나 걷거나 앉아 있을 때 C를 누르면 교체됩니다. 한 번 바꾸면 2초 동안은 다시 못 바꿉니다. 들어오는 캐릭터는 위에서 뛰어내리는데, 내려오는 동안에는 맞을 수 있습니다.
- 맞은 데미지의 절반은 체력바에 빨갛게 남고, 쉬는 동안 천천히 돌아옵니다. 쉬는 캐릭터는 내공도 빨리 찹니다.
- 궁극기 게이지는 팀이 같이 씁니다.
- 한 명이라도 쓰러지면 그 라운드는 집니다. 시간이 다 되면 팀 전체 체력이 많은 쪽이 이깁니다.
- CPU도 체력이 떨어지면 교체합니다.

### 라운드
- 한 라운드 60초. KO를 시키거나, 시간이 끝났을 때 체력 비율이 높으면 이깁니다. 먼저 2라운드를 이기면 승리.
- 게이지는 다음 라운드로 이어집니다.

## 조작

1인 플레이 키입니다. 2인 대전 키는 아래 표에 있습니다.

| 키 | 동작 |
|---|---|
| ← → | 이동 |
| ↑ | 점프 |
| ↓ | 앉기 |
| A / S / Z / X | 약손 / 강손 / 약발 / 강발. ↓를 누른 채면 앉아서, 점프 중이면 공중에서 |
| A + Z 동시 (가까이서) | 잡기. 잡힌 직후 A + Z를 누르면 풀 수 있음 (커맨드 잡기는 못 풂) |
| 상대 반대쪽 | 가드. 앉은 발차기는 앉아서, 점프 공격은 서서 막아야 함 |
| →→ / ←← | 대시 / 백대시 |
| C | 교체 (TAG BATTLE) |
| Esc | 경기 중에는 일시정지, 시작 화면에서는 게임 종료 |

2인 대전 기본 키 (KEY CONFIG에서 바꿀 수 있고, 이미 쓰는 키를 고르면 두 키가 서로 바뀝니다)

| | 이동 | 약손 / 강손 | 약발 / 강발 |
|---|---|---|---|
| 1P | W A S D | R / T | F / G |
| 2P | 방향키 | K / L | , / . |

- 경기 중에 다른 창을 클릭하면 자동으로 일시정지됩니다. S나 RESUME으로 이어서 하면 됩니다.
- 게임 중에는 셀을 클릭할 수 없습니다. 게임을 끝내면 다시 됩니다.
- 키보드에 따라 여러 키를 한꺼번에 누르면 일부가 안 먹힐 수 있습니다.

## 캐릭터

P는 A 또는 S, K는 Z 또는 X입니다. →는 상대가 있는 쪽입니다.

| 캐릭터 | 스타일 | 필살기 1 | 필살기 2 | 궁극기 (게이지 MAX) |
|---|---|---|---|---|
| 단우 | 올라운더 | ↓↘→P 청풍장 (장풍) | →↓↘P 학비상 (무적 대공) | ↓→↓→P 천풍만장 (5연타 장풍) |
| 철산 | 잡기 | →↓←P 금강파쇄 (커맨드 잡기) | ↓↘→P 철산고 (돌진) | ↓→↓→P 나한복마 (4연타 후 내리꽂기) |
| 무영 | 검객 | ↓↘→P 검기 (칼바람) | ↓↙←P 섬광참 (돌진 베기) | ↓→↓→P 만검귀종 (연속 베기) |
| 설화 | 발차기 | ↓↙←K 연환비각 (3연타 돌진) | →↓↘K 선풍각 (공중제비 대공) | ↓→↓→K 낙화유수 (연속 차기) |
| 흑림 | 견제 | ↓↘→P 흑사침 (독침) | ↓↙←K 귀영보 (등 뒤로 순간이동) | ↓→↓→P 천마강림 (순간이동 장법) |
| 홍련 | 부채 | ↓↘→P 부채 던지기 (돌아오는 부채) | →↓↘K 연화선 (회전 차기) | ↓→↓→P 홍련업화 (상대 발밑 불기둥) |
| 천마 | 아케이드 보스 | ↓↘→P 마공탄 | ↓↙←K 순간이동 | ↓→↓→P 천마무황 |

- 내공(초록 바): 필살기 1·2를 쓰면 줄고 시간이 지나면 다시 찹니다. 모자라면 커맨드를 넣어도 일반 공격이 나가고 바가 빨갛게 깜빡입니다.
- 궁극기 게이지(파란 바): 때리거나 맞으면 찹니다.
- 필살기는 막아도 체력이 조금 깎입니다.
- 상대가 공격을 내미는 순간에 맞히면 COUNTER가 떠서 데미지와 경직이 늘어납니다.
- 콤보는 3타째부터 데미지가 10%씩 줄어듭니다.

## 콤보

실제로 이어지는 걸 확인한 콤보입니다. 숫자는 단우에게 넣었을 때 데미지이고, `-`는 그 캐릭터로는 안 이어진다는 뜻입니다. 필살기와 궁극기 커맨드는 위 표를 보세요.

약공격(A·Z, 앉아서 A·Z) 다음에는 다른 기본기가 최대 3번까지 이어집니다. 기본기가 맞거나 막히면 필살기나 궁극기로 끊고 이어 쓸 수 있습니다.

| 입력 | 단우 | 철산 | 무영 | 설화 | 흑림 | 홍련 | 천마 |
|---|---|---|---|---|---|---|---|
| ↓Z → ↓X (다리 쓸기, 다운) | 120 | 150 | 118 | 101 | 132 | 118 | 150 |
| A → A → S | 141 | - | - | 99 | - | - | - |
| A → S | - | 160 | - | - | - | - | - |
| S → 필살기1 | 160 | - | 145 | 146 | 155 | 150 | 180 |
| S → 필살기2 | - | 240 | 185 | 150 | - | - | - |
| A → A → 필살기1 | 123 | - | - | 112 | 119 | 123 | - |
| A → A → 필살기2 | - | 183 | 140 | - | - | - | - |
| S → 궁극기 | 294 | - | 415 | 365 | 410 | 330 | 475 |
| 궁극기만 (가까이서) | - | 380 | - | - | - | - | - |
| 점프 X → (착지) A → S | 170 | - | - | - | - | - | - |

- 철산의 금강파쇄와 궁극기는 잡기라서 콤보 끝에 붙이지 않고 따로 씁니다.
- 흑림의 귀영보는 순간이동이라 콤보에 넣지 않습니다.

밸런스는 CPU끼리 싸운 승률로 맞췄고, 지금 전체 승률은 44~56%입니다. 기록은 [docs/BALANCE.md](docs/BALANCE.md)에 있습니다.

</details>

<details>
<summary><b>English (how to play)</b></summary>

## Game modes

Pick SINGLE PLAYER or MULTIPLAYER on the title screen. In menus, ↑↓ moves, A confirms and S goes back.

| Menu | Mode | What it is |
|---|---|---|
| SINGLE PLAYER | ARCADE | Pick a fighter and a level, beat the other five, then face the final boss Cheonma. Each fighter has an ending |
| SINGLE PLAYER | TAG BATTLE | Two-fighter teams against a CPU team. Press C to switch |
| SINGLE PLAYER | TRAINING | No timer. Health refills after a combo; inner energy and super gauge stay full |
| MULTIPLAYER | VERSUS | 1P vs 2P on one keyboard |
| MULTIPLAYER | KEY CONFIG | Change 1P / 2P keys. Saved to `data/keys.txt` |

### Arcade
- The five opponents come in a random order. From the fourth one the CPU gets one level harder; the boss is always HARD.
- Lose and you get CONTINUE?. Press A within 9 seconds to retry the same fight; press S or let it run out and it's game over.
- Cheonma has 1400 HP and a stronger version of Heukrim's moves (demon orb, teleport, super Cheonma Muhwang).

### TAG BATTLE
- Pick your two fighters, then the CPU's two.
- Press C while standing, walking or crouching to switch. After a switch you have to wait 2 seconds. The incoming fighter drops in from above and can be hit on the way down.
- Half of the damage you take stays on the health bar in red and slowly comes back while that fighter rests. A resting fighter also refills inner energy faster.
- The super gauge is shared by the team.
- If either fighter is knocked out, the round is lost. On time-out, the team with more total health wins.
- The CPU switches too when its fighter is low.

### Rounds
- 60 seconds per round. Win by KO, or by having more health (as a ratio) when time runs out. First to two rounds wins.
- The super gauge carries over to the next round.

## Controls

Single-player keys. Two-player keys are in the second table.

| Key | Action |
|---|---|
| ← → | Move |
| ↑ | Jump |
| ↓ | Crouch |
| A / S / Z / X | Light punch / heavy punch / light kick / heavy kick. Hold ↓ for crouching attacks; in the air for jump attacks |
| A + Z together (close) | Throw. Press A + Z right after being grabbed to break it (command grabs can't be broken) |
| Away from the opponent | Block. Crouching kicks must be blocked low, jump attacks high |
| →→ / ←← | Dash / back dash |
| C | Switch (TAG BATTLE) |
| Esc | Pause during a fight, quit on the title screen |

Two-player default keys (change them in KEY CONFIG; picking a key that's already used swaps the two)

| | Move | Light / heavy punch | Light / heavy kick |
|---|---|---|---|
| 1P | W A S D | R / T | F / G |
| 2P | Arrow keys | K / L | , / . |

- Clicking another window during a fight pauses the game. Press S or pick RESUME to continue.
- You can't click cells while the game is running. That comes back when you quit.
- Some keyboards drop keys when many are held at once.

## Fighters

P is A or S, K is Z or X. → means toward the opponent.

| Fighter | Style | Special 1 | Special 2 | Super (full gauge) |
|---|---|---|---|---|
| Danwoo | All-rounder | ↓↘→P Cheongpungjang (palm blast) | →↓↘P Hakbisang (invincible uppercut) | ↓→↓→P Cheonpung Manjang (5-hit blast) |
| Cheolsan | Grappler | →↓←P Geumgang Pashwae (command grab) | ↓↘→P Cheolsango (shoulder rush) | ↓→↓→P Nahan Bokma (4 hits, then a slam) |
| Muyeong | Swordsman | ↓↘→P Geomgi (sword wave) | ↓↙←P Seomgwangcham (dash slash) | ↓→↓→P Mangeom Gwijong (flurry of cuts) |
| Seolhwa | Kicker | ↓↙←K Yeonhwan Bigak (3-hit rush) | →↓↘K Seonpunggak (flip kick) | ↓→↓→K Nakhwa Yusu (kick barrage) |
| Heukrim | Zoner | ↓↘→P Heuksachim (poison needles) | ↓↙←K Gwiyeongbo (teleport behind) | ↓→↓→P Cheonma Gangrim (teleport palm) |
| Hongryeon | Fans | ↓↘→P Fan throw (comes back) | →↓↘K Yeonhwaseon (spinning kick) | ↓→↓→P Hongryeon Eophwa (fire pillar under the opponent) |
| Cheonma | Arcade boss | ↓↘→P Demon orb | ↓↙←K Teleport | ↓→↓→P Cheonma Muhwang |

- Inner energy (green bar): specials 1 and 2 use it and it refills over time. Without enough of it you get a normal attack instead, and the bar flashes red.
- Super gauge (blue bar): fills when you hit or get hit.
- Specials do a little damage even when blocked.
- Hitting an opponent while their attack is coming out is a COUNTER: more damage and longer stun.
- From the third hit of a combo, each hit does 10% less damage.

## Combos

Combos checked to work. Numbers are damage against Danwoo; `-` means it doesn't connect for that fighter. Special and super inputs are in the fighter table.

A light attack (A·Z, standing or crouching) can chain into another normal, up to three. When a normal hits or is blocked, you can cancel it into a special or super.

| Input | Danwoo | Cheolsan | Muyeong | Seolhwa | Heukrim | Hongryeon | Cheonma |
|---|---|---|---|---|---|---|---|
| ↓Z → ↓X (sweep, knockdown) | 120 | 150 | 118 | 101 | 132 | 118 | 150 |
| A → A → S | 141 | - | - | 99 | - | - | - |
| A → S | - | 160 | - | - | - | - | - |
| S → special 1 | 160 | - | 145 | 146 | 155 | 150 | 180 |
| S → special 2 | - | 240 | 185 | 150 | - | - | - |
| A → A → special 1 | 123 | - | - | 112 | 119 | 123 | - |
| A → A → special 2 | - | 183 | 140 | - | - | - | - |
| S → super | 294 | - | 415 | 365 | 410 | 330 | 475 |
| Super alone (close) | - | 380 | - | - | - | - | - |
| Jump X → (land) A → S | 170 | - | - | - | - | - | - |

- Cheolsan's command grab and super are grabs, so use them on their own rather than at the end of a combo.
- Heukrim's Gwiyeongbo is a teleport and doesn't go in combos.

Balance is tuned with CPU-vs-CPU matches; overall win rates are currently 44-56%. Notes are in [docs/BALANCE.md](docs/BALANCE.md).

</details>

<details>
<summary><b>개발자용 / For developers</b></summary>

<details>
<summary>한국어</summary>

### 문제가 생겼을 때
- 엑셀이 입력을 받지 않으면 VBE 직접 실행 창(Ctrl+G)에 `GameLoop.RestoreExcel`을 입력합니다.
- 상태 표시줄에 FPS, 틱 수, 프레임당 칠한 횟수, 렌더 시간이 나옵니다.

### 디버그 키 (TRAINING에서만)
| 키 | 동작 |
|---|---|
| Tab | 조작 캐릭터 바꾸기 |
| D | 상대 바꾸기: CPU → 가만히 → 가드만 → 공격만 |
| H / J | 상대 피격 / 다운 모션 강제 재생 |
| F1 | 판정 박스 표시 (초록 = 맞는 판정, 빨강 = 공격 판정) |

전체 도트 확인: Alt+F8 → `ShowGallery`

### 데이터 고치기
- 기술 발생·지속·후딜·데미지·판정 박스·내공 소모는 `data/moves.txt`에서 고칩니다. 코드는 안 건드려도 됩니다.
- `data/`의 스프라이트, 배경, 이펙트는 생성기가 만든 파일입니다. 포즈나 색은 `tools/`의 스크립트를 고친 뒤 다시 생성합니다 (`python tools/gen_sprites.py`).
- 기술의 여러 프레임 모션은 `tools/gen_sprites.py`의 `MOVE_ANIMS`에 있습니다 (`<기술>_pre` 발생, `_seq` 지속·공중, `_rec` 후딜).
- `.bas` 파일에는 영어만 씁니다. VBE가 CP949로 읽어서 한글 주석이 깨집니다.
- VBE에서 코드를 직접 고쳤다면 `ExportSrc`를 실행한 뒤 커밋합니다.

### 폴더 구조
```
excel-murim/
  ExcelMurim.xlsm     엑셀에서 직접 만듦 (git에 없음)
  src/                VBA 모듈 (.bas)
    GameLoop.bas      고정 타임스텝 루프, 엑셀 설정/복구
    Game.bas          상태 전환, 한 틱 진행과 그리기, 일시정지 메뉴
    KeyInput.bas      키 입력 (GetAsyncKeyState), 창 포커스 확인
    KeyConfig.bas     2인 키 설정 화면, data/keys.txt 저장
    Renderer.bas      셀 픽셀 버퍼, 스프라이트, 바뀐 셀만 칠하기
    Data.bas          data/*.txt 읽기
    Screens.bas       시작, 캐릭터·난이도·무대 선택 화면
    Arcade.bas        아케이드 대진, 컨티뉴, 엔딩
    Tag.bas           2 VS 2 교체, 대기 회복
    Match.bas         라운드, 타이머, 훈련장 회복
    Fighter.bas       캐릭터 상태 머신, 입력, 커맨드 판정, 궁극기 연출
    Combat.bas        공격 판정, 가드, 데미지, 콤보, 카운터, KO
    CpuAI.bas         CPU 판단, 난이도 3단계
    Projectile.bas    장풍, 부채, 불기둥
    FX.bas            타격 이펙트
    Hud.bas           체력·내공·게이지 바
    PixelFont.bas     픽셀 폰트
    WinApi.bas        Sleep (엑셀 전용)
    Gallery.bas       스프라이트 모아보기
    DevTools.bas      src와 워크북 동기화 (ImportSrc / ExportSrc)
  data/               실행 중에 읽는 데이터
    <캐릭터>.txt      스프라이트
    stage_*.txt       무대 배경 5종
    fx.txt            투사체 이펙트
    moves.txt         기술 데이터
  tools/              데이터 생성기 (Python)
    gen_sprites.py    관절 리그로 도트 스프라이트 생성
    gen_stages.py     무대 배경
    gen_fx.py         투사체 이펙트
    anim_preview.py   기술 애니메이션 미리보기
    lint_vba.py       엑셀/LibreOffice 호환 검사
    harness/          LibreOffice로 VBA 코드를 그대로 돌리는 테스트
  docs/               세팅 안내, 밸런스 기록, 미리보기 PNG
```

### 엑셀 없이 테스트하기
`src/*.bas`를 LibreOffice의 VBA 호환 모드로 그대로 돌립니다. 키 입력은 `tools/harness/scripts/*.txt` 스크립트가 대신하고, 결과는 `tools/harness/out/<이름>.log`(틱별 상태)와 `.gif`로 나옵니다. 엑셀에서의 속도와 실제 키보드 입력은 여기서 확인할 수 없습니다.

```
python3 -I tools/harness/run_harness.py            # 시나리오 전체
python3 tools/harness/strip.py combo 8 9 10 22      # 특정 틱 화면을 PNG로
python3 tools/lint_vba.py                           # 호환 검사
```

`lint_vba.py`가 잡아주는 규칙:
- 모듈 수준 선언(Const, 변수)은 첫 프로시저보다 위에 둡니다 (엑셀에서 컴파일 오류).
- Collection은 `.Item(key)`로 읽습니다.
- 한 줄 If의 양쪽에 인자 있는 Sub 호출을 쓰지 않습니다.
- Sub 인자를 괄호로 시작하지 않습니다.
- 함수가 `As String()`을 돌려주지 않고 `As Variant`를 씁니다.
- `Font`, `Scale`, `Line`, `Input`처럼 엑셀이 쓰는 이름을 모듈이나 변수 이름으로 쓰지 않습니다.

### 진행 상황
- [x] 루프, 입력, 렌더링
- [x] 캐릭터 6명, 무대 5종
- [x] 공격 판정, 콤보, 잡기, KO
- [x] 필살기, 궁극기, 대시, 카운터, CPU 난이도
- [x] 라운드, 메뉴, 일시정지
- [x] 아케이드와 최종 보스
- [x] 2 VS 2 태그, 2인 대전, 키 설정
- [ ] 태그 콤보·교체 공격, 효과음, 기술표

</details>

<details>
<summary>English</summary>

### If something goes wrong
- If Excel stops taking input, type `GameLoop.RestoreExcel` in the VBE Immediate window (Ctrl+G).
- The status bar shows FPS, ticks per second, paint calls per frame and render time.

### Debug keys (TRAINING only)
| Key | Action |
|---|---|
| Tab | Switch which fighter you control |
| D | Opponent: CPU → idle → block only → attack only |
| H / J | Force the opponent's hit / knockdown animation |
| F1 | Show boxes (green = hurtbox, red = hitbox) |

All sprites at once: Alt+F8 → `ShowGallery`

### Editing data
- Startup, active and recovery frames, damage, hitboxes and inner-energy cost live in `data/moves.txt`. No code changes needed.
- Sprites, stages and effects in `data/` are generated. Change the scripts in `tools/` and regenerate (`python tools/gen_sprites.py`).
- Multi-frame move animations are in `MOVE_ANIMS` in `tools/gen_sprites.py` (`<move>_pre` startup, `_seq` active/airborne, `_rec` recovery).
- Keep `.bas` files ASCII only. The VBE reads them in the system code page, so non-ASCII comments break.
- If you edit code in the VBE, run `ExportSrc` before committing.

### Folder layout
```
excel-murim/
  ExcelMurim.xlsm     created in Excel (not in git)
  src/                VBA modules (.bas)
    GameLoop.bas      fixed-timestep loop, Excel session setup/restore
    Game.bas          game state, one tick and one frame, pause menu
    KeyInput.bas      keyboard (GetAsyncKeyState), window focus check
    KeyConfig.bas     two-player key settings, data/keys.txt
    Renderer.bas      cell pixel buffer, sprites, repaint only changed cells
    Data.bas          reads data/*.txt
    Screens.bas       title, fighter / level / stage select
    Arcade.bas        arcade ladder, continue, endings
    Tag.bas           2 VS 2 switching and recovery
    Match.bas         rounds, timer, training refills
    Fighter.bas       fighter state machine, input, command motions, super cinematics
    Combat.bas        hits, blocking, damage, combos, counters, KO
    CpuAI.bas         CPU decisions, three levels
    Projectile.bas    blasts, fans, fire pillars
    FX.bas            hit effects
    Hud.bas           health, inner energy and super bars
    PixelFont.bas     pixel fonts
    WinApi.bas        Sleep (Excel only)
    Gallery.bas       sprite sheet viewer
    DevTools.bas      sync src with the workbook (ImportSrc / ExportSrc)
  data/               data read at run time
    <fighter>.txt     sprites
    stage_*.txt       5 stages
    fx.txt            projectile effects
    moves.txt         move data
  tools/              generators (Python)
    gen_sprites.py    joint rig to pixel sprites
    gen_stages.py     stage backgrounds
    gen_fx.py         projectile effects
    anim_preview.py   move animation preview
    lint_vba.py       Excel / LibreOffice compatibility check
    harness/          runs the VBA code in LibreOffice for tests
  docs/               setup guide, balance notes, preview PNGs
```

### Testing without Excel
`src/*.bas` runs unchanged in LibreOffice's VBA compatibility mode. Scripts in `tools/harness/scripts/*.txt` stand in for the keyboard, and each run writes `tools/harness/out/<name>.log` (state per tick) and a `.gif`. Excel speed and real keyboard input can't be checked this way.

```
python3 -I tools/harness/run_harness.py            # all scenarios
python3 tools/harness/strip.py combo 8 9 10 22      # PNG of chosen ticks
python3 tools/lint_vba.py                           # compatibility check
```

Rules `lint_vba.py` checks:
- Module-level declarations (Const, variables) go above the first procedure (Excel compile error otherwise).
- Read Collections with `.Item(key)`.
- No one-line If with Sub calls taking arguments in both branches.
- Don't start a Sub argument with a parenthesis.
- Return `As Variant` instead of `As String()`.
- Don't name modules or variables after Excel names like `Font`, `Scale`, `Line`, `Input`.

### Progress
- [x] Loop, input, rendering
- [x] Six fighters, five stages
- [x] Hits, combos, throws, KO
- [x] Specials, supers, dashes, counters, CPU levels
- [x] Rounds, menus, pause
- [x] Arcade and final boss
- [x] 2 VS 2 tag, two-player versus, key config
- [ ] Tag combos and tag attacks, sound, move list

</details>

</details>
