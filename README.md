# Excel Fighter

엑셀 VBA로 만드는 실시간 2D 격투 게임 (Windows 엑셀, `.xlsm`)

## 폴더 구조

```
excel-fighter/
  ExcelFighter.xlsm   <- 엑셀에서 직접 생성
  src/                <- VBA 모듈(.bas), git으로 관리
    GameLoop.bas      고정 타임스텝 루프, 세션 설정/복구
    KeyInput.bas      GetAsyncKeyState 입력 (held/pressed), 포커스 없으면 무시
    Renderer.bas      셀 픽셀 프레임버퍼, 스프라이트(좌우 반전), 바뀐 셀만 칠하기
    Data.bas          data/*.txt 로더 (스프라이트, 애니메이션, 스테이지)
    Game.bas          메뉴 + 순수 게임 로직 (GameInit / GameTick / GameRender) - 하니스와 공유
    CpuAI.bas         CPU: 반응 지연·확률 기반 판단, 실제 커맨드 입력으로 필살기 사용, 난이도 3단계
    Projectile.bas    장풍 (상쇄, 다단히트)
    Font.bas          3x5 픽셀 폰트
    Fighter.bas       상태 머신: 대기/걷기/앉기/점프/공격/가드/피격/다운/잡기, 콤보 체인, 입력 버퍼
    Combat.bas        히트박스 vs 허트박스, 상단/하단/중단 가드 판정, 데미지, 히트스톱, KO
    FX.bas            히트/가드 스파크
    Hud.bas           체력 바(빨간 잔상), KO 표시, 입력 모니터
    WinApi.bas        Sleep (엑셀 전용)
    Gallery.bas       모든 스프라이트를 Gallery 시트에 펼쳐서 확인
    DevTools.bas      src <-> 워크북 동기화 (Import/Export)
  data/               <- 게임 데이터 (텍스트, 엑셀이 실행 시 읽음)
    danwoo.txt, cheolsan.txt, stage_sanmun.txt
  tools/              <- 데이터 생성기 (Python)
    gen_sprites.py    관절 리그 -> 도트 스프라이트
    gen_stage.py      스테이지 배경
    harness/          LibreOffice로 실제 VBA 코드를 헤드리스 실행하는 테스트 하니스
  docs/preview/       <- 생성된 스프라이트/배경 미리보기 PNG
```

> `data/`의 스프라이트는 생성기 출력물이라 직접 고치지 않는다. 포즈나 색을 바꾸려면 `tools/`의 스크립트를 수정한 뒤 `python tools/gen_sprites.py`를 실행한다.

> `.bas` 파일은 ASCII만 쓴다. VBE가 시스템 코드페이지(CP949)로 읽기 때문에 한글 주석을 넣으면 깨진다.

## 최초 세팅 (1회)

1. 엑셀에서 새 통합 문서를 만들고 `src/`와 같은 폴더에 `ExcelFighter.xlsm`으로 저장한다.
   (OneDrive 동기화 폴더에 두면 `ThisWorkbook.Path`가 URL이 되므로 로컬 경로를 쓴다.)
2. 파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정
   - **"VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음"**을 켠다.
3. Alt+F11 → 파일 > 파일 가져오기 → `src/DevTools.bas`
4. Alt+F8로 `DevTools.ImportSrc`를 실행하면 나머지 모듈을 가져온다.

이후로는 코드가 바뀔 때마다 `ImportSrc`만 다시 실행하면 된다.
VBE에서 코드를 직접 고쳤다면 `ExportSrc`를 실행한 뒤 커밋한다.

## 실행

Alt+F8 → `StartGame` → 메뉴에서 ↑↓로 항목, ←→로 값 변경, A로 시작

- FIGHTER: 단우 / 철산 (상대는 CPU가 나머지 캐릭터)
- LEVEL: 1 EASY / 2 NORMAL / 3 HARD
- MODE: VS CPU / WATCH CPU (CPU끼리 대전 관전)

| 키 | 동작 |
|---|---|
| ← → | 이동 |
| ↑ | 점프 |
| ↓ | 앉기 |
| A / S / Z / X | 약P / 강P / 약K / 강K (↓ 누르고 = 앉아 공격, 점프 중 = 공중 공격) |
| A + Z 동시 (근접) | 잡기 |
| 뒤로 (상대가 공격 중일 때) | 가드 (↓ 같이 = 앉아 가드). 하단(앉은 킥)은 앉아서, 점프 공격은 서서 막아야 함 |
| 약공격 적중 후 다시 공격 | 체인 콤보 (최대 3연결, 예: A → A → S) |
| →→ / ←← | 대시 / 백대시 |
| 잡힌 직후 A + Z | 잡기 풀기 (일반 잡기만) |
| 공격 적중/가드 후 필살기 커맨드 | 필살기 캔슬 (일반기 → 필살기 → 궁극기) |

필살기 (방향은 상대 쪽이 →, P = A 또는 S)

| 캐릭터 | 커맨드 | 기술 |
|---|---|---|
| 단우 | ↓↘→ + P | 청풍장 (장풍) |
| 단우 | →↓↘ + P | 학비상 (무적 대공) |
| 단우 | ↓→↓→ + P (게이지 MAX) | 궁극기 천풍만장 (5연타 대형 장풍) |
| 철산 | →↓← + P (근접) | 금강파쇄 (커맨드 잡기, 풀 수 없음) |
| 철산 | ↓↘→ + P | 철산고 (어깨 돌진, 다운) |
| 철산 | ↓→↓→ + P (게이지 MAX) | 궁극기 나한복마 (원거리 잡기) |

게이지는 때리거나 맞으면 찬다. 필살기는 가드해도 깎기 데미지가 들어가고, 상대 공격 발생 중에 맞히면 COUNTER(데미지·경직 증가),
3타째부터 콤보 데미지가 10%씩 줄어든다.
| Esc | 종료 |

디버그 키

| 키 | 동작 |
|---|---|
| Tab | 조작 캐릭터 전환 (단우 <-> 철산) |
| D | 상대 전환: CPU → 가만히 → 가드 더미 → 공격 더미 |
| H / J | 상대 피격 / 다운 모션 강제 재생 |
| F1 | 판정 박스 표시 (초록 = 피격 판정, 빨강 = 공격 판정) |

기술별 발생/지속/후딜/데미지/판정 박스는 `data/moves.txt`에서 수정한다 (코드 수정 불필요, `StartGame` 다시 실행).

전체 도트 확인: Alt+F8 → `ShowGallery`

- 상태 표시줄에 `FPS | ticks/s | paint calls/frame`이 표시된다.
- 엑셀이 입력을 받지 않게 되면 VBE 직접 실행 창(Ctrl+G)에 `GameLoop.RestoreExcel`을 입력한다.

## 헤드리스 테스트 (엑셀 없이)

`src/*.bas`를 LibreOffice의 VBA 호환 모드로 그대로 실행한다. 키 입력은 스크립트(`tools/harness/scripts/*.txt`)로 대신하고,
결과는 `tools/harness/out/<이름>.log`(틱별 상태)와 `.gif`(화면)로 나온다. 엑셀 속도와 실제 키 입력은 검증하지 못한다.

```
python3 -I tools/harness/run_harness.py            # 전체 시나리오
python3 tools/harness/strip.py combo 8 9 10 22      # 특정 틱 화면을 PNG로
```

LibreOffice 호환을 위해 지키는 규칙: Collection은 `.Item(key)`로 조회, 한 줄 If의 양쪽에 인자 있는 Sub 호출 금지,
`FX.Spawn (a) \ 2, ...`처럼 괄호로 시작하는 인자 금지, 함수 반환형 `As String()` 대신 `As Variant`.

## 단계

- [x] ① 루프 + 입력 + 박스 이동
- [x] ② 캐릭터 그리기 + 스테이지 배경
- [x] ③ 상태 머신 (공격 / 가드 / 피격 / 다운)
- [x] ④ 공격 판정, 체력 바, 콤보 체인, 잡기, KO
- [ ] ⑤ 필살기·궁극기·대시·카운터 + CPU 난이도 3단계 + 메뉴 ← **현재** (헤드리스 시나리오 15개 통과, 엑셀 확인 대기)
- [ ] ⑥ 라운드 시스템 (3판 2선승, 체력 바, 타이머)
