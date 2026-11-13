# 새 엑셀 파일로 처음부터 세팅하기

경로 예시: `C:\Python\26-2\기계학습\STREET` (본인 경로로 바꿔서 읽기)

## 0. 준비
1. 열려 있는 엑셀 창을 **모두 닫는다**.
2. PowerShell을 열고 최신 코드를 받는다.
   ```
   cd "C:\Python\26-2\기계학습\STREET"
   git pull
   ```
3. 탐색기로 `STREET` 폴더를 열어 확인한다.
   - `src` 폴더: `.bas` 파일 **15개** (Combat, CpuAI, Data, DevTools, FX, Fighter, Gallery, Game, GameLoop, Hud, KeyInput, PixelFont, Projectile, Renderer, WinApi)
   - `data` 폴더: `cheolsan.txt`, `danwoo.txt`, `fx.txt`, `moves.txt`, `stage_sanmun.txt`
   - 예전에 만든 `.xlsm` 파일(`통합 문서1.xlsm` 등)은 지우거나 다른 폴더로 옮긴다.

## 1. 엑셀 보안 설정 (PC당 한 번)
1. 엑셀 실행 → **빈 통합 문서**.
2. **파일 → 옵션 → 보안 센터 → [보안 센터 설정...]**
3. 왼쪽 **매크로 설정**:
   - "알림이 포함된 VBA 매크로 사용 안 함" 선택 (기본값)
   - 아래쪽 **"VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음"** 체크 ← `ImportSrc`에 필요
4. (선택) 왼쪽 **신뢰할 수 있는 위치 → [새 위치 추가...]** → `STREET` 폴더 선택 → 확인.
   이렇게 하면 열 때마다 "콘텐츠 사용"을 누르지 않아도 된다.
5. **확인 → 확인**.

## 2. 새 워크북 만들고 저장
1. 1단계에서 연 빈 통합 문서에서 **F12** (다른 이름으로 저장).
2. 위치: **`STREET` 폴더 바로 안** (`src`, `data` 폴더가 보이는 곳. `src` 안이 아님).
3. 파일 이름: `ExcelMurim`
4. 파일 형식: **Excel 매크로 사용 통합 문서 (\*.xlsm)** ← 반드시. `.xlsx`면 코드가 저장되지 않는다.
5. **저장**. 탐색기에서 `STREET\ExcelMurim.xlsm` 이 `src`, `data`와 나란히 있는지 확인.

## 3. DevTools 가져오기 (워크북당 한 번)
1. 엑셀에서 **Alt+F11** → VBA 편집기.
2. **Ctrl+R** → 왼쪽에 프로젝트 탐색기. `VBAProject (ExcelMurim.xlsm)` 이 보여야 한다.
3. 프로젝트 탐색기에서 `VBAProject (ExcelMurim.xlsm)`을 한 번 클릭해 선택.
4. 메뉴 **파일 → 파일 가져오기...** → `STREET\src\DevTools.bas` → 열기.
5. `VBAProject (ExcelMurim.xlsm)` 아래 **모듈** 폴더에 `DevTools`가 생겼는지 확인.

## 4. 나머지 모듈 가져오기
1. 엑셀 창으로 돌아가 **Alt+F8** (매크로).
2. 목록에서 **`ImportSrc`** 선택 → **실행**.
3. "15 file(s) imported from ...\src\" 메시지 → 확인.
4. VBA 편집기 프로젝트 탐색기의 **모듈** 폴더에 15개 모듈이 있는지 확인.

## 5. 컴파일 확인
1. VBA 편집기 메뉴 **디버그 → VBAProject 컴파일**.
2. 아무 창도 안 뜨면 성공. (메뉴가 회색이 되면 이미 컴파일된 상태)
3. 오류가 나면 **메시지 창 + 표시된 줄**이 보이게 스크린샷을 찍어 보낸다.
4. 엑셀로 돌아가 **Ctrl+S** 저장.

## 6. 실행
1. 엑셀 시트를 한 번 클릭해 엑셀 창이 선택된 상태로 만든다. (다른 창이 앞에 있으면 키 입력을 무시한다)
2. **Alt+F8 → `StartGame` → 실행**.
3. 첫 화면을 칠하는 데 1~2초 걸린다. 메뉴가 나오면 ↑↓ 항목, ←→ 값, **A** 시작, **Esc** 종료.
4. 상태 표시줄(엑셀 맨 아래)에 FPS 등이 표시된다.
5. 엑셀이 입력을 안 받게 되면: Alt+F11 → **Ctrl+G** (직접 실행 창) → `GameLoop.RestoreExcel` 입력 → Enter.

## 이후 업데이트 받을 때
1. 엑셀은 열어 둔 채로 PowerShell에서 `git pull`.
2. 엑셀에서 **Alt+F8 → `ImportSrc`** (src에서 사라진 옛 모듈도 자동으로 지워진다).
3. **디버그 → VBAProject 컴파일** → Ctrl+S → `StartGame`.
4. `DevTools.bas` 자체가 바뀐 경우에만 3단계를 다시 한다 (그때는 따로 알려줌).
