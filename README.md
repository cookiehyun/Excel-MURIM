# Excel Fighter

엑셀 VBA로 만드는 실시간 2D 격투 게임 (Windows 엑셀, `.xlsm`)

## 폴더 구조

```
excel-fighter/
  ExcelFighter.xlsm   <- 엑셀에서 직접 생성
  src/                <- VBA 모듈(.bas), git으로 관리
    GameLoop.bas      고정 타임스텝 루프, 세션 설정/복구
    KeyInput.bas      GetAsyncKeyState 입력 (held/pressed), 포커스 없으면 무시
    Renderer.bas      셀 픽셀 프레임버퍼, 바뀐 셀만 칠하기, 가로 런 병합
    Fighter.bas       1단계용 테스트 박스 (걷기/점프/앉기)
    DevTools.bas      src <-> 워크북 동기화 (Import/Export)
```

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

Alt+F8 → `StartGame`

| 키 | 동작 |
|---|---|
| ← → | 이동 |
| ↑ | 점프 |
| ↓ | 앉기 |
| A / S / Z / X | 약P / 강P / 약K / 강K (1단계에서는 박스 색만 바뀜) |
| Esc | 종료 |

- 상태 표시줄에 `FPS | ticks/s | paint calls/frame`이 표시된다.
- 엑셀이 입력을 받지 않게 되면 VBE 직접 실행 창(Ctrl+G)에 `GameLoop.RestoreExcel`을 입력한다.

## 단계

- [ ] ① 루프 + 입력 + 박스 이동 ← **현재**
- [ ] ② 캐릭터 그리기 (숨긴 시트의 스프라이트 데이터)
- [ ] ③ 상태 머신
- [ ] ④ 공격 판정 (히트박스/허트박스, 프레임 데이터)
- [ ] ⑤ CPU AI (3단계)
- [ ] ⑥ 라운드 시스템 (3판 2선승, 체력 바, 타이머)
