# 00. COM 기본설정 보일러플레이트

HWP COM 객체 생성·종료, 메시지박스 자동 처리, 새 문서 만들기. 모든 HWP 자동화 스크립트의 첫 머리에 들어가는 정형 코드.

---

## 왜 이 boilerplate가 필요한가

| 라인 | 안 쓰면 무슨 일이 | 왜 |
|---|---|---|
| `EnsureDispatch("HWPFrame.HwpObject")` | COM 객체 생성 자체가 안 됨 | `Dispatch` 대신 `gencache.EnsureDispatch`를 써야 IntelliSense·early-binding이 적용되어 `HParameterSet.HCharShape` 같은 속성이 노출됨 |
| `RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")` | `InsertPicture`나 `SaveAs` 호출 시 "외부 파일 접근 권한" 보안 경고가 떠서 스크립트가 멈춤 | Hancom 보안모듈을 사전 등록해 경고를 끔. **반드시 try/except로 감쌀 것** — 일부 환경에선 모듈명이 다르거나 등록 실패해도 동작에는 지장 없음 |
| `SetMessageBoxMode(0x2FF1)` | 저장 덮어쓰기·암호 입력 등 모든 MessageBox에서 멈춤 | 모든 메시지박스를 자동으로 기본값 처리 |
| `XHwpWindows.Item(0).Visible = False` | 한글 창이 화면에 떠서 깜빡임 | 헤드리스 동작 |
| `hwp.HAction.Run("FileNew")` | 빈 문서가 없어서 첫 `InsertText`가 실패 | 새 빈 문서 생성 |

---

## 표준 헤더 (그대로 복사)

```python
# -*- coding: utf-8 -*-
import os
import win32com.client as win32

# ── 경로 (raw string 필수: 백슬래시 이스케이프 함정 회피) ──
BASE = r"C:\Users\<USER>\Desktop\<폴더>"
OUT_HWP = os.path.join(BASE, "결과물.hwp")
OUT_PDF = os.path.join(BASE, "결과물.pdf")

# ── COM 초기화 ──
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
try:
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
except Exception:
    pass  # 일부 환경에서 모듈 미설치 — 무시 가능
hwp.XHwpWindows.Item(0).Visible = False
hwp.SetMessageBoxMode(0x2FF1)  # 모든 메시지박스 자동 처리
hwp.HAction.Run("FileNew")
```

---

## 표준 푸터 (그대로 복사)

```python
# ── 저장 ──
hwp.SaveAs(OUT_HWP, "HWP", "")
print("Saved:", OUT_HWP)

# ── (선택) PDF 변환으로 시각 검증 — 04-페이지와-검증.md 참고 ──
hwp.HAction.GetDefault("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.HParameterSet.HFileOpenSave.filename = OUT_PDF
hwp.HParameterSet.HFileOpenSave.Format = "PDF"
hwp.HAction.Execute("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
print("PDF:", OUT_PDF)

# ── 종료 (반드시 호출 — 안 부르면 hwp.exe 백그라운드에 남음) ──
hwp.Quit()
```

---

## 자주 막히는 지점

- **경로 한글**: 한글 폴더명 자체는 문제 없으나, 경로에 줄바꿈/탭이 들어가면 `SaveAs`가 silently 실패. `os.path.join`을 일관되게 사용.
- **`Visible = False`인데 창이 잠깐 뜸**: 정상 — Dispatch 직후 `Visible` 적용까지 한 frame 지연. 사용자에게는 보이지만 문제는 없음.
- **`Quit` 호출 후에도 `hwp.exe`가 살아있음**: 다른 win32com 객체가 같은 프로세스를 잡고 있는 경우. `del hwp; gc.collect()`까지 해야 안전.
- **`SetMessageBoxMode`의 `0x2FF1`**: bit 마스크. 상위 12bit가 어떤 메시지를 자동 처리할지 지정, 하위 4bit가 응답 종류. 이 값은 "모두 OK/예 응답"으로 검증된 상수 — 외우지 말고 그대로 사용.

---

## 다음 단계

- 본문을 쓰려면 → [01-글자-렌더링.md](./01-글자-렌더링.md)
- 헬퍼 함수만 복사해서 시작하려면 → [05-검증된-헬퍼함수.md](./05-검증된-헬퍼함수.md)
