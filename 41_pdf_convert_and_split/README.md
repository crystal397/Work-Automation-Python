# 폴더 PDF 변환 + 분할 도구 — 배포판 빌드 가이드

폴더 안의 Word/Excel/PowerPoint/한글/이미지/텍스트 파일을 모두 PDF로 변환하고,
200MB가 넘는 PDF는 200MB 이하 단위로 자동 분할하는 Windows용 도구입니다.

배포 형태: **단일 .exe 파일** (PyInstaller `--onefile`)
지원 사용 방식: GUI · 드래그&드롭 · 우클릭 메뉴 · CLI

---

## 빠른 시작 (사용자 입장)

배포된 ZIP 파일에는 다음 3개가 들어 있습니다:

```
convert_and_split.exe               ← 본체
register_context_menu.bat           ← 우클릭 메뉴 등록 (선택)
unregister_context_menu.bat         ← 우클릭 메뉴 제거 (선택)
```

원하는 방식으로 사용하세요:

| 방법 | 동작 |
|------|------|
| **더블클릭** | GUI 창이 뜸 → "찾아보기"로 폴더 선택 또는 폴더 드래그&드롭 |
| **드래그&드롭** | EXE 위에 폴더를 끌어다 놓으면 그 폴더로 자동 채워진 GUI가 뜸 |
| **우클릭 메뉴** | `register_context_menu.bat` 실행 후, 탐색기에서 폴더 우클릭 |
| **CLI** | `convert_and_split.exe --cli "C:\folder"` |

---

## 빌드 방법 (개발자 입장)

### 1) 사전 준비

- **Python 3.10 이상** 이 PATH에 설치되어 있어야 합니다.
- 빌드는 **Windows에서만** 가능합니다 (pywin32, COM 사용).
- 다음 파일들을 같은 폴더에 둡니다:
  ```
  convert_and_split.py
  build.bat
  register_context_menu.bat
  unregister_context_menu.bat
  ```

### 2) 빌드

`build.bat`를 더블클릭하거나 cmd에서 실행하세요.

```cmd
build.bat
```

이 배치 파일이 자동으로 처리합니다:
1. `pyinstaller`, `pywin32`, `pypdf`, `pillow`, `reportlab`, `tkinterdnd2` 설치
2. 이전 빌드 산출물(`build/`, `dist/`, `.spec`) 정리
3. PyInstaller로 단일 EXE 빌드

빌드가 끝나면 **`dist\convert_and_split.exe`** 파일이 생깁니다.

### 3) 배포 패키지 만들기

`dist\convert_and_split.exe` 와 두 개의 `.bat` 파일을 한 폴더에 모아
ZIP으로 압축하면 그게 배포판입니다.

```
배포판.zip
├── convert_and_split.exe
├── register_context_menu.bat
└── unregister_context_menu.bat
```

---

## 환경별 요구사항 (실행하는 PC에 설치되어 있어야 함)

| 기능 | 필요한 외부 프로그램 |
|------|---------------------|
| Word(.doc/.docx) 변환 | Microsoft Word |
| Excel(.xls/.xlsx) 변환 | Microsoft Excel |
| PowerPoint(.ppt/.pptx) 변환 | Microsoft PowerPoint |
| 한글(.hwp/.hwpx) 변환 | 한컴오피스 (HWP COM 자동화) |
| 이미지 변환 | (EXE에 내장됨) |
| 텍스트 변환 | (EXE에 내장됨) |
| PDF 분할 | (EXE에 내장됨) |

> Office가 설치되지 않은 PC에서는 해당 형식만 변환 실패로 표시되고,
> 나머지 작업은 정상 진행됩니다.

---

## 우클릭 메뉴 등록 (선택사항)

`register_context_menu.bat`를 실행하면 탐색기 폴더 우클릭 메뉴에
**"PDF 변환 + 200MB 분할"** 항목이 추가됩니다.

- **관리자 권한 불필요** — 현재 사용자(`HKEY_CURRENT_USER`)에만 등록됩니다.
- EXE가 이동/삭제되면 메뉴는 동작하지 않습니다. EXE 위치를 옮긴 후엔
  재등록(`register_context_menu.bat` 다시 실행)이 필요합니다.
- **Windows 11**에서는 우클릭 후 [추가 옵션 표시](또는 Shift+우클릭)를
  눌러야 보일 수 있습니다.
- 제거는 `unregister_context_menu.bat` 로 합니다.

---

## 알려진 사항 / 주의

- **Windows Defender / SmartScreen 경고**
  PyInstaller로 만든 서명되지 않은 EXE는 처음 실행 시 SmartScreen이
  경고할 수 있습니다. "추가 정보 → 실행"으로 통과 가능합니다.
  배포 규모가 커지면 코드 사이닝 인증서로 서명하는 것이 좋습니다.

- **PowerPoint 변환 중 창 깜빡임**
  COM 특성상 PPT는 완전한 백그라운드 실행이 어렵습니다.

- **EXE 크기**
  reportlab, pillow, pywin32까지 포함되어 **약 30~50MB** 정도가 됩니다.
  더 작게 만들고 싶다면 `--exclude-module` 로 안 쓰는 모듈을 제외하세요.

- **첫 실행 속도**
  `--onefile`은 임시 폴더에 압축을 푸는 방식이라 첫 실행이 2~5초 느립니다.
  속도가 중요하면 `build.bat`에서 `--onefile`을 빼고 `--onedir`로 바꾸세요
  (대신 EXE 옆에 라이브러리 폴더가 같이 생깁니다).

- **드래그&드롭 미지원 환경**
  `tkinterdnd2`는 빌드 시 자동 포함되지만, 만약 실패하면 GUI 우상단에
  "드래그&드롭 비활성" 표시가 뜨고 "찾아보기" 버튼만 사용 가능합니다.

---

## 트러블슈팅

**Q. EXE 실행 시 아무 반응 없음**
A. `--windowed` 옵션 때문에 오류 메시지가 콘솔에 안 뜹니다.
   `build.bat`에서 `--windowed`를 임시로 빼고 다시 빌드해 콘솔에서
   에러를 확인하세요.

**Q. "Word.Application을 만들 수 없습니다" 오류**
A. 해당 PC에 Microsoft Word가 설치되어 있지 않거나,
   COM 등록이 깨진 상태입니다. Office 복구 설치를 시도해 보세요.

**Q. 한글 변환 시 보안 모듈 경고**
A. 한컴오피스 보안 정책 때문입니다. 한컴오피스 설치 폴더에
   `FilePathCheckerModule.dll` 을 두면 경고 없이 자동화됩니다.

**Q. 분할 결과가 200MB를 약간 넘음**
A. PDF 압축 특성상 페이지 단위 측정값과 실제 저장 크기에 미세한 오차가
   생길 수 있습니다. 보수적으로 자르려면 `convert_and_split.py`의
   `SIZE_LIMIT` 값을 `195 * 1024 * 1024` 등으로 줄이세요.
