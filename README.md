# xExcel Viewer 2.5.13

**2.5.13:** LEFT FILES Tab-key focus fix. When the file list owns focus, `Tab` / `Shift+Tab` are now intentionally ignored so focus does not jump unpredictably through toolbar controls. Use `Ctrl+Right` or `Alt+2` to move to WORKBOOK; Tab remains Excel-style cell navigation only inside WORKBOOK.

**2.5.12:** RIGHT-pane stale-workbook cleanup. Selecting a folder or a non-Excel file on the LEFT now clears any previously displayed workbook immediately, cancels pending workbook loads, and returns the RIGHT pane to its blank placeholder so an old Excel image cannot remain on screen.

**2.5.11:** Search-to-file keyboard navigation update. After typing in the LEFT Search box, pressing **Up** or **Down** immediately applies the current filter, leaves the search box, focuses the filtered FILES list, and moves file selection so browsing can continue without reaching for the mouse.

이 버전은 **왼쪽의 수많은 Excel 파일을 빠르게 고르고, 오른쪽에서 그 Excel의 전체 내용을 직접 보고/편집/복사**하는 한 가지 업무 흐름에 집중합니다.

## 핵심 변화

- 기존 mDIR-P 2.26.15의 안정판 구조와 Excel Virtual Grid 코드를 기준으로 다시 단순화했습니다.
- 예전처럼 오른쪽에 `first 35 rows x 9 columns` Preview를 표시하지 않습니다.
- Excel 파일을 왼쪽에서 선택하면 **같은 창 오른쪽**에 전체 Workbook이 열립니다.
- 모든 Sheet 탭을 사용할 수 있습니다.
- 사용 영역 전체를 가상 그리드(Virtual Grid)로 스크롤합니다. 화면에 보이는 부분을 중심으로 그리므로 큰 파일도 전체 셀 UI를 한꺼번에 만들지 않습니다.
- `.xlsx/.xlsm/.xltx/.xltm`은 셀 편집과 저장을 지원합니다.
- `.xls`는 전체 열람/복사는 가능하지만 안전 때문에 읽기 전용입니다.
- Embedded image를 표시하고 선택한 이미지를 `Ctrl+C`로 Windows Clipboard에 복사할 수 있습니다.

## 화면 구성

**왼쪽:** 파일 브라우저
- `Ctrl+클릭` / `Shift+클릭` 멀티 선택
- **오른쪽 마우스 버튼을 누른 채 위/아래로 드래그**하면 시작 파일을 유지하면서 지나간 파일을 선택에 추가 (mDIR 방식)
- 빠르게 드래그해도 중간 행을 빠뜨리지 않고, 위/아래 가장자리에서는 자동 스크롤
- `Del`로 선택 항목을 Windows 휴지통으로 이동
- 폴더 이동 / Back / Up / Home / Drive 버튼
- Name / Ext / Size / Modified
- 이름 Filter
- Excel only 필터
- Excel 파일을 한 번 선택하면 약 0.28초 뒤 오른쪽에서 Workbook 로드

**오른쪽:** Full Excel Workspace
- Workbook의 모든 Sheet
- 전체 행/열 스크롤
- 셀/범위 선택 후 `Ctrl+C`
- `F2` 또는 셀 더블클릭으로 직접 편집
- 상단 Formula Bar에서도 값 편집
- `Ctrl+V`로 탭/행 구조를 유지한 범위 Paste
- `Delete`로 선택 범위 지우기
- `Ctrl+S` 저장
- Save As
- Embedded image 클릭 후 `Ctrl+C`
- 이미지 우클릭 → Copy Image / Save Image As
- Zoom 50%~200%
- Open in Excel 버튼

## 화면 테마

상단 오른쪽 `Theme:`에서 세 가지 모드를 선택할 수 있습니다.

- `Bright`: 기존 밝은 테마
- `Dark`: xExcel의 파일 목록, 툴바, 입력창, 탭, Link Manager, 상태바, 경로바와 Excel row/column header를 어둡게 표시
- `System`: Windows의 **앱 모드**를 따라 Bright/Dark를 자동 전환

선택한 테마는 `%USERPROFILE%\.xexcel-viewer.json`에 저장됩니다. `System`을 선택하면 프로그램 실행 중 Windows 테마 변경도 주기적으로 확인하여 반영합니다.

중요: **Excel Workbook의 실제 셀 색상과 서식은 원본을 우선합니다.** Dark 테마라고 해서 흰색 Excel 셀을 검게 뒤집지 않습니다. 이것은 오래된 buyer/spec sheet를 원본 MS Excel과 최대한 비슷하게 보여주기 위한 의도입니다.

## 저장 안전장치

원본에 처음 `Ctrl+S`로 덮어쓸 때 같은 폴더의 `xExcel_Backup` 폴더에 타임스탬프 백업을 한 번 자동 생성합니다.

예:

```text
D:\Buyer\010620-14 TJ.xlsx
D:\Buyer\xExcel_Backup\010620-14 TJ-20260917-153000.xlsx
```

오래된 `.xls` 파일은 xExcel Viewer에서 직접 저장하지 않습니다. 필요하면 `Open in Excel`로 편집하거나 새 `.xlsx`로 정리하는 것을 권장합니다.

## 설치

1. ZIP 전체 압축 해제
2. `install_xExcel.bat` 더블클릭
3. 바탕화면 `xExcel Viewer` 실행

기본 설치 위치:

```text
%LOCALAPPDATA%\xExcel Viewer
```

Python 3.11 이상이 필요합니다. 설치 후 프로그램은 `pythonw.exe`로 실행되므로 별도 Terminal 창이 뜨지 않습니다.

## 단축키

- `F2` / Double Click: 셀 편집
- `Ctrl+C`: 셀 범위 또는 선택된 embedded image 복사
- `Ctrl+V`: 셀 범위 붙여넣기
- `Delete`: 선택 셀 지우기
- `Ctrl+S`: 저장
- `Ctrl++ / Ctrl+- / Ctrl+0`: Zoom
- 왼쪽 파일 목록 `F5`: 새로고침

## 기준

- Base: mDIR-P 2.26.15 안정판
- xExcel Viewer: 2.5.13
- 목표: File manager가 아니라 **대량의 오래된 Excel buyer/work-sheet를 빠르게 열고, 전체 내용/이미지를 찾아 다른 자료로 정리하는 전용 작업 프로그램**


## 2.1 keyboard workflow

- **Mouse click**: click the left file list or the right workbook to activate that pane.
- **Ctrl+Left**: go directly to the left FILES pane.
- **Ctrl+Right**: go directly to the right WORKBOOK pane.
- **Alt+1**: secondary shortcut for the left FILES pane.
- **Alt+2**: secondary shortcut for the right WORKBOOK pane.
- **F6**: legacy pane-toggle shortcut kept for compatibility.
- **Tab / Shift+Tab (LEFT FILES)**: disabled; focus stays in the file list. Use `Ctrl+Right` or `Alt+2` to move to WORKBOOK.
- **Tab / Shift+Tab (WORKBOOK)**: move to the next / previous Excel cell; Tab is not used to switch panes.
- **Left FILES pane**: Up/Down selects the previous/next file. Home/End and PageUp/PageDown are supported.
- **Right WORKBOOK pane**: arrow keys move the active cell. Shift+arrow extends the selection.
- **Sheets**: every worksheet is listed in the horizontal Sheets bar. Click any sheet name, use the arrow buttons, or press Ctrl+PageUp/Ctrl+PageDown.

Workbook loading is deliberately prevented from stealing focus from the left pane, so you can keep moving through many files with the keyboard while their contents load on the right.


## 2.2.0 Excel-layout fidelity update

오른쪽 WORKBOOK 패널이 원본 Excel과 다르게 뒤엉켜 보이던 가장 큰 원인은 **숨김 행/열을 화면에서 실제로 숨기지 않고 모두 폭/높이를 주어 그렸던 것**입니다. 이 버전은 Excel의 저장된 레이아웃 정보를 더 충실히 반영합니다.

- 숨김 Row / Column은 0px로 처리하고 행/열 Header에서도 건너뜁니다.
- Column width / Row height / 기본 크기를 Excel 값에 맞춰 계산합니다.
- `Show Gridlines` 설정을 반영합니다. 원본 Excel에서 Gridline이 꺼져 있으면 xExcel에서도 빈 셀 격자를 강제로 그리지 않습니다.
- Cell border의 thin / medium / thick / double을 표시합니다.
- Font name / size / bold / italic / underline, Fill color, horizontal/vertical alignment, indent, wrap을 반영합니다.
- Excel의 기본 세로 정렬(bottom)과 숫자의 기본 오른쪽 정렬을 반영합니다.
- Theme / indexed color를 가능한 범위에서 해석합니다.
- Embedded image의 anchor offset 및 one-cell/two-cell anchor 크기를 반영합니다.
- 방향키와 Tab 이동 시 숨김 행/열을 건너뜁니다.
- 일반 텍스트는 빈 인접 셀 쪽으로 자연스럽게 보이도록 cell fill과 text를 분리 렌더링합니다.

이 렌더러는 Excel 자체를 내장한 것은 아니므로 모든 Office 기능을 100% pixel-identical하게 재현하지는 않습니다. 하지만 오래된 buyer/spec worksheet에서 흔한 **숨김 행·열, 병합 셀, 이미지, 테이블 border, 사용자 지정 폭/높이** 때문에 레이아웃이 무너지는 문제를 우선 해결하는 버전입니다.

## mDIR-compatible Link Manager (2.4.0)

The top shortcut bar and **Edit Links** window now use the same link model as mDIR. Links are shown in their saved order and can be **Folder, File, Program, Web, Action, or Command**. Each link also stores a target pane (`active`, `left/files`, or `right/workbook`) and optional JSON arguments.

The Link Manager uses the mDIR layout: a table with **Name / Type / URL-Path-Action / Pane**, then editable **Name, Type/Pane, Target, Arguments** fields, followed by **Add, Remove, Move Up, Move Down, Browse File, Browse Folder, Save, Cancel**. An **Import mDIR** button copies the complete mDIR link list, including non-folder items such as GitHub, program launchers, and `powershell_here`.

Links are stored in `%USERPROFILE%\.xexcel-viewer-links.json` using the mDIR-compatible JSON schema. If xExcel 2.3.x contains only the old folder-only list and it matches the folder subset of `%USERPROFILE%\.mdir-p-shortcuts.json`, 2.4.0 and later restore the complete mDIR link list on load.

Supported placeholders include `{home}`, `{current}`, `{selected}`, `{left}`, `{right}`, `{left_selected}`, `{right_selected}`, and `{project}`. The mDIR-style green path bar remains clickable by directory name, and `Ctrl+L` still opens direct path entry.

### Left-pane filename search
Use the small `▼` button at the right edge of the green path bar to reopen recently visited folders. xExcel keeps up to 30 unique directories in most-recent-first order in `%USERPROFILE%\.xexcel-viewer.json`; the list survives restarts. Choose any path to jump there, or select **Clear Folder History** to remove previous entries while keeping the current folder.

Use the `Search:` box above the file list to filter the current folder instantly. Matching is case-insensitive. Separate terms with spaces to require all terms, e.g. `010719 TJ`. `Ctrl+F` focuses the search box, `Enter` selects the first visible result, and `Esc` or the `×` button clears the filter. xExcel caches the current directory listing so typing does not rescan a large folder for every character.

## LEFT-pane multi-selection and Delete

- Click: select one item.
- Ctrl+click: add/remove individual items.
- Shift+click: select a continuous range.
- Del: move the complete selection to the Recycle Bin after confirmation.
- Enter: open a single selected file with its Windows-associated application; folders navigate into the folder.

Deletion is intentionally recycle-bin only; xExcel does not silently switch to permanent deletion when recycling fails.


## GitHub / 개발 검사

저장소에는 Windows GitHub Actions CI가 포함되어 있습니다. Push 또는 Pull Request 시 Python 3.11 / 3.12 / 3.13에서 다음을 자동 검사합니다.

```text
compileall
unittest
python -m mdir --check
python -m build
clean release ZIP packaging
```

로컬 최종 검사:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m mdir --check
.\.venv\Scripts\python.exe -m build
.\.venv\Scripts\python.exe tools\package_release.py
```

`tools/package_release.py`는 `__pycache__`, build/dist, virtual environment, 임시 파일을 제외한 설치용 ZIP과 `SHA256SUMS.txt`를 만듭니다.
