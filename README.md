# Zoteley

**Zotero · Mendeley Cite · EndNote 원고 인용 변환** · [English below](#english)

Word 원고의 인용을 Zotero · Mendeley Cite · EndNote 형식 사이에서 바꿉니다(6방향).
서로 다른 문헌 관리 프로그램을 쓰는 공저자가 같은 원고에서 인용을 편집할 수 있게 하는 것이 목적입니다.

- 원본 파일은 바꾸지 않고 `원고_zotero.docx`, `원고_mendeley.docx`, `원고_endnote.docx` 사본을 만듭니다.
- 상대 라이브러리에 없는 문헌은 Mendeley에 추가하거나 Zotero에 넣을 수 있습니다. EndNote는 가져오기 파일(XML)을 만들어 줍니다.
- 화면에 보이는 인용 표기와 본문은 그대로 둡니다.

> Zotero, Mendeley, Elsevier, EndNote(Clarivate)와 관계없는 개인 프로젝트입니다. 각 이름은 해당 소유자의 상표입니다.

## 내려받기

[Releases](https://github.com/hspark-rm/zoteley/releases/latest)에서 내려받습니다.

| 컴퓨터 | 파일 |
|---|---|
| Mac (M1·M2·M3·M4) | `Zoteley-mac-apple-silicon.dmg` |
| Mac (Intel) | `Zoteley-mac-intel.dmg` |
| Windows 10·11 | `Zoteley-windows.zip` |

### 처음 열 때 보안 경고

코드 서명을 하지 않은 무료 앱이라 처음 한 번 경고가 뜹니다.

- **Mac**: dmg를 열고 Zoteley를 응용 프로그램 폴더로 끌어다 놓습니다. 앱을 열면 경고가 뜹니다 → **시스템 설정 > 개인정보 보호 및 보안** 아래쪽의 **그래도 열기**를 누릅니다.
- **Windows**: zip을 풀고 `Zoteley.exe`를 실행합니다. "Windows의 PC 보호" 창이 뜨면 **추가 정보 > 실행**을 누릅니다.

## 처음 설정 (한 번만)

1. **Mendeley 앱 등록** (무료, Mendeley 계정 필요)
   1. Zoteley의 설정에서 **개발자 사이트 열기**를 누릅니다(dev.mendeley.com).
   2. **Register a new app**에서 이름은 아무것이나, **Redirect URL**에는 `http://localhost:53682/callback`을 넣습니다.
   3. **Generate secret**을 누르고 약관에 동의한 뒤 등록합니다.
   4. 받은 **client ID**와 **secret**을 Zoteley 설정에 넣고 **저장하고 로그인**을 누릅니다.
2. **Zotero 데이터 폴더**: 보통 홈 폴더의 `Zotero`라 자동으로 찾습니다. 다르면 Zotero 설정 > 고급 > 파일 및 폴더의 위치를 고릅니다.
3. **EndNote 라이브러리 (선택)**: EndNote로 바꿀 때 내 라이브러리 문헌과 연결하려면 `.enl` 파일을 고릅니다. 기본값은 문서 폴더의 `My EndNote Library.enl`입니다.

Mendeley 설정은 Mendeley로 바꿀 때만 필요합니다.

## 쓰는 법

1. Word 원고(.docx)를 Zoteley 창에 끌어다 놓습니다.
2. 바꿀 형식의 버튼(**Zotero 원고로 / Mendeley 원고로 / EndNote 원고로**)을 누릅니다. 원고의 지금 형식은 자동으로 알아냅니다.
3. 끝나면 **Word로 열기**를 누릅니다.
   - Mendeley 원고: Word의 Mendeley Cite에서 참고문헌을 한 번 새로고침합니다.
   - Zotero 원고: Word의 Zotero 탭에서 **Refresh**를 한 번 누릅니다.
   - EndNote 원고: Word의 EndNote 탭에서 **Update Citations and Bibliography**를 한 번 누릅니다.

변환하기 전에 Word에서 원고를 닫고, 변경 내용 추적은 모두 수락해 두세요.

## 한계

- 각주 안의 인용은 바꾸지 않습니다.
- 인용 표기를 손으로 고친 경우 새로고침하면 원래 표기로 돌아갈 수 있습니다.
- Zotero에 없는 문헌을 넣을 때는 Zotero가 실행 중이어야 하며, Zotero에서 지금 선택된 컬렉션에 들어갑니다.

## 개인정보

- 원고와 라이브러리는 내 컴퓨터에서만 처리합니다.
- 네트워크는 Mendeley 공식 API(`api.mendeley.com`)와 내 컴퓨터의 Zotero에만 연결합니다. EndNote 라이브러리는 사본을 읽기만 합니다.
- Mendeley 키와 로그인 정보는 내 컴퓨터의 사용자 폴더에만 저장합니다.

---

<a id="english"></a>
## English

Zoteley converts citations in Word manuscripts between **Zotero**, **Mendeley Cite** and **EndNote** (all six directions), so co-authors using different reference managers can edit citations in the same manuscript.

- The original file is never modified; a copy (`*_zotero.docx`, `*_mendeley.docx` or `*_endnote.docx`) is created.
- References missing on the other side can be added to Mendeley or imported into Zotero.
- Visible citation text and body text are preserved.

> Independent project, not affiliated with Zotero, Mendeley, Elsevier, or EndNote (Clarivate). All names are trademarks of their owners.

### Download

Get the latest build from [Releases](https://github.com/hspark-rm/zoteley/releases/latest): `Zoteley-mac-apple-silicon.dmg`, `Zoteley-mac-intel.dmg`, or `Zoteley-windows.zip`.

The app is not code-signed, so you will see a warning the first time:
- **macOS**: open the app once, then go to **System Settings > Privacy & Security** and click **Open Anyway**.
- **Windows**: on "Windows protected your PC", click **More info > Run anyway**.

### First-time setup

1. **Register a Mendeley app** (free): in Zoteley's Settings click **Open developer site** → **Register a new app** → set **Redirect URL** to `http://localhost:53682/callback` → **Generate secret** → register. Enter the client ID and secret in Zoteley and click **Save and sign in**.
2. **Zotero data folder**: usually `~/Zotero` and detected automatically.

### Usage

Drop a `.docx` onto the window and click **To Zotero**, **To Mendeley** or **To EndNote**; the current format is detected automatically. Open the result in Word and refresh once: Mendeley Cite (refresh bibliography), Zotero (**Refresh**), EndNote (**Update Citations and Bibliography**). Mendeley setup is only needed when converting to Mendeley; an EndNote library (`.enl`) is optional and links citations to your records.

### Limitations

Citations inside footnotes are not converted. Manually edited citation text may be reset when refreshed. Importing into Zotero requires Zotero to be running.

### Command line

`zms.py` provides the same functions from a terminal (`python3 zms.py -h`).

### License

MIT
