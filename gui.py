"""Zoteley — Zotero ↔ Mendeley Cite 원고 변환 데스크톱 앱 (macOS · Windows).

화면만 맡는다. 변환·인증·API 는 zms.py 가 하고, 이 파일은 zms.run([...]) 의
종료 코드와 KEY=값 출력만 읽는다.
"""
import os
import shutil
import subprocess
import sys
import threading
import queue
from pathlib import Path

APP = "Zoteley"


def app_dir():
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / APP
    if os.name == "nt":
        return Path(os.environ.get("APPDATA") or Path.home()) / APP
    return Path.home() / ".config" / "zoteley"


# zms 는 가져올 때 설정·작업 폴더를 정하므로, 가져오기 전에 OS별 폴더를 알려 준다.
APP_DIR = app_dir()
APP_DIR.mkdir(parents=True, exist_ok=True)
os.environ.setdefault("ZMS_CONF", str(APP_DIR))
os.environ.setdefault("ZMS_DATA", str(APP_DIR / "data"))
_old = Path.home() / ".config" / "zotero-mendeley-sync"
if not (APP_DIR / "config.json").exists() and (_old / "config.json").exists():
    # 명령줄 도구로 이미 설정했다면 그대로 가져온다
    for n in ("config.json", "token.json"):
        if (_old / n).exists():
            shutil.copy2(_old / n, APP_DIR / n)

import zms  # noqa: E402
from PySide6.QtCore import QEvent, QLocale, QObject, Qt, QUrl, Signal  # noqa: E402
from PySide6.QtGui import QDesktopServices, QFont, QGuiApplication, QIcon, QPixmap  # noqa: E402
from PySide6.QtWidgets import (  # noqa: E402
    QApplication, QComboBox, QDialog, QFileDialog, QFormLayout, QFrame, QGroupBox, QHBoxLayout, QLabel,
    QLineEdit, QMessageBox, QPushButton, QVBoxLayout, QWidget)

BASE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
ASSETS = BASE / "assets"
DEV_SITE = "https://dev.mendeley.com/myapps.html"
HELP_URL = "https://github.com/hspark-rm/zoteley#readme"

# ---------- 문구 (한국어 · 영어) ----------

KO = QLocale.system().name().startswith("ko")
T = {
    "subtitle": ("Zotero ↔ Mendeley Cite 원고 변환", "Convert Word manuscripts between Zotero and Mendeley Cite"),
    "drop": ("Word 원고(.docx)를 여기에 끌어다 놓으세요", "Drop a Word manuscript (.docx) here"),
    "choose": ("파일 선택…", "Choose file…"),
    "send": ("Mendeley로 보내기", "Send to Mendeley"),
    "send_sub": ("Zotero 원고 → Mendeley 원고", "Zotero → Mendeley Cite"),
    "recv": ("Zotero로 받기", "Receive into Zotero"),
    "recv_sub": ("Mendeley 원고 → Zotero 원고", "Mendeley Cite → Zotero"),
    "settings": ("설정", "Settings"),
    "help": ("도움말", "Help"),
    "kind_zotero": ("Zotero 인용 {n}개", "{n} Zotero citations"),
    "kind_mendeley": ("Mendeley 인용 {n}개", "{n} Mendeley citations"),
    "kind_both": ("Zotero 인용 {z}개 · Mendeley 인용 {m}개", "{z} Zotero · {m} Mendeley citations"),
    "kind_none": ("인용이 없는 원고입니다", "No citations found in this manuscript"),
    "only_docx": ("Word 원고(.docx)만 바꿀 수 있습니다.", "Only Word manuscripts (.docx) can be converted."),
    "working_send": ("Mendeley 라이브러리를 확인하고 바꾸는 중…", "Checking your Mendeley library and converting…"),
    "working_recv": ("Zotero 라이브러리와 맞춰 보는 중…", "Matching against your Zotero library…"),
    "working_import": ("Zotero에 문헌을 넣는 중…", "Adding references to Zotero…"),
    "done_send": ("인용 {n}개를 Mendeley 형식으로 바꿨습니다.", "Converted {n} citations to Mendeley format."),
    "done_send_added": ("Mendeley에 없던 문헌 {n}건을 라이브러리에 추가했습니다.",
                        "Added {n} references that were missing from Mendeley."),
    "done_send_tip": ("Word에서 Mendeley Cite를 열고 참고문헌을 한 번 새로고침하세요.",
                      "Open Mendeley Cite in Word and refresh the bibliography once."),
    "done_recv": ("인용 {n}개를 Zotero 형식으로 바꿨습니다.", "Converted {n} citations to Zotero format."),
    "done_recv_unl": ("Zotero에 없는 문헌 {n}건은 문서 안 정보로 넣었습니다.",
                      "{n} references not in Zotero were kept as embedded data."),
    "done_recv_tip": ("Word의 Zotero 메뉴에서 Refresh를 한 번 누르세요.", "Click Refresh in Word's Zotero tab once."),
    "new_file": ("새 파일: {f}", "New file: {f}"),
    "open_word": ("Word로 열기", "Open in Word"),
    "reveal": ("폴더에서 보기", "Show in folder"),
    "ok": ("확인", "OK"),
    "cancel": ("취소", "Cancel"),
    "missing_q": ("Zotero에 없는 문헌이 {n}건 있습니다.\n\nZotero에 넣으면 이 문헌들도 Zotero에서 편집할 수 있습니다. "
                  "문헌은 Zotero에서 지금 선택된 컬렉션에 들어갑니다.",
                  "{n} cited references are not in your Zotero library.\n\nAdd them to Zotero so you can edit them "
                  "there? They go into the collection currently selected in Zotero."),
    "add_and_convert": ("Zotero에 넣고 바꾸기", "Add to Zotero and convert"),
    "convert_only": ("넣지 않고 바꾸기", "Convert without adding"),
    "zotero_off": ("Zotero가 열려 있어야 문헌을 넣을 수 있습니다. Zotero를 연 뒤 다시 시도하세요.",
                   "Zotero must be running to add references. Open Zotero and try again."),
    "retry": ("다시 시도", "Try again"),
    "error": ("작업 중 오류가 났습니다.", "Something went wrong."),
    "need_login": ("Mendeley 로그인이 필요합니다.", "You need to sign in to Mendeley."),
    "need_setup": ("처음 한 번 설정이 필요합니다.", "First-time setup is needed."),
    "status_ok": ("Mendeley 연결됨 · Zotero: {z}", "Mendeley connected · Zotero: {z}"),
    "status_nologin": ("Mendeley 로그인 필요 · Zotero: {z}", "Mendeley sign-in needed · Zotero: {z}"),
    "status_nokeys": ("Mendeley 설정 필요", "Mendeley setup needed"),
    # 설정 창
    "s_title": ("설정", "Settings"),
    "s_mendeley": ("1. Mendeley 연결", "1. Connect Mendeley"),
    "s_guide": ("Mendeley 개발자 사이트에서 앱을 한 번 등록해야 합니다(무료).\n"
                "① '개발자 사이트 열기' → Register a new app\n"
                "② Redirect URL에 아래 주소를 그대로 넣기\n"
                "③ Generate secret → 약관 동의 → 등록\n"
                "④ 받은 client ID와 secret을 아래에 넣고 '저장하고 로그인'",
                "Register a (free) app on the Mendeley developer site once.\n"
                "① 'Open developer site' → Register a new app\n"
                "② Paste the Redirect URL below exactly\n"
                "③ Generate secret → accept terms → register\n"
                "④ Enter the client ID and secret below and click 'Save and sign in'"),
    "s_open_dev": ("개발자 사이트 열기", "Open developer site"),
    "s_copy": ("복사", "Copy"),
    "s_copied": ("복사함", "Copied"),
    "s_cid": ("client ID", "Client ID"),
    "s_secret": ("client secret", "Client secret"),
    "s_secret_keep": ("바꾸지 않으려면 비워 두세요", "Leave empty to keep the current one"),
    "s_save_login": ("저장하고 로그인", "Save and sign in"),
    "s_login": ("다시 로그인", "Sign in again"),
    "s_zotero": ("2. Zotero 데이터 폴더", "2. Zotero data folder"),
    "s_zotero_tip": ("Zotero 설정 > 고급 > 파일 및 폴더의 '데이터 디렉터리 위치'", "Zotero Settings > Advanced > Files and Folders > Data Directory Location"),
    "s_browse": ("찾아보기…", "Browse…"),
    "s_target": ("3. Mendeley로 보낼 때 문헌을 넣을 곳", "3. Where to add references when sending to Mendeley"),
    "s_personal": ("내 개인 라이브러리", "My personal library"),
    "s_load_groups": ("그룹 불러오기", "Load groups"),
    "s_close": ("닫기", "Close"),
    "s_state_ok": ("로그인됨", "Signed in"),
    "s_state_no": ("로그인 안 됨", "Not signed in"),
    "s_need_cid": ("client ID를 입력하세요.", "Enter the client ID."),
    "s_need_secret": ("client secret을 입력하세요.", "Enter the client secret."),
    "s_bad_zotero": ("이 폴더에 zotero.sqlite가 없습니다.", "zotero.sqlite was not found in this folder."),
    # 로그인 창
    "l_title": ("Mendeley 로그인", "Sign in to Mendeley"),
    "l_text": ("브라우저에서 Mendeley 로그인 창이 열렸습니다.\n로그인하고 '허용'을 누르면 자동으로 이어집니다.",
               "A Mendeley sign-in page opened in your browser.\nSign in and click 'Allow' — this window continues automatically."),
    "l_paste": ("브라우저에 오류 화면이 남으면, 그 창의 주소 전체를 복사해 붙여 넣으세요:",
                "If the browser shows an error page, copy its full address and paste it here:"),
    "l_apply": ("적용", "Apply"),
}


def tr(key, **kw):
    s = T[key][0 if KO else 1]
    return s.format(**kw) if kw else s


def value(out, key):
    for line in out.splitlines():
        if line.startswith(key + "="):
            return line[len(key) + 1:]
    return ""


def last_line(out):
    lines = [x for x in out.strip().splitlines() if x.strip()]
    return lines[-1] if lines else ""


# ---------- 백그라운드 실행 ----------

class Runner(QObject):
    """zms 명령을 작업 스레드에서 돌리고, 끝나면 (코드, 출력) 을 화면 스레드로 넘긴다."""
    done = Signal(int, str)

    def start(self, argv):
        threading.Thread(target=lambda: self.done.emit(*zms.run(argv)), daemon=True).start()


def run_blocking(parent, argv, message=None):
    """짧은 작업을 돌리는 동안 커서를 바꾸고 이벤트를 처리하며 기다린다."""
    result = {}
    ev = threading.Event()

    def work():
        result["r"] = zms.run(argv)
        ev.set()

    QGuiApplication.setOverrideCursor(Qt.WaitCursor)
    if message and parent is not None and hasattr(parent, "set_busy"):
        parent.set_busy(message)
    threading.Thread(target=work, daemon=True).start()
    while not ev.wait(0.03):
        QApplication.processEvents()
    QGuiApplication.restoreOverrideCursor()
    if message and parent is not None and hasattr(parent, "set_busy"):
        parent.set_busy(None)
    return result["r"]


def reveal(path):
    if sys.platform == "darwin":
        subprocess.run(["open", "-R", path])
    elif os.name == "nt":
        subprocess.run(["explorer", "/select,", os.path.normpath(path)])
    else:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(path).parent)))


def show_error(parent, code, out):
    box = QMessageBox(parent)
    box.setIcon(QMessageBox.Warning)
    box.setWindowTitle(APP)
    box.setText(tr("error"))
    box.setInformativeText(last_line(out))
    box.setDetailedText(out)
    box.exec()


# ---------- 로그인 창 ----------

class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("l_title"))
        self.setMinimumWidth(520)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(tr("l_text")))
        hint = QLabel(tr("l_paste"))
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray;")
        lay.addWidget(hint)
        row = QHBoxLayout()
        self.paste = QLineEdit()
        self.paste.setPlaceholderText("http://localhost:53682/callback?code=…")
        apply_btn = QPushButton(tr("l_apply"))
        apply_btn.clicked.connect(self.apply_paste)
        row.addWidget(self.paste)
        row.addWidget(apply_btn)
        lay.addLayout(row)
        cancel = QPushButton(tr("cancel"))
        cancel.clicked.connect(self.cancel)
        lay.addWidget(cancel, alignment=Qt.AlignRight)
        self.result_code, self.result_out = None, ""
        zms.PASTE_QUEUE = queue.Queue()
        zms.CANCEL.clear()
        self.runner = Runner()
        self.runner.done.connect(self.finished_auth)
        self.runner.start(["auth", "--timeout", "300"])

    def apply_paste(self):
        if self.paste.text().strip():
            zms.PASTE_QUEUE.put(self.paste.text().strip())

    def cancel(self):
        zms.CANCEL.set()

    def reject(self):  # Esc·창 닫기도 취소로 처리해 작업 스레드를 끝낸다
        zms.CANCEL.set()

    def finished_auth(self, code, out):
        zms.PASTE_QUEUE = None
        self.result_code, self.result_out = code, out
        self.done(0)


def login(parent):
    dlg = LoginDialog(parent)
    dlg.exec()
    if dlg.result_code == 0:
        return True
    if dlg.result_code not in (zms.EXIT_CANCEL, None):
        show_error(parent, dlg.result_code, dlg.result_out)
    return False


# ---------- 설정 창 ----------

class SettingsDialog(QDialog):
    def __init__(self, parent=None, welcome=False):
        super().__init__(parent)
        self.setWindowTitle(tr("s_title"))
        self.setMinimumWidth(560)
        lay = QVBoxLayout(self)
        if welcome:
            w = QLabel(tr("need_setup"))
            w.setStyleSheet("font-weight: 600;")
            lay.addWidget(w)
        conf = zms.conf_raw()

        g1 = QGroupBox(tr("s_mendeley"))
        f1 = QVBoxLayout(g1)
        guide = QLabel(tr("s_guide"))
        guide.setWordWrap(True)
        f1.addWidget(guide)
        row = QHBoxLayout()
        self.redirect = QLineEdit(conf.get("redirect_uri") or zms.REDIRECT_DEFAULT)
        self.redirect.setReadOnly(True)
        copy = QPushButton(tr("s_copy"))
        copy.clicked.connect(lambda: (QGuiApplication.clipboard().setText(self.redirect.text()),
                                      copy.setText(tr("s_copied"))))
        dev = QPushButton(tr("s_open_dev"))
        dev.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(DEV_SITE)))
        row.addWidget(QLabel("Redirect URL"))
        row.addWidget(self.redirect)
        row.addWidget(copy)
        f1.addLayout(row)
        f1.addWidget(dev, alignment=Qt.AlignLeft)
        form = QFormLayout()
        self.cid = QLineEdit(conf.get("client_id") or "")
        self.secret = QLineEdit()
        self.secret.setEchoMode(QLineEdit.Password)
        if conf.get("client_secret"):
            self.secret.setPlaceholderText(tr("s_secret_keep"))
        form.addRow(tr("s_cid"), self.cid)
        form.addRow(tr("s_secret"), self.secret)
        f1.addLayout(form)
        row2 = QHBoxLayout()
        self.login_state = QLabel()
        save = QPushButton(tr("s_save_login"))
        save.setDefault(True)
        save.clicked.connect(self.save_and_login)
        row2.addWidget(self.login_state)
        row2.addStretch()
        row2.addWidget(save)
        f1.addLayout(row2)
        lay.addWidget(g1)

        g2 = QGroupBox(tr("s_zotero"))
        f2 = QVBoxLayout(g2)
        row3 = QHBoxLayout()
        self.zdir = QLineEdit(str(zms.zotero_db().parent))
        self.zdir.setReadOnly(True)
        browse = QPushButton(tr("s_browse"))
        browse.clicked.connect(self.choose_zotero)
        row3.addWidget(self.zdir)
        row3.addWidget(browse)
        f2.addLayout(row3)
        tip = QLabel(tr("s_zotero_tip"))
        tip.setStyleSheet("color: gray;")
        f2.addWidget(tip)
        lay.addWidget(g2)

        g3 = QGroupBox(tr("s_target"))
        f3 = QHBoxLayout(g3)
        self.group = QComboBox()
        self.group.addItem(tr("s_personal"), "")
        if conf.get("group"):
            self.group.addItem(conf["group"], conf["group"])
            self.group.setCurrentIndex(1)
        self.group.currentIndexChanged.connect(self.save_group)
        load = QPushButton(tr("s_load_groups"))
        load.clicked.connect(self.load_groups)
        f3.addWidget(self.group, 1)
        f3.addWidget(load)
        lay.addWidget(g3)

        close = QPushButton(tr("s_close"))
        close.clicked.connect(self.accept)
        lay.addWidget(close, alignment=Qt.AlignRight)
        self.refresh_state()

    def refresh_state(self):
        ok = (zms.CONF_DIR / "token.json").exists()
        self.login_state.setText(("✓ " if ok else "• ") + tr("s_state_ok" if ok else "s_state_no"))
        self.login_state.setStyleSheet("color: #1a7f37;" if ok else "color: #b42318;")

    def save_and_login(self):
        cid = self.cid.text().strip()
        sec = self.secret.text().strip()
        if not cid:
            QMessageBox.warning(self, APP, tr("s_need_cid"))
            return
        if not sec and not zms.conf_raw().get("client_secret"):
            QMessageBox.warning(self, APP, tr("s_need_secret"))
            return
        # secret 은 명령줄이 아닌 프로세스 환경으로만 넘기고 바로 지운다
        if sec:
            os.environ["ZMS_CLIENT_SECRET"] = sec
        try:
            code, out = zms.run(["config", "--client-id", cid])
        finally:
            os.environ.pop("ZMS_CLIENT_SECRET", None)
        if code != 0:
            show_error(self, code, out)
            return
        self.secret.clear()
        login(self)
        self.refresh_state()

    def choose_zotero(self):
        d = QFileDialog.getExistingDirectory(self, tr("s_zotero"), self.zdir.text())
        if not d:
            return
        if not (Path(d) / "zotero.sqlite").exists():
            QMessageBox.warning(self, APP, tr("s_bad_zotero"))
            return
        code, out = zms.run(["config", "--zotero-dir", d])
        if code == 0:
            self.zdir.setText(d)
        else:
            show_error(self, code, out)

    def load_groups(self):
        code, out = run_blocking(self, ["groups"])
        if code == zms.EXIT_AUTH and login(self):
            code, out = run_blocking(self, ["groups"])
        if code != 0:
            show_error(self, code, out)
            return
        current = self.group.currentData()
        self.group.blockSignals(True)
        self.group.clear()
        self.group.addItem(tr("s_personal"), "")
        for line in out.splitlines():
            if line.startswith("GROUP=") and len(line) > 6:
                self.group.addItem(line[6:], line[6:])
        i = max(0, self.group.findData(current))
        self.group.setCurrentIndex(i)
        self.group.blockSignals(False)

    def save_group(self):
        zms.run(["config", "--group", self.group.currentData() or ""])


# ---------- 원고 놓는 칸 ----------

class DropZone(QFrame):
    def __init__(self, on_file):
        super().__init__()
        self.on_file = on_file
        self.setAcceptDrops(True)
        self.setMinimumHeight(220)
        self.setObjectName("drop")
        self.normal_style = ("#drop { border: 2px dashed #c9ced6; border-radius: 16px; background: palette(base); }")
        self.hover_style = ("#drop { border: 2px dashed #2f6feb; border-radius: 16px; background: rgba(47,111,235,0.06); }")
        self.setStyleSheet(self.normal_style)
        lay = QVBoxLayout(self)
        lay.setAlignment(Qt.AlignCenter)
        art = ASSETS / "dropzone.png"
        if art.exists():
            pic = QLabel()
            pic.setPixmap(QPixmap(str(art)).scaledToHeight(110, Qt.SmoothTransformation))
            pic.setAlignment(Qt.AlignCenter)
            lay.addWidget(pic)
        self.label = QLabel(tr("drop"))
        self.label.setAlignment(Qt.AlignCenter)
        f = self.label.font()
        f.setPointSize(f.pointSize() + 2)
        self.label.setFont(f)
        lay.addWidget(self.label)
        btn = QPushButton(tr("choose"))
        btn.clicked.connect(self.choose)
        lay.addWidget(btn, alignment=Qt.AlignCenter)

    def choose(self):
        path, _ = QFileDialog.getOpenFileName(self, tr("choose"), str(Path.home()), "Word (*.docx)")
        if path:
            self.on_file(path)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self.setStyleSheet(self.hover_style)

    def dragLeaveEvent(self, e):
        self.setStyleSheet(self.normal_style)

    def dropEvent(self, e):
        self.setStyleSheet(self.normal_style)
        urls = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        if urls:
            self.on_file(urls[0])


# ---------- 주 창 ----------

class Main(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP)
        self.setMinimumSize(600, 560)
        self.path = None
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 16)
        lay.setSpacing(14)

        head = QHBoxLayout()
        icon = ASSETS / "icon.png"
        if icon.exists():
            ic = QLabel()
            ic.setPixmap(QPixmap(str(icon)).scaled(56, 56, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            head.addWidget(ic)
        titles = QVBoxLayout()
        t = QLabel(APP)
        tf = QFont(t.font())
        tf.setPointSize(tf.pointSize() + 9)
        tf.setBold(True)
        t.setFont(tf)
        sub = QLabel(tr("subtitle"))
        sub.setStyleSheet("color: gray;")
        titles.addWidget(t)
        titles.addWidget(sub)
        head.addLayout(titles)
        head.addStretch()
        lay.addLayout(head)

        self.drop = DropZone(self.set_file)
        lay.addWidget(self.drop)

        self.file_label = QLabel()
        self.file_label.setWordWrap(True)
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setMinimumHeight(44)
        lay.addWidget(self.file_label)

        acts = QHBoxLayout()
        self.send_btn = self.action_button(tr("send"), tr("send_sub"), "#a4161a")
        self.recv_btn = self.action_button(tr("recv"), tr("recv_sub"), "#2f6feb")
        self.send_btn.clicked.connect(self.send)
        self.recv_btn.clicked.connect(self.receive)
        acts.addWidget(self.send_btn)
        acts.addWidget(self.recv_btn)
        lay.addLayout(acts)

        self.busy = QLabel()
        self.busy.setAlignment(Qt.AlignCenter)
        self.busy.setStyleSheet("color: #2f6feb;")
        lay.addWidget(self.busy)

        foot = QHBoxLayout()
        self.status = QLabel()
        self.status.setStyleSheet("color: gray;")
        help_btn = QPushButton(tr("help"))
        help_btn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(HELP_URL)))
        set_btn = QPushButton(tr("settings"))
        set_btn.clicked.connect(self.open_settings)
        foot.addWidget(self.status, 1)
        foot.addWidget(help_btn)
        foot.addWidget(set_btn)
        lay.addLayout(foot)
        self.update_actions(None)
        self.refresh_status()

    def action_button(self, title, sub, color):
        b = QPushButton(f"{title}\n{sub}")
        b.setMinimumHeight(64)
        b.setStyleSheet(
            f"QPushButton {{ background: {color}; color: white; border: none; border-radius: 12px; "
            f"font-weight: 600; padding: 8px; }} QPushButton:disabled {{ background: #d0d4da; color: #7a808a; }}")
        return b

    def set_busy(self, msg):
        self.busy.setText(msg or "")
        for w in (self.send_btn, self.recv_btn, self.drop):
            w.setEnabled(msg is None)
        if msg is None:
            self.update_actions(getattr(self, "kind", None))

    def refresh_status(self):
        _, out = zms.run(["status"])
        z = value(out, "ZOTERO_DIR")
        if value(out, "MENDELEY_KEYS") != "ok":
            self.status.setText(tr("status_nokeys"))
        elif value(out, "MENDELEY_LOGIN") != "ok":
            self.status.setText(tr("status_nologin", z=z))
        else:
            self.status.setText(tr("status_ok", z=z))
        return out

    def ensure_setup(self):
        out = self.refresh_status()
        if value(out, "MENDELEY_KEYS") != "ok" or value(out, "ZOTERO_DB") != "ok":
            SettingsDialog(self, welcome=True).exec()
            out = self.refresh_status()
        if value(out, "MENDELEY_KEYS") != "ok" or value(out, "ZOTERO_DB") != "ok":
            return False
        if value(out, "MENDELEY_LOGIN") != "ok":
            ok = login(self)
            self.refresh_status()
            return ok
        return True

    def open_settings(self):
        SettingsDialog(self).exec()
        self.refresh_status()

    def update_actions(self, kind):
        self.kind = kind
        self.send_btn.setEnabled(kind in ("zotero", "both"))
        self.recv_btn.setEnabled(kind in ("mendeley", "both"))

    def set_file(self, path):
        if not path.lower().endswith(".docx"):
            QMessageBox.information(self, APP, tr("only_docx"))
            return
        self.path = path
        code, out = zms.run(["detect", path])
        if code != 0:
            show_error(self, code, out)
            return
        kind = value(out, "KIND")
        name = Path(path).name
        if kind == "zotero":
            info = tr("kind_zotero", n=value(out, "ZOTERO"))
        elif kind == "mendeley":
            info = tr("kind_mendeley", n=value(out, "MENDELEY"))
        elif kind == "both":
            info = tr("kind_both", z=value(out, "ZOTERO"), m=value(out, "MENDELEY"))
        else:
            info = tr("kind_none")
        self.file_label.setText(f"<b>{name}</b><br><span style='color:gray'>{info}</span>")
        self.update_actions(kind)

    def run_with_login(self, argv, message):
        code, out = run_blocking(self, argv, message)
        if code == zms.EXIT_AUTH:
            QMessageBox.information(self, APP, tr("need_login"))
            if login(self):
                self.refresh_status()
                code, out = run_blocking(self, argv, message)
        return code, out

    def send(self):
        if not self.path or not self.ensure_setup():
            return
        code, out = self.run_with_login(["pull"], tr("working_send"))
        if code == 0:
            code, out = self.run_with_login(["to-mendeley", self.path, "--add-missing"], tr("working_send"))
        if code != 0:
            show_error(self, code, out)
            return
        msg = tr("done_send", n=value(out, "CITES"))
        if value(out, "ADDED") not in ("", "0"):
            msg += "\n" + tr("done_send_added", n=value(out, "ADDED"))
        self.show_result(msg + "\n\n" + tr("done_send_tip"), value(out, "OUT"))

    def receive(self):
        if not self.path or not self.ensure_setup():
            return
        code, out = run_blocking(self, ["to-zotero", self.path], tr("working_recv"))
        if code == zms.EXIT_MISSING:
            box = QMessageBox(self)
            box.setWindowTitle(APP)
            box.setIcon(QMessageBox.Question)
            box.setText(tr("missing_q", n=value(out, "MISSING")))
            add = box.addButton(tr("add_and_convert"), QMessageBox.AcceptRole)
            only = box.addButton(tr("convert_only"), QMessageBox.ActionRole)
            box.addButton(tr("cancel"), QMessageBox.RejectRole)
            box.setDefaultButton(add)
            box.exec()
            if box.clickedButton() not in (add, only):
                return
            if box.clickedButton() is add and not self.import_to_zotero(value(out, "MISSING_FILE")):
                return
            code, out = run_blocking(self, ["to-zotero", self.path, "--allow-unlinked"], tr("working_recv"))
        if code != 0:
            show_error(self, code, out)
            return
        msg = tr("done_recv", n=value(out, "CITES"))
        if value(out, "UNLINKED") not in ("", "0"):
            msg += "\n" + tr("done_recv_unl", n=value(out, "UNLINKED"))
        self.show_result(msg + "\n\n" + tr("done_recv_tip"), value(out, "OUT"))

    def import_to_zotero(self, json_path):
        while True:
            code, out = run_blocking(self, ["zotero-import", json_path], tr("working_import"))
            if code == 0:
                return True
            if code != zms.EXIT_ZOTERO_OFF:
                show_error(self, code, out)
                return False
            r = QMessageBox.question(self, APP, tr("zotero_off"), QMessageBox.Retry | QMessageBox.Cancel)
            if r != QMessageBox.Retry:
                return False

    def show_result(self, msg, out_path):
        box = QMessageBox(self)
        box.setWindowTitle(APP)
        box.setIcon(QMessageBox.Information)
        box.setText(msg)
        box.setInformativeText(tr("new_file", f=Path(out_path).name))
        word = box.addButton(tr("open_word"), QMessageBox.AcceptRole)
        folder = box.addButton(tr("reveal"), QMessageBox.ActionRole)
        box.addButton(tr("ok"), QMessageBox.RejectRole)
        box.setDefaultButton(word)
        box.exec()
        if box.clickedButton() is word:
            QDesktopServices.openUrl(QUrl.fromLocalFile(out_path))
        elif box.clickedButton() is folder:
            reveal(out_path)


class App(QApplication):
    """macOS 에서 Dock 아이콘에 파일을 끌어다 놓으면 FileOpen 이벤트로 들어온다."""

    def __init__(self, argv):
        super().__init__(argv)
        self.window = None
        self.pending = None

    def event(self, e):
        if e.type() == QEvent.FileOpen:
            if self.window:
                self.window.set_file(e.file())
            else:
                self.pending = e.file()
            return True
        return super().event(e)


def main():
    app = App(sys.argv)
    app.setApplicationName(APP)
    if (ASSETS / "icon.png").exists():
        app.setWindowIcon(QIcon(str(ASSETS / "icon.png")))
    w = Main()
    app.window = w
    w.show()
    # Windows 에서는 exe 아이콘에 끌어다 놓은 파일이 명령줄 인자로 들어온다
    first = app.pending or next((a for a in sys.argv[1:] if a.lower().endswith(".docx")), None)
    if first:
        w.set_file(first)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
