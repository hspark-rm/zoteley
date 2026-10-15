#!/usr/bin/env python3
"""Zotero → Mendeley 문헌 정보 반영 도구.

명령:
  config client ID·secret 입력(secret 은 화면에 안 보임)
  auth   브라우저로 Mendeley 로그인, 토큰 저장
  pull   Mendeley 개인·그룹 문헌을 data/backup_<시각>.json 으로 백업
  diff   최신 백업과 Zotero(sqlite 사본)를 짝지어 data/diff_<시각>.csv 생성
  apply  diff CSV에서 approve=y 행만 PATCH (기본 1건, --all 로 전체)
  harvest <docx>  공저자 원고 인용 중 내게 없는 문헌 → Zotero 가져오기 파일 + Mendeley 추가(--post)
  relink  <docx>  원고 사본(_relinked.docx)의 인용을 대상 라이브러리 문헌 ID로 다시 연결
  to-mendeley <docx>  Zotero 원고 → Mendeley Cite 원고(_mendeley.docx)
  to-zotero   <docx>  Mendeley Cite 원고 → Zotero 원고(_zotero.docx)

설정: ~/.config/zotero-mendeley-sync/config.json
      {"client_id": ..., "client_secret": ..., "redirect_uri": "http://localhost:5000/callback"}
"""
import argparse
import base64
import csv
import datetime as dt
import email.utils
import html
import http.server
import json
import os
import re
import secrets
import select
import shutil
import socket
import sqlite3
import sys
import tempfile
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import uuid
import webbrowser
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

API = "https://api.mendeley.com"
DOC_TYPE = "application/vnd.mendeley-document.1+json"
# Cloudflare 가 urllib 기본 User-Agent 를 error 1010 으로 막는다(2026-10-13 확인).
VERSION = "1.0.0"
UA = f"Zoteley/{VERSION} (+https://github.com/hspark-rm/zoteley)"
CONF_DIR = Path(os.environ.get("ZMS_CONF") or Path.home() / ".config" / "zotero-mendeley-sync")
# 앱으로 실행하면 앱 꾸러미 안이 아니라 사용자 폴더에 작업 파일을 둔다(ZMS_DATA).
DATA = Path(os.environ.get("ZMS_DATA") or Path(__file__).resolve().parent / "data")
DEFAULT_ZOTERO_DIR = Path.home() / "Zotero"
REDIRECT_DEFAULT = "http://localhost:53682/callback"
# 앱(AppleScript)이 오류 종류를 가려 다음 창을 띄우도록 종료 코드를 나눈다.
EXIT_MISSING, EXIT_AUTH, EXIT_CONFIG, EXIT_ZOTERO_OFF, EXIT_CANCEL = 3, 6, 7, 8, 9
# 앱(GUI)이 로그인 창에서 붙여 넣은 주소와 취소 신호를 넘겨주는 통로. 명령줄에서는 쓰지 않는다.
PASTE_QUEUE = None
CANCEL = __import__("threading").Event()

# Zotero itemType → Mendeley type. 대응이 없는 유형은 type 을 건드리지 않는다.
TYPE_MAP = {
    "journalArticle": "journal", "book": "book", "bookSection": "book_section",
    "conferencePaper": "conference_proceedings", "report": "report", "thesis": "thesis",
    "webpage": "web_page", "magazineArticle": "magazine_article",
    "newspaperArticle": "newspaper_article", "patent": "patent", "statute": "statute",
    "computerProgram": "computer_program", "encyclopediaArticle": "encyclopedia_article",
    "case": "case", "film": "film", "bill": "bill", "hearing": "hearing",
    "preprint": "working_paper", "document": "generic",
}
SOURCE_FIELDS = ["publicationTitle", "bookTitle", "proceedingsTitle", "websiteTitle",
                 "blogTitle", "encyclopediaTitle", "dictionaryTitle", "forumTitle", "programTitle"]
# Mendeley API 길이 제한. 넘으면 잘라 넣지 않고 그 필드를 건너뛴다(잘린 값은 틀린 값이므로).
MAX_LEN = {"title": 500, "source": 255, "volume": 20, "pages": 50, "abstract": 10000}
ID_KEYS = ["doi", "issn", "isbn", "pmid"]
FIELDS = ["title", "authors", "editors", "year", "month", "day", "source", "volume",
          "issue", "pages", "abstract", "language", "type"] + [f"identifiers.{k}" for k in ID_KEYS]
CSV_COLS = ["approve", "status", "kind", "library", "title", "field", "mendeley_value",
            "zotero_value", "match", "mendeley_id", "zotero_key", "last_modified",
            "mendeley_json", "zotero_json"]


# ---------- 인증 ----------

def fail(code, msg):
    print(msg, file=sys.stderr)
    sys.exit(code)


def conf_raw():
    p = CONF_DIR / "config.json"
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def write_conf(conf):
    CONF_DIR.mkdir(parents=True, exist_ok=True)
    fd = os.open(CONF_DIR / "config.json", os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(conf, f, indent=2)


def zotero_db():
    return Path(conf_raw().get("zotero_dir") or DEFAULT_ZOTERO_DIR) / "zotero.sqlite"


def default_group(a):
    return getattr(a, "group", None) or conf_raw().get("group") or None


def load_conf():
    c = conf_raw()
    for k in ("client_id", "client_secret", "redirect_uri"):
        if not c.get(k) or str(c[k]).startswith("<"):
            fail(EXIT_CONFIG, f"Mendeley 설정({k})이 비어 있음 — 설정에서 입력하세요.")
    return c


def cmd_config(a):
    """설정을 저장한다. 옵션을 주면(앱) 묻지 않고 바꾸고, 없으면 터미널에서 묻는다.
    secret 은 명령줄에 남지 않게 환경변수 ZMS_CLIENT_SECRET 으로만 받는다."""
    secret_file = os.environ.get("ZMS_CLIENT_SECRET_FILE")
    if secret_file and Path(secret_file).exists():
        # 앱은 secret 을 명령줄 대신 권한 600 임시 파일로 넘긴다. 읽은 즉시 지운다.
        os.environ["ZMS_CLIENT_SECRET"] = Path(secret_file).read_text(encoding="utf-8").strip()
        Path(secret_file).unlink()
    if a.client_id or a.zotero_dir or a.group is not None or os.environ.get("ZMS_CLIENT_SECRET"):
        conf = conf_raw()
        if a.client_id and a.client_id != conf.get("client_id"):
            conf["client_id"] = a.client_id.strip()
            (CONF_DIR / "token.json").unlink(missing_ok=True)  # 앱이 바뀌면 이전 토큰은 못 쓴다
        if os.environ.get("ZMS_CLIENT_SECRET"):
            conf["client_secret"] = os.environ["ZMS_CLIENT_SECRET"].strip()
        if a.zotero_dir:
            if not (Path(a.zotero_dir) / "zotero.sqlite").exists():
                fail(EXIT_CONFIG, f"{a.zotero_dir} 에 zotero.sqlite 가 없음 — Zotero 데이터 폴더를 고르세요.")
            conf["zotero_dir"] = a.zotero_dir
        if a.group is not None:
            conf["group"] = a.group or None
        conf.setdefault("redirect_uri", REDIRECT_DEFAULT)
        write_conf(conf)
        print("저장됨")
        return
    import getpass
    CONF_DIR.mkdir(parents=True, exist_ok=True)
    conf = {**conf_raw(), "client_id": input("client ID: ").strip(),
            "client_secret": getpass.getpass("client secret (입력 내용은 안 보임): ").strip(),
            "redirect_uri": input("redirect URL [http://localhost:53682/callback]: ").strip()}
    # 방향키 같은 제어 문자가 섞이거나 비어 있으면 기본 주소를 쓴다.
    if not re.fullmatch(r"http://[\w.\[\]:-]+(/\S*)?", conf["redirect_uri"]):
        conf["redirect_uri"] = "http://localhost:53682/callback"
    write_conf(conf)
    (CONF_DIR / "token.json").unlink(missing_ok=True)  # 앱이 바뀌면 이전 토큰은 못 쓴다
    print(f"저장: {CONF_DIR / 'config.json'}")


def save_token(tok):
    p = CONF_DIR / "token.json"
    fd = os.open(p, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(tok, f)


def token_request(conf, data):
    basic = base64.b64encode(f"{conf['client_id']}:{conf['client_secret']}".encode()).decode()
    req = urllib.request.Request(
        API + "/oauth/token", data=urllib.parse.urlencode(data).encode(),
        headers={"Content-Type": "application/x-www-form-urlencoded", "Authorization": "Basic " + basic,
                 "User-Agent": UA})
    with urllib.request.urlopen(req) as r:
        tok = json.load(r)
    tok["obtained_at"] = time.time()
    return tok


def cmd_auth(a):
    conf = load_conf()
    ru = urllib.parse.urlparse(conf["redirect_uri"])
    if ru.scheme != "http" or ru.hostname not in ("localhost", "127.0.0.1"):
        sys.exit("redirect_uri 는 http://localhost:<포트>/... 형식이어야 함 (앱 등록 값과 같아야 함)")
    state = secrets.token_urlsafe(16)
    got = {}

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            got.update({k: v[0] for k, v in q.items()})
            self.send_response(200)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write("Mendeley 로그인 완료. 이 창을 닫아도 됩니다.".encode())

        def log_message(self, *args):
            pass

    class Server6(http.server.HTTPServer):
        address_family = socket.AF_INET6

    # macOS AirPlay 수신(ControlCenter)이 *:5000 을 IPv4·IPv6 모두 점유한다. 브라우저는 localhost 를
    # ::1 로 열 수 있으므로 127.0.0.1 과 ::1 을 따로 묶어야 AirPlay 대신 이 서버가 응답한다.
    port = ru.port or 80
    addrs = [("127.0.0.1", http.server.HTTPServer), ("::1", Server6)] if ru.hostname == "localhost" \
        else [(ru.hostname, http.server.HTTPServer)]
    servers = []
    for host, cls in addrs:
        try:
            servers.append(cls((host, port), Handler))
        except OSError as e:
            print(f"  {host}:{port} 대기 실패: {e}")
    if not servers:
        print(f"  {port}번 포트를 쓸 수 없어 주소 붙여 넣기로만 받음")
    url = API + "/oauth/authorize?" + urllib.parse.urlencode({
        "client_id": conf["client_id"], "redirect_uri": conf["redirect_uri"],
        "response_type": "code", "scope": "all", "state": state})
    print("브라우저에서 Mendeley에 로그인하세요:\n" + url)
    print("\n로그인 뒤 브라우저에 오류(403 등)가 보이면, 그 창 주소창의 주소 전체를 복사해 여기에 붙여 넣고 Enter.")
    webbrowser.open(url)
    # 콜백 서버로 받거나, 사용자가 붙여 넣은 주소에서 code 를 꺼낸다.
    # 다른 프로그램(AirPlay 등)이 포트를 가로채도 붙여 넣기로 끝낼 수 있게 하기 위함이다.
    use_stdin = os.name != "nt" and PASTE_QUEUE is None and sys.stdin is not None and sys.stdin.isatty()
    inputs = [sys.stdin] if use_stdin else []
    deadline = time.time() + a.timeout if a.timeout else None
    while "code" not in got and "error" not in got:
        if CANCEL.is_set():
            fail(EXIT_CANCEL, "로그인을 취소함")
        if PASTE_QUEUE is not None and not PASTE_QUEUE.empty():
            q = urllib.parse.parse_qs(urllib.parse.urlparse(PASTE_QUEUE.get()).query)
            got.update({k: v[0] for k, v in q.items()})
            continue
        left = deadline - time.time() if deadline else None
        if left is not None and left <= 0:
            fail(EXIT_AUTH, "로그인 시간이 지나 중단함 — 다시 시도하세요.")
        wait = 0.5 if left is None else min(0.5, left)
        if not servers and not inputs:
            time.sleep(wait)
            continue
        ready, _, _ = select.select(servers + inputs, [], [], wait)
        for r in ready:
            if r is sys.stdin:
                line = sys.stdin.readline()
                if not line:  # 입력이 닫혔으면(EOF) 더 읽지 않는다
                    inputs.clear()
                    continue
                line = line.strip()
                q = urllib.parse.parse_qs(urllib.parse.urlparse(line).query)
                if "code" in q or "error" in q:
                    got.update({k: v[0] for k, v in q.items()})
                else:
                    print("주소에 code 가 없음. 로그인 뒤 주소창의 주소 전체를 붙여 넣으세요.")
            else:
                r.handle_request()
    for srv in servers:
        srv.server_close()
    if "error" in got:
        sys.exit(f"인증 거부: {got}")
    if got.get("state") != state:
        sys.exit("state 불일치 — 다른 요청이 섞였을 수 있어 중단함")
    try:
        tok = token_request(conf, {"grant_type": "authorization_code", "code": got["code"],
                                   "redirect_uri": conf["redirect_uri"]})
    except urllib.error.HTTPError as e:
        fail(EXIT_AUTH, f"토큰 발급 실패({e.code}): {e.read().decode(errors='replace')[:300]}\n"
                        "client ID·secret·redirect URL 이 Mendeley 앱 등록 값과 같은지 확인하세요.")
    save_token(tok)
    print("토큰 저장:", CONF_DIR / "token.json")


def access_token():
    p = CONF_DIR / "token.json"
    if not p.exists():
        fail(EXIT_AUTH, "Mendeley 로그인이 필요함")
    tok = json.loads(p.read_text(encoding="utf-8"))
    if time.time() < tok["obtained_at"] + tok.get("expires_in", 3600) - 120:
        return tok["access_token"]
    if not tok.get("refresh_token"):
        fail(EXIT_AUTH, "Mendeley 로그인이 만료됨")
    conf = load_conf()
    try:
        new = token_request(conf, {"grant_type": "refresh_token", "refresh_token": tok["refresh_token"],
                                   "redirect_uri": conf["redirect_uri"]})
    except urllib.error.HTTPError as e:
        fail(EXIT_AUTH, f"Mendeley 로그인 갱신 실패({e.code})")
    # 갱신 응답에 refresh_token 이 빠지면 이전 것을 계속 쓴다.
    new.setdefault("refresh_token", tok["refresh_token"])
    save_token(new)
    return new["access_token"]


# ---------- API ----------

def api(method, path, params=None, body=None, headers=None, accept=DOC_TYPE):
    url = path if path.startswith("http") else API + path
    if params:
        url += "?" + urllib.parse.urlencode(params)
    data = json.dumps(body).encode() if body is not None else None
    for attempt in range(5):
        h = {"Authorization": "Bearer " + access_token(), "Accept": accept, "User-Agent": UA}
        if data is not None:
            h["Content-Type"] = DOC_TYPE
        h.update(headers or {})
        req = urllib.request.Request(url, data=data, method=method, headers=h)
        try:
            with urllib.request.urlopen(req) as r:
                raw = r.read()
                return r.status, r.headers, (json.loads(raw) if raw else None)
        except urllib.error.HTTPError as e:
            # 호출 한도(429)와 서버 오류(5xx)만 기다렸다가 다시 시도한다.
            if (e.code == 429 or e.code >= 500) and attempt < 4:
                time.sleep(int(e.headers.get("Retry-After") or 5 * 2 ** attempt))
                continue
            raise


def next_link(link):
    for part in (link or "").split(","):
        m = re.match(r'\s*<([^>]+)>\s*;\s*rel="?next"?', part)
        if m:
            return m.group(1)
    return None


def get_all(path, params, accept=DOC_TYPE):
    out, url = [], path
    while url:
        _, h, js = api("GET", url, params=params, accept=accept)
        out += js or []
        params = None  # next 링크에 쿼리가 이미 들어 있다
        url = next_link(h.get("Link"))
    return out


def list_groups():
    for path, accept in (("/groups/v2", "application/vnd.mendeley-group-list+json"),
                         ("/groups", "application/vnd.mendeley-group.1+json")):
        try:
            return get_all(path, {"limit": 500}, accept=accept)
        except urllib.error.HTTPError as e:
            print(f"  그룹 목록 {path} 실패({e.code})")
    print("  그룹 목록을 못 받아 개인 라이브러리만 백업함")
    return []


def cmd_pull(a):
    DATA.mkdir(exist_ok=True)
    groups = list_groups()
    docs = {d["id"]: d for d in get_all("/documents", {"view": "all", "limit": 500})}
    print(f"개인 라이브러리: {len(docs)}건")
    for g in groups:
        gd = get_all("/documents", {"view": "all", "limit": 500, "group_id": g["id"]})
        print(f"그룹 '{g.get('name')}': {len(gd)}건")
        docs.update({d["id"]: d for d in gd})
    ts = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    out = DATA / f"backup_{ts}.json"
    out.write_text(json.dumps({
        "pulled_at": dt.datetime.now(dt.timezone.utc).isoformat(),
        "groups": [{"id": g["id"], "name": g.get("name")} for g in groups],
        "documents": list(docs.values())}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"백업 {len(docs)}건 → {out}")


# ---------- Zotero ----------

def norm_doi(s):
    s = (s or "").strip().lower()
    s = re.sub(r"^(https?://(dx\.)?doi\.org/|doi:\s*)", "", s)
    return s if s.startswith("10.") else ""


def parse_date(s):
    # Zotero 날짜는 "2022-04-00 04/2022" 처럼 SQL 날짜가 앞에 붙은 다중 형식이다.
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s or "")
    if not m:
        y = re.search(r"\b(1[5-9]\d{2}|20\d{2})\b", s or "")
        return (int(y.group(1)), None, None) if y else (None, None, None)
    y, mo, d = (int(x) for x in m.groups())
    return (y or None, mo or None, d or None)


def load_zotero():
    tmp = Path(tempfile.mkdtemp()) / "zotero.sqlite"
    db = zotero_db()
    if not db.exists():
        fail(EXIT_CONFIG, f"Zotero 데이터 파일 {db} 이 없음 — 설정에서 Zotero 폴더를 고르세요.")
    shutil.copy2(db, tmp)
    wal = db.with_name(db.name + "-wal")
    if wal.exists():
        shutil.copy2(wal, tmp.with_name(tmp.name + "-wal"))
    con = sqlite3.connect(tmp)
    items = {}
    for iid, key, typ in con.execute(
            """select i.itemID, i.key, t.typeName from items i join itemTypes t using(itemTypeID)
               where t.typeName not in ('attachment','note','annotation')
                 and i.itemID not in (select itemID from deletedItems)"""):
        items[iid] = {"id": iid, "key": key, "type": typ, "fields": {}, "creators": []}
    for iid, name, val in con.execute(
            """select d.itemID, f.fieldName, v.value from itemData d
               join fields f using(fieldID) join itemDataValues v using(valueID)"""):
        if iid in items:
            items[iid]["fields"][name] = val
    for iid, first, last, mode, ctype in con.execute(
            """select ic.itemID, c.firstName, c.lastName, c.fieldMode, ct.creatorType
               from itemCreators ic join creators c using(creatorID)
               join creatorTypes ct using(creatorTypeID) order by ic.itemID, ic.orderIndex"""):
        if iid in items:
            items[iid]["creators"].append({"type": ctype, "first": "" if mode == 1 else (first or ""),
                                           "last": last or ""})
    con.close()
    shutil.rmtree(tmp.parent, ignore_errors=True)
    return list(items.values())


def zotero_to_mendeley(it):
    f, m = it["fields"], {}
    extra = f.get("extra", "")

    def put(k, v):
        if v in (None, "", []):
            return
        if isinstance(v, str) and len(v) > MAX_LEN.get(k, 10 ** 9):
            return
        m[k] = v

    put("title", f.get("title"))
    for role, key in (("author", "authors"), ("editor", "editors")):
        put(key, [{"first_name": c["first"], "last_name": c["last"]}
                  for c in it["creators"] if c["type"] == role])
    y, mo, d = parse_date(f.get("date"))
    put("year", y)
    put("month", mo)
    put("day", d)
    put("source", next((f[k] for k in SOURCE_FIELDS if f.get(k)), None))
    for k in ("volume", "issue", "pages", "language"):
        put(k, f.get(k))
    put("abstract", f.get("abstractNote"))
    put("type", TYPE_MAP.get(it["type"]))
    # DOI 필드가 없는 유형은 Zotero가 extra 에 "DOI: ..." 로 둔다.
    doi = norm_doi(f.get("DOI")) or norm_doi((re.search(r"^DOI:\s*(\S+)", extra, re.M | re.I) or [None, ""])[1])
    pmid = (re.search(r"^PMID:\s*(\d+)", extra, re.M) or [None, ""])[1]
    ids = {"doi": doi, "issn": re.split(r"[,;\s]+", f.get("ISSN", "").strip())[0], "isbn": f.get("ISBN", ""), "pmid": pmid}
    for k in ID_KEYS:
        put(f"identifiers.{k}", ids[k])
    return m


# ---------- 비교 ----------

def nfc(s):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(s))).strip()


def alnum(s):
    return re.sub(r"[\W_]+", "", unicodedata.normalize("NFKC", s or "").casefold())


def norm_title(s):
    return re.sub(r"[\W_]+", "", unicodedata.normalize("NFKC", s or "").casefold())


def names(v):
    return [(nfc(x.get("last_name", "")), nfc(x.get("first_name", ""))) for x in v or []]


def show(field, v):
    if v in (None, "", []):
        return ""
    if field in ("authors", "editors"):
        return "; ".join(f"{l}, {f}" if f else l for l, f in names(v))
    return nfc(v)


def mendeley_value(doc, field):
    if field.startswith("identifiers."):
        return (doc.get("identifiers") or {}).get(field.split(".", 1)[1])
    return doc.get(field)


def classify(field, mv, zv):
    """차이 종류. None 이면 같은 값으로 본다."""
    if field in ("authors", "editors"):
        a, b = names(mv), names(zv)
        if a == b:
            return None
        if not a:
            return "fill"
        if [(alnum(l), alnum(f)) for l, f in a] == [(alnum(l), alnum(f)) for l, f in b]:
            return "format"  # 아포스트로피·마침표 등 문장부호만 다름
        if [alnum(x[0]) for x in a] == [alnum(x[0]) for x in b]:
            return "firstname"  # 성 순서는 같고 이름(이니셜 등)만 다름
        return "change"
    a, b = show(field, mv), show(field, zv)
    if field == "pages":
        a, b = re.sub(r"[‐-―−]", "-", a), re.sub(r"[‐-―−]", "-", b)
    if field.startswith("identifiers.") or field == "language":
        a, b = a.lower(), b.lower()
    if a == b:
        return None
    if not a:
        return "fill"
    if a.casefold() == b.casefold():
        return "case"
    if alnum(a) == alnum(b):
        return "format"  # 공백·하이픈·문장부호만 다름 (예: ISSN 0141-9870 / 01419870)
    return "change"


def match_key(title, year, first_last):
    t = norm_title(title)
    return (t, year, first_last.casefold()) if t and year and first_last else None


def doc_key(doc):
    """Mendeley 문헌의 짝짓기 키: DOI 가 있으면 DOI, 없으면 (제목, 연도, 첫 저자 성)."""
    doi = norm_doi((doc.get("identifiers") or {}).get("doi"))
    if doi:
        return ("doi", doi)
    k = match_key(doc.get("title"), doc.get("year"), ((doc.get("authors") or [{}])[0]).get("last_name", ""))
    return ("tyl", k) if k else None


def latest(pattern):
    files = sorted(DATA.glob(pattern))
    if not files:
        sys.exit(f"data/{pattern} 없음")
    return files[-1]


def cmd_diff(a):
    backup_path = Path(a.backup) if a.backup else latest("backup_*.json")
    backup = json.loads(backup_path.read_text(encoding="utf-8"))
    gname = {g["id"]: g["name"] for g in backup["groups"]}
    zitems = load_zotero()
    by_doi, by_tyl = defaultdict(list), defaultdict(list)
    for it in zitems:
        tm = zotero_to_mendeley(it)
        it["target"] = tm
        if tm.get("identifiers.doi"):
            by_doi[tm["identifiers.doi"]].append(it)
        k = match_key(tm.get("title"), tm.get("year"), (tm.get("authors") or [{}])[0].get("last_name", ""))
        if k:
            by_tyl[k].append(it)

    # 같은 라이브러리 안에서 같은 문헌이 여러 번 있으면(D5) 건너뛴다.
    mkey = doc_key
    lib_count = Counter((d.get("group_id"), mkey(d)) for d in backup["documents"] if mkey(d))

    rows, stats = [], Counter()
    for doc in backup["documents"]:
        lib = gname.get(doc.get("group_id"), doc.get("group_id") or "개인")
        base = {"library": lib, "title": nfc(doc.get("title", ""))[:120], "mendeley_id": doc["id"],
                "last_modified": doc.get("last_modified", "")}
        k = mkey(doc)
        if not k:
            stats["짝짓기 정보 없음"] += 1
            continue
        cands = by_doi.get(k[1], []) if k[0] == "doi" else by_tyl.get(k[1], [])
        if not cands:
            stats["Zotero에 없음"] += 1
            continue
        if lib_count[(doc.get("group_id"), k)] > 1:
            stats["Mendeley 중복(건너뜀)"] += 1
            rows.append({**base, "status": "skip", "kind": "mendeley_duplicate", "match": k[0]})
            continue
        if len(cands) > 1:
            stats["Zotero 중복(건너뜀)"] += 1
            rows.append({**base, "status": "skip", "kind": "zotero_duplicate", "match": k[0],
                         "zotero_key": ",".join(c["key"] for c in cands)})
            continue
        z = cands[0]
        stats[f"짝지음({k[0]})"] += 1
        for field in FIELDS:
            zv = z["target"].get(field)
            if zv in (None, "", []):
                continue  # Zotero에 없는 값으로 Mendeley 값을 지우지 않는다
            mv = mendeley_value(doc, field)
            kind = classify(field, mv, zv)
            if kind:
                stats[f"차이 {kind}"] += 1
                rows.append({**base, "status": "diff", "kind": kind, "match": k[0], "zotero_key": z["key"],
                             "field": field, "mendeley_value": show(field, mv), "zotero_value": show(field, zv),
                             "mendeley_json": json.dumps(mv, ensure_ascii=False),
                             "zotero_json": json.dumps(zv, ensure_ascii=False)})
    ts = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    out = DATA / f"diff_{ts}.csv"
    with open(out, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig: Excel 한글 깨짐 방지
        w = csv.DictWriter(f, CSV_COLS)
        w.writeheader()
        w.writerows({c: r.get(c, "") for c in CSV_COLS} for r in rows)
    out.with_suffix(".meta.json").write_text(json.dumps({"backup": str(backup_path)}), encoding="utf-8")
    for k, v in sorted(stats.items()):
        print(f"  {k}: {v}")
    print(f"차이표 → {out}\n승인할 행의 approve 칸에 y 를 넣고 저장한 뒤 apply 를 실행하세요.")


# ---------- 반영 ----------

def http_date(iso):
    t = dt.datetime.fromisoformat(iso.replace("Z", "+00:00"))
    # HTTP 날짜는 초 단위라서, 내림하면 서버가 '그 뒤에 수정됨'으로 오판한다. 올림한다.
    if t.microsecond:
        t = t.replace(microsecond=0) + dt.timedelta(seconds=1)
    return email.utils.format_datetime(t.astimezone(dt.timezone.utc), usegmt=True)


def cmd_apply(a):
    csv_path = Path(a.csv) if a.csv else latest("diff_*.csv")
    backup = json.loads(Path(json.loads(csv_path.with_suffix(".meta.json").read_text(encoding="utf-8"))["backup"]).read_text(encoding="utf-8"))
    bdocs = {d["id"]: d for d in backup["documents"]}
    with open(csv_path, encoding="utf-8-sig") as f:
        rows = [r for r in csv.DictReader(f)
                if r["status"] == "diff" and r["approve"].strip().lower() in ("y", "yes", "o", "1")]
    plan = defaultdict(list)
    for r in rows:
        plan[r["mendeley_id"]].append(r)
    ids = list(plan) if a.all else list(plan)[:a.limit]
    print(f"승인 행 {len(rows)}개, 문헌 {len(plan)}건 중 {len(ids)}건 처리" + (" (dry-run)" if a.dry_run else ""))
    log_path = DATA / f"apply_{dt.datetime.now().strftime('%Y%m%d-%H%M%S')}.jsonl"
    result = Counter()
    with open(log_path, "a", encoding="utf-8") as log:
        for doc_id in ids:
            old = bdocs[doc_id]
            try:
                _, _, cur = api("GET", f"/documents/{doc_id}", params={"view": "all"})
            except urllib.error.HTTPError as e:
                result[f"조회 실패 {e.code}"] += 1
                log.write(json.dumps({"id": doc_id, "status": f"get_{e.code}"}) + "\n")
                continue
            if cur.get("last_modified") != old.get("last_modified"):
                result["백업 후 수정됨(건너뜀)"] += 1
                log.write(json.dumps({"id": doc_id, "status": "changed_since_pull"}) + "\n")
                continue
            body = {}
            for r in plan[doc_id]:
                val = json.loads(r["zotero_json"])
                if r["field"].startswith("identifiers."):
                    # identifiers 는 통째로 바뀌므로 기존 식별자(scopus 등)를 살려서 합친다.
                    body.setdefault("identifiers", dict(cur.get("identifiers") or {}))[r["field"].split(".", 1)[1]] = val
                else:
                    body[r["field"]] = val
            before = {k: cur.get(k) for k in body}
            if a.dry_run:
                print(json.dumps({"id": doc_id, "title": cur.get("title"), "patch": body}, ensure_ascii=False, indent=1))
                continue
            try:
                _, _, new = api("PATCH", f"/documents/{doc_id}", body=body,
                                headers={"If-Unmodified-Since": http_date(old["last_modified"])})
            except urllib.error.HTTPError as e:
                status = {412: "공저자가 먼저 수정함", 403: "수정 권한 없음"}.get(e.code, f"오류 {e.code}")
                result[status] += 1
                log.write(json.dumps({"id": doc_id, "status": f"patch_{e.code}", "patch": body,
                                      "error": e.read().decode(errors="replace")[:500]}, ensure_ascii=False) + "\n")
                continue
            ok = all(classify(k, (new or {}).get(k), v) is None for k, v in body.items() if k != "identifiers")
            result["반영·확인" if ok else "반영(값 불일치)"] += 1
            log.write(json.dumps({"id": doc_id, "status": "ok" if ok else "mismatch", "before": before,
                                  "patch": body, "last_modified": (new or {}).get("last_modified")},
                                 ensure_ascii=False) + "\n")
    for k, v in result.items():
        print(f"  {k}: {v}")
    if not a.dry_run:
        print(f"기록 → {log_path}")


# ---------- Word 원고 (Mendeley Cite) ----------

CITE_PREFIX = "MENDELEY_CITATION_v3_"
CSL_TYPE_MAP = {
    "article-journal": "journal", "book": "book", "chapter": "book_section",
    "paper-conference": "conference_proceedings", "report": "report", "thesis": "thesis",
    "webpage": "web_page", "article-magazine": "magazine_article",
    "article-newspaper": "newspaper_article", "patent": "patent", "legislation": "statute",
    "software": "computer_program", "entry-encyclopedia": "encyclopedia_article",
    "legal_case": "case", "motion_picture": "film", "bill": "bill", "hearing": "hearing",
    "article": "working_paper",
}


def cite_encode(obj):
    # Mendeley Cite 와 같은 직렬화(공백 없는 JSON, UTF-8, base64). 원본 인용 38/38개 바이트 일치 확인.
    return base64.b64encode(json.dumps(obj, separators=(",", ":"), ensure_ascii=False).encode()).decode()


def cite_decode(b64):
    return json.loads(base64.b64decode(b64))


def xml_attr_escape(s):
    # Word 는 이 속성에서 & < " 만 바꾸고 > 는 그대로 둔다(원본 파일 확인).
    return s.replace("&", "&amp;").replace("<", "&lt;").replace('"', "&quot;")


def docx_parts(zin):
    # 인용은 본문 외에 각주·머리글에도 들어갈 수 있다.
    return [n for n in zin.namelist() if n.startswith("word/") and n.endswith(".xml")
            and "webextension" not in n]


def read_docx_citations(path):
    items = {}
    with zipfile.ZipFile(path) as zin:
        for n in docx_parts(zin):
            for b in re.findall(CITE_PREFIX + r"([A-Za-z0-9+/=]+)", zin.read(n).decode("utf-8")):
                for it in cite_decode(b)["citationItems"]:
                    items.setdefault(it["id"], it)
    return list(items.values())


def csl_year(d):
    parts = ((d.get("issued") or {}).get("date-parts") or [[]])[0]
    try:
        return int(parts[0])
    except (IndexError, TypeError, ValueError):
        return None


def csl_keys(d):
    a = (d.get("author") or [{}])[0]
    k = match_key(d.get("title"), csl_year(d), a.get("family") or a.get("literal") or "")
    return {x for x in (("doi", norm_doi(d.get("DOI"))), ("tyl", k)) if x[1]}


def doc_keys(doc):
    """Mendeley 문헌의 키 전부. 한쪽만 DOI 가 있어도 제목 키로 같은 문헌을 알아보게 한다."""
    a = (doc.get("authors") or [{}])[0]
    k = match_key(doc.get("title"), doc.get("year"), a.get("last_name", ""))
    return {x for x in (("doi", norm_doi((doc.get("identifiers") or {}).get("doi"))), ("tyl", k)) if x[1]}


def zotero_keys():
    keys = set()
    for it in load_zotero():
        t = zotero_to_mendeley(it)
        if t.get("identifiers.doi"):
            keys.add(("doi", t["identifiers.doi"]))
        k = match_key(t.get("title"), t.get("year"), (t.get("authors") or [{}])[0].get("last_name", ""))
        if k:
            keys.add(("tyl", k))
    return keys


def csl_to_mendeley(d):
    def people(lst):
        return [{"first_name": p.get("given", ""), "last_name": p.get("family") or p.get("literal")}
                for p in lst or [] if p.get("family") or p.get("literal")]

    m = {"title": nfc(d.get("title", ""))[:MAX_LEN["title"]], "type": CSL_TYPE_MAP.get(d.get("type"), "generic")}
    for src, dst in (("author", "authors"), ("editor", "editors")):
        if people(d.get(src)):
            m[dst] = people(d.get(src))
    parts = ((d.get("issued") or {}).get("date-parts") or [[]])[0]
    for k, v in zip(("year", "month", "day"), parts):
        try:
            m[k] = int(v)
        except (TypeError, ValueError):
            pass
    for src, dst in (("container-title", "source"), ("volume", "volume"), ("issue", "issue"),
                     ("page", "pages"), ("abstract", "abstract"), ("language", "language")):
        v = d.get(src)
        if v and len(str(v)) <= MAX_LEN.get(dst, 10 ** 9):
            m[dst] = str(v)
    ids = {"doi": norm_doi(d.get("DOI")), "issn": d.get("ISSN", ""), "isbn": d.get("ISBN", ""),
           "pmid": d.get("PMID", "")}
    ids = {k: v for k, v in ids.items() if v}
    if ids:
        m["identifiers"] = ids
    return m


def target_library(backup, group_name):
    """--group 이름이 있으면 그 그룹, 없으면 개인 라이브러리 문헌만 돌려준다."""
    if not group_name:
        return None, [d for d in backup["documents"] if not d.get("group_id")]
    gid = next((g["id"] for g in backup["groups"] if g["name"] == group_name), None)
    if not gid:
        sys.exit(f"그룹 '{group_name}' 없음. 백업에 있는 그룹: {[g['name'] for g in backup['groups']]}")
    return gid, [d for d in backup["documents"] if d.get("group_id") == gid]


def cmd_harvest(a):
    """공저자 원고의 인용 중 내 Zotero·대상 Mendeley 라이브러리에 없는 문헌을 찾아 넣는다(S2)."""
    DATA.mkdir(exist_ok=True)
    items = [it for it in read_docx_citations(a.docx) if it["itemData"].get("title")]
    zk = zotero_keys()
    backup = json.loads(latest("backup_*.json").read_text(encoding="utf-8"))
    gid, docs = target_library(backup, default_group(a))
    lk = set().union(*(doc_keys(d) for d in docs)) if docs else set()
    miss_z = [it["itemData"] for it in items if not (csl_keys(it["itemData"]) & zk)]
    miss_m = [it["itemData"] for it in items if not (csl_keys(it["itemData"]) & lk)]
    lib = default_group(a) or "개인"
    print(f"원고 인용 문헌 {len(items)}건 | Zotero에 없음 {len(miss_z)} | Mendeley({lib})에 없음 {len(miss_m)}")
    ts = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    stem = Path(a.docx).stem[:40]
    zpath = DATA / f"harvest_{stem}_{ts}_zotero.json"
    zpath.write_text(json.dumps(miss_z, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Zotero 가져오기용 CSL-JSON → {zpath}\n  Zotero: 파일 > 가져오기 에서 이 파일을 고르면 새 컬렉션으로 들어감")
    if not a.post:
        for d in miss_m[:10]:
            print(f"  Mendeley 추가 예정: {d.get('title', '')[:80]}")
        print("Mendeley에 실제로 넣으려면 --post (기본 1건, --all 전체)")
        return
    todo = miss_m if a.all else miss_m[:a.limit]
    log_path = DATA / f"harvest_{stem}_{ts}_posted.jsonl"
    with open(log_path, "a", encoding="utf-8") as log:
        for d in todo:
            body = csl_to_mendeley(d)
            if gid:
                body["group_id"] = gid
            try:
                _, _, new = api("POST", "/documents", body=body)
                log.write(json.dumps({"status": "ok", "id": new["id"], "title": body["title"]}, ensure_ascii=False) + "\n")
                print(f"  추가: {new['id']} {body['title'][:70]}")
            except urllib.error.HTTPError as e:
                log.write(json.dumps({"status": f"post_{e.code}", "body": body,
                                      "error": e.read().decode(errors="replace")[:500]}, ensure_ascii=False) + "\n")
                print(f"  실패({e.code}): {body['title'][:70]}")
    print(f"기록 → {log_path}\n추가한 문헌으로 relink 하려면 먼저 pull 을 다시 실행하세요.")


def cmd_relink(a):
    """원고 사본의 인용을 대상 라이브러리의 같은 문헌 ID로 바꾼다(S3). 원본은 건드리지 않는다."""
    src = Path(a.docx)
    out = src.with_name(src.stem + "_relinked.docx")
    if out.exists():
        sys.exit(f"{out} 이미 있음 — 덮어쓰지 않음")
    backup = json.loads(latest("backup_*.json").read_text(encoding="utf-8"))
    _, docs = target_library(backup, default_group(a))
    idx = defaultdict(set)
    for d in docs:
        for k in doc_keys(d):
            idx[k].add(d["id"])

    def target_for(item_data):
        for k in sorted(csl_keys(item_data), key=lambda k: k[0] != "doi"):  # DOI 를 먼저 본다
            if len(idx.get(k, ())) == 1:
                return next(iter(idx[k]))
            if len(idx.get(k, ())) > 1:
                return "dup"
        return None

    stats, tag_map, id_map = Counter(), {}, {}
    with zipfile.ZipFile(src) as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}
        parts = docx_parts(zin)
    for n in parts:
        for b in set(re.findall(CITE_PREFIX + r"([A-Za-z0-9+/=]+)", data[n].decode("utf-8"))):
            if b in tag_map:
                continue
            c, changed = cite_decode(b), False
            for i, it in enumerate(c["citationItems"]):
                t = target_for(it["itemData"])
                if t is None:
                    stats["대상 라이브러리에 없음"] += 1
                elif t == "dup":
                    stats["대상 라이브러리에 중복(건너뜀)"] += 1
                elif t == it["id"]:
                    stats["이미 연결됨"] += 1
                else:
                    # 내 라이브러리에서 직접 넣은 인용과 같은 모양으로 만든다(uris·legacyDesktopId 없음).
                    c["citationItems"][i] = {"id": t, "itemData": {**it["itemData"], "id": t}, "isTemporary": False}
                    id_map[it["id"]] = t
                    stats["연결 바꿈"] += 1
                    changed = True
            if changed:
                tag_map[b] = cite_encode(c)
    for k, v in stats.items():
        print(f"  인용 항목 {k}: {v}")
    if not tag_map:
        sys.exit("바꿀 인용이 없어 파일을 만들지 않음")
    for n in parts:
        text = data[n].decode("utf-8")
        for old, new in tag_map.items():
            text = text.replace(CITE_PREFIX + old, CITE_PREFIX + new)
        data[n] = text.encode("utf-8")
    # 추가기능 속성 MENDELEY_CITATIONS 에도 인용 전체가 복사돼 있어 같이 바꿔야 한다.
    for n in data:
        if not (n.startswith("word/webextensions/") and n.endswith(".xml")):
            continue
        text = data[n].decode("utf-8")
        m = re.search(r'(<we:property name="MENDELEY_CITATIONS" value=")([^"]*)(")', text)
        if not m:
            continue
        raw = m.group(2)
        if xml_attr_escape(html.unescape(raw)) != raw:
            sys.exit(f"{n} 의 문자 이스케이프 방식을 재현하지 못해 중단함")
        cites = json.loads(html.unescape(raw))
        for c in cites:
            old = c.get("citationTag", "")[len(CITE_PREFIX):]
            if old in tag_map:
                c["citationItems"] = cite_decode(tag_map[old])["citationItems"]
                c["citationTag"] = CITE_PREFIX + tag_map[old]
        new_raw = xml_attr_escape(json.dumps(cites, separators=(",", ":"), ensure_ascii=False))
        data[n] = (text[:m.start(2)] + new_raw + text[m.end(2):]).encode("utf-8")
    with zipfile.ZipFile(out, "w") as zout:
        for i in infos:
            zout.writestr(i, data[i.filename])
    after = read_docx_citations(out)  # 새 파일을 다시 읽어 인용이 모두 해석되는지 확인
    print(f"새 파일 → {out} (인용 문헌 {len(after)}건 재확인)")
    DATA.mkdir(exist_ok=True)
    log = DATA / f"relink_{src.stem[:40]}_{dt.datetime.now().strftime('%Y%m%d-%H%M%S')}.json"
    log.write_text(json.dumps({"source": str(src), "output": str(out), "id_map": id_map}, indent=1), encoding="utf-8")
    print(f"ID 대응표 → {log}")


# ---------- 원고 변환: Zotero 필드 ↔ Mendeley Cite 콘텐츠 컨트롤 ----------

MENDELEY_CITE_ADDIN = "WA104382081"
CITE_ITEM_EXTRA = ("locator", "label", "prefix", "suffix", "suppress-author", "author-only")
CT_WEBEXT = "application/vnd.ms-office.webextension+xml"
CT_TASKPANES = "application/vnd.ms-office.webextensiontaskpanes+xml"
CT_CUSTOM = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
REL_TASKPANES = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
REL_WEBEXT = "http://schemas.microsoft.com/office/2011/relationships/webextension"
REL_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
CUSTOM_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"


def xml_text_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def load_docx(path):
    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        return infos, {i.filename: zin.read(i.filename) for i in infos}


def save_docx(out, infos, data):
    names = [i.filename for i in infos]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in infos:
            if i.filename in data:  # None 으로 지운 부분은 빼고 쓴다
                zout.writestr(i, data[i.filename])
        for n in data:
            if n not in names and data[n] is not None:
                zout.writestr(n, data[n])
    # 모든 XML 부분이 다시 읽히는지 확인한다(깨진 파일을 넘기지 않기 위함).
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(out) as z:
        for n in z.namelist():
            if n.endswith(".xml") or n.endswith(".rels"):
                ET.fromstring(z.read(n))


def out_path(src, suffix):
    """결과 파일 경로. 같은 이름이 있으면 덮어쓰지 않고 ' 2', ' 3' 을 붙인다."""
    base = Path(src).with_name(Path(src).stem + suffix)
    out, n = base.with_suffix(".docx"), 2
    while out.exists():
        out, n = base.with_name(f"{base.name} {n}.docx"), n + 1
    return out


def run_start(x, pos):
    """pos 를 포함하는 <w:r> 의 시작 위치."""
    return max(m.start() for m in re.finditer(r"<w:r[ >]", x[:pos]))


def run_end(x, pos):
    return x.index("</w:r>", pos) + len("</w:r>")


def para_start(x, pos):
    return max(m.start() for m in re.finditer(r"<w:p[ >]", x[:pos]))


def para_end(x, pos):
    return x.index("</w:p>", pos) + len("</w:p>")


def runs_text(xml):
    return "".join(html.unescape(t) for t in re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", xml))


def first_rpr(xml):
    m = re.search(r"<w:rPr>.*?</w:rPr>", xml, re.S)
    return m.group(0) if m else ""


def find_fields(x, marker):
    """marker(예: ZOTERO_ITEM)가 들어 있는 필드들의 (시작, separate 런 끝, end 런 시작, 끝, 명령문)."""
    out = []
    for m in re.finditer(r'<w:fldChar w:fldCharType="begin"[^>]*/>', x):
        sep = x.find('w:fldCharType="separate"', m.end())
        end = x.find('w:fldCharType="end"', m.end())
        nxt = x.find('w:fldCharType="begin"', m.end())
        if sep < 0 or end < 0 or sep > end or (0 <= nxt < end):
            continue  # 중첩 필드나 결과 없는 필드는 다루지 않는다
        instr = html.unescape("".join(re.findall(r"<w:instrText[^>]*>(.*?)</w:instrText>", x[m.end():sep], re.S)))
        if marker not in instr:
            continue
        out.append((run_start(x, m.start()), run_end(x, sep), run_start(x, end), run_end(x, end), instr))
    return out


def find_sdts(x, tag_prefix):
    """tag 가 tag_prefix 로 시작하는 콘텐츠 컨트롤들의 (시작, 끝, tag 값, sdtPr, sdtContent 안쪽)."""
    out = []
    for m in re.finditer(r'<w:tag w:val="(' + re.escape(tag_prefix) + r'[^"]*)"/>', x):
        s = max(k.start() for k in re.finditer(r"<w:sdt[ >]", x[:m.start()]))
        e = x.index("</w:sdt>", m.end()) + len("</w:sdt>")
        pr = re.search(r"<w:sdtPr>.*?</w:sdtPr>", x[s:e], re.S).group(0)
        inner = re.search(r"<w:sdtContent>(.*)</w:sdtContent>", x[s:e], re.S).group(1)
        out.append((s, e, m.group(1), pr, inner))
    return out


def apply_edits(x, edits):
    for s, e, rep in sorted(edits, key=lambda t: t[0], reverse=True):
        x = x[:s] + rep + x[e:]
    return x


def custom_props(data):
    return data.get("docProps/custom.xml", b"").decode("utf-8")


def set_custom_props(data, remove_prefix, add):
    """docProps/custom.xml 에서 remove_prefix 속성을 지우고 add(이름→문자열)를 넣는다. 없으면 만든다."""
    c = custom_props(data)
    if not c:
        c = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.'
             'openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.'
             'org/officeDocument/2006/docPropsVTypes"></Properties>')
        add_override(data, "/docProps/custom.xml", CT_CUSTOM)
        add_root_rel(data, REL_CUSTOM, "docProps/custom.xml")
    c = re.sub(r'<property [^>]*name="' + re.escape(remove_prefix) + r'[^"]*"[^>]*>.*?</property>', "", c, flags=re.S)
    pid = max([int(p) for p in re.findall(r'pid="(\d+)"', c)] or [1])
    props = ""
    for name, val in add.items():
        pid += 1
        props += (f'<property fmtid="{CUSTOM_FMTID}" pid="{pid}" name="{name}">'
                  f"<vt:lpwstr>{xml_text_escape(val)}</vt:lpwstr></property>")
    c = c.replace("</Properties>", props + "</Properties>")
    data["docProps/custom.xml"] = c.encode("utf-8")


def add_override(data, part, ctype):
    ct = data["[Content_Types].xml"].decode("utf-8")
    if f'PartName="{part}"' not in ct:
        ct = ct.replace("</Types>", f'<Override PartName="{part}" ContentType="{ctype}"/></Types>')
        data["[Content_Types].xml"] = ct.encode("utf-8")


def add_root_rel(data, rtype, target):
    r = data["_rels/.rels"].decode("utf-8")
    if f'Target="{target}"' in r:
        return
    ids = [int(i) for i in re.findall(r'Id="rId(\d+)"', r)] or [0]
    r = r.replace("</Relationships>", f'<Relationship Id="rId{max(ids) + 1}" Type="{rtype}" Target="{target}"/></Relationships>')
    data["_rels/.rels"] = r.encode("utf-8")


def mendeley_webext_part(data):
    """Mendeley Cite 추가기능 부분의 경로. 문서에 없으면 새로 만든다."""
    for n in data:
        if re.match(r"word/webextensions/webextension\d+\.xml$", n) and MENDELEY_CITE_ADDIN in data[n].decode("utf-8"):
            return n
    nums = [int(m.group(1)) for n in data for m in [re.match(r"word/webextensions/webextension(\d+)\.xml$", n)] if m]
    n = f"word/webextensions/webextension{max(nums or [0]) + 1}.xml"
    guid = "{" + str(uuid.uuid4()).upper() + "}"
    data[n] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<we:webextension xmlns:we="http://schemas.'
               f'microsoft.com/office/webextensions/webextension/2010/11" id="{guid}"><we:reference id="{MENDELEY_CITE_ADDIN}" '
               f'version="1.55.1.0" store="Omex" storeType="OMEX"/><we:alternateReferences><we:reference id="{MENDELEY_CITE_ADDIN}" '
               f'version="1.55.1.0" store="{MENDELEY_CITE_ADDIN}" storeType="OMEX"/></we:alternateReferences><we:properties>'
               '</we:properties><we:bindings/><we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
               'relationships"/></we:webextension>').encode("utf-8")
    add_override(data, "/" + n, CT_WEBEXT)
    tp, tpr = "word/webextensions/taskpanes.xml", "word/webextensions/_rels/taskpanes.xml.rels"
    if tp not in data:
        data[tp] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<wetp:taskpanes xmlns:wetp="http://schemas.'
                    'microsoft.com/office/webextensions/taskpanes/2010/11"></wetp:taskpanes>').encode("utf-8")
        data[tpr] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.'
                     'openxmlformats.org/package/2006/relationships"></Relationships>').encode("utf-8")
        add_override(data, "/" + tp, CT_TASKPANES)
        add_root_rel(data, REL_TASKPANES, tp)
    rels = data[tpr].decode("utf-8")
    used = [int(i) for i in re.findall(r'Id="rId(\d+)"', rels)] or [0]
    rid = f"rId{max(used) + 1}"
    data[tpr] = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{REL_WEBEXT}" '
                             f'Target="{Path(n).name}"/></Relationships>').encode("utf-8")
    data[tp] = data[tp].decode("utf-8").replace(
        "</wetp:taskpanes>", '<wetp:taskpane dockstate="right" visibility="0" width="350" row="0"><wetp:webextensionref '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/></wetp:taskpane>'
        "</wetp:taskpanes>").encode("utf-8")
    return n


def set_webext_props(data, part, props):
    w = data[part].decode("utf-8")
    for name, val in props.items():
        w = re.sub(r'<we:property name="' + re.escape(name) + r'" value="[^"]*"/>', "", w)
        w = w.replace("</we:properties>", f'<we:property name="{name}" value="{xml_attr_escape(val)}"/></we:properties>')
        w = w.replace("<we:properties/>", f'<we:properties><we:property name="{name}" value="{xml_attr_escape(val)}"/></we:properties>')
    data[part] = w.encode("utf-8")


def style_info(style_id):
    """CSL 스타일의 제목과 인용 형식(author-date, numeric 등). 못 받으면 주소 끝부분으로 대신한다."""
    name = style_id.rstrip("/").split("/")[-1]
    try:
        req = urllib.request.Request(style_id.replace("http://", "https://"), headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=15) as r:
            csl = r.read().decode("utf-8", errors="replace")
        title = html.unescape(re.search(r"<title>(.*?)</title>", csl, re.S).group(1).strip())
        fmt = re.search(r'citation-format="([^"]+)"', csl)
        return title, fmt.group(1) if fmt else "author-date"
    except Exception:
        return name, "author-date"


def cmd_to_mendeley(a):
    """Zotero 로 인용한 원고 사본을 Mendeley Cite 원고로 바꾼다."""
    src, out = Path(a.docx), out_path(a.docx, "_mendeley")
    infos, data = load_docx(src)
    x = data["word/document.xml"].decode("utf-8")
    for n in data:
        if n != "word/document.xml" and n.startswith("word/") and b"ZOTERO_ITEM" in (data[n] or b""):
            print(f"  주의: {n} 의 Zotero 인용(각주 등)은 바꾸지 않음")
    backup = json.loads(latest("backup_*.json").read_text(encoding="utf-8"))
    gid, docs = target_library(backup, default_group(a))
    idx = defaultdict(set)
    for d in docs:
        for k in doc_keys(d):
            idx[k].add(d["id"])

    fields = find_fields(x, "ZOTERO_ITEM")
    cites = []
    for f in fields:
        instr = f[4]
        cites.append(json.loads(instr[instr.index("{"):instr.rindex("}") + 1]))
    # 같은 Zotero 문헌은 한 번만 짝짓는다.
    zitems = {}
    for c in cites:
        for it in c["citationItems"]:
            zitems.setdefault((it.get("uris") or [str(it.get("id"))])[0], it["itemData"])
    mid, missing = {}, []
    for zk, d in zitems.items():
        hits = set().union(*(idx.get(k, set()) for k in csl_keys(d))) if csl_keys(d) else set()
        if len(hits) == 1:
            mid[zk] = next(iter(hits))
        else:
            missing.append((zk, d, "중복" if hits else "없음"))
    lib = default_group(a) or "개인"
    print(f"Zotero 인용 {len(fields)}개, 문헌 {len(zitems)}건 | Mendeley({lib})에서 찾음 {len(mid)} | 못 찾음 {len(missing)}")
    if missing and not a.add_missing:
        for _, d, why in missing:
            print(f"  {why}: {d.get('title', '')[:80]}")
        fail(EXIT_MISSING, "못 찾은 문헌을 Mendeley에 먼저 넣으려면 --add-missing 을 붙여 다시 실행하세요. 파일은 만들지 않음.")
    for zk, d, why in missing:
        if why == "중복":
            sys.exit(f"Mendeley에 같은 문헌이 여러 건 있어 고를 수 없음: {d.get('title', '')[:80]}")
        body = csl_to_mendeley(d)
        if gid:
            body["group_id"] = gid
        _, _, new = api("POST", "/documents", body=body)
        mid[zk] = new["id"]
        DATA.mkdir(exist_ok=True)
        with open(DATA / f"to_mendeley_added_{src.stem[:40]}.jsonl", "a", encoding="utf-8") as log:  # 되돌릴 때 휴지통으로 옮길 목록
            log.write(json.dumps({"id": new["id"], "title": body["title"]}, ensure_ascii=False) + "\n")
        print(f"  Mendeley에 추가: {new['id']} {body['title'][:70]}")

    edits, mcites = [], []
    for (s, sep_end, end_start, e, _), c in zip(fields, cites):
        result = x[sep_end:end_start]
        text = runs_text(result)
        sup = 'w:vertAlign w:val="superscript"' in result
        items = []
        for it in c["citationItems"]:
            m_id = mid[(it.get("uris") or [str(it.get("id"))])[0]]
            new_it = {"id": m_id, "itemData": {**it["itemData"], "id": m_id}, "isTemporary": False}
            new_it.update({k: it[k] for k in CITE_ITEM_EXTRA if k in it})
            items.append(new_it)
        obj = {"citationID": f"MENDELEY_CITATION_{uuid.uuid4()}", "properties": {"noteIndex": 0}, "isEdited": False,
               "manualOverride": {"citeprocText": f"<sup>{text}</sup>" if sup else text,
                                  "isManuallyOverridden": False, "manualOverrideText": ""},
               "citationItems": items}
        tag = CITE_PREFIX + cite_encode(obj)
        mcites.append({**{k: obj[k] for k in ("citationID", "properties", "isEdited", "manualOverride")},
                       "citationTag": tag, "citationItems": items})
        sdt = (f'<w:sdt><w:sdtPr>{first_rpr(result)}<w:tag w:val="{tag}"/><w:id w:val="{secrets.randbelow(2 ** 31)}"/>'
               f"</w:sdtPr><w:sdtContent>{result}</w:sdtContent></w:sdt>")
        edits.append((s, e, sdt))
    for s, sep_end, end_start, e, _ in find_fields(x, "ZOTERO_BIBL"):
        ps, pe = para_start(x, s), para_end(x, end_start)
        inner = x[ps:s] + x[sep_end:end_start] + x[e:pe]  # 필드 표시 런만 빼고 단락은 그대로 둔다
        edits.append((ps, pe, f'<w:sdt><w:sdtPr><w:tag w:val="MENDELEY_BIBLIOGRAPHY"/><w:id w:val="'
                              f'{secrets.randbelow(2 ** 31)}"/></w:sdtPr><w:sdtContent>{inner}</w:sdtContent></w:sdt>'))
        print("  참고문헌 목록을 Mendeley 형식으로 바꿈")
    data["word/document.xml"] = apply_edits(x, edits).encode("utf-8")

    pref = html.unescape("".join(v for _, v in sorted(re.findall(
        r'name="ZOTERO_PREF_(\d+)"><vt:lpwstr>(.*?)</vt:lpwstr>', custom_props(data), re.S), key=lambda t: int(t[0]))))
    sm = re.search(r'<style id="([^"]+)"(?: locale="([^"]*)")?', pref)
    style_id, locale = (sm.group(1), sm.group(2) or "en-US") if sm else ("http://www.zotero.org/styles/apa", "en-US")
    title, fmt = style_info(style_id)
    part = mendeley_webext_part(data)
    set_webext_props(data, part, {
        "MENDELEY_BIBLIOGRAPHY_IS_DIRTY": "true",
        "MENDELEY_BIBLIOGRAPHY_LAST_MODIFIED": str(int(time.time() * 1000)),
        "MENDELEY_CITATIONS": json.dumps(mcites, separators=(",", ":"), ensure_ascii=False),
        "MENDELEY_CITATIONS_LOCALE_CODE": json.dumps(locale),
        "MENDELEY_CITATIONS_STYLE": json.dumps({"id": style_id.replace("http://", "https://"), "title": title,
                                                "format": fmt, "defaultLocale": None, "isLocaleCodeValid": True},
                                               separators=(",", ":"), ensure_ascii=False)})
    if custom_props(data):
        set_custom_props(data, "ZOTERO_PREF_", {})  # Zotero 가 이 문서를 자기 문서로 보지 않게 한다
    save_docx(out, infos, data)
    after = read_docx_citations(out)
    print(f"스타일: {title} ({fmt}, {locale})\n새 파일 → {out} (Mendeley 인용 문헌 {len(after)}건 재확인)")
    print(f"CITES={len(fields)}\nADDED={len(missing)}\nOUT={out}")


def zotero_index():
    """Zotero 문헌을 짝짓기 키로 찾는 색인과 로컬 라이브러리 키."""
    con = sqlite3.connect(zotero_db().as_uri() + "?mode=ro&immutable=1", uri=True)
    local_key = con.execute("select value from settings where setting='account' and key='localUserKey'").fetchone()[0]
    con.close()
    idx = defaultdict(set)
    for it in load_zotero():
        t = zotero_to_mendeley(it)
        a = (t.get("authors") or [{}])[0].get("last_name", "")
        for k in (("doi", t.get("identifiers.doi")), ("tyl", match_key(t.get("title"), t.get("year"), a))):
            if k[1]:
                idx[k].add((it["id"], it["key"]))
    return idx, local_key


def cmd_to_zotero(a):
    """Mendeley Cite 로 인용한 원고 사본을 Zotero 원고로 바꾼다."""
    src, out = Path(a.docx), out_path(a.docx, "_zotero")
    infos, data = load_docx(src)
    x = data["word/document.xml"].decode("utf-8")
    idx, local_key = zotero_index()
    sdts = find_sdts(x, CITE_PREFIX)
    cites = [cite_decode(t[len(CITE_PREFIX):]) for _, _, t, _, _ in sdts]
    found, missing = {}, {}
    for c in cites:
        for it in c["citationItems"]:
            d = it["itemData"]
            hits = set().union(*(idx.get(k, set()) for k in csl_keys(d))) if csl_keys(d) else set()
            if len(hits) == 1:
                found[it["id"]] = next(iter(hits))
            else:
                missing[it["id"]] = d
    print(f"Mendeley 인용 {len(sdts)}개 | Zotero에서 찾음 {len(found)} | 못 찾음 {len(missing)}")
    if missing:
        DATA.mkdir(exist_ok=True)
        mp = DATA / f"to_zotero_missing_{src.stem[:40]}.json"
        mp.write_text(json.dumps(list(missing.values()), ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"  못 찾은 문헌 CSL-JSON → {mp}\n  Zotero: 파일 > 가져오기 로 넣은 뒤 다시 실행하면 연결됨")
        print(f"MISSING={len(missing)}\nMISSING_FILE={mp}")
        if not a.allow_unlinked:
            fail(EXIT_MISSING, "연결 없이 문헌 정보만 넣어 바꾸려면 --allow-unlinked 를 붙이세요. 파일은 만들지 않음.")

    def zitem(it):
        d = dict(it["itemData"])
        z = {k: it[k] for k in CITE_ITEM_EXTRA if k in it}
        if it["id"] in found:
            iid, key = found[it["id"]]
            d["id"] = iid
            return {"id": iid, "uris": [f"http://zotero.org/users/local/{local_key}/items/{key}"], "itemData": d, **z}
        # Zotero 는 라이브러리에 없는 인용도 uris 가 있어야 문서 안 정보로 다룬다(integration.js loadItemData).
        # mendeley.com 주소를 쓰면 Mendeley 가져오기 안내 창이 뜨므로 중립 URN 을 쓴다.
        return {"id": d.get("id", it["id"]), "uris": [f"urn:zotero-mendeley-sync:mendeley:{it['id']}"], "itemData": d, **z}

    def field(rpr, instr, content):
        r = lambda body: f"<w:r>{rpr}{body}</w:r>"
        return (r('<w:fldChar w:fldCharType="begin"/>') + r(f'<w:instrText xml:space="preserve">{xml_text_escape(instr)}</w:instrText>')
                + r('<w:fldChar w:fldCharType="separate"/>') + content + r('<w:fldChar w:fldCharType="end"/>'))

    edits = []
    for (s, e, _, pr, inner), c in zip(sdts, cites):
        text = runs_text(inner)
        obj = {"citationID": secrets.token_hex(4), "properties": {"formattedCitation": text, "plainCitation": text,
                                                                 "noteIndex": 0},
               "citationItems": [zitem(it) for it in c["citationItems"]],
               "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}
        instr = " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(obj, ensure_ascii=False) + " "
        edits.append((s, e, field(first_rpr(pr), instr, inner)))
    for s, e, _, pr, inner in find_sdts(x, "MENDELEY_BIBLIOGRAPHY"):
        paras = list(re.finditer(r"<w:p[ >]", inner))
        if not paras:
            print("  주의: 참고문헌 목록에 단락이 없어 바꾸지 않음")
            continue
        r = lambda body: f"<w:r>{body}</w:r>"
        begin = (r('<w:fldChar w:fldCharType="begin"/>') + r('<w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL '
                 '{"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY </w:instrText>') + r('<w:fldChar w:fldCharType="separate"/>'))
        p0 = paras[0].start()
        open_end = inner.index(">", p0) + 1
        pp = re.match(r"<w:pPr>.*?</w:pPr>|<w:pPr/>", inner[open_end:], re.S)
        ins = open_end + (pp.end() if pp else 0)
        inner = inner[:ins] + begin + inner[ins:]
        last = inner.rindex("</w:p>")
        inner = inner[:last] + r('<w:fldChar w:fldCharType="end"/>') + inner[last:]
        edits.append((s, e, inner))
        print("  참고문헌 목록을 Zotero 형식으로 바꿈")
    data["word/document.xml"] = apply_edits(x, edits).encode("utf-8")

    style = {"id": "http://www.zotero.org/styles/apa"}
    locale = "en-US"
    for n in data:
        if re.match(r"word/webextensions/webextension\d+\.xml$", n) and MENDELEY_CITE_ADDIN in data[n].decode("utf-8"):
            w = data[n].decode("utf-8")
            sm = re.search(r'name="MENDELEY_CITATIONS_STYLE" value="([^"]*)"', w)
            lm = re.search(r'name="MENDELEY_CITATIONS_LOCALE_CODE" value="([^"]*)"', w)
            style = json.loads(html.unescape(sm.group(1))) if sm else style
            locale = json.loads(html.unescape(lm.group(1))) if lm else locale
            set_webext_props(data, n, {"MENDELEY_CITATIONS": "[]"})  # Mendeley Cite 에 옛 인용 목록이 남지 않게
    sid = style["id"].replace("https://www.zotero.org", "http://www.zotero.org")
    pref = (f'<data data-version="3" zotero-version="7.0"><session id="{secrets.token_urlsafe(6)[:8]}"/>'
            f'<style id="{sid}" locale="{locale}" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
            '<prefs><pref name="fieldType" value="Field"/></prefs></data>')
    chunks = [pref[i:i + 255] for i in range(0, len(pref), 255)]
    set_custom_props(data, "ZOTERO_PREF_", {f"ZOTERO_PREF_{i + 1}": ch for i, ch in enumerate(chunks)})
    save_docx(out, infos, data)
    n_after = len(find_fields(data["word/document.xml"].decode("utf-8"), "ZOTERO_ITEM"))
    print(f"스타일: {sid} ({locale})\n새 파일 → {out} (Zotero 인용 {n_after}개 재확인)")
    print(f"CITES={n_after}\nUNLINKED={len(missing)}\nOUT={out}")


# ---------- 앱용 보조 명령 ----------

ZOTERO_CONNECTOR = "http://127.0.0.1:23119/connector"
CSL_TO_ZOTERO_TYPE = {
    "article-journal": "journalArticle", "book": "book", "chapter": "bookSection",
    "paper-conference": "conferencePaper", "report": "report", "thesis": "thesis",
    "webpage": "webpage", "article-magazine": "magazineArticle", "article-newspaper": "newspaperArticle",
    "article": "preprint", "entry-encyclopedia": "encyclopediaArticle",
}
# Zotero 유형마다 '실린 곳' 필드 이름이 다르다. 맞지 않는 필드는 Zotero 가 버리므로 유형별로 고른다.
CONTAINER_FIELD = {"journalArticle": "publicationTitle", "magazineArticle": "publicationTitle",
                   "newspaperArticle": "publicationTitle", "bookSection": "bookTitle",
                   "conferencePaper": "proceedingsTitle", "encyclopediaArticle": "encyclopediaTitle",
                   "webpage": "websiteTitle", "preprint": "repository"}
DOI_TYPES = {"journalArticle", "conferencePaper", "preprint", "report", "thesis", "book", "bookSection"}


def cmd_status(a):
    c = conf_raw()
    print(f"MENDELEY_KEYS={'ok' if c.get('client_id') and c.get('client_secret') else 'missing'}")
    print(f"MENDELEY_LOGIN={'ok' if (CONF_DIR / 'token.json').exists() else 'missing'}")
    print(f"ZOTERO_DIR={zotero_db().parent}")
    print(f"ZOTERO_DB={'ok' if zotero_db().exists() else 'missing'}")
    print(f"GROUP={c.get('group') or ''}")
    print(f"REDIRECT={c.get('redirect_uri') or REDIRECT_DEFAULT}")


def cmd_detect(a):
    """원고가 Zotero 인용인지 Mendeley 인용인지 알려 준다."""
    z = m = 0
    with zipfile.ZipFile(a.docx) as zin:
        for n in docx_parts(zin):
            x = zin.read(n).decode("utf-8")
            z += len(find_fields(x, "ZOTERO_ITEM"))
            m += len(find_sdts(x, CITE_PREFIX))
    print(f"ZOTERO={z}\nMENDELEY={m}")
    print("KIND=" + ("both" if z and m else "zotero" if z else "mendeley" if m else "none"))


def cmd_groups(a):
    for g in list_groups():
        print("GROUP=" + g.get("name", ""))


def csl_to_zotero_item(d):
    t = CSL_TO_ZOTERO_TYPE.get(d.get("type"), "document")
    item = {"itemType": t, "title": d.get("title", ""), "creators": [], "tags": [], "attachments": []}
    for role, ctype in (("author", "author"), ("editor", "editor")):
        for p in d.get(role) or []:
            if p.get("family"):
                item["creators"].append({"creatorType": ctype, "firstName": p.get("given", ""), "lastName": p["family"]})
            elif p.get("literal"):
                item["creators"].append({"creatorType": ctype, "lastName": p["literal"], "fieldMode": 1})
    parts = [str(x) for x in (((d.get("issued") or {}).get("date-parts") or [[]])[0]) if x not in (None, "")]
    if parts:
        item["date"] = "-".join(parts[:1] + [x.zfill(2) for x in parts[1:3]])
    if d.get("container-title") and t in CONTAINER_FIELD:
        item[CONTAINER_FIELD[t]] = d["container-title"]
    for src, dst in (("volume", "volume"), ("issue", "issue"), ("page", "pages"), ("abstract", "abstractNote"),
                     ("language", "language"), ("URL", "url"), ("ISSN", "ISSN"), ("ISBN", "ISBN"),
                     ("publisher", "publisher")):
        if d.get(src):
            item[dst] = str(d[src])
    extra = []
    if d.get("DOI"):
        if t in DOI_TYPES and t not in ("report", "thesis", "book", "bookSection"):
            item["DOI"] = d["DOI"]
        else:
            extra.append(f"DOI: {d['DOI']}")  # DOI 필드가 없는 유형은 Zotero 관례대로 extra 에 둔다
    if d.get("PMID"):
        extra.append(f"PMID: {d['PMID']}")
    if extra:
        item["extra"] = "\n".join(extra)
    return item


def cmd_zotero_import(a):
    """CSL-JSON 문헌을 실행 중인 Zotero 에 넣는다(커넥터 API). Zotero 에서 지금 고른 컬렉션에 들어간다."""
    try:
        urllib.request.urlopen(urllib.request.Request(ZOTERO_CONNECTOR + "/ping", headers={"User-Agent": UA}), timeout=3)
    except Exception:
        fail(EXIT_ZOTERO_OFF, "Zotero 가 실행 중이 아님 — Zotero 를 연 뒤 다시 시도하세요.")
    items = [csl_to_zotero_item(d) for d in json.loads(Path(a.json).read_text(encoding="utf-8")) if d.get("title")]
    for i in range(0, len(items), 50):  # 한 번에 너무 많이 보내지 않는다
        body = json.dumps({"items": items[i:i + 50], "uri": "https://www.mendeley.com/",
                           "sessionID": "zms-" + secrets.token_hex(6)}).encode()
        req = urllib.request.Request(ZOTERO_CONNECTOR + "/saveItems", data=body, headers={
            "Content-Type": "application/json", "X-Zotero-Connector-API-Version": "3", "User-Agent": UA})
        with urllib.request.urlopen(req, timeout=120) as r:
            if r.status not in (200, 201):
                fail(1, f"Zotero 저장 실패({r.status})")
    print(f"IMPORTED={len(items)}")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    cf = sub.add_parser("config", help="설정 저장 (옵션 없으면 터미널에서 묻기)")
    cf.add_argument("--client-id")
    cf.add_argument("--zotero-dir")
    cf.add_argument("--group", help="기본 Mendeley 그룹 이름 (빈 문자열이면 개인 라이브러리)")
    cf.set_defaults(fn=cmd_config)
    au = sub.add_parser("auth")
    au.add_argument("--timeout", type=int, default=0, help="로그인 대기 초 (0이면 무한)")
    au.set_defaults(fn=cmd_auth)
    sub.add_parser("status").set_defaults(fn=cmd_status)
    de = sub.add_parser("detect")
    de.add_argument("docx")
    de.set_defaults(fn=cmd_detect)
    sub.add_parser("groups").set_defaults(fn=cmd_groups)
    zi = sub.add_parser("zotero-import", help="CSL-JSON 문헌을 실행 중인 Zotero 에 넣기")
    zi.add_argument("json")
    zi.set_defaults(fn=cmd_zotero_import)
    sub.add_parser("pull").set_defaults(fn=cmd_pull)
    d = sub.add_parser("diff")
    d.add_argument("--backup", help="백업 JSON (기본: 최신)")
    d.set_defaults(fn=cmd_diff)
    ap = sub.add_parser("apply")
    ap.add_argument("--csv", help="차이표 CSV (기본: 최신)")
    ap.add_argument("--limit", type=int, default=1, help="처리할 문헌 수 (기본 1)")
    ap.add_argument("--all", action="store_true", help="승인된 문헌 전체 처리")
    ap.add_argument("--dry-run", action="store_true", help="보낼 내용만 출력")
    ap.set_defaults(fn=cmd_apply)
    hv = sub.add_parser("harvest", help="공저자 원고에서 내게 없는 문헌 찾아 넣기")
    hv.add_argument("docx")
    hv.add_argument("--group", help="Mendeley 대상 그룹 이름 (기본: 개인 라이브러리)")
    hv.add_argument("--post", action="store_true", help="Mendeley에 실제로 추가")
    hv.add_argument("--limit", type=int, default=1)
    hv.add_argument("--all", action="store_true")
    hv.set_defaults(fn=cmd_harvest)
    rl = sub.add_parser("relink", help="원고 사본의 인용을 대상 라이브러리 문헌으로 다시 연결")
    rl.add_argument("docx")
    rl.add_argument("--group", help="Mendeley 대상 그룹 이름 (기본: 개인 라이브러리)")
    rl.set_defaults(fn=cmd_relink)
    tm = sub.add_parser("to-mendeley", help="Zotero 원고 사본 → Mendeley Cite 원고(_mendeley.docx)")
    tm.add_argument("docx")
    tm.add_argument("--group", help="Mendeley 대상 그룹 이름 (기본: 개인 라이브러리)")
    tm.add_argument("--add-missing", action="store_true", help="Mendeley에 없는 문헌을 먼저 추가")
    tm.set_defaults(fn=cmd_to_mendeley)
    tz = sub.add_parser("to-zotero", help="Mendeley Cite 원고 사본 → Zotero 원고(_zotero.docx)")
    tz.add_argument("docx")
    tz.add_argument("--allow-unlinked", action="store_true", help="Zotero에 없는 문헌도 문서 안 정보로 바꿈")
    tz.set_defaults(fn=cmd_to_zotero)
    a = p.parse_args()
    CONF_DIR.mkdir(parents=True, exist_ok=True)
    a.fn(a)


def run(argv):
    """앱(GUI)에서 명령을 실행한다. 표준 출력이 없는 창 앱에서도 동작하도록 출력을 모아 돌려준다."""
    import contextlib
    import io
    buf = io.StringIO()
    old_argv = sys.argv
    sys.argv = ["zms"] + list(argv)
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            try:
                main()
                code = 0
            except SystemExit as e:
                code = e.code if isinstance(e.code, int) else (0 if e.code is None else 1)
                if isinstance(e.code, str):
                    print(e.code)
            except urllib.error.URLError as e:
                print(f"네트워크 오류: {e}")
                code = 1
            except Exception as e:  # 앱이 죽지 않게 마지막 줄에 오류를 남긴다
                import traceback
                traceback.print_exc()
                print(f"오류: {e}")
                code = 1
    finally:
        sys.argv = old_argv
    return code, buf.getvalue()


if __name__ == "__main__":
    main()
