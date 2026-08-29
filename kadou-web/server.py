# -*- coding: utf-8 -*-
"""
稼動日報 Web アプリ  ローカルサーバー

    python server.py                      # config.json の設定で起動
    python server.py --src "U:\\...\\★新・26年稼動" --year 2026
    python server.py --no-build           # 再読込せず前回のデータで起動

標準ライブラリのみで動く（データ生成には xlrd が必要）。
社内 PC のブラウザから http://127.0.0.1:8765/ で開く。

【元ファイル保護】
  ・稼動日報フォルダは読み取りのみ。書き込み先はこのアプリのフォルダに限定。
  ・画面で入力した「今年の動向 / 代替対策 / 対策通し数」は memo.json に保存する。
    元の .xls には一切書き戻さない。
"""
import argparse
import json
import os
import secrets
import socket
import sys
import threading
import traceback
import webbrowser
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).resolve().parent
WEB = HERE / 'web'
CONFIG = HERE / 'config.json'
MEMO = HERE / 'memo.json'
MEMO_PREV = HERE / 'memo_前回.json'
# アプリのフォルダを入れ替えても記入欄が消えないよう、ユーザーフォルダにも控える
MEMO_HOME = Path(os.environ.get('APPDATA') or Path.home()) / 'kadou-web' / 'memo.json'
MANIFEST = WEB / 'years.json'

# 稼動日報の置き場所。実在するフォルダを全部読み、年ごとにまとめる。
#   ・デスクトップの「☆平版印刷課稼働日報」（今年ぶん＋過去分）
#   ・U: の「☆第二工場日報・稼動報告」（同じものの元）
# 年フォルダ（★新・NN年稼動 / NN年稼働）は配下から探すので、年が変わって
# 入れ物のフォルダ名が「13年～26年稼働」に変わっても、「27年稼働」が増えても、
# ここを直す必要はない。U: の実体は \\ntfham001\Users（ショートカットから確認）。
# U: が割り当てられていないPCでも開けるようUNCパスも並べてある。
# 同じ年が複数のフォルダにあるときは、先に書いたほうから1つだけ読む。
_U = r'製造本部\第2工場\【第二工場】\☆第二工場日報・稼動報告\☆第二工場日報・稼動報告'
DEFAULT_CONFIG = {
    'src': [
        str(Path.home() / 'Desktop' / '☆平版印刷課稼働日報'),
        str(Path.home() / 'Desktop' / '13年稼動～25年稼動'),
        'U:\\' + _U,
        r'\\ntfham001\Users' + '\\' + _U,
    ],
    'year': None,        # None = 見つかった年を全部読む
    'port': 8765,
    # 稼働率の分母から外す会社の休業日（お盆・年末年始など）。1行に1つ。
    #   2026-08-13 ／ 08-13（毎年） ／ 08-13..15（範囲）
    #   12-29..01-03（年をまたぐ範囲）／ 後ろに覚え書き可
    # 土日祝は書かなくても自動で外れる。
    'closed': [],
}


_QUOTES = '"' + "'" + '“”「」'      # 貼り付けで付いてくる引用符


def clean_src(s):
    r"""フォルダのパスの前後の空白と引用符を落とす

    エクスプローラーの「パスのコピー」は "C:\..." と引用符付きで入るため、
    そのまま設定に入れると実在するフォルダでも見つからなくなる。
    build.clean_src と同じ処理（xlrd 無しでも起動できるよう別に持つ）。
    """
    return str(s).strip().strip(_QUOTES)


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if CONFIG.exists():
        try:
            cfg.update(json.loads(CONFIG.read_text(encoding='utf-8')))
        except Exception:                                        # noqa: BLE001
            print('config.json を読めませんでした。既定値で起動します。')
    return cfg


def save_config(cfg):
    """設定を config.json に書く（フォルダのパスは引用符を落として揃える）"""
    src = cfg.get('src')
    if isinstance(src, list):
        cfg['src'] = [x for x in (clean_src(c) for c in src if c) if x]
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')


# ── 共有モード（社内サーバーに置いて、みんなで使うとき） ──
# 自分のPCで使うとき（既定）は、これまでどおり 127.0.0.1 だけで待ち受け、
# フォルダの指定も自由にできる。
# --host で外から届くようにしたときは自動で共有モードになり、
#   ・合い言葉（パスワード）を求める
#   ・稼動日報フォルダの変更と「フォルダを選ぶ」を止める
# 誰でもサーバー上の好きなフォルダを読めてしまわないようにするため。
SHARED = False
PASSWORD = ''
# サーバーにしているPC自身（127.0.0.1）からの操作は、これまでどおり扱う。
# その画面を開けるのはPCの前にいる人だけなので、合い言葉を求めず、
# 稼動日報フォルダの変更もできるようにする。テストのときだけ False にする。
TRUST_LOCAL = True
SESSIONS = set()
PWFILE = HERE / 'password.txt'


def load_password(arg=None):
    """合い言葉を、指定・環境変数・password.txt の順に探す"""
    if arg:
        return arg.strip()
    env = os.environ.get('KADOU_PASSWORD')
    if env:
        return env.strip()
    if PWFILE.exists():
        return PWFILE.read_text(encoding='utf-8').strip()
    return ''


def read_json(path):
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except Exception:                                            # noqa: BLE001
        return None


def load_memo():
    """記入欄の内容を読む。中身の多い方を採る

    アプリのフォルダ（memo.json）と、PCのユーザーフォルダ（MEMO_HOME）の両方に
    同じものを書いている。アプリを入れ替えたり作り直したりしてフォルダ側が
    無くなっても、ユーザーフォルダ側から戻せるようにするため。
    """
    a, b = read_json(MEMO), read_json(MEMO_HOME)
    if a is None and b is None:
        if MEMO.exists() or MEMO_HOME.exists():
            print('記入欄の保存ファイルを読めませんでした。空で開始します。')
        return {}
    if a is None:
        print('記入欄の内容を %s から戻しました。' % MEMO_HOME)
        return b
    if b is not None and len(b) > len(a):
        print('記入欄の内容を %s から戻しました（%d件）。' % (MEMO_HOME, len(b)))
        return b
    return a


_memo_lock = threading.Lock()


def write_atomic(path, text):
    """書きかけで壊れないよう、一時ファイルに書いてから置き換える"""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + '.tmp')
    tmp.write_text(text, encoding='utf-8')
    os.replace(tmp, path)


def save_memo(memo):
    """記入欄の内容を保存する

    消えると手入力がやり直しになるため、3か所に残す。
      ・memo.json           … アプリのフォルダ
      ・memo_前回.json      … 上書きする前の内容（打ち間違いの戻し用）
      ・ユーザーフォルダ     … アプリを入れ替えても残る控え
    """
    with _memo_lock:
        body = json.dumps(memo, ensure_ascii=False, indent=1)
        if MEMO.exists():
            try:
                write_atomic(MEMO_PREV, MEMO.read_text(encoding='utf-8'))
            except OSError:
                pass
        write_atomic(MEMO, body)
        try:
            write_atomic(MEMO_HOME, body)
        except OSError as e:                                     # noqa: BLE001
            print('控えを %s に書けませんでした: %s' % (MEMO_HOME, e))


_pick_lock = threading.Lock()
_build_lock = threading.Lock()   # 読み込みは一度に1つだけ


def pick_folder(initial=None):
    """稼動日報フォルダを選ぶダイアログを出し、選ばれたパスを返す

    サーバーはこのPCの上で動いているので、ブラウザの「フォルダを選ぶ」から
    Windows のフォルダ選択ダイアログを出せる。パスを手で書き写さずに済む。
    選ばずに閉じたときは None。
    """
    import tkinter                                               # noqa: PLC0415
    from tkinter import filedialog                               # noqa: PLC0415

    start = clean_src(initial or '')
    while start and not Path(start).is_dir():                    # 上の階層まで戻って開く
        parent = str(Path(start).parent)
        start = '' if parent == start else parent

    root = tkinter.Tk()
    root.withdraw()
    root.attributes('-topmost', True)                            # ブラウザの後ろに隠れないように
    try:
        got = filedialog.askdirectory(title='稼動日報フォルダを選んでください',
                                      initialdir=start or str(Path.home() / 'Desktop'),
                                      mustexist=True)
    finally:
        root.destroy()
    return str(Path(got)) if got else None                       # 区切りを \ に揃える


def src_candidates(cfg):
    """稼動日報フォルダの一覧（実在するものだけが実際に読まれる）"""
    src = cfg.get('src')
    lst = list(src) if isinstance(src, list) else ([src] if src else [])
    lst += list(cfg.get('srcAlt') or [])            # 旧い config.json との互換
    return [x for x in (clean_src(c) for c in lst if c) if x]


def print_sources(cands):
    """どのフォルダを読むかを表示する

    名前の表記ゆれ（全角チルダ「～」と波ダッシュ「〜」など）を吸収して
    別名で見つかった場合は、実際に読むフォルダ名も出す。
    """
    sys.path.insert(0, str(HERE))
    import build                                                 # noqa: PLC0415
    for c in cands:
        p = Path(c)
        if p.is_dir():
            print('    %s' % c)
            continue
        hit = build.match_by_name(p)
        if hit is not None:
            print('    %s\n      → %s として読み込みます' % (c, hit))
        else:
            print('    %s  （見つかりません）' % c)


def rebuild(cfg):
    """稼動日報フォルダを読み直して web/ のデータを作り直す

    year が空なら、フォルダに入っている年を全部作る（13年〜25年のような
    複数年フォルダに対応するため）。
    """
    sys.path.insert(0, str(HERE))
    import build                                                 # noqa: PLC0415
    year = cfg.get('year')
    return build.build(src_candidates(cfg), int(year) if year else None, WEB,
                       cfg.get('closed'))


LOGIN_HTML = """<!doctype html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>稼動日報 印刷実績ビューア</title>
<style>
 body{font-family:"Yu Gothic UI","Hiragino Kaku Gothic ProN",system-ui,sans-serif;
   background:#f1f1ef; color:#1a1a1a; display:flex; min-height:100vh; margin:0;
   align-items:center; justify-content:center}
 form{background:#fff; border:1px solid #dcdfe3; padding:28px 26px; min-width:280px}
 h1{font-size:15px; margin:0 0 18px}
 input{width:100%; box-sizing:border-box; font-size:16px; padding:9px 10px;
   border:1px solid #9ba1a8; border-radius:2px}
 button{margin-top:14px; width:100%; font-size:15px; padding:9px; cursor:pointer;
   background:#3a4450; color:#fff; border:0; border-radius:2px}
 p{color:#992222; font-size:12.5px; margin:12px 0 0; min-height:1em}
</style></head><body>
<form onsubmit="go(event)">
  <h1>稼動日報 印刷実績ビューア</h1>
  <input id="p" type="password" placeholder="合い言葉" autofocus autocomplete="current-password">
  <button type="submit">開く</button>
  <p id="m">{msg}</p>
</form>
<script>
function go(e){
  e.preventDefault();
  fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({password:document.getElementById('p').value})})
   .then(function(r){return r.json();})
   .then(function(j){ if(j.ok){ location.reload(); }
     else { document.getElementById('m').textContent = j.error || '入れませんでした'; } })
   .catch(function(){ document.getElementById('m').textContent='つながりませんでした'; });
}
</script></body></html>"""


class Handler(SimpleHTTPRequestHandler):
    cfg = None

    # HTTP/1.1 で応答する。既定の HTTP/1.0 だと応答のたびに接続を切るため、
    # data_<年>.json のような大きなファイル（10MB前後）を送っている途中で
    # 切断が先に届き、ブラウザ側が「読み込み中のまま止まる」ことがあった。
    protocol_version = 'HTTP/1.1'

    def __init__(self, *a, **kw):
        super().__init__(*a, directory=str(WEB), **kw)

    def log_message(self, fmt, *args):
        pass                                     # コンソールを静かに保つ

    def end_headers(self):
        # ブラウザにためこませない。index.html / app.js / style.css は名前が
        # 変わらないため、キャッシュが効くとアプリを更新しても古い画面が出続ける
        # （実際に、更新したのに「何も変わっていない」状態になった）。
        # 手元のフォルダを読むだけなので、毎回読み直しても遅くならない。
        self.send_header('Cache-Control', 'no-store, must-revalidate')
        super().end_headers()

    # ── 応答ヘルパ ──
    def _json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        n = int(self.headers.get('Content-Length') or 0)
        return json.loads(self.rfile.read(n) or b'{}')

    # ── 合い言葉 ──
    def token(self):
        for part in (self.headers.get('Cookie') or '').split(';'):
            k, _, v = part.strip().partition('=')
            if k == 'kadou':
                return v
        return ''

    def here(self):
        """サーバーにしているPC自身から開いているか"""
        return TRUST_LOCAL and self.client_address[0] in ('127.0.0.1', '::1')

    def guest(self):
        """ほかの端末から見ている状態か（共有中に外から来た人）"""
        return SHARED and not self.here()

    def allowed(self):
        """共有モードでは、合い言葉を入れた人だけが中を見られる"""
        return (not self.guest()) or (not PASSWORD) or (self.token() in SESSIONS)

    def send_login(self, msg=''):
        body = LOGIN_HTML.replace('{msg}', msg).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_bytes(self, data, ctype, gzipped=False):
        self.send_response(200)
        self.send_header('Content-Type', ctype)
        if gzipped:
            self.send_header('Content-Encoding', 'gzip')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    # ── API ──
    def do_GET(self):
        if not self.allowed():
            return self.send_login()
        if self.path.startswith('/api/memo'):
            return self._json(load_memo())
        if self.path.startswith('/api/config'):
            return self._json(dict(Handler.cfg, shared=self.guest()))
        # data_<年>.json は10MB前後あるので、縮めて送る（社内LANでも効く）
        name = self.path.split('?')[0].lstrip('/')
        if name.endswith('.json') and 'gzip' in (self.headers.get('Accept-Encoding') or ''):
            f = Path(self.directory) / name
            if f.is_file():
                import gzip                                       # noqa: PLC0415
                return self.send_bytes(gzip.compress(f.read_bytes(), 6),
                                       'application/json; charset=utf-8', True)
        return super().do_GET()

    def do_POST(self):
        try:
            if self.path.startswith('/api/login'):
                if self._body().get('password', '') == PASSWORD and PASSWORD:
                    tok = secrets.token_urlsafe(24)
                    SESSIONS.add(tok)
                    body = json.dumps({'ok': True}).encode('utf-8')
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Content-Length', str(len(body)))
                    self.send_header('Set-Cookie',
                                     'kadou=%s; Path=/; HttpOnly; SameSite=Lax' % tok)
                    self.end_headers()
                    self.wfile.write(body)
                    return
                return self._json({'ok': False, 'error': '合い言葉が違います。'})

            if not self.allowed():
                return self._json({'ok': False, 'error': 'もう一度開き直してください。'}, 401)

            if self.path.startswith('/api/memo'):
                # {key: {trend, plan, tsu}} を差分マージする
                memo = load_memo()
                for k, v in self._body().items():
                    if v and any(str(x).strip() for x in v.values()):
                        memo[k] = v
                    else:
                        memo.pop(k, None)
                save_memo(memo)
                return self._json({'ok': True, 'count': len(memo)})

            if self.path.startswith('/api/xlsx'):
                # 画面の表をそのまま Excel ファイルにして返す
                b = self._body()
                import xlsx                                      # noqa: PLC0415
                data = xlsx.build(b.get('rows') or [],
                                  b.get('sheet') or '集計')
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-'
                                                 'officedocument.spreadsheetml.sheet')
                self.send_header('Content-Length', str(len(data)))
                self.send_header('Content-Disposition', 'attachment; filename="table.xlsx"')
                self.end_headers()
                self.wfile.write(data)
                return

            if self.path.startswith('/api/pick'):
                if self.guest():
                    return self._json({'ok': False,
                                       'error': '共有で使っているときは、フォルダを'
                                                'この画面から選べません。'})
                # ダイアログは1つずつ（2つ開くと閉じられなくなるため）
                if not _pick_lock.acquire(blocking=False):
                    return self._json({'ok': False,
                                       'error': 'フォルダを選ぶ画面がすでに開いています。'})
                try:
                    return self._json({'ok': True,
                                       'path': pick_folder(self._body().get('initial'))})
                except Exception as e:                           # noqa: BLE001
                    traceback.print_exc()
                    return self._json({'ok': False,
                                       'error': 'フォルダを選ぶ画面を開けませんでした（%s）。'
                                                'お手数ですが欄に直接ご記入ください。' % e})
                finally:
                    _pick_lock.release()

            if self.path.startswith('/api/rebuild'):
                b = self._body()
                if self.guest():
                    # ほかの端末からは、読むフォルダを変えさせない。
                    # 画面から変えられると、サーバー上の好きなフォルダを
                    # 読めてしまうため。
                    b = {}
                if not _build_lock.acquire(blocking=False):
                    return self._json({'ok': False,
                                       'error': 'いま読み込み中です。'
                                                '終わるまでお待ちください。'})
                if b.get('src') is not None:
                    # 画面からは1行1フォルダで来る
                    v = b['src']
                    if isinstance(v, str):
                        v = [x.strip() for x in v.splitlines() if x.strip()]
                    if v:
                        Handler.cfg['src'] = v
                        Handler.cfg.pop('srcAlt', None)
                # 会社の休業日（空欄可）。画面からは1行1日で来る
                if 'closed' in b:
                    v = b['closed']
                    if isinstance(v, str):
                        v = [x.strip() for x in v.splitlines() if x.strip()]
                    Handler.cfg['closed'] = list(v or [])
                # 年は空欄可（空 = フォルダに入っている年を全部）
                if 'year' in b:
                    Handler.cfg['year'] = int(b['year']) if str(b['year']).strip() else None
                save_config(Handler.cfg)
                try:
                    m = rebuild(Handler.cfg)
                finally:
                    _build_lock.release()
                return self._json({'ok': True,
                                   'generated': m['generated'],
                                   'years': [y['year'] for y in m['years']],
                                   'records': sum(y['records'] for y in m['years']),
                                   'warnings': m['warnings']})
        except Exception as e:                                   # noqa: BLE001
            traceback.print_exc()
            return self._json({'ok': False, 'error': str(e)}, 500)
        return self._json({'ok': False, 'error': 'unknown endpoint'}, 404)


def my_ip():
    """ほかの端末から届くアドレスを調べる（案内に出すだけ）"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s2:
            s2.connect(('8.8.8.8', 80))          # つながらなくてもよい
            return s2.getsockname()[0]
    except OSError:
        return socket.gethostname()


def free_port(port):
    for p in range(port, port + 20):
        with socket.socket() as s:
            if s.connect_ex(('127.0.0.1', p)) != 0:
                return p
    return port


def main():
    cfg = load_config()
    p = argparse.ArgumentParser(description='稼動日報 Web アプリを起動します')
    p.add_argument('--src', action='append',
                   help='稼動日報フォルダ（複数指定可: --src A --src B）')
    p.add_argument('--year', type=int, help='対象年')
    p.add_argument('--port', type=int, help='ポート番号')
    p.add_argument('--no-build', action='store_true', help='読み直さず前回のデータで起動')
    p.add_argument('--no-browser', action='store_true', help='ブラウザを自動で開かない')
    p.add_argument('--host', default='127.0.0.1',
                   help='待ち受けるアドレス（社内で共有するときは 0.0.0.0）')
    p.add_argument('--password', help='共有するときの合い言葉'
                                      '（省略時は KADOU_PASSWORD か password.txt）')
    a = p.parse_args()

    # 自分のPC以外から届くようにしたときは、自動で共有モードにする
    global SHARED, PASSWORD                                      # noqa: PLW0603
    SHARED = a.host not in ('127.0.0.1', 'localhost')
    PASSWORD = load_password(a.password)
    if SHARED and not PASSWORD:
        raise SystemExit(
            '合い言葉が設定されていないため、共有で起動できません。\n'
            '  次のどれかで決めてください。\n'
            '    ・password.txt に1行で書く（%s）\n'
            '    ・環境変数 KADOU_PASSWORD に入れる\n'
            '    ・--password で渡す' % PWFILE)
    for k in ('src', 'year', 'port'):
        if getattr(a, k):
            cfg[k] = getattr(a, k)
    save_config(cfg)

    if not a.no_build:
        print('稼動日報を読み込んでいます …')
        print_sources(src_candidates(cfg))
        try:
            rebuild(cfg)
            print('読み込み完了\n')
        except SystemExit as e:
            print('  [エラー] %s' % e)
            if not MANIFEST.exists():
                print('\n  フォルダの場所が違う可能性があります。')
                print('  画面右上の「データ更新」からフォルダを指定し直すこともできます。\n')
        except Exception as e:                                   # noqa: BLE001
            traceback.print_exc()
            print('  [エラー] 読み込みに失敗しました: %s' % e)

    Handler.cfg = cfg
    port = free_port(int(cfg['port'])) if not SHARED else int(cfg['port'])
    url = 'http://127.0.0.1:%d/' % port
    try:
        srv = ThreadingHTTPServer((a.host, port), partial(Handler))
    except OSError as e:
        raise SystemExit(
            '%d番のポートを使えませんでした（%s）。\n'
            '  ・すでにこのアプリが動いていないか確かめてください\n'
            '  ・別の番号にするには config.json の "port" を変えてください'
            % (port, e))
    print('=' * 60)
    print('  稼動日報 印刷実績ビューア%s'
          % ('（%d年）' % cfg['year'] if cfg.get('year') else ''))
    print('  ブラウザで開いてください →  %s' % url)
    if SHARED:
        print('  ほかの端末からは →  http://%s:%d/' % (my_ip(), port))
        print('  合い言葉を入れると開きます（フォルダの変更はできません）')
    print('  終了するには この画面で Ctrl+C を押してください')
    print('=' * 60)
    if not a.no_browser:
        threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print('\n終了しました。')


if __name__ == '__main__':
    main()
