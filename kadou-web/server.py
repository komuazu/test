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
DATA = WEB / 'data.json'

# ショートカット（.lnk）から確認した実際のパス。
#   リンク先: U:\製造本部\第2工場\【第二工場】\☆第二工場日報・稼動報告\☆第二工場日報・稼動報告
#   U: の実体: \\ntfham001\Users
# U: が割り当てられていないPCでも開けるよう、UNCパスを予備として持たせる。
_BASE = r'製造本部\第2工場\【第二工場】\☆第二工場日報・稼動報告\☆第二工場日報・稼動報告\★新・26年稼動'
DEFAULT_CONFIG = {
    'src': 'U:\\' + _BASE,
    'srcAlt': [r'\\ntfham001\Users' + '\\' + _BASE],
    'year': 2026,
    'port': 8765,
}


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if CONFIG.exists():
        try:
            cfg.update(json.loads(CONFIG.read_text(encoding='utf-8')))
        except Exception:                                        # noqa: BLE001
            print('config.json を読めませんでした。既定値で起動します。')
    return cfg


def save_config(cfg):
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')


def load_memo():
    if MEMO.exists():
        try:
            return json.loads(MEMO.read_text(encoding='utf-8'))
        except Exception:                                        # noqa: BLE001
            print('memo.json を読めませんでした。空で開始します。')
    return {}


_memo_lock = threading.Lock()


def save_memo(memo):
    """memo.json を安全に置き換える（書きかけで壊れないよう一時ファイル経由）"""
    with _memo_lock:
        tmp = MEMO.with_suffix('.json.tmp')
        tmp.write_text(json.dumps(memo, ensure_ascii=False, indent=1), encoding='utf-8')
        os.replace(tmp, MEMO)


def src_candidates(cfg):
    """稼動日報フォルダの候補（本命 → 予備のUNCパス）"""
    return [cfg.get('src')] + list(cfg.get('srcAlt') or [])


def rebuild(cfg):
    """稼動日報フォルダを読み直して web/data.json を作り直す"""
    sys.path.insert(0, str(HERE))
    import build                                                 # noqa: PLC0415
    return build.build(src_candidates(cfg), int(cfg['year']), DATA)


class Handler(SimpleHTTPRequestHandler):
    cfg = None

    def __init__(self, *a, **kw):
        super().__init__(*a, directory=str(WEB), **kw)

    def log_message(self, fmt, *args):
        pass                                     # コンソールを静かに保つ

    # ── 応答ヘルパ ──
    def _json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        n = int(self.headers.get('Content-Length') or 0)
        return json.loads(self.rfile.read(n) or b'{}')

    # ── API ──
    def do_GET(self):
        if self.path.startswith('/api/memo'):
            return self._json(load_memo())
        if self.path.startswith('/api/config'):
            return self._json(Handler.cfg)
        return super().do_GET()

    def do_POST(self):
        try:
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

            if self.path.startswith('/api/rebuild'):
                b = self._body()
                if b.get('src'):
                    Handler.cfg['src'] = b['src']
                if b.get('year'):
                    Handler.cfg['year'] = int(b['year'])
                save_config(Handler.cfg)
                data = rebuild(Handler.cfg)
                return self._json({'ok': True,
                                   'generated': data['generated'],
                                   'records': len(data['records']),
                                   'warnings': data['warnings']})
        except Exception as e:                                   # noqa: BLE001
            traceback.print_exc()
            return self._json({'ok': False, 'error': str(e)}, 500)
        return self._json({'ok': False, 'error': 'unknown endpoint'}, 404)


def free_port(port):
    for p in range(port, port + 20):
        with socket.socket() as s:
            if s.connect_ex(('127.0.0.1', p)) != 0:
                return p
    return port


def main():
    cfg = load_config()
    p = argparse.ArgumentParser(description='稼動日報 Web アプリを起動します')
    p.add_argument('--src', help='稼動日報フォルダ')
    p.add_argument('--year', type=int, help='対象年')
    p.add_argument('--port', type=int, help='ポート番号')
    p.add_argument('--no-build', action='store_true', help='読み直さず前回のデータで起動')
    p.add_argument('--no-browser', action='store_true', help='ブラウザを自動で開かない')
    a = p.parse_args()
    for k in ('src', 'year', 'port'):
        if getattr(a, k):
            cfg[k] = getattr(a, k)
    save_config(cfg)

    if not a.no_build:
        print('稼動日報を読み込んでいます … %s' % cfg['src'])
        try:
            rebuild(cfg)
            print('読み込み完了\n')
        except SystemExit as e:
            print('  [エラー] %s' % e)
            if not DATA.exists():
                print('\n  フォルダの場所が違う可能性があります。')
                print('  画面右上の「データ更新」からフォルダを指定し直すこともできます。\n')
        except Exception as e:                                   # noqa: BLE001
            traceback.print_exc()
            print('  [エラー] 読み込みに失敗しました: %s' % e)

    Handler.cfg = cfg
    port = free_port(int(cfg['port']))
    url = 'http://127.0.0.1:%d/' % port
    srv = ThreadingHTTPServer(('127.0.0.1', port), partial(Handler))
    print('=' * 60)
    print('  稼動日報 %d年 印刷実績ビューア' % cfg['year'])
    print('  ブラウザで開いてください →  %s' % url)
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
