#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Piiicker — локальный сервер ревизии. Запуск: python server.py"""
import csv, json, os, re, secrets, socket, sqlite3, threading, time
from datetime import datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

HOST, PORT = '0.0.0.0', 8000
BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, 'piiicker.db')
EMP = os.path.join(BASE, 'employees.csv')
CAT = os.path.join(BASE, 'catalog.csv')
BLOCKED = {'server.py', 'employees.csv', 'catalog.csv', 'piiicker.db',
           'piiicker.db-wal', 'piiicker.db-shm'}

PICK_RE = re.compile(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3}$')

lock = threading.Lock()
sessions = {}      # token -> (code, name)
active_dev = {}    # code -> device_id
catalog = {'columns': [], 'rows': {}}


# ---------- Время ----------
def now_ms():
    t = time.time()
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(t)) + '.%03d' % int((t % 1) * 1000)


def seconds_between(t1, t2):
    f = '%Y-%m-%d %H:%M:%S.%f'
    try:
        return round((datetime.strptime(t2, f) - datetime.strptime(t1, f)).total_seconds(), 3)
    except Exception:
        return None


# ---------- Устройство ----------
def parse_ua(ua):
    if not ua:
        return ''
    os_ = next((v for k, v in [('Windows', 'Windows'), ('Android', 'Android'),
                               ('iPhone', 'iPhone'), ('iPad', 'iPad'),
                               ('Mac OS', 'macOS'), ('Linux', 'Linux')]
                if k in ua), '?')
    br = next((v for k, v in [('Edg/', 'Edge'), ('Chrome/', 'Chrome'),
                              ('Firefox/', 'Firefox'), ('Safari/', 'Safari')]
               if k in ua), '?')
    return os_ + '/' + br


# ---------- Сотрудники ----------
def load_employees():
    res = {}
    if not os.path.exists(EMP):
        return res
    with open(EMP, encoding='utf-8-sig') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            parts = line.split(';', 1) if ';' in line else line.split(',', 1)
            code = parts[0].strip()
            name = parts[1].strip() if len(parts) > 1 else ''
            if code:
                res[code] = name
    return res


# ---------- Каталог ----------
def load_catalog():
    global catalog
    if not os.path.exists(CAT):
        catalog = {'name_col': '', 'columns': [], 'rows': {}}
        return
    text = None
    for enc in ('utf-8-sig', 'cp1251'):
        try:
            with open(CAT, encoding=enc, newline='') as f:
                text = f.read()
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        print('!! Не удалось прочитать catalog.csv — неизвестная кодировка.')
        catalog = {'name_col': '', 'columns': [], 'rows': {}}
        return

    import io
    r = csv.reader(io.StringIO(text), delimiter=';')
    header = next(r, None)
    if not header or len(header) < 2:
        print('!! catalog.csv: нужно минимум 2 столбца (код и имя).')
        catalog = {'name_col': '', 'columns': [], 'rows': {}}
        return

    header = [h.strip() for h in header]
    # Позиция 0 — код (ключ), позиция 1 — имя (показывается всегда).
    # Всё, что дальше, — столбцы для меню в compare.html.
    name_col = header[1]
    other = header[2:]

    rows = {}
    for row in r:
        if not row or not row[0].strip():
            continue
        code = row[0].strip()
        item = {}
        for i in range(1, len(header)):
            item[header[i]] = row[i].strip() if i < len(row) else ''
        rows[code] = item

    catalog = {'name_col': name_col, 'columns': other, 'rows': rows}
    
# ---------- База ----------
def db_init():
    c = sqlite3.connect(DB)
    c.execute('PRAGMA journal_mode=WAL')
    c.executescript('''
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pick_time TEXT NOT NULL,
            employee_code TEXT NOT NULL,
            employee_name TEXT NOT NULL,
            value TEXT NOT NULL,
            display_name TEXT DEFAULT '',
            UNIQUE(pick_time, employee_code, value)
        );
        CREATE TABLE IF NOT EXISTS compare_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            check_time TEXT NOT NULL,
            employee_code TEXT NOT NULL,
            employee_name TEXT NOT NULL,
            device_id TEXT DEFAULT '',
            value TEXT NOT NULL,
            display_name TEXT DEFAULT '',
            op TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            server_time TEXT NOT NULL,
            pick_time TEXT,
            delay_sec REAL,
            employee_code TEXT,
            employee_name TEXT,
            device_id TEXT,
            device_info TEXT,
            op TEXT NOT NULL,
            detail TEXT
        );
    ''')
    c.commit()
    c.close()


def db(q, params=(), fetch=False):
    with lock:
        c = sqlite3.connect(DB, timeout=10)
        try:
            with c:
                cur = c.execute(q, params)
                return cur.fetchall() if fetch else cur.lastrowid
        finally:
            c.close()


def log(op, code='', name='', device_id='', device_info='', pick_time=None, detail=''):
    st = now_ms()
    delay = seconds_between(pick_time, st) if pick_time else None
    db('INSERT INTO events(server_time,pick_time,delay_sec,employee_code,'
       'employee_name,device_id,device_info,op,detail) VALUES(?,?,?,?,?,?,?,?,?)',
       (st, pick_time, delay, code, name, device_id, device_info, op, detail))


# ---------- Сеть ----------
def default_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        return s.getsockname()[0]
    except Exception:
        return None
    finally:
        s.close()


def local_ips():
    ips = []
    ip = default_ip()
    if ip and not ip.startswith('127.'):
        ips.append(ip)
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            x = info[4][0]
            if not x.startswith(('127.', '169.254.')) and x not in ips:
                ips.append(x)
    except Exception:
        pass
    return ips


# ---------- HTTP ----------
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _json(self, obj, code=200, cookie=None):
        b = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(b)))
        if cookie:
            self.send_header('Set-Cookie', cookie)
        self.end_headers()
        self.wfile.write(b)

    def _file(self, path):
        if not os.path.isfile(path):
            return self.send_error(404)
        if path.endswith('.html'):
            mime = 'text/html; charset=utf-8'
        elif path.endswith('.css'):
            mime = 'text/css; charset=utf-8'
        else:
            mime = 'application/octet-stream'
        with open(path, 'rb') as f:
            data = f.read()
        self.send_response(200)
        self.send_header('Content-Type', mime)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(data)

    def _redirect(self, to):
        self.send_response(302)
        self.send_header('Location', to)
        self.end_headers()

    def _session(self):
        for p in self.headers.get('Cookie', '').split(';'):
            p = p.strip()
            if p.startswith('emp='):
                return sessions.get(p[4:])
        return None

    def _did(self):
        return self.headers.get('X-Device-Id', '')[:32]

    def _ua(self):
        return parse_ua(self.headers.get('User-Agent', ''))

    def do_GET(self):
        p = urlparse(self.path).path
        if p in ('/', '/login.html'):
            return self._file(os.path.join(BASE, 'login.html'))
        if p == '/tsd':
            if not self._session():
                return self._redirect('/')
            return self._file(os.path.join(BASE, 'tsd.html'))
        if p == '/compare':
            if not self._session():
                return self._redirect('/')
            return self._file(os.path.join(BASE, 'compare.html'))
        if p == '/shared.css':
            return self._file(os.path.join(BASE, 'shared.css'))
        if p == '/api/state':
            s = self._session()
            if not s:
                return self._json({'ok': False, 'error': 'unauthorized'}, 401)
            code, name = s
            rows = db('SELECT pick_time, value, display_name FROM logs '
                      'WHERE employee_code=? ORDER BY id DESC LIMIT 500',
                      (code,), fetch=True)
            total = db('SELECT COUNT(*) FROM logs WHERE employee_code=?',
                       (code,), fetch=True)[0][0]
            return self._json({'ok': True, 'employee_code': code, 'employee_name': name,
                               'count': total,
                               'entries': [{'time': r[0], 'value': r[1],
                                            'display': r[2] or ''} for r in rows]})
        if p == '/api/catalog':
            if not self._session():
                return self._json({'ok': False, 'error': 'unauthorized'}, 401)
            return self._json({'ok': True,
                   'name_col': catalog['name_col'],
                   'columns': catalog['columns'],
                   'rows': catalog['rows']})
        if p == '/api/ping':
            return self._json({'ok': True})
        safe = p.lstrip('/')
        if '..' in safe or safe in BLOCKED:
            return self.send_error(403)
        fp = os.path.join(BASE, safe)
        if os.path.isfile(fp):
            return self._file(fp)
        self.send_error(404)

    def do_POST(self):
        try:
            n = int(self.headers.get('Content-Length') or 0)
            data = json.loads(self.rfile.read(n) or b'{}')
        except Exception:
            data = {}
        path = urlparse(self.path).path
        fn = {'/api/login': self._login, '/api/logout': self._logout,
              '/api/log': self._log, '/api/undo': self._undo,
              '/api/check': self._check, '/api/mismatch': self._mismatch,
              '/api/diag': self._diag}.get(path)
        if not fn:
            return self.send_error(404)
        fn(data)

    def _login(self, data):
        code = str(data.get('code', '')).strip()
        did, info, ip = self._did(), self._ua(), self.client_address[0]
        allowed = load_employees()
        if code and code in allowed:
            name = allowed[code]
            prev = active_dev.get(code)
            if prev and prev != did:
                log('multi_login', code, name, did, info, detail='prev=%s' % prev)
            active_dev[code] = did
            token = secrets.token_urlsafe(24)
            sessions[token] = (code, name)
            log('auth_ok', code, name, did, info, detail=ip)
            return self._json({'ok': True, 'employee_code': code, 'employee_name': name},
                              cookie='emp=%s; Path=/; HttpOnly; SameSite=Lax' % token)
        log('auth_fail', code or '(empty)', '', did, info, detail=ip)
        self._json({'ok': False, 'error': 'Неверный код сотрудника'}, 401)

    def _logout(self, data=None):
        s = self._session()
        for p in self.headers.get('Cookie', '').split(';'):
            p = p.strip()
            if p.startswith('emp='):
                sessions.pop(p[4:], None)
        if s:
            log('logout', s[0], s[1], self._did(), self._ua())
            active_dev.pop(s[0], None)
        self._json({'ok': True}, cookie='emp=; Path=/; Max-Age=0')

    def _log(self, data):
        s = self._session()
        if not s:
            return self._json({'ok': False, 'error': 'unauthorized'}, 401)
        code, name = s
        did, info = self._did(), self._ua()

        items = data.get('entries')
        if items is None:
            items = [{'time': data.get('time', ''), 'value': data.get('value', ''),
                      'display': data.get('display', '')}]

        saved = 0
        for it in items[:500]:
            value = str(it.get('value', '')).strip()
            pt = str(it.get('time', '')).strip()
            disp = str(it.get('display', ''))[:200]
            if not value:
                continue
            if not PICK_RE.match(pt):
                log('bad_time', code, name, did, info, detail=pt[:64])
                continue
            try:
                db('INSERT INTO logs(pick_time,employee_code,employee_name,value,display_name) '
                   'VALUES(?,?,?,?,?)', (pt, code, name, value, disp))
                log('add', code, name, did, info, pick_time=pt, detail=value)
                saved += 1
            except sqlite3.IntegrityError:
                log('duplicate', code, name, did, info, pick_time=pt, detail=value)
        self._json({'ok': True, 'saved': saved})

    def _undo(self, data=None):
        s = self._session()
        if not s:
            return self._json({'ok': False, 'error': 'unauthorized'}, 401)
        code, name = s
        row = db('''DELETE FROM logs WHERE id IN (
                        SELECT id FROM logs WHERE employee_code=?
                        ORDER BY id DESC LIMIT 1)
                    RETURNING pick_time, value''', (code,), fetch=True)
        if not row:
            return self._json({'ok': False, 'error': 'empty'}, 404)
        pt, val = row[0]
        log('undo', code, name, self._did(), self._ua(), pick_time=pt, detail=val)
        self._json({'ok': True, 'removed': val})

    def _check(self, data):
        return self._compare(data, 'check')

    def _mismatch(self, data):
        return self._compare(data, 'mismatch')

    def _compare(self, data, op):
        s = self._session()
        if not s:
            return self._json({'ok': False, 'error': 'unauthorized'}, 401)
        code, name = s
        did, info = self._did(), self._ua()
        value = str(data.get('value', '')).strip()
        disp = str(data.get('display', ''))[:200]
        ct = str(data.get('time', '')).strip()
        if not value:
            return self._json({'ok': False, 'error': 'empty'}, 400)
        if not PICK_RE.match(ct):
            log('bad_time', code, name, did, info, detail=ct[:64])
            ct = now_ms()
        db('INSERT INTO compare_logs(check_time,employee_code,employee_name,'
           'device_id,value,display_name,op) VALUES(?,?,?,?,?,?,?)',
           (ct, code, name, did, value, disp, op))
        log(op, code, name, did, info, pick_time=ct, detail=value)
        self._json({'ok': True})

    def _diag(self, data):
        s = self._session()
        if not s:
            return self._json({'ok': False, 'error': 'unauthorized'}, 401)
        code, name = s
        did, info = self._did(), self._ua()
        for ev in (data.get('events') or [])[:200]:
            op = str(ev.get('op', ''))[:32]
            if op not in ('offline_start', 'offline_end'):
                continue
            log(op, code, name, did, info, detail=str(ev.get('detail', ''))[:200])
        self._json({'ok': True})


class Srv(ThreadingHTTPServer):
    allow_reuse_address = True


def main():
    db_init()
    load_catalog()
    emps = load_employees()
    print('Сотрудников в базе : %d' % len(emps))
    print('Строк в каталоге   : %d' % len(catalog['rows']))
    print('Имя (2-й столбец)  : %s' % (catalog['name_col'] or '(нет)'))
    print('Для compare        : %s' % (', '.join(catalog['columns']) or '(нет)'))
    print('База данных        : %s' % DB)
    print()
    print('=' * 62)
    print(' ССЫЛКА ДЛЯ СОТРУДНИКА:')
    for ip in (local_ips() or ['?']):
        print('     http://%s:%d/' % (ip, PORT))
    print('=' * 62)
    if emps:
        print()
        print(' СПИСОК СОТРУДНИКОВ (код ; имя):')
        w = max(len(c) for c in emps)
        for c, n in emps.items():
            print('   %-*s ; %s' % (w, c, n or '(без имени)'))
    print()
    print(' На этом компьютере: http://localhost:%d/' % PORT)
    print(' Ctrl+C — остановка.')
    try:
        Srv((HOST, PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        print('\nОстановлено.')


if __name__ == '__main__':
    main()