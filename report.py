#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Piiicker Report — генератор Excel-отчёта из базы piiicker.db.

Превращает плоский лог пиков в таблицу «первичный + вторичные»:
  - Первичный ключ открывает новую группу.
  - Все последующие вторичные принадлежат этой группе.
  - Каждый вторичный разворачивается в отдельную строку
    с дублированием первичного.
  - Вторичный без первичного отбрасывается.
  - Первичный без вторичных даёт одну строку с пустым вторичным.

Правила классификации задаются в GUI и сохраняются
в report_config.json рядом со скриптом.
"""

import json
import os
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False


# ---------- Пути ----------
BASE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE, 'piiicker.db')
CFG_PATH = os.path.join(BASE, 'report_config.json')


# ---------- Настройки по умолчанию ----------
DEFAULT_CFG = {
    'primary': {'min': 6, 'max': 999, 'kind': 'numbers'},
    'secondaries': [
        {'min': 1, 'max': 5, 'kind': 'numbers'},
    ],
}

KINDS = [
    ('numbers', 'только цифры'),
    ('text',    'только текст'),
    ('mixed',   'текст + цифры'),
]

HEADERS = ['Время', 'Сотрудник', 'Код', 'Наименование', 'Значение', 'Описание']
COL_WIDTHS = [22, 22, 20, 34, 15, 24]


# ---------- Конфиг ----------
def load_cfg():
    if not os.path.exists(CFG_PATH):
        return json.loads(json.dumps(DEFAULT_CFG))
    try:
        with open(CFG_PATH, encoding='utf-8') as f:
            cfg = json.load(f)
        if 'primary' not in cfg or not isinstance(cfg['primary'], dict):
            cfg['primary'] = dict(DEFAULT_CFG['primary'])
        if 'secondaries' not in cfg or not isinstance(cfg['secondaries'], list):
            cfg['secondaries'] = []
        return cfg
    except Exception:
        return json.loads(json.dumps(DEFAULT_CFG))


def save_cfg(cfg):
    with open(CFG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ---------- Классификация ----------
def kind_match(value, kind):
    """Проверяет, соответствует ли значение типу символов."""
    if not value:
        return False
    digits = sum(1 for c in value if c.isdigit())
    if kind == 'numbers':
        return digits == len(value)
    if kind == 'text':
        return digits == 0
    if kind == 'mixed':
        return 0 < digits < len(value)
    return False


def rule_match(value, rule):
    """Значение попадает в правило: длина в диапазоне и тип совпал."""
    try:
        lo = int(rule.get('min', 0))
        hi = int(rule.get('max', 99999))
    except (TypeError, ValueError):
        return False
    n = len(value)
    if n < lo or n > hi:
        return False
    return kind_match(value, rule.get('kind', 'numbers'))


def classify(value, cfg):
    """'P' — первичный, 'S' — вторичный, None — аномалия."""
    if rule_match(value, cfg['primary']):
        return 'P'
    for rule in cfg.get('secondaries', []):
        if rule_match(value, rule):
            return 'S'
    return None


# ---------- Чтение и разбор ----------
def read_logs():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError('База данных не найдена:\n' + DB_PATH)
    conn = sqlite3.connect(DB_PATH)
    try:
        rows = conn.execute(
            'SELECT pick_time, employee_name, value, display_name '
            'FROM logs ORDER BY pick_time ASC, id ASC'
        ).fetchall()
    finally:
        conn.close()
    return rows


def build_table(rows, cfg):
    """Разворачивает плоский лог в таблицу. Возвращает (строки, отброшено)."""
    groups = []
    current = None
    skipped = 0

    for pick_time, employee, value, display in rows:
        kind = classify(value or '', cfg)
        if kind == 'P':
            current = {
                'time': pick_time or '',
                'employee': employee or '',
                'value': value or '',
                'display': display or '',
                'children': [],
            }
            groups.append(current)
        elif kind == 'S':
            if current is None:
                skipped += 1
                continue
            current['children'].append((value or '', display or ''))
        else:
            skipped += 1

    table = []
    for g in groups:
        if g['children']:
            for v, d in g['children']:
                table.append([g['time'], g['employee'],
                              g['value'], g['display'], v, d])
        else:
            table.append([g['time'], g['employee'],
                          g['value'], g['display'], '', ''])

    return table, skipped


# ---------- Экспорт ----------
def export_xlsx(table, path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Отчёт'
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in table:
        ws.append(row)
    for i, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = 'A2'
    wb.save(path)


# ---------- Виджет правила ----------
def make_rule_widget(parent, rule, on_remove):
    f = ttk.Frame(parent)
    vars_ = {
        'min':  tk.IntVar(value=int(rule.get('min', 1))),
        'max':  tk.IntVar(value=int(rule.get('max', 999))),
        'kind': tk.StringVar(value=rule.get('kind', 'numbers')),
    }
    ttk.Label(f, text='Длина: от').pack(side='left')
    ttk.Entry(f, textvariable=vars_['min'], width=5).pack(side='left', padx=2)
    ttk.Label(f, text='до').pack(side='left')
    ttk.Entry(f, textvariable=vars_['max'], width=5).pack(side='left', padx=2)
    ttk.Label(f, text='    Тип:').pack(side='left')
    for val, txt in KINDS:
        ttk.Radiobutton(f, text=txt, variable=vars_['kind'],
                        value=val).pack(side='left', padx=4)
    if on_remove:
        ttk.Button(f, text='×', width=3, command=on_remove).pack(side='right')
    return f, vars_


# ---------- GUI ----------
class App:
    def __init__(self, root):
        self.root = root
        root.title('Piiicker Report')
        root.geometry('820x600')
        root.minsize(720, 500)

        self.sec_rows = []   # [(frame, vars), ...]
        self.status_var = tk.StringVar(value='Готово')

        self._build_ui()
        self._load_into_ui(load_cfg())

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill='both', expand=True)

        ttk.Label(main, text='Piiicker Report',
                  font=('', 14, 'bold')).pack(anchor='w')
        ttk.Label(main, text='Преобразование лога пиков в таблицу',
                  foreground='#666').pack(anchor='w', pady=(0, 12))

        ttk.Label(main, text='База данных: ' + DB_PATH,
                  foreground='#888', wraplength=780).pack(anchor='w', pady=(0, 12))

        # Первичный ключ
        box1 = ttk.LabelFrame(main, text='Первичный ключ', padding=10)
        box1.pack(fill='x', pady=(0, 10))
        self.primary_frame, self.primary_vars = make_rule_widget(box1, {}, None)
        self.primary_frame.pack(fill='x')

        # Вторичные ключи
        box2 = ttk.LabelFrame(main, text='Вторичные ключи', padding=10)
        box2.pack(fill='both', expand=True, pady=(0, 10))
        self.sec_container = ttk.Frame(box2)
        self.sec_container.pack(fill='both', expand=True)
        ttk.Button(box2, text='+ Добавить вторичный ключ',
                   command=lambda: self.add_sec_row({})
                   ).pack(anchor='w', pady=(8, 0))

        # Кнопки действий
        btns = ttk.Frame(main)
        btns.pack(fill='x', pady=(4, 0))
        ttk.Button(btns, text='📊 Сформировать отчёт',
                   command=self.generate).pack(side='left')
        ttk.Button(btns, text='💾 Сохранить настройки',
                   command=self.save_only).pack(side='left', padx=8)
        ttk.Button(btns, text='↺ Сбросить',
                   command=self.reset_defaults).pack(side='left')

        # Статус
        status = ttk.Frame(self.root, relief='sunken', padding=(10, 4))
        status.pack(fill='x', side='bottom')
        ttk.Label(status, textvariable=self.status_var,
                  foreground='#555').pack(anchor='w')

    # --- работа с правилами ---
    def _set_rule_vars(self, vars_, rule):
        try:
            vars_['min'].set(int(rule.get('min', 1)))
            vars_['max'].set(int(rule.get('max', 999)))
        except Exception:
            vars_['min'].set(1)
            vars_['max'].set(999)
        vars_['kind'].set(rule.get('kind', 'numbers'))

    def _load_into_ui(self, cfg):
        self._set_rule_vars(self.primary_vars, cfg.get('primary', {}))
        for frame, _ in self.sec_rows:
            frame.destroy()
        self.sec_rows = []
        for rule in cfg.get('secondaries', []):
            self.add_sec_row(rule)

    def add_sec_row(self, rule):
        holder = {}
        frame, vars_ = make_rule_widget(
            self.sec_container, rule,
            on_remove=lambda: self._remove_sec_frame(holder.get('f'))
        )
        holder['f'] = frame
        frame.pack(fill='x', pady=2)
        self._set_rule_vars(vars_, rule)
        self.sec_rows.append((frame, vars_))

    def _remove_sec_frame(self, frame):
        if frame is None:
            return
        for i, (f, _) in enumerate(self.sec_rows):
            if f is frame:
                self.sec_rows.pop(i)
                f.destroy()
                return

    def collect_cfg(self):
        def safe_int(var, default):
            try:
                return int(var.get())
            except (tk.TclError, ValueError):
                return default
        primary = {
            'min':  safe_int(self.primary_vars['min'], 1),
            'max':  safe_int(self.primary_vars['max'], 999),
            'kind': self.primary_vars['kind'].get() or 'numbers',
        }
        secondaries = []
        for _, v in self.sec_rows:
            secondaries.append({
                'min':  safe_int(v['min'], 1),
                'max':  safe_int(v['max'], 999),
                'kind': v['kind'].get() or 'numbers',
            })
        return {'primary': primary, 'secondaries': secondaries}

    # --- действия ---
    def save_only(self):
        try:
            save_cfg(self.collect_cfg())
            self.status_var.set('Настройки сохранены в report_config.json')
        except Exception as e:
            messagebox.showerror('Ошибка', 'Не удалось сохранить настройки:\n' + str(e))

    def reset_defaults(self):
        if not messagebox.askyesno('Сброс',
                'Сбросить все настройки к значениям по умолчанию?'):
            return
        self._load_into_ui(DEFAULT_CFG)
        self.status_var.set('Настройки сброшены. Не забудьте сохранить.')

    def generate(self):
        if not HAVE_OPENPYXL:
            messagebox.showerror(
                'Нет библиотеки openpyxl',
                'Для формирования xlsx установите библиотеку:\n\n'
                '    pip install openpyxl'
            )
            return

        cfg = self.collect_cfg()
        try:
            save_cfg(cfg)
        except Exception:
            pass

        self.status_var.set('Чтение базы…')
        self.root.update_idletasks()
        try:
            rows = read_logs()
        except Exception as e:
            messagebox.showerror('Ошибка', str(e))
            self.status_var.set('Ошибка чтения базы')
            return

        if not rows:
            messagebox.showinfo('Пусто', 'В базе нет ни одной записи в logs.')
            self.status_var.set('Нет данных')
            return

        table, skipped = build_table(rows, cfg)
        if not table:
            messagebox.showinfo('Пусто',
                'Ни одна запись не подошла под правила.\n'
                'Проверьте настройки первичного и вторичных ключей.')
            self.status_var.set('Нет данных после фильтрации')
            return

        path = filedialog.asksaveasfilename(
            title='Куда сохранить отчёт',
            defaultextension='.xlsx',
            initialfile='report.xlsx',
            initialdir=BASE,
            filetypes=[('Excel', '*.xlsx')],
        )
        if not path:
            self.status_var.set('Отменено')
            return

        try:
            export_xlsx(table, path)
        except Exception as e:
            messagebox.showerror('Ошибка',
                'Не удалось сохранить отчёт:\n' + str(e))
            self.status_var.set('Ошибка записи')
            return

        self.status_var.set(
            'Готово: %d строк, отброшено пиков: %d. Файл: %s'
            % (len(table), skipped, os.path.basename(path))
        )
        messagebox.showinfo(
            'Готово',
            'Отчёт сохранён:\n%s\n\nСтрок: %d\nОтброшено пиков: %d'
            % (path, len(table), skipped)
        )


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()