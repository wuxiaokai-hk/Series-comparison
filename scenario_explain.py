#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scenario_explain.py - explain why two runs of an equation-based model (scenario A and B) differ.

Works with any model written as equations (EViews syntax: Y = formula), with or without
geographies, and with the same or different equations in the two scenarios (--eq-a / --eq-b).

Standard library only (Python 3.8+). No pip installs, no openpyxl/pandas.

Commands
  report    interactive HTML report: key changes, where the scenario starts,
            driver decomposition for each variable, path tracing X -> Y
  top       print the biggest changes (optionally save all stats to CSV)
  explain   contribution of each input (and the add-factor) to the gap of VAR
  addfactors  formula vs add-factor share of the gap, for every equation
  impact    can X affect Y, and by how much? (re-solves the equations with only X changed)
  attribute split the gap of Y across every scenario shock upstream of it
  influence if X changes, how many variables move and by how much
  path      the channels through which X reaches Y
  serve     the report in your browser, plus exact impact / attribution on demand
  info      what was loaded (layout detected, periods, equations parsed)

Examples
  python scenario_explain.py report  --a base.csv --b stress.csv --eq model.txt -o explain.html
  python scenario_explain.py report  --a base.csv --b stress.csv --eq-a eq_base.txt --eq-b eq_stress.txt
  python scenario_explain.py top     --a base.csv --b stress.csv -n 40
  python scenario_explain.py explain CET1 --a base.csv --b stress.csv --eq model.txt
  python scenario_explain.py path UNEMP CET1 --a base.csv --b stress.csv --eq model.txt

Run any command with -h for all options.
"""
import argparse
import csv
import datetime
import json
import math
import os
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from array import array
from collections import Counter, defaultdict, deque

NAN = float('nan')
NA_STRINGS = {'NA', 'N/A', '#N/A', 'NAN', 'ND', 'NULL', 'NONE', '-', '.', '..',
              '#VALUE!', '#DIV/0!', '#REF!', '#NUM!'}

try:
    csv.field_size_limit(min(sys.maxsize, 2 ** 31 - 1))
except Exception:  # pragma: no cover
    pass


def _num(cell):
    """Parse a numeric cell; missing / text -> NaN."""
    if cell is None:
        return NAN
    if isinstance(cell, (int, float)):
        v = float(cell)
        return v if math.isfinite(v) else NAN
    t = str(cell).strip()
    if not t or t.upper() in NA_STRINGS:
        return NAN
    t = t.replace(',', '')
    try:
        v = float(t)
    except ValueError:
        return NAN
    return v if math.isfinite(v) else NAN


def _clean(c):
    return '' if c is None else str(c).strip()


def _isnan(v):
    return v != v


# --------------------------------------------------------------------------
# Periods
# --------------------------------------------------------------------------
MONTHS = {m: i + 1 for i, m in enumerate(
    'JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC'.split())}
_RE_Q1 = re.compile(r'^(\d{4})\s*[-\s]?\s*[Qq]\s*([1-4])$')
_RE_Q2 = re.compile(r'^[Qq]([1-4])[\s\-/:]*(\d{4})$')
_RE_M1 = re.compile(r'^(\d{4})\s*[Mm]\s*(\d{1,2})$')
_RE_YM = re.compile(r'^(\d{4})-(\d{1,2})$')
_RE_COLON = re.compile(r'^(\d{4}):(\d{1,2})$')
_RE_ISO = re.compile(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T].*)?$')
_RE_US = re.compile(r'^(\d{1,2})/(\d{1,2})/(\d{2,4})(?:\s.*)?$')
_RE_Y = re.compile(r'^(\d{4})\s*[Aa]?$')
_RE_MON = re.compile(r'^([A-Za-z]{3})[a-z]*[\s\-]+(\d{2}|\d{4})$')


def _date_key(y, mo, d):
    try:
        datetime.date(y, mo, d)
    except ValueError:
        return None
    return ((y, 5, mo * 100 + d), '%04d-%02d-%02d' % (y, mo, d))


def parse_period(cell, allow_serial=False):
    """Return (sortkey, label) if the cell looks like a date/period, else None."""
    if cell is None:
        return None
    t = str(cell).strip()
    if not t or len(t) > 30:
        return None
    m = _RE_Q1.match(t)
    if m:
        y, q = int(m.group(1)), int(m.group(2))
        return ((y, 2, q), '%dQ%d' % (y, q))
    m = _RE_Q2.match(t)
    if m:
        q, y = int(m.group(1)), int(m.group(2))
        return ((y, 2, q), '%dQ%d' % (y, q))
    m = _RE_M1.match(t) or _RE_YM.match(t)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return ((y, 3, mo), '%dM%02d' % (y, mo))
        return None
    m = _RE_COLON.match(t)
    if m:
        y, s = int(m.group(1)), int(m.group(2))
        if 1 <= s <= 12:
            return ((y, 4, s), '%d:%d' % (y, s))
        return None
    m = _RE_ISO.match(t)
    if m:
        return _date_key(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = _RE_US.match(t)
    if m:
        mo, d, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000 if y < 50 else 1900
        return _date_key(y, mo, d)
    m = _RE_MON.match(t)
    if m and m.group(1).upper() in MONTHS:
        y = int(m.group(2))
        if y < 100:
            y += 2000 if y < 50 else 1900
        mo = MONTHS[m.group(1).upper()]
        return ((y, 3, mo), '%dM%02d' % (y, mo))
    m = _RE_Y.match(t)
    if m:
        y = int(m.group(1))
        if 1800 <= y <= 2200:
            return ((y, 1, 0), str(y))
        return None
    if allow_serial:  # Excel date serial number (xlsx headers)
        try:
            v = float(t)
        except ValueError:
            return None
        if 15000 < v < 80000 and abs(v - round(v)) < 1e-9:
            d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(round(v)))
            return _date_key(d.year, d.month, d.day)
    return None


def infer_ppy(keys):
    """Periods per year, used by @PCHY/@PCY."""
    if not keys:
        return 4
    kind = Counter(k[1] for k in keys).most_common(1)[0][0]
    if kind == 2:
        return 4
    if kind == 3:
        return 12
    if kind == 1:
        return 1
    if kind == 4:
        return 12 if max(k[2] for k in keys if k[1] == 4) > 4 else 4
    ds = sorted(datetime.date(k[0], k[2] // 100, k[2] % 100) for k in keys if k[1] == 5)
    if len(ds) < 2:
        return 4
    gaps = sorted((b - a).days for a, b in zip(ds, ds[1:]))
    g = gaps[len(gaps) // 2]
    if g < 45:
        return 12
    if g < 200:
        return 4
    return 1


# --------------------------------------------------------------------------
# Reading files (CSV / TSV / XLSX) with the standard library
# --------------------------------------------------------------------------
_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def _col_index(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def read_xlsx(path, sheet=None):
    """Minimal .xlsx reader (zipfile + xml). Returns list of rows of strings."""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if 'xl/sharedStrings.xml' in names:
            for _, el in ET.iterparse(z.open('xl/sharedStrings.xml')):
                if el.tag == _NS + 'si':
                    shared.append(''.join(t.text or '' for t in el.iter(_NS + 't')))
                    el.clear()
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap = {r.get('Id'): r.get('Target') for r in rels}
        sheets = []
        for s in wb.iter(_NS + 'sheet'):
            tgt = (relmap.get(s.get(_RNS + 'id')) or '').lstrip('/')
            if not tgt.startswith('xl/'):
                tgt = 'xl/' + tgt
            sheets.append((s.get('name'), tgt))
        if not sheets:
            raise ValueError('%s: no worksheets found' % path)
        target = sheets[0][1]
        if sheet is not None:
            found = None
            for i, (nm, tg) in enumerate(sheets):
                if nm.lower() == str(sheet).lower() or str(i + 1) == str(sheet):
                    found = tg
            if not found:
                raise ValueError('%s: sheet %r not found; sheets are %s'
                                 % (path, sheet, [s[0] for s in sheets]))
            target = found
        grid = []
        for _, el in ET.iterparse(z.open(target)):
            if el.tag != _NS + 'row':
                continue
            row = []
            for c in el.findall(_NS + 'c'):
                ref = c.get('r')
                j = _col_index(ref) if ref else len(row)
                t = c.get('t')
                val = ''
                if t == 's':
                    v = c.find(_NS + 'v')
                    if v is not None and v.text is not None:
                        val = shared[int(v.text)]
                elif t == 'inlineStr':
                    is_ = c.find(_NS + 'is')
                    if is_ is not None:
                        val = ''.join(x.text or '' for x in is_.iter(_NS + 't'))
                else:
                    v = c.find(_NS + 'v')
                    if v is not None and v.text is not None:
                        val = v.text
                        if t not in ('str', 'e', 'b'):
                            try:
                                f = float(val)
                                val = str(int(f)) if f.is_integer() and abs(f) < 1e15 else repr(f)
                            except ValueError:
                                pass
                if j >= len(row):
                    row.extend([''] * (j - len(row) + 1))
                row[j] = val
            r = el.get('r')
            ri = int(r) - 1 if r else len(grid)
            while len(grid) < ri:
                grid.append([])
            grid.append(row)
            el.clear()
    return grid


def read_grid(path, sheet=None):
    """Return (rows, is_xlsx)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xlsm'):
        return read_xlsx(path, sheet), True
    if ext == '.xls':
        raise SystemExit('%s: old .xls format is not supported - save it as .xlsx or .csv' % path)
    with open(path, 'r', encoding='utf-8-sig', errors='replace', newline='') as f:
        sample = f.read(65536)
        f.seek(0)
        delim = ','
        if ext in ('.tsv', '.tab'):
            delim = '\t'
        else:
            try:
                delim = csv.Sniffer().sniff(sample, delimiters=',\t;|').delimiter
            except csv.Error:
                pass
        return [row for row in csv.reader(f, delimiter=delim)], False


# --------------------------------------------------------------------------
# Scenario data
# --------------------------------------------------------------------------
NAME_HDRS = {'mnemonic', 'mnemonics', 'series', 'variable', 'name', 'code', 'ticker',
             'id', 'series id', 'series_id', 'mnem', 'series name', 'var'}


def _is_desc_hdr(h):
    h = h.lower().rstrip(':').strip()
    return ('desc' in h or 'title' in h or h in ('concept', 'label', 'long name', 'longname'))


class Scenario:
    def __init__(self, label):
        self.label = label
        self.blocks = []        # list of lists of period sortkeys
        self.series = {}        # NAME -> (block id, array of floats)
        self.desc = {}
        self.keys = {}          # sortkey -> label
        self.notes = []
        self.dups = 0

    def put(self, name, bid, arr, desc=''):
        name = name.strip().upper()
        if not name:
            return
        if name in self.series:
            self.dups += 1
        self.series[name] = (bid, arr)
        if desc and name not in self.desc:
            self.desc[name] = desc


def _pick_name_col(grid, start, candidates, hdr):
    for j in candidates:
        if j < len(hdr) and hdr[j] in NAME_HDRS:
            return j
    best, bestscore = None, -1
    sample = grid[start:start + 60]
    for j in candidates:
        score = 0
        for r in sample:
            c = _clean(r[j]) if j < len(r) else ''
            if c and _isnan(_num(c)):
                score += 1
        if score > bestscore:
            best, bestscore = j, score
    return best


def _increasing(cands):
    """Keep period cells that form an increasing run (rejects numbers that merely look like dates)."""
    out, last = [], None
    for j, p in cands:
        if not p:
            continue
        if last is None or p[0] > last:
            out.append((j, p))
            last = p[0]
    # a real date axis has (almost) no out-of-order cells
    n = sum(1 for _, p in cands if p)
    return out if n and len(out) >= 0.9 * n else []


def _ingest(sc, grid, serial, path):
    nonempty = [i for i, r in enumerate(grid) if any(_clean(c) for c in r)]
    if not nonempty:
        raise SystemExit('%s: file is empty' % path)
    h = nonempty[0]
    hdr = [_clean(c).lower() for c in grid[h]]

    def find(cands):
        for j, c in enumerate(hdr):
            if c in cands:
                return j
        return -1
    vcol = find({'value', 'val', 'data', 'obs_value', 'amount'})
    dcol = find({'date', 'period', 'time', 'obs', 'datetime', 'quarter', 'month', 'year'})
    scol = find(NAME_HDRS)
    if vcol >= 0 and dcol >= 0 and scol >= 0 and len([c for c in hdr if c]) <= 10:
        return _ingest_long(sc, grid, h, scol, dcol, vcol, serial, path)

    rows_hdr = None
    for i in nonempty[:25]:
        cols = _increasing([(j, parse_period(c, serial)) for j, c in enumerate(grid[i])])
        if len(cols) >= 3:
            rows_hdr = (i, cols)
            break
    cols_best = None
    for c in range(4):
        idx = _increasing([(i, parse_period(r[c], serial) if c < len(r) else None)
                           for i, r in enumerate(grid)])
        if len(idx) >= 3 and (cols_best is None or len(idx) > len(cols_best[1])):
            cols_best = (c, idx)
    nr = len(rows_hdr[1]) if rows_hdr else 0
    nc = len(cols_best[1]) if cols_best else 0
    if nr == 0 and nc == 0:
        raise SystemExit('%s: could not find a row or column of dates/periods '
                         '(e.g. 2026Q1, 2026-03-31, 2026:1).' % path)
    if nr >= nc:
        _ingest_rows(sc, grid, rows_hdr[0], rows_hdr[1], path)
    else:
        _ingest_cols(sc, grid, cols_best[0], cols_best[1], path)


def _ingest_rows(sc, grid, i, cols, path):
    for _, p in cols:
        sc.keys[p[0]] = p[1]
    datecols = set(j for j, _ in cols)
    hdr = [_clean(c).lower() for c in grid[i]]
    width = max(len(r) for r in grid[i:i + 50])
    nondate = [j for j in range(width) if j not in datecols]
    if not nondate:
        raise SystemExit('%s: no column with series names found' % path)
    namecol = _pick_name_col(grid, i + 1, nondate, hdr)
    desccol = next((j for j in nondate if j < len(hdr) and _is_desc_hdr(hdr[j])), None)
    bid = len(sc.blocks)
    sc.blocks.append([p[0] for _, p in cols])
    n = 0
    jlist = [j for j, _ in cols]
    for r in grid[i + 1:]:
        name = _clean(r[namecol]) if namecol < len(r) else ''
        if not name:
            continue
        L = len(r)
        arr = array('d', (_num(r[j]) if j < L else NAN for j in jlist))
        if all(_isnan(v) for v in arr):
            continue
        desc = _clean(r[desccol]) if desccol is not None and desccol < L else ''
        sc.put(name, bid, arr, desc)
        n += 1
    sc.notes.append('%s: series in rows, %d series x %d periods (%s .. %s), names from column %d'
                    % (os.path.basename(path), n, len(cols), cols[0][1][1], cols[-1][1][1], namecol + 1))


def _ingest_cols(sc, grid, c, idx, path):
    for _, p in idx:
        sc.keys[p[0]] = p[1]
    first = idx[0][0]
    pre = [i for i in range(first) if any(_clean(x) for x in grid[i])]
    if not pre:
        raise SystemExit('%s: dates are in column %d but there is no header row with series names'
                         % (path, c + 1))
    namerow, descrow = None, None
    for i in pre:
        lab = _clean(grid[i][c] if c < len(grid[i]) else '').lower().rstrip(':').strip()
        if not lab and c > 0:
            lab = _clean(grid[i][0]).lower().rstrip(':').strip()
        if lab in NAME_HDRS and namerow is None:
            namerow = i
        elif _is_desc_hdr(lab) and descrow is None:
            descrow = i
    if namerow is None:
        best = -1
        for i in pre:
            sc_ = sum(1 for j, x in enumerate(grid[i]) if j > c and _clean(x) and _isnan(_num(x)))
            if sc_ > best:
                best, namerow = sc_, i
    names = grid[namerow]
    descs = grid[descrow] if descrow is not None else []
    bid = len(sc.blocks)
    sc.blocks.append([p[0] for _, p in idx])
    rows = [grid[i] for i, _ in idx]
    n = 0
    for j in range(c + 1, len(names)):
        name = _clean(names[j])
        if not name:
            continue
        arr = array('d', (_num(r[j]) if j < len(r) else NAN for r in rows))
        if all(_isnan(v) for v in arr):
            continue
        sc.put(name, bid, arr, _clean(descs[j]) if j < len(descs) else '')
        n += 1
    sc.notes.append('%s: series in columns, %d series x %d periods (%s .. %s), names from row %d'
                    % (os.path.basename(path), n, len(idx), idx[0][1][1], idx[-1][1][1], namerow + 1))


def _ingest_long(sc, grid, h, scol, dcol, vcol, serial, path):
    data = defaultdict(dict)
    for r in grid[h + 1:]:
        if max(scol, dcol, vcol) >= len(r):
            continue
        p = parse_period(r[dcol], serial)
        name = _clean(r[scol])
        if not p or not name:
            continue
        sc.keys[p[0]] = p[1]
        data[name][p[0]] = _num(r[vcol])
    keys = sorted(set(k for d in data.values() for k in d))
    bid = len(sc.blocks)
    sc.blocks.append(keys)
    for name, d in data.items():
        sc.put(name, bid, array('d', (d.get(k, NAN) for k in keys)))
    sc.notes.append('%s: long format (series, date, value), %d series' % (os.path.basename(path), len(data)))


def load_scenario(paths, label, sheet=None):
    sc = Scenario(label)
    for p in paths:
        grid, serial = read_grid(p, sheet)
        _ingest(sc, grid, serial, p)
    if sc.dups:
        sc.notes.append('%d duplicate series names (last one kept)' % sc.dups)
    return sc


def load_meta(path):
    grid, _ = read_grid(path)
    hdr = [_clean(c).lower() for c in grid[0]]
    nc = next((j for j, h in enumerate(hdr) if h in NAME_HDRS), 0)
    dc = next((j for j, h in enumerate(hdr) if _is_desc_hdr(h)), 1)
    out = {}
    for r in grid[1:]:
        if nc < len(r) and dc < len(r) and _clean(r[nc]):
            out[_clean(r[nc]).upper()] = _clean(r[dc])
    return out


# --------------------------------------------------------------------------
# EViews-style equations: tokenizer, parser, evaluator
# --------------------------------------------------------------------------
class ParseError(Exception):
    pass


class EvalError(Exception):
    pass


_TOKRE = re.compile(r'''\s*(?:
    (?P<str>"[^"]*")
  | (?P<num>(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?)
  | (?P<name>@?[A-Za-z_][A-Za-z0-9_$.%#]*)
  | (?P<op>\*\*|<=|>=|<>|[-+*/^(),=<>])
)''', re.X)

FUNCS = {'LOG', 'LN', 'EXP', 'ABS', 'SQR', 'SQRT', 'D', 'DLOG', 'LOG10', 'INV', 'LAG'}
EXCLUDE = {'C', 'AR', 'MA', 'SAR', 'SMA', 'PDL', 'NA'}


def tokenize(s):
    toks, pos, n = [], 0, len(s)
    while pos < n:
        if s[pos].isspace():
            pos += 1
            continue
        m = _TOKRE.match(s, pos)
        if not m or m.end() == pos:
            raise ParseError('unexpected character %r at "%s"' % (s[pos], s[max(0, pos - 15):pos + 10]))
        if m.group('str') is not None:
            toks.append(('str', m.group('str')[1:-1]))
        elif m.group('num') is not None:
            toks.append(('num', m.group('num')))
        elif m.group('name') is not None:
            toks.append(('name', m.group('name')))
        else:
            op = m.group('op')
            toks.append(('op', '^' if op == '**' else op))
        pos = m.end()
    return toks


class _Parser:
    def __init__(self, toks):
        self.t = toks
        self.i = 0

    def peek(self, k=0):
        j = self.i + k
        return self.t[j] if j < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expect(self, v):
        if self.take()[1] != v:
            raise ParseError('expected %r' % v)

    def parse(self):
        n = self.full()
        if self.i != len(self.t):
            raise ParseError('unexpected %r' % (self.peek()[1],))
        return n

    def full(self):   # a or b
        n = self.andx()
        while self.peek()[0] == 'name' and self.peek()[1].upper() == 'OR':
            self.take()
            n = ('bin', 'OR', n, self.andx())
        return n

    def andx(self):   # a and b
        n = self.cmp()
        while self.peek()[0] == 'name' and self.peek()[1].upper() == 'AND':
            self.take()
            n = ('bin', 'AND', n, self.cmp())
        return n

    def cmp(self):    # a < b, a = b, ...
        n = self.expr()
        while self.peek()[1] in ('<', '>', '<=', '>=', '=', '<>') and self.peek()[0] == 'op':
            op = self.take()[1]
            n = ('bin', op, n, self.expr())
        return n

    def expr(self):
        n = self.term()
        while self.peek()[1] in ('+', '-'):
            op = self.take()[1]
            n = ('bin', op, n, self.term())
        return n

    def term(self):
        n = self.unary()
        while self.peek()[1] in ('*', '/'):
            op = self.take()[1]
            n = ('bin', op, n, self.unary())
        return n

    def unary(self):
        if self.peek()[1] in ('-', '+'):
            op = self.take()[1]
            u = self.unary()
            return ('neg', u) if op == '-' else u
        return self.power()

    def power(self):
        b = self.atom()
        if self.peek()[1] == '^':
            self.take()
            return ('bin', '^', b, self.unary())
        return b

    def _try_lag(self):
        # NAME ( [+-] int )
        t0, t1, t2, t3 = self.peek(0), self.peek(1), self.peek(2), self.peek(3)
        if t0[1] != '(':
            return None
        if t1[0] == 'num' and t2[1] == ')' and re.fullmatch(r'\d+', t1[1]):
            self.i += 3
            return int(t1[1])
        if t1[1] in ('-', '+') and t2[0] == 'num' and t3[1] == ')' and re.fullmatch(r'\d+', t2[1]):
            self.i += 4
            return -int(t2[1]) if t1[1] == '-' else int(t2[1])
        return None

    def atom(self):
        kind, val = self.take()
        if kind == 'num':
            return ('num', float(val))
        if kind == 'str':
            return ('str', val)
        if kind == 'op' and val == '(':
            n = self.full()
            self.expect(')')
            return n
        if kind == 'name':
            up = val.upper()
            is_func = up in FUNCS or up.startswith('@')
            if self.peek()[1] == '(':
                if not is_func:
                    lag = self._try_lag()
                    if lag is not None:
                        return ('coef', lag) if up == 'C' else ('var', up, lag)
                self.take()
                args = []
                if self.peek()[1] != ')':
                    args.append(self.full())
                    while self.peek()[1] == ',':
                        self.take()
                        args.append(self.full())
                self.expect(')')
                return ('call', up, args)
            if up.startswith('@'):
                return ('call', up, [])
            return ('var', up, 0)
        raise ParseError('unexpected %r' % (val,))


def parse_expr(toks):
    if not toks:
        raise ParseError('empty expression')
    p = _Parser(toks)
    try:
        return p.parse()
    except ParseError as e:
        near = ' '.join(v for _, v in toks[max(0, p.i - 4):p.i + 2])
        raise ParseError('%s, near "%s"' % (e, near))


def ast_vars(node, out=None):
    """Ordered unique variable names in an expression."""
    if out is None:
        out = []
    k = node[0]
    if k == 'var':
        if node[1] not in out and node[1] not in EXCLUDE:
            out.append(node[1])
    elif k == 'neg':
        ast_vars(node[1], out)
    elif k == 'bin':
        ast_vars(node[2], out)
        ast_vars(node[3], out)
    elif k == 'call':
        for a in node[2]:
            ast_vars(a, out)
    return out


def ast_has(node, kinds):
    k = node[0]
    if k in kinds:
        return True
    if k == 'neg':
        return ast_has(node[1], kinds)
    if k == 'bin':
        return ast_has(node[2], kinds) or ast_has(node[3], kinds)
    if k == 'call':
        return any(ast_has(a, kinds) for a in node[2])
    return False


def _const(node):
    if node[0] == 'num':
        return node[1]
    if node[0] == 'neg' and node[1][0] == 'num':
        return -node[1][1]
    raise EvalError('expected a constant argument')


def _log(v):
    return math.log(v) if v > 0 else NAN


def _exp(v):
    try:
        return math.exp(v)
    except OverflowError:
        return NAN


def _pow(x, y):
    if x != x or y != y:
        return NAN
    if x < 0 and not float(y).is_integer():
        return NAN
    try:
        return float(x ** y)
    except (OverflowError, ZeroDivisionError, ValueError):
        return NAN


def _diff(xs, k):
    return [xs[i] - xs[i - k] for i in range(k, len(xs))]


# periods of the loaded data, set by Model, used by date functions (@DURING, @ELEM, @SEAS, ...)
PERIODS = {'keys': [], 'index': {}}


def set_periods(keys):
    PERIODS['keys'] = list(keys)
    PERIODS['index'] = {k: i for i, k in enumerate(keys)}


def _period_idx(text, end=False):
    """'2020q1' -> index of that period; a year like '2020' -> its first (or last) period."""
    t = text.strip()
    p = parse_period(t)
    if p is None:
        raise EvalError('cannot read date %r' % text)
    key, idx, keys = p[0], PERIODS['index'], PERIODS['keys']
    if key in idx:
        return idx[key]
    # coarser date (a year) or a date outside the data: nearest period
    cands = [i for i, k in enumerate(keys) if k[0] == key[0]] if key[1] == 1 else []
    if cands:
        return cands[-1] if end else cands[0]
    if keys and key < keys[0]:
        return 0 if not end else -1
    return len(keys) if not end else len(keys) - 1


def _sample_mask(text, lo, hi):
    """EViews sample string '2020q1 2020q4 2022q1 2022q2' (pairs; '@first'/'@last' allowed)."""
    toks = text.replace(',', ' ').split()
    n = len(PERIODS['keys'])
    if len(toks) == 1:
        toks = toks * 2
    if len(toks) % 2:
        raise EvalError('odd number of dates in sample %r' % text)
    out = [0.0] * (hi - lo + 1)
    for a, b in zip(toks[::2], toks[1::2]):
        ia = 0 if a.lower() == '@first' else _period_idx(a)
        ib = n - 1 if b.lower() == '@last' else _period_idx(b, end=True)
        for t in range(max(lo, ia), min(hi, ib) + 1):
            out[t - lo] = 1.0
    return out


def _sarg(node, fn):
    if node[0] != 'str':
        raise EvalError('%s needs a date string argument' % fn)
    return node[1]


def _period_field(t, field):
    keys = PERIODS['keys']
    if not 0 <= t < len(keys):
        return NAN
    y, kind, sub = keys[t]
    if field == 'YEAR':
        return float(y)
    if kind == 2:
        q, m = sub, sub * 3
    elif kind == 3 or kind == 4:
        m = sub
        q = (m - 1) // 3 + 1
    elif kind == 5:
        m = sub // 100
        q = (m - 1) // 3 + 1
    else:
        q = m = NAN
    return float(q if field == 'QUARTER' else m)


def evaluate(node, get, lo, hi, ppy):
    """Evaluate an expression for period indices lo..hi; returns list of floats (NaN = missing)."""
    k = node[0]
    if k == 'str':
        raise EvalError('text %r used as a number' % node[1])
    if k == 'num':
        return [node[1]] * (hi - lo + 1)
    if k == 'var':
        if node[1] == 'NA':
            return [NAN] * (hi - lo + 1)
        s = get(node[1])
        a, b = lo + node[2], hi + node[2]
        n = len(s)
        if a >= 0 and b < n:
            return list(s[a:b + 1])
        return [s[t] if 0 <= t < n else NAN for t in range(a, b + 1)]
    if k == 'neg':
        return [-v for v in evaluate(node[1], get, lo, hi, ppy)]
    if k == 'bin':
        op = node[1]
        x = evaluate(node[2], get, lo, hi, ppy)
        y = evaluate(node[3], get, lo, hi, ppy)
        if op == '+':
            return [p + q for p, q in zip(x, y)]
        if op == '-':
            return [p - q for p, q in zip(x, y)]
        if op == '*':
            return [p * q for p, q in zip(x, y)]
        if op == '/':
            return [p / q if q != 0 else NAN for p, q in zip(x, y)]
        if op == '^':
            return [_pow(p, q) for p, q in zip(x, y)]
        if op in ('<', '>', '<=', '>=', '=', '<>', 'AND', 'OR'):
            f = {'<': lambda a, b: a < b, '>': lambda a, b: a > b, '<=': lambda a, b: a <= b,
                 '>=': lambda a, b: a >= b, '=': lambda a, b: a == b, '<>': lambda a, b: a != b,
                 'AND': lambda a, b: bool(a) and bool(b), 'OR': lambda a, b: bool(a) or bool(b)}[op]
            return [NAN if (p != p or q != q) else (1.0 if f(p, q) else 0.0) for p, q in zip(x, y)]
        raise EvalError('operator %s' % op)
    if k == 'coef':
        raise EvalError('equation has unsubstituted coefficient C(%d)' % node[1])
    if k == 'call':
        f = node[1].lstrip('@')
        args = node[2]
        if f in ('LOG', 'LN'):
            return [_log(v) for v in evaluate(args[0], get, lo, hi, ppy)]
        if f == 'EXP':
            return [_exp(v) for v in evaluate(args[0], get, lo, hi, ppy)]
        if f == 'ABS':
            return [abs(v) for v in evaluate(args[0], get, lo, hi, ppy)]
        if f in ('SQR', 'SQRT'):
            return [math.sqrt(v) if v >= 0 else NAN for v in evaluate(args[0], get, lo, hi, ppy)]
        if f in ('D', 'DLOG'):
            order = int(_const(args[1])) if len(args) > 1 else 1
            seas = int(_const(args[2])) if len(args) > 2 else 0
            xs = evaluate(args[0], get, lo - order - seas, hi, ppy)
            if f == 'DLOG':
                xs = [_log(v) for v in xs]
            for _ in range(order):
                xs = _diff(xs, 1)
            if seas:
                xs = _diff(xs, seas)
            return xs
        if f in ('PCH', 'PC', 'PCHY', 'PCY', 'PCHA', 'PCA'):
            lag = ppy if f in ('PCHY', 'PCY') else 1
            xs = evaluate(args[0], get, lo - lag, hi, ppy)
            out = [xs[i] / xs[i - lag] - 1 if xs[i - lag] != 0 else NAN for i in range(lag, len(xs))]
            if f in ('PCHA', 'PCA'):
                out = [_pow(1 + v, ppy) - 1 for v in out]
            if f in ('PC', 'PCY', 'PCA'):
                out = [100 * v for v in out]
            return out
        if f in ('MOVAV', 'MOVSUM'):
            n = int(_const(args[1]))
            xs = evaluate(args[0], get, lo - n + 1, hi, ppy)
            out = [sum(xs[i - n + 1:i + 1]) for i in range(n - 1, len(xs))]
            return [v / n for v in out] if f == 'MOVAV' else out
        if f == 'TREND':
            base = _period_idx(_sarg(args[0], '@TREND')) if args else 0
            return [float(t - base) for t in range(lo, hi + 1)]
        if f in ('DURING', 'ISPERIOD'):
            return _sample_mask(_sarg(args[0], '@' + f), lo, hi)
        if f in ('BEFORE', 'AFTER'):
            i0 = _period_idx(_sarg(args[0], '@' + f))
            return [1.0 if (t < i0 if f == 'BEFORE' else t >= i0) else 0.0 for t in range(lo, hi + 1)]
        if f == 'SEAS':
            q = _const(args[0])
            fld = 'MONTH' if ppy == 12 else 'QUARTER'
            return [1.0 if _period_field(t, fld) == q else 0.0 for t in range(lo, hi + 1)]
        if f in ('YEAR', 'QUARTER', 'MONTH'):
            return [_period_field(t, f) for t in range(lo, hi + 1)]
        if f == 'ELEM':
            i0 = _period_idx(_sarg(args[1], '@ELEM'))
            v = evaluate(args[0], get, i0, i0, ppy)[0]
            return [v] * (hi - lo + 1)
        if f in ('RECODE', 'IIF'):
            c = evaluate(args[0], get, lo, hi, ppy)
            a = evaluate(args[1], get, lo, hi, ppy)
            b = evaluate(args[2], get, lo, hi, ppy)
            return [NAN if cc != cc else (aa if cc else bb) for cc, aa, bb in zip(c, a, b)]
        if f in ('NAN', 'NA2'):   # @nan(x, y): x, with missing values replaced by y
            a = evaluate(args[0], get, lo, hi, ppy)
            b = evaluate(args[1], get, lo, hi, ppy)
            return [bb if aa != aa else aa for aa, bb in zip(a, b)]
        if f == 'ISNA':
            return [1.0 if v != v else 0.0 for v in evaluate(args[0], get, lo, hi, ppy)]
        if f in ('MAX', 'MIN') and len(args) == 2 and args[1][0] != 'str':
            a = evaluate(args[0], get, lo, hi, ppy)
            b = evaluate(args[1], get, lo, hi, ppy)
            g = max if f == 'MAX' else min
            return [NAN if (x != x or y != y) else g(x, y) for x, y in zip(a, b)]
        if f in ('ROUND', 'FLOOR', 'CEILING', 'CEIL', 'INT') and len(args) == 1:
            g = {'ROUND': _round_half_away, 'FLOOR': math.floor, 'CEILING': math.ceil, 'CEIL': math.ceil, 'INT': int}[f]
            return [float(g(v)) if v == v else NAN for v in evaluate(args[0], get, lo, hi, ppy)]
        if f in ('LOG10',):
            return [math.log10(v) if v > 0 else NAN for v in evaluate(args[0], get, lo, hi, ppy)]
        if f in ('INV',):
            return [1.0 / v if v else NAN for v in evaluate(args[0], get, lo, hi, ppy)]
        if f in ('LAG',):
            n = int(_const(args[1])) if len(args) > 1 else 1
            return evaluate(args[0], get, lo - n, hi - n, ppy)
        r = _eviews_extra(f, args, get, lo, hi, ppy)
        if r is not None:
            return r
        raise EvalError('unsupported function %s' % node[1])
    raise EvalError('bad node')


# ---------------------------------------------------------------------------
# The rest of the EViews series functions (Command Reference, "Operator and Function Reference")
# ---------------------------------------------------------------------------
def _round_half_away(v):
    return math.copysign(math.floor(abs(v) + 0.5), v)


def _nz(xs):
    return [v for v in xs if v == v]


def _mean(xs):
    return sum(xs) / len(xs) if xs else NAN


def _var(xs, ddof):
    n = len(xs)
    if n - ddof <= 0:
        return NAN
    m = sum(xs) / n
    return sum((v - m) ** 2 for v in xs) / (n - ddof)


def _median(xs):
    xs = sorted(xs)
    n = len(xs)
    if not n:
        return NAN
    return xs[n // 2] if n % 2 else (xs[n // 2 - 1] + xs[n // 2]) / 2


def _quantile(xs, q):   # Cleveland definition (as EViews)
    xs = sorted(xs)
    n = len(xs)
    if not n:
        return NAN
    pos = q * n - 0.5
    if pos <= 0:
        return xs[0]
    if pos >= n - 1:
        return xs[-1]
    i = int(math.floor(pos))
    return xs[i] + (pos - i) * (xs[i + 1] - xs[i])


def _skew(xs):
    n = len(xs)
    if n < 2:
        return NAN
    m = sum(xs) / n
    sd = math.sqrt(sum((v - m) ** 2 for v in xs) / n)
    return sum(((v - m) / sd) ** 3 for v in xs) / n if sd else NAN


def _kurt(xs):
    n = len(xs)
    if n < 2:
        return NAN
    m = sum(xs) / n
    sd = math.sqrt(sum((v - m) ** 2 for v in xs) / n)
    return sum(((v - m) / sd) ** 4 for v in xs) / n if sd else NAN


def _cov(xs, ys, ddof):
    pr = [(a, b) for a, b in zip(xs, ys) if a == a and b == b]
    n = len(pr)
    if n - ddof <= 0:
        return NAN
    ma, mb = sum(a for a, _ in pr) / n, sum(b for _, b in pr) / n
    return sum((a - ma) * (b - mb) for a, b in pr) / (n - ddof)


def _cor(xs, ys):
    c = _cov(xs, ys, 0)
    pr = [(a, b) for a, b in zip(xs, ys) if a == a and b == b]
    va, vb = _var([a for a, _ in pr], 0), _var([b for _, b in pr], 0)
    return c / math.sqrt(va * vb) if va and vb and va == va and vb == vb else NAN


_STAT1 = {   # name -> function of the non-missing values
    'SUM': lambda x: sum(x), 'AV': _mean, 'MEAN': _mean, 'MEDIAN': _median,
    'STDEV': lambda x: math.sqrt(_var(x, 1)) if len(x) > 1 else NAN, 'STDEVS': lambda x: math.sqrt(_var(x, 1)) if len(x) > 1 else NAN,
    'STDEVP': lambda x: math.sqrt(_var(x, 0)) if x else NAN,
    'VAR': lambda x: _var(x, 0), 'VARP': lambda x: _var(x, 0), 'VARS': lambda x: _var(x, 1),
    'MAX': lambda x: max(x) if x else NAN, 'MIN': lambda x: min(x) if x else NAN,
    'SUMSQ': lambda x: sum(v * v for v in x), 'SKEW': _skew, 'KURT': _kurt,
    'PROD': lambda x: math.prod(x) if x else NAN,
}
_STAT2 = {'COV': lambda x, y: _cov(x, y, 0), 'COVP': lambda x, y: _cov(x, y, 0), 'COVS': lambda x, y: _cov(x, y, 1),
          'COR': _cor, 'INNER': lambda x, y: sum(a * b for a, b in zip(x, y) if a == a and b == b)}


def _sample_vals(args, k, get, ppy):
    """Whole-workfile values of args[0] (and args[1] for two-series stats), restricted to an optional sample."""
    T = len(PERIODS['keys'])
    xs = [evaluate(a, get, 0, T - 1, ppy) for a in args[:k]]
    if len(args) > k:
        m = _sample_mask(_sarg(args[k], 'sample'), 0, T - 1)
        xs = [[v if mm else NAN for v, mm in zip(x, m)] for x in xs]
    return xs


def _normcdf(x):
    return 0.5 * math.erfc(-x / math.sqrt(2))


def _normq(p):
    if not 0 < p < 1:
        return NAN if not (p == 0 or p == 1) else (-float('inf') if p == 0 else float('inf'))
    lo, hi = -40.0, 40.0
    for _ in range(200):
        mid = (lo + hi) / 2
        if _normcdf(mid) < p:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def _elem1(f):
    """Element-by-element functions of one argument."""
    E = {
        'FACT': lambda v: float(math.factorial(int(v))) if v >= 0 and v == int(v) and v < 171 else NAN,
        'FACTLOG': lambda v: math.lgamma(v + 1) if v >= 0 else NAN,
        'EXPM1': math.expm1, 'LOG1P': lambda v: math.log1p(v) if v > -1 else NAN,
        'LOG1MEXP': lambda v: math.log(-math.expm1(v)) if v < 0 else NAN,
        'SIGN': lambda v: float((v > 0) - (v < 0)),
        'LOGIT': lambda v: 1 / (1 + math.exp(-v)) if v > -700 else 0.0,
        'GAMMA': lambda v: math.gamma(v) if v > 0 or v != int(v) else NAN,
        'GAMMALOG': lambda v: math.lgamma(v) if v > 0 else NAN,
        'SIN': math.sin, 'COS': math.cos, 'TAN': math.tan,
        'ASIN': lambda v: math.asin(v) if -1 <= v <= 1 else NAN, 'ACOS': lambda v: math.acos(v) if -1 <= v <= 1 else NAN,
        'ATAN': math.atan,
        'CNORM': _normcdf, 'DNORM': lambda v: math.exp(-v * v / 2) / math.sqrt(2 * math.pi), 'QNORM': _normq,
        'CLOGISTIC': lambda v: 1 / (1 + math.exp(-v)) if v > -700 else 0.0,
        'DLOGISTIC': lambda v: math.exp(-v) / (1 + math.exp(-v)) ** 2 if abs(v) < 700 else 0.0,
        'QLOGISTIC': lambda v: math.log(v / (1 - v)) if 0 < v < 1 else NAN,
        'ERF': math.erf, 'ERFC': math.erfc,
        'ISNA': None,
    }
    return E.get(f)


def _eviews_extra(f, args, get, lo, hi, ppy):
    n_out = hi - lo + 1
    T = len(PERIODS['keys'])

    def ev(a, l=lo, h=hi):
        return evaluate(a, get, l, h, ppy)

    def safe(g, xs):
        out = []
        for v in xs:
            try:
                out.append(g(v) if v == v else NAN)
            except (ValueError, OverflowError, ZeroDivisionError):
                out.append(NAN)
        return out

    # numeric constants
    if f in ('PI',) and not args:
        return [math.pi] * n_out
    if f in ('E',) and not args:
        return [math.e] * n_out
    if f in ('NA',) and not args:
        return [NAN] * n_out
    if f == 'OBSNUM':
        return [float(t + 1) for t in range(lo, hi + 1)]
    if f == 'TRENDC':
        return [float(t) for t in range(lo, hi + 1)]
    if f == 'TRENDBR':
        b = _period_idx(_sarg(args[0], '@TRENDBR'))
        return [float(max(0, t - b)) for t in range(lo, hi + 1)]
    # one-argument element functions
    g = _elem1(f)
    if g is not None and len(args) == 1:
        return safe(g, ev(args[0]))
    # two-argument element functions
    if f in ('MOD', 'POW', 'POWM1', 'POW1PM1', 'LOGX', 'ATAN2', 'EQNA', 'NEQNA'):
        a, b = ev(args[0]), ev(args[1])
        out = []
        for x, y in zip(a, b):
            if f == 'EQNA':
                out.append(1.0 if (x == y or (x != x and y != y)) else 0.0)
                continue
            if f == 'NEQNA':
                out.append(0.0 if (x == y or (x != x and y != y)) else 1.0)
                continue
            if x != x or y != y:
                out.append(NAN)
                continue
            try:
                out.append({'MOD': lambda: math.fmod(x, y), 'POW': lambda: _pow(x, y), 'POWM1': lambda: _pow(x, y) - 1,
                            'POW1PM1': lambda: _pow(1 + x, y) - 1, 'LOGX': lambda: math.log(x) / math.log(y),
                            'ATAN2': lambda: math.atan2(x, y)}[f]())
            except (ValueError, ZeroDivisionError, OverflowError):
                out.append(NAN)
        return out
    if f in ('ROUND', 'FLOOR', 'CEILING', 'CEIL') and len(args) == 2:
        x, n = ev(args[0]), _const(args[1])
        if f == 'ROUND':     # documented: n=2 rounds to the nearest 100, n=-2 to the nearest 0.01
            sc = 10.0 ** n
            return [_round_half_away(v / sc) * sc if v == v else NAN for v in x]
        sc = 10.0 ** n       # documented: @floor(@pi,2) = 3.14
        g2 = math.floor if f == 'FLOOR' else math.ceil
        return [g2(v * sc) / sc if v == v else NAN for v in x]
    if f == 'BOUNDS':
        x, a, b = ev(args[0]), ev(args[1]), ev(args[2])
        return [NAN if NAN in (v,) or v != v else min(max(v, aa), bb) for v, aa, bb in zip(x, a, b)]
    if f == 'BETWEEN':
        x, a, b = ev(args[0]), ev(args[1]), ev(args[2])
        return [NAN if v != v else (1.0 if aa <= v <= bb else 0.0) for v, aa, bb in zip(x, a, b)]
    if f == 'INLIST':
        vals = {float(t) for t in _sarg(args[1], '@INLIST').replace(',', ' ').split()}
        return [NAN if v != v else (1.0 if v in vals else 0.0) for v in ev(args[0])]
    if f == 'IFF':
        c, a, b = ev(args[0]), ev(args[1]), ev(args[2])
        return [NAN if cc != cc else (aa if cc else bb) for cc, aa, bb in zip(c, a, b)]
    if f == 'BRIDGE':   # fill NAs with the last non-missing value
        xs = ev(args[0], 0, hi)
        last = NAN
        out = []
        for v in xs:
            if v == v:
                last = v
            out.append(last)
        return out[lo:] if lo >= 0 else [NAN] * (-lo) + out
    if f == 'YTD':      # cumulative sum within the calendar year
        xs = ev(args[0], 0, hi)
        out, acc, yr = [], 0.0, None
        for t, v in enumerate(xs):
            y = PERIODS['keys'][t][0] if t < T else None
            if y != yr:
                acc, yr = 0.0, y
            acc += v if v == v else 0.0
            out.append(acc)
        return out[lo:] if lo >= 0 else [NAN] * (-lo) + out
    if f in ('DEMEAN', 'DETREND'):
        x = _sample_vals(args, 1, get, ppy)[0]
        full = ev(args[0], 0, T - 1)
        ts = [t for t, v in enumerate(x) if v == v]
        if f == 'DEMEAN':
            m = _mean(_nz(x))
            res = [v - m for v in full]
        else:
            if len(ts) < 2:
                return [NAN] * n_out
            mt, mv = _mean(ts), _mean([x[t] for t in ts])
            b = sum((t - mt) * (x[t] - mv) for t in ts) / sum((t - mt) ** 2 for t in ts)
            res = [v - (mv + b * (t - mt)) for t, v in enumerate(full)]
        return [res[t] if 0 <= t < T else NAN for t in range(lo, hi + 1)]
    # descriptive statistics over the workfile or a sample -> the same value in every period
    if f in _STAT1 and (len(args) == 1 or (len(args) == 2 and args[1][0] == 'str')):
        v = _STAT1[f](_nz(_sample_vals(args, 1, get, ppy)[0]))
        return [v] * n_out
    if f in ('OBS', 'NAS'):
        x = _sample_vals(args, 1, get, ppy)[0]
        m = _sample_mask(_sarg(args[1], 'sample'), 0, T - 1) if len(args) > 1 else [1] * T
        c = sum(1 for v, mm in zip(x, m) if mm and (v == v) == (f == 'OBS'))
        return [float(c)] * n_out
    if f in _STAT2 and (len(args) == 2 or (len(args) == 3 and args[2][0] == 'str')):
        x, y = _sample_vals(args, 2, get, ppy)
        return [_STAT2[f](x, y)] * n_out
    if f in ('QUANTILE',):
        x = _sample_vals([args[0]] + list(args[2:]), 1, get, ppy)[0]
        return [_quantile(_nz(x), _const(args[1]))] * n_out
    if f in ('FIRST', 'LAST'):
        x = _nz(_sample_vals(args, 1, get, ppy)[0])
        return [(x[0] if f == 'FIRST' else x[-1]) if x else NAN] * n_out
    # moving (rolling) statistics: @mov* propagate NAs, @m* skip them
    m = re.match(r'^(MOV|M)(SUM|AV|AVC|STDEV|STDEVS|STDEVP|VAR|VARS|VARP|MAX|MIN|MEDIAN|SUMSQ|SKEW|KURT|OBS|NAS|COV|COVS|COVP|COR|INNER)$', f)
    if m and f not in ('MAX', 'MIN', 'MEDIAN'):
        prop, st = m.group(1) == 'MOV', m.group(2)
        if st in _STAT2:
            n = int(_const(args[2]))
            xs = [ev(args[0], lo - n + 1, hi), ev(args[1], lo - n + 1, hi)]
            out = []
            for i in range(n - 1, len(xs[0])):
                a, b = xs[0][i - n + 1:i + 1], xs[1][i - n + 1:i + 1]
                out.append(NAN if prop and any(v != v for v in a + b) else _STAT2[st](a, b))
            return out
        n = int(_const(args[1]))
        if st == 'AVC':     # centred; even n: window n+1 with half weight at both ends
            if n % 2:
                h = n // 2
                w = [1.0] * n
            else:
                h = n // 2
                w = [0.5] + [1.0] * (n - 1) + [0.5]
            xs = ev(args[0], lo - h, hi + h)
            out = []
            for i in range(n_out):
                win = xs[i:i + len(w)]
                if prop and any(v != v for v in win):
                    out.append(NAN)
                    continue
                pr = [(v, ww) for v, ww in zip(win, w) if v == v]
                out.append(sum(v * ww for v, ww in pr) / sum(ww for _, ww in pr) if pr else NAN)
            return out
        xs = ev(args[0], lo - n + 1, hi)
        out = []
        for i in range(n - 1, len(xs)):
            win = xs[i - n + 1:i + 1]
            if st == 'NAS':
                out.append(float(sum(1 for v in win if v != v)))
                continue
            if st == 'OBS':
                out.append(float(sum(1 for v in win if v == v)))
                continue
            if prop and any(v != v for v in win):
                out.append(NAN)
                continue
            nz = _nz(win)
            out.append(_STAT1['MEAN' if st == 'AV' else st](nz) if nz else NAN)
        return out
    # cumulative statistics (NAs skipped), forward @cum* or backward @cumb*, optional sample
    m = re.match(r'^CUM(B?)(SUM|PROD|MEAN|STDEV|STDEVS|STDEVP|VAR|VARS|VARP|MAX|MIN|SUMSQ|OBS|NAS|MEDIAN|QUANTILE)$', f)
    if m:
        back, st = m.group(1) == 'B', m.group(2)
        q = _const(args[1]) if st == 'QUANTILE' else None
        rest = args[2:] if st == 'QUANTILE' else args[1:]
        full = ev(args[0], 0, T - 1)
        mask = _sample_mask(_sarg(rest[0], 'sample'), 0, T - 1) if rest else [1.0] * T
        idx = [t for t in range(T) if mask[t]]
        res = [NAN] * T
        seq = idx[::-1] if back else idx
        acc = []
        for t in seq:
            v = full[t]
            if st == 'NAS':
                acc.append(v)
                res[t] = float(sum(1 for z in acc if z != z))
                continue
            if v == v:
                acc.append(v)
            if st == 'OBS':
                res[t] = float(len(acc))
            elif st == 'QUANTILE':
                res[t] = _quantile(acc, q)
            elif acc:
                res[t] = _STAT1['MEAN' if st == 'MEAN' else st](acc)
        return [res[t] if 0 <= t < T else NAN for t in range(lo, hi + 1)]
    return None


def ast_text(node):
    k = node[0]
    if k == 'num':
        v = node[1]
        return ('%d' % v) if v.is_integer() else ('%g' % v)
    if k == 'var':
        return node[1] if node[2] == 0 else '%s(%d)' % (node[1], node[2])
    if k == 'coef':
        return 'C(%d)' % node[1]
    if k == 'neg':
        return '-' + ast_text(node[1])
    if k == 'bin':
        return '%s%s%s' % (ast_text(node[2]), node[1], ast_text(node[3]))
    if k == 'call':
        return '%s(%s)' % (node[1], ','.join(ast_text(a) for a in node[2]))
    if k == 'str':
        return '"%s"' % node[1]
    return '?'


class Equation:
    __slots__ = ('target', 'lhs', 'rhs', 'lhs_vars', 'rhs_vars', 'drivers', 'text',
                 'kind', 'evaluable', 'source', 'why')

    def __init__(self, target, drivers, text, kind, source, lhs=None, rhs=None,
                 lhs_vars=(), rhs_vars=(), evaluable=False, why=''):
        self.target = target
        self.drivers = list(drivers)
        self.text = text
        self.kind = kind
        self.source = source
        self.lhs, self.rhs = lhs, rhs
        self.lhs_vars, self.rhs_vars = list(lhs_vars), list(rhs_vars)
        self.evaluable = evaluable
        self.why = why


_NAME_SCAN = re.compile(r'(?<![\w$.@%#])([A-Za-z_][A-Za-z0-9_$.%#]*)')


def _scan_names(s):
    out = []
    for m in _NAME_SCAN.finditer(s):
        nm = m.group(1).upper()
        if nm in FUNCS or nm in EXCLUDE or nm in out:
            continue
        out.append(nm)
    return out


def _split_terms(s):
    terms, depth, cur = [], 0, ''
    for ch in s:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        if ch.isspace() and depth == 0:
            if cur:
                terms.append(cur)
            cur = ''
        else:
            cur += ch
    if cur:
        terms.append(cur)
    return terms


_SKIP_START = re.compile(
    r'^(for|next|if|endif|else|while|wend|smpl|include|call|subroutine|endsub|string|scalar|'
    r'vector|matrix|show|freeze|delete|wfopen|wfsave|pageselect|assign|solve|output|close|exit|'
    r'stop|logmode|model|copy|fetch|store|group|graph|table|text|coef|sample|pagecreate|wfcreate|'
    r'open|save|print|setmaxerrs|mode|tic|toc|statusline)\b', re.I)
_SPEC = re.compile(r'^(?:equation\s+)?(?:[A-Za-z_][\w$]*\.)?(ls|tsls|gmm|nls|arch|binary|logit|probit|'
                   r'glm|ardl|qreg|robustls|cointreg)\b(?:\([^)]*\))?\s*', re.I)


def _strip_comment(s):
    """Drop a trailing EViews comment (' ...) that is not inside a quoted string."""
    inq = False
    for i, ch in enumerate(s):
        if ch == '"':
            inq = not inq
        elif ch == "'" and not inq:
            return s[:i].rstrip()
    return s


def _drop_error_terms(s):
    """Remove EViews error-term options "[ar(1)=0.8, uncond, ...]" and the + that joins them to the formula.
    Scans character by character: brackets inside quotes and a missing closing ] are handled.
    Returns (formula without them, list of the removed [..] parts)."""
    out, found, i, n = [], [], 0, len(s)
    while i < n:
        if s[i] != '[':
            out.append(s[i]); i += 1; continue
        j, q = i + 1, None                      # find the matching ]
        while j < n:
            c = s[j]
            if q:
                if c == q:
                    q = None
            elif c in '"\'':
                q = c
            elif c == ']':
                break
            j += 1
        found.append(s[i:j + 1].strip() if j < n else s[i:].strip() + ']')
        head = ''.join(out).rstrip()
        if head.endswith('+'):                   # "... + [..]"
            head = head[:-1].rstrip()
            out = [head + ' ']
            i = j + 1
        else:                                    # "[..] + ..." or "= [..]"
            out = [head + ' ']
            k = j + 1
            while k < n and s[k].isspace():
                k += 1
            i = k + 1 if k < n and s[k] == '+' else j + 1
    return re.sub(r'\s{2,}', ' ', ''.join(out)).strip(), found


def parse_equation_line(raw, source):
    """Parse one logical line of EViews text. Returns Equation, ('addlink', y, af) or None."""
    s = raw.strip()
    if not s or s[0] in "'#;!%{" or s.startswith('//'):
        return None
    # characters pasted from Word / Excel
    s = (s.replace('\u2212', '-').replace('\u2013', '-').replace('\u2014', '-').replace('\u00a0', ' ')
          .replace('\u201c', '"').replace('\u201d', '"').replace('\t', ' ')
          .replace('\uff0b', '+').replace('\uff3b', '[').replace('\uff3d', ']').replace('\uff08', '(')
          .replace('\uff09', ')').replace('\uff1d', '=').replace('\uff0c', ',').replace('\u3000', ' '))
    s = re.sub('[\u200b\u200c\u200d\u2060\ufeff\u00ad]', '', s)    # invisible characters
    s = _strip_comment(s)
    shown = s.strip()                       # what the user sees, error-term options included
    # error terms written as an added term: "... + [ar(1)=0.13, ...]" or "[ar(1)=.1] + ..."
    s, _br = _drop_error_terms(s)
    s = re.sub(r'[+\-*/^]\s*$', '', s.strip()).strip()   # a dangling operator left at the end
    s = re.sub(r'=\s*\+', '= ', s)
    s = re.sub(r"\s{2,}", " ", s)
    if not s or s.startswith(':'):
        return None
    m = re.match(r'^@ADD\w*\s*(?:\([^)]*\))?\s+(\S+)\s+(\S+)', s, re.I)
    if m:
        return ('addlink', m.group(1).upper(), m.group(2).upper())
    if _SKIP_START.match(s):
        return None
    kind = 'equation'
    m = re.match(r'^@IDENTITY\s+', s, re.I)
    if m:
        kind, s = 'identity', s[m.end():]
    m = re.match(r'^(series|genr|frml)\s+', s, re.I)
    if m:
        kind, s = 'identity', s[m.end():]
    spec = False
    m = _SPEC.match(s)
    if m:
        spec, s = True, s[m.end():]
        if s.lower().startswith('equation '):
            return None
    if s.startswith('@'):
        return None
    text = s.strip()
    br = _drop_error_terms(shown)[1]
    if br:
        text = text + '   ' + ' '.join(b.strip() for b in br)
    why = ''
    try:
        toks = tokenize(s)
    except ParseError as e:
        toks = None
        why = str(e)
    if toks is not None:
        depth, eqi = 0, -1
        for i, (kd, v) in enumerate(toks):
            if kd != 'op':
                continue
            if v == '(':
                depth += 1
            elif v == ')':
                depth -= 1
            elif v == '=' and depth == 0:
                eqi = i
                break
        if eqi > 0:
            try:
                lhs = parse_expr(toks[:eqi])
                rhs = parse_expr(toks[eqi + 1:])
            except ParseError as e:
                lhs = rhs = None
                why = str(e)
                if eqi + 1 >= len(toks):
                    why = 'nothing after "="'

            if lhs is not None:
                lv, rv = ast_vars(lhs), ast_vars(rhs)
                if not lv:
                    return None
                tgt = lv[0]
                drivers = rv + [v for v in lv[1:] if v not in rv]
                coef = ast_has(rhs, ('coef',)) or ast_has(lhs, ('coef',))
                return Equation(tgt, drivers, text, kind, source, lhs, rhs, lv, rv,
                                evaluable=not coef,
                                why='has unsubstituted coefficients C(i)' if coef else '')
        elif eqi < 0 and spec:
            terms = _split_terms(s)
            names = []
            for t in terms:
                try:
                    names.append(ast_vars(parse_expr(tokenize(t))))
                except ParseError:
                    names.append(_scan_names(t))
            if not names or not names[0]:
                return None
            tgt = names[0][0]
            drivers = []
            for lst in names[1:] + [names[0][1:]]:
                for v in lst:
                    if v not in drivers:
                        drivers.append(v)
            return Equation(tgt, drivers, text, 'spec', source,
                            why='specification only (no coefficients)')
    # fallback: dependencies only
    if '=' not in s:
        return None
    left, right = s.split('=', 1)
    lv, rv = _scan_names(left), _scan_names(right)
    if not lv:
        return None
    return Equation(lv[0], rv + [v for v in lv[1:] if v not in rv], text, kind, source,
                    why='could not parse: %s' % (why or 'unknown problem'))


def _eq_sig(e):
    """What counts as 'the same equation': same parsed formula (spacing, case and error-term options ignored)."""
    if e.lhs is not None and e.rhs is not None:
        return ast_text(e.lhs) + '=' + ast_text(e.rhs)
    return re.sub(r'\s+', '', _drop_error_terms(e.text)[0]).upper()


def _logical_lines(text):
    out, buf = [], ''
    for line in text.splitlines():
        st = line.strip()
        if not st:
            if buf:
                out.append(buf)
                buf = ''
            continue
        if buf and (buf.count('(') > buf.count(')') or buf.count('[') > buf.count(']') or re.search(r'[-+*/^(,=]\s*$', buf)
                    or re.match(r'^[-+*/^]', st)) and not st.startswith(("'", ':', '@')):
            buf += ' ' + st
        else:
            if buf:
                out.append(buf)
            buf = st
    if buf:
        out.append(buf)
    return out


EQ_EXT = {'.txt', '.prg', '.eq', '.eqs', '.mod', '.model', '.eviews', '.text'}


def load_equations(paths):
    eqs, addlinks, stats, fails = {}, {}, Counter(), []
    files = []
    for p in paths:
        if os.path.isdir(p):
            for root, _, fs in os.walk(p):
                for f in sorted(fs):
                    if os.path.splitext(f)[1].lower() in EQ_EXT:
                        files.append(os.path.join(root, f))
        else:
            files.append(p)
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        if ext in ('.csv', '.tsv', '.xlsx', '.xlsm'):
            grid, _ = read_grid(f)
            lines = [str(c) for r in grid for c in r if c and ('=' in str(c) or _SPEC.match(str(c).strip()))]
        else:
            with open(f, 'r', encoding='utf-8-sig', errors='replace') as fh:
                lines = _logical_lines(fh.read())
        stats['files'] += 1
        for ln in lines:
            stats['lines'] += 1
            r = parse_equation_line(ln, os.path.basename(f))
            if r is None:
                if '=' in ln and len(fails) < 30 and not _SKIP_START.match(ln.strip()):
                    fails.append(ln[:160])
                continue
            if isinstance(r, tuple):
                addlinks[r[1]] = r[2]
                continue
            stats['parsed'] += 1
            prev = eqs.get(r.target)
            if prev is not None:
                stats['duplicate targets'] += 1
                if prev.evaluable and not r.evaluable:
                    continue
            eqs[r.target] = r
    stats['equations'] = len(eqs)
    stats['evaluable'] = sum(1 for e in eqs.values() if e.evaluable)
    return eqs, addlinks, stats, fails


# --------------------------------------------------------------------------
# The comparison engine
# --------------------------------------------------------------------------
def _sum(xs):
    return sum(v for v in xs if v == v)


def _sumabs(xs):
    return sum(abs(v) for v in xs if v == v)


def ast_contemp(node, out=None):
    """Variables used at lag 0 (same-period dependencies)."""
    if out is None:
        out = set()
    k = node[0]
    if k == 'var':
        if node[2] == 0:
            out.add(node[1])
    elif k == 'neg':
        ast_contemp(node[1], out)
    elif k == 'bin':
        ast_contemp(node[2], out)
        ast_contemp(node[3], out)
    elif k == 'call':
        for a in node[2]:
            ast_contemp(a, out)
    return out


def lhs_form(eq):
    l, t = eq.lhs, eq.target
    if l == ('var', t, 0):
        return 'level'
    if l[0] == 'call' and len(l[2]) == 1 and l[2][0] == ('var', t, 0):
        f = l[1].lstrip('@')
        if f == 'DLOG':
            return 'dlog'
        if f in ('LOG', 'LN'):
            return 'log'
        if f == 'D':
            return 'd'
    return 'generic'


def parse_shock(s):
    """'1%' -> ('pct', 1.0); '0.25' or '+0.25' -> ('abs', 0.25); None -> None (use scenario B)."""
    if not s:
        return None
    s = s.strip()
    if s.endswith('%'):
        return ('pct', float(s[:-1]))
    return ('abs', float(s))


class Model:
    def __init__(self, A, B, eqs=None, addlinks=None, start=None, end=None, tol=1e-6,
                 geo_sep='auto', adj_threshold=0.25, meta=None, verbose=True, af_suffix='', exog=None,
                 eqs_b=None):
        t0 = time.time()
        self.A, self.B = A, B
        self.tol = tol
        self.geo_sep = geo_sep
        self.adj_threshold = adj_threshold
        self.verbose = verbose
        keys = sorted(set(A.keys) | set(B.keys))
        self.keys = keys
        self.labels = [A.keys.get(k) or B.keys.get(k) for k in keys]
        self.ppy = infer_ppy(keys)
        T = self.T = len(keys)
        kidx = {k: i for i, k in enumerate(keys)}
        set_periods(keys)
        self.a = self._align(A, kidx, T)
        self.b = self._align(B, kidx, T)
        self.desc = dict(B.desc)
        self.desc.update(A.desc)
        if meta:
            self.desc.update(meta)
        self.eqs = dict(eqs or {})
        self._detect_geo(geo_sep, af_suffix)
        # variables the modeller set exogenously: keep their equation for display, but treat them as
        # assumptions (no equation): never re-solved, so shocks stop there
        # scenario B may have its own equations: compare them with A's, variable by variable
        self.eqs_b, self.eq_status, self.eqs_b_all = {}, {}, None
        if eqs_b is not None:
            self.eqs_b_all = dict(eqs_b)
            for n in set(self.eqs) | set(eqs_b):
                ea, eb = self.eqs.get(n), eqs_b.get(n)
                if ea is not None and eb is not None:
                    if _eq_sig(ea) != _eq_sig(eb):
                        self.eq_status[n] = 'changed'
                        if eb.evaluable:
                            self.eqs_b[n] = eb
                elif ea is not None:
                    self.eq_status[n] = 'only_a'      # B sets it directly: shows up as its add-factor
                else:
                    self.eq_status[n] = 'only_b'      # A sets it directly: use B's formula for both
                    self.eqs[n] = eb
        self.exog = {}
        self.exog_missing = []
        for n in (exog or ()):
            if n in self.eqs:
                self.exog[n] = self.eqs.pop(n)
                self.eqs_b.pop(n, None)
            else:
                self.exog_missing.append(n)
        self.addlinks = addlinks or {}
        self.has_eqs = bool(self.eqs)
        self.common = sorted(set(self.a) & set(self.b))
        self.only_a = sorted(set(self.a) - set(self.b))
        self.only_b = sorted(set(self.b) - set(self.a))

        self.first_div = {}
        for n in self.common:
            fd = self._first_div(self.a[n], self.b[n])
            if fd is not None:
                self.first_div[n] = fd
        lab2i = {lab: i for i, lab in enumerate(self.labels)}

        def find_label(s):
            if s is None:
                return None
            if s in lab2i:
                return lab2i[s]
            p = parse_period(s)
            if p and p[0] in kidx:
                return kidx[p[0]]
            raise SystemExit('period %r not found; data runs %s .. %s'
                             % (s, self.labels[0], self.labels[-1]))
        ws = find_label(start)
        we = find_label(end)
        if ws is None:
            ws = min(self.first_div.values()) if self.first_div else 0
        if we is None:
            we = T - 1
        if we < ws:
            raise SystemExit('end period is before start period')
        self.ws, self.we = ws, we

        # graph
        self.parents = {}
        for t, e in self.eqs.items():
            ds = [d for d in e.drivers if d != t]
            if t in self.eqs_b:
                ds += [d for d in self.eqs_b[t].drivers if d != t and d not in ds]
            self.parents[t] = ds
        self.children = defaultdict(list)
        for t, ps in self.parents.items():
            for p in ps:
                self.children[p].append(t)

        # add-factors: series named like <target>_A, or linked with @ADD
        self.af_suffixes = [s.strip().upper() for s in (af_suffix or '').split(',') if s.strip()]
        allnames = set(self.a) | set(self.b)
        self.af, self.af_series = {}, set()
        for t, e in self.eqs.items():
            for d in e.drivers:
                if self.is_af_name(d):
                    self.af_series.add(d)
                    self.af.setdefault(t, d)
            if t in self.af:
                continue
            for c in self._af_cands(t):
                if c in allnames:
                    self.af[t] = c
                    self.af_series.add(c)
                    break

        self.stats = {}
        for n in self.common:
            st = self._stats(n)
            if st:
                self.stats[n] = st
        self.differs = {n for n, st in self.stats.items() if st['differs']}
        self.differs_any = set(self.first_div)
        self._resid, self._dd, self._lform, self._contemp = {}, {}, {}, {}
        self.load_secs = time.time() - t0

    # ---- helpers
    def _align(self, sc, kidx, T):
        out = {}
        maps = [[kidx[k] for k in blk] for blk in sc.blocks]
        template = array('d', [NAN]) * T
        for name, (bid, arr) in sc.series.items():
            full = array('d', template)
            for i, v in zip(maps[bid], arr):
                full[i] = v
            out[name] = full
        return out

    def is_af_name(self, d):
        for s in self.af_suffixes:
            if d.endswith(s):
                return True
            if self.geo_sep and self.geo_sep in d and d.rsplit(self.geo_sep, 1)[0].endswith(s):
                return True
        return False

    def _af_cands(self, t):
        out = []
        if t in self.addlinks:
            out.append(self.addlinks[t])
        for s in self.af_suffixes:
            out.append(t + s)
            if self.geo_sep and self.geo_sep in t:
                st, g = t.rsplit(self.geo_sep, 1)
                out.append(st + s + self.geo_sep + g)
        return out

    def _first_div(self, a, b):
        tol = self.tol
        for t in range(self.T):
            x, y = a[t], b[t]
            if x == x and y == y and abs(y - x) > tol * max(1.0, abs(x)):
                return t
        return None

    def _scale(self, a, b=None):
        for s in (a, b):
            if s is None:
                continue
            ds = [s[t] - s[t - 1] for t in range(1, len(s)) if s[t] == s[t] and s[t - 1] == s[t - 1]]
            if len(ds) >= 3:
                m = sum(ds) / len(ds)
                sd = math.sqrt(sum((d - m) ** 2 for d in ds) / (len(ds) - 1))
                if sd > 1e-12:
                    return sd
        lv = [abs(v) for v in a if v == v]
        base = (sum(lv) / len(lv)) if lv else 0.0
        return max(base * 0.01, 1e-12)

    def _stats(self, n):
        a, b = self.a[n], self.b[n]
        maxabs, tmax, last, cnt = 0.0, -1, -1, 0
        for t in range(self.ws, self.we + 1):
            x, y = a[t], b[t]
            if x != x or y != y:
                continue
            cnt += 1
            last = t
            d = abs(y - x)
            if d > maxabs:
                maxabs, tmax = d, t
        if cnt == 0:
            return None
        if tmax < 0:
            tmax = last
        dmax = b[tmax] - a[tmax]
        dend = b[last] - a[last]
        pmax = dmax / abs(a[tmax]) * 100 if a[tmax] != 0 else NAN
        pend = dend / abs(a[last]) * 100 if a[last] != 0 else NAN
        differs = maxabs > self.tol * max(1.0, abs(a[tmax])) and n in self.first_div
        return dict(maxabs=maxabs, dmax=dmax, tmax=tmax, dend=dend, pmax=pmax, pend=pend,
                    tend=last, score=min(1e6, maxabs / self._scale(a, b)) if differs else 0.0, differs=differs,
                    first=self.first_div.get(n))

    def _base_name(self, n):
        """Add-factor series are grouped with the variable they adjust (FC_IUSA_A -> FC_IUSA)."""
        for suf in getattr(self, 'af_suffixes', ()):
            if n.endswith(suf) and n[:-len(suf)] and (n[:-len(suf)] in self.a or n[:-len(suf)] in self.b
                                                     or n[:-len(suf)] in self.eqs):
                return n[:-len(suf)]
        return n

    def _detect_geo(self, sep, af_suffix):
        """Find the geography separator and the geography codes.
        A name's last token (after the separator) counts as a geography code only if the same token
        ends the names of at least 3 different concepts, e.g. FGDP_ICAN, FLBR_ICAN, FHPI_ICAN.
        So A_B_ICAN -> concept A_B, geography ICAN; but FRATE_10Y -> no geography."""
        names = set(self.a) | set(self.b) | set(self.eqs)
        afs = {x.strip().upper().lstrip('_.') for x in (af_suffix or '').split(',') if x.strip()}
        best = None
        for sp in (['.', '_'] if sep in (None, '', 'auto') else [sep]):
            stems = defaultdict(set)
            for n in names:
                if sp in n:
                    st, g = n.rsplit(sp, 1)
                    if g and st and g not in afs:
                        stems[g].add(st)
            codes = {g for g, ss in stems.items() if len(ss) >= 3}
            cover = sum(1 for n in names if sp in n and n.rsplit(sp, 1)[1] in codes)
            if best is None or cover > best[2]:
                best = (sp, codes, cover)
        self.geo_sep, self.geo_codes, self.geo_cover = best if best and best[2] else (
            (sep if sep not in (None, '', 'auto') else '.'), set(), 0)

    def geo(self, n):
        n = self._base_name(n)
        if self.geo_sep and self.geo_sep in n:
            g = n.rsplit(self.geo_sep, 1)[1]
            if g in self.geo_codes:
                return g
        return ''

    def stem(self, n):
        n = self._base_name(n)
        if self.geo_sep and self.geo_sep in n:
            st, g = n.rsplit(self.geo_sep, 1)
            if g in self.geo_codes:
                return st
        return n

    def series(self, n, which):
        s = (self.a if which == 'a' else self.b).get(n)
        if s is None:
            s = (self.b if which == 'a' else self.a).get(n)
        return s

    def af_label(self, n):
        st = self.eq_status.get(n)
        if st == 'changed':
            return 'equation changed: %s' % n
        if st == 'only_a':
            return 'equation dropped in B: %s' % n
        if st == 'only_b':
            return 'equation added in B: %s' % n
        return 'add-factor on %s' % n

    def kind(self, n):
        """X potential exogenous (no equation), F add-factor series, Q equation differs between the
        scenarios, A equation whose change comes mostly from its add-factor/residual, E explained by its
        equation, N not evaluable, - no equations."""
        if not self.has_eqs:
            return '-'
        if n in self.eq_status and n in self.eqs and n in self.differs_any:
            return 'Q'
        if n in self.af_series and n not in self.eqs:
            return 'F'
        eq = self.eqs.get(n)
        if eq is None:
            return 'X'
        if not eq.evaluable:
            return 'N'
        d = self.dyn(n)
        if d is None or 'error' in d:
            return 'N'
        if n in self.differs and d['af_share'] >= self.adj_threshold:
            return 'A'
        return 'E'

    def _getter(self, which):
        prim, sec = (self.a, self.b) if which == 'a' else (self.b, self.a)

        def get(n):
            s = prim.get(n)
            if s is None:
                s = sec.get(n)
            if s is None:
                raise EvalError('series %s is not in the data' % n)
            return s
        return get

    # ---- residuals (= add-factor + equation error) in each scenario
    def resid(self, n, which):
        key = (n, which)
        r = self._resid.get(key)
        if r is None:
            eq = self.eqs_b[n] if (which == 'b' and n in self.eqs_b) else self.eqs[n]
            g = self._getter(which)
            l = evaluate(eq.lhs, g, self.ws, self.we, self.ppy)
            rh = evaluate(eq.rhs, g, self.ws, self.we, self.ppy)
            r = ([x - y for x, y in zip(l, rh)], l)
            self._resid[key] = r
        return r[0]

    def resid_changed(self, n):
        if n in self.eqs_b or self.eq_status.get(n) in ('only_a', 'only_b'):
            return True
        ra = self.resid(n, 'a')
        rb = self.resid(n, 'b')
        la = self._resid[(n, 'a')][1]
        for x, y, l in zip(ra, rb, la):
            if x == x and y == y and abs(y - x) > self.tol * max(1.0, abs(l) if l == l else 1.0):
                return True
        return False

    # ---- model simulation (solve equations period by period)
    def _form(self, n, eq=None):
        key = (n, id(eq)) if eq is not None else n
        f = self._lform.get(key)
        if f is None:
            eq = eq or self.eqs[n]
            # Y also appears on the right-hand side in the same quarter (e.g. @MOVAV(Y,4)):
            # both sides must be solved together
            f = 'implicit' if n in ast_contemp(eq.rhs) else lhs_form(eq)
            self._lform[key] = f
        return f

    def _cont(self, n):
        c = self._contemp.get(n)
        if c is None:
            e = self.eqs[n]
            c = ast_contemp(e.rhs) | (ast_contemp(e.lhs) - {n})
            if n in self.eqs_b:
                e2 = self.eqs_b[n]
                c = c | ast_contemp(e2.rhs) | (ast_contemp(e2.lhs) - {n})
            self._contemp[n] = c
        return c

    def _order(self, nodes):
        ns = set(nodes)
        preds = {n: [p for p in self._cont(n) if p in ns and p != n] for n in nodes}
        succ = defaultdict(list)
        indeg = {n: len(preds[n]) for n in nodes}
        for n, ps in preds.items():
            for p in ps:
                succ[p].append(n)
        q = deque(sorted(n for n in nodes if indeg[n] == 0))
        order = []
        while q:
            u = q.popleft()
            order.append(u)
            for v in succ[u]:
                indeg[v] -= 1
                if indeg[v] == 0:
                    q.append(v)
        done = set(order)
        cyclic = [n for n in nodes if n not in done]
        return order, cyclic

    def _solve(self, n, get, arr, t, res, eq=None):
        eq = eq or self.eqs[n]
        f = self._form(n, eq if eq is not self.eqs.get(n) else None)
        if f == 'implicit':
            y = self._secant_implicit(eq, get, arr, t, res)
            arr[t] = y
            return y
        target = evaluate(eq.rhs, get, t, t, self.ppy)[0] + res
        if f == 'level':
            y = target
        elif f == 'dlog':
            y = arr[t - 1] * _exp(target) if t > 0 else NAN
        elif f == 'log':
            y = _exp(target)
        elif f == 'd':
            y = arr[t - 1] + target if t > 0 else NAN
        else:
            y = self._secant(eq, get, arr, t, target)
        arr[t] = y
        return y

    def _secant_implicit(self, eq, get, arr, t, res):
        """Solve LHS(y) - RHS(y) - res = 0 when y appears on both sides in the same quarter."""
        def fn(y):
            arr[t] = y
            return evaluate(eq.lhs, get, t, t, self.ppy)[0] - evaluate(eq.rhs, get, t, t, self.ppy)[0] - res
        y0 = arr[t]
        if y0 != y0:
            y0 = arr[t - 1] if t > 0 and arr[t - 1] == arr[t - 1] else 1.0
        f0 = fn(y0)
        if f0 != f0:
            return NAN
        if abs(f0) < 1e-13 * (1 + abs(y0)):
            return y0
        y1 = y0 * (1 + 1e-4) + 1e-8
        f1 = fn(y1)
        for _ in range(60):
            if f1 != f1 or f1 == f0:
                break
            y2 = y1 - f1 * (y1 - y0) / (f1 - f0)
            y0, f0, y1 = y1, f1, y2
            f1 = fn(y1)
            if abs(f1) < 1e-12 * (1 + abs(y1)) or abs(y1 - y0) < 1e-13 * (1 + abs(y1)):
                break
        return y1

    def _secant(self, eq, get, arr, t, target):
        def fn(y):
            arr[t] = y
            return evaluate(eq.lhs, get, t, t, self.ppy)[0] - target
        y0 = arr[t]
        if y0 != y0:
            y0 = arr[t - 1] if t > 0 and arr[t - 1] == arr[t - 1] else 1.0
        f0 = fn(y0)
        if f0 != f0:
            return NAN
        if abs(f0) < 1e-13 * (1 + abs(target)):
            return y0
        y1 = y0 * (1 + 1e-4) + 1e-8
        f1 = fn(y1)
        for _ in range(50):
            if f1 != f1 or f1 == f0:
                break
            y2 = y1 - f1 * (y1 - y0) / (f1 - f0)
            y0, f0, y1 = y1, f1, y2
            f1 = fn(y1)
            if abs(f1) < 1e-12 * (1 + abs(target)) or abs(y1 - y0) < 1e-13 * (1 + abs(y1)):
                break
        return y1

    def simulate(self, nodes, override=None, resid_b=()):
        """Re-solve `nodes` over the window. Everything else is fixed at scenario-A values,
        except series in `override` (name -> array). Residuals/add-factors come from A,
        except for equations in `resid_b`, which use scenario B's."""
        override = override or {}
        nodes = [n for n in nodes if n in self.eqs and self.eqs[n].evaluable and n not in override]
        sim = {}
        for n in nodes:
            s = self.series(n, 'b' if n in resid_b else 'a')
            if s is None:
                raise EvalError('series %s is not in the data' % n)
            sim[n] = array('d', s)
        res = {n: self.resid(n, 'b' if n in resid_b else 'a') for n in nodes}
        # switching a variable's add-factor to B also switches its formula to B's, where they differ
        eqb = {n: self.eqs_b[n] for n in nodes if n in resid_b and n in self.eqs_b}
        A, B = self.a, self.b

        def get(name):
            s = sim.get(name)
            if s is None:
                s = override.get(name)
                if s is None:
                    s = A.get(name)
                    if s is None:
                        s = B.get(name)
                        if s is None:
                            raise EvalError('series %s is not in the data' % name)
            return s
        order, cyclic = self._order(nodes)
        ws = self.ws
        for t in range(ws, self.we + 1):
            i = t - ws
            for n in order:
                if res[n][i] == res[n][i]:      # NaN residual = equation cannot be computed here (e.g. no lag
                    self._solve(n, get, sim[n], t, res[n][i], eqb.get(n))   # yet): keep the data value
            if cyclic:
                for _ in range(200):
                    mx = 0.0
                    for n in cyclic:
                        if res[n][i] != res[n][i]:
                            continue
                        old = sim[n][t]
                        new = self._solve(n, get, sim[n], t, res[n][i], eqb.get(n))
                        if new == new and old == old:
                            mx = max(mx, abs(new - old) / (1.0 + abs(old)))
                    if mx < 1e-11:
                        break
        return sim

    # ---- contributions of each driver to the gap (B - A) of one variable
    def dyn(self, n, periods=False):
        d = self._dd.get(n)
        if d is not None and (not periods or 'error' in d or 'contrib' in d):
            return d
        eq = self.eqs.get(n)
        if eq is None or not eq.evaluable:
            return None
        try:
            d = self._dyn(eq, periods)
        except EvalError as e:
            d = {'error': str(e)}
        except IndexError:
            d = {'error': 'a function has too few arguments, e.g. @movav needs (series, n)'}
        except (ValueError, ZeroDivisionError, OverflowError, TypeError) as e:
            d = {'error': 'evaluation failed: %s' % e}
        self._dd[n] = d
        return d

    def _dyn(self, eq, periods):
        n = eq.target
        ws, we = self.ws, self.we
        A, Bv = self.series(n, 'a'), self.series(n, 'b')
        if A is None:
            raise EvalError('series %s is not in the data' % n)
        rng = range(ws, we + 1)
        gap = [Bv[t] - A[t] for t in rng]
        base = self.simulate([n])[n]
        fit = max([abs(base[t] - A[t]) / (1.0 + abs(A[t])) for t in rng if base[t] == base[t] and A[t] == A[t]] or [NAN])
        if fit != fit:
            raise EvalError('equation gives missing values over the window')
        drivers = []
        for d in eq.rhs_vars + eq.lhs_vars[1:]:
            if d != n and d in self.differs_any and d not in drivers:
                drivers.append(d)
        contrib = {}
        af = None
        if n in self.eqs_b:
            # the equation itself differs: average every effect over the two equations, so inputs that
            # only B's (or only A's) formula uses still get their share
            for d in eq.rhs_vars + eq.lhs_vars[1:] + self.eqs_b[n].rhs_vars + self.eqs_b[n].lhs_vars[1:]:
                if d != n and d in self.differs_any and d not in drivers:
                    drivers.append(d)
            allb = {d: self.series(d, 'b') for d in drivers}
            base_b = self.simulate([n], resid_b={n})[n]                 # B equation, A inputs
            full_a = self.simulate([n], override=allb)[n]               # A equation, B inputs
            full_b = self.simulate([n], override=allb, resid_b={n})[n]  # B equation, B inputs (= B data)
            for d in drivers:
                sa = self.simulate([n], override={d: allb[d]})[n]
                sb = self.simulate([n], override={d: allb[d]}, resid_b={n})[n]
                contrib[d] = [0.5 * ((sa[t] - base[t]) + (sb[t] - base_b[t])) for t in rng]
            af = [0.5 * ((base_b[t] - base[t]) + (full_b[t] - full_a[t])) for t in rng]
        else:
            for d in drivers:
                s = self.simulate([n], override={d: self.series(d, 'b')})[n]
                contrib[d] = [s[t] - base[t] for t in rng]
        if af is None and self.resid_changed(n):
            s = self.simulate([n], resid_b={n})[n]
            af = [s[t] - base[t] for t in rng]
            big = max([abs(g) for g in gap if g == g] + [0.0])
            lvl = max([abs(A[t]) for t in rng if A[t] == A[t]] + [1.0])
            if max([abs(v) for v in af if v == v] + [0.0]) < max(1e-3 * big, self.tol * lvl):
                af = None  # rounding noise, not a real add-factor change
        inter = []
        for i, g in enumerate(gap):
            c = sum(v[i] for v in contrib.values() if v[i] == v[i])
            if af and af[i] == af[i]:
                c += af[i]
            inter.append(g - c)
        iend = max([i for i, g in enumerate(gap) if g == g] or [-1])
        if iend < 0:
            raise EvalError('no overlapping data in the window')
        ge = gap[iend]

        def sh(v):
            return v / ge if abs(ge) > 1e-12 * max(1.0, abs(A[ws + iend])) else None
        end = {d: c[iend] for d, c in contrib.items()}
        afe = af[iend] if af else 0.0
        absd = sum(_sumabs(c) for c in contrib.values())
        absaf = _sumabs(af) if af else 0.0
        out = dict(target=n, lhs=ast_text(eq.lhs), iend=iend, gap_end=ge, end=end, af_end=afe,
                   inter_end=inter[iend], shares={d: sh(v) for d, v in end.items()},
                   af_share_end=sh(afe) if af else None, formula_share_end=sh(sum(end.values())),
                   af_share=absaf / (absaf + absd) if absaf + absd > 1e-15 else 0.0,
                   af_name=self.af.get(n), af_changed=af is not None, fit_err=fit,
                   pgap=gap, pf=[sum(c[i] for c in contrib.values() if c[i] == c[i]) for i in range(len(gap))],
                   paf=af)
        big = max([abs(g) for g in gap if g == g] + [0.0])
        lim = max(1e-3 * big, 1e-12)
        out['pshare'] = {d: [int(round(1000 * c[i] / g)) if (g == g and c[i] == c[i] and abs(g) > lim) else None
                             for i, g in enumerate(gap)] for d, c in contrib.items()}
        out['pafshare'] = [int(round(1000 * v / g)) if (af and g == g and v == v and abs(g) > lim) else None
                           for v, g in zip(af or [], gap)] if af else None
        if periods:
            out.update(contrib=contrib, af=af, inter=inter, gap=gap)
        return out

    def change_class(self, n):
        """D = changed directly (an assumption, an add-factor series, or its equation's add-factor changed);
        M = moved only through its formula (same equation, same add-factor, its inputs moved);
        U = cannot tell (equation not evaluable); '' = no equations loaded."""
        if not self.has_eqs:
            return ''
        k = self.kind(n)
        if k in ('X', 'F', 'Q'):
            return 'D'
        if k == 'N':
            return 'U'
        d = self.dyn(n)
        return 'D' if d and d.get('af_changed') else 'M'

    def share(self, child, parent):
        d = self.dyn(child)
        if not d or 'error' in d:
            return None
        return d['shares'].get(parent)

    # ---- graph walks
    def descendants(self, x):
        seen, q = {x}, deque([x])
        while q:
            u = q.popleft()
            for v in self.children.get(u, ()):
                if v not in seen:
                    seen.add(v)
                    q.append(v)
        seen.discard(x)
        return seen

    def ancestors(self, y):
        seen, q = {y}, deque([y])
        while q:
            u = q.popleft()
            for v in self.parents.get(u, ()):
                if v not in seen:
                    seen.add(v)
                    q.append(v)
        seen.discard(y)
        return seen

    # ---- impact of X on Y (multi-equation simulation)
    def impact(self, x, y=None, shock=None):
        """Re-run the equations between X and Y with only X changed (to its scenario-B path,
        or by a hypothetical shock). Returns (effects dict name -> list over window, info)."""
        desc = self.descendants(x)
        if y is not None:
            nodes = desc & (self.ancestors(y) | {y})
        else:
            nodes = desc
        ev = [n for n in nodes if n in self.eqs and self.eqs[n].evaluable]
        blocked = sorted(n for n in nodes if n not in self.eqs or not self.eqs[n].evaluable)
        xa = self.series(x, 'a')
        if xa is None:
            raise EvalError('%s is not in the data' % x)
        if shock is None:
            xo = self.series(x, 'b')
        else:
            xo = array('d', xa)
            for t in range(self.ws, self.we + 1):
                xo[t] = xa[t] * (1 + shock[1] / 100.0) if shock[0] == 'pct' else xa[t] + shock[1]
        base = self.simulate(ev)
        sim = self.simulate(ev, override={x: xo})
        rng = range(self.ws, self.we + 1)
        eff = {n: [sim[n][t] - base[n][t] for t in rng] for n in ev}
        eff[x] = [xo[t] - xa[t] for t in rng]
        return eff, {'nodes': len(ev), 'blocked': blocked}

    def sources(self, y):
        """Scenario shocks upstream of Y: changed assumptions / add-factor series / data for
        equations that cannot be evaluated, and equations whose add-factor (residual) changed."""
        anc = self.ancestors(y) | {y}
        out = []
        for s in anc:
            ev = s in self.eqs and self.eqs[s].evaluable
            if not ev and s in self.differs_any and s != y:
                out.append(('data', s))
            elif ev and self.resid_changed(s):
                d = self.dyn(s)
                if d and 'error' not in d and d['af_changed']:
                    out.append(('addfactor', s))
        return out, anc

    def attribute(self, y):
        """Split Y's gap into the effect of each upstream scenario shock."""
        srcs, anc = self.sources(y)
        if y not in self.eqs or not self.eqs[y].evaluable:
            raise EvalError('%s has no evaluable equation' % y)
        nodes = [n for n in anc if n in self.eqs and self.eqs[n].evaluable]
        base = self.simulate(nodes)[y]
        rng = range(self.ws, self.we + 1)
        res = []
        for kind, s in srcs:
            sub = (self.descendants(s) | {s}) & anc if kind == 'addfactor' else self.descendants(s) & anc
            sub = [n for n in sub if n in self.eqs and self.eqs[n].evaluable]
            if kind == 'data':
                sim = self.simulate(sub, override={s: self.series(s, 'b')})
            else:
                sim = self.simulate(sub, resid_b={s})
            yy = sim.get(y, base)
            res.append((kind, s, [yy[t] - base[t] for t in rng]))
        allb = {s for k, s in srcs if k == 'data'}
        sim = self.simulate(nodes, override={s: self.series(s, 'b') for s in allb},
                            resid_b={s for k, s in srcs if k == 'addfactor'})
        joint = [sim[y][t] - base[t] for t in rng]
        A, Bv = self.series(y, 'a'), self.series(y, 'b')
        gap = [Bv[t] - A[t] for t in rng]
        return res, joint, gap

    # ---- influence: how many variables a change in X reaches, and by how much
    def reach_counts(self):
        """Number of variables each variable feeds, directly or indirectly (via the equations)."""
        if getattr(self, '_reach', None) is not None:
            return self._reach
        nodes = sorted(set(self.children) | set(self.eqs) | set(self.a) | set(self.b))
        idx = {n: i for i, n in enumerate(nodes)}
        ch = [[idx[c] for c in self.children.get(n, ()) if c in idx] for n in nodes]
        # iterative Tarjan SCC (components come out in reverse topological order)
        N = len(nodes)
        index, low, onst, comp = [None] * N, [0] * N, [False] * N, [-1] * N
        stack, comps, counter = [], [], 0
        for root in range(N):
            if index[root] is not None:
                continue
            work = [(root, 0)]
            while work:
                v, pi = work.pop()
                if pi == 0:
                    index[v] = low[v] = counter
                    counter += 1
                    stack.append(v)
                    onst[v] = True
                recurse = False
                for k in range(pi, len(ch[v])):
                    w = ch[v][k]
                    if index[w] is None:
                        work.append((v, k + 1))
                        work.append((w, 0))
                        recurse = True
                        break
                    elif onst[w]:
                        low[v] = min(low[v], index[w])
                if recurse:
                    continue
                if low[v] == index[v]:
                    c = []
                    while True:
                        w = stack.pop()
                        onst[w] = False
                        comp[w] = len(comps)
                        c.append(w)
                        if w == v:
                            break
                    comps.append(c)
                if work:
                    u = work[-1][0]
                    low[u] = min(low[u], low[v])
        # reach bitsets per component, children components come first
        reach = [0] * len(comps)
        for ci, c in enumerate(comps):
            bits = 0
            for v in c:
                for w in ch[v]:
                    cw = comp[w]
                    bits |= reach[cw] | (1 << w)
            if len(c) > 1:
                for v in c:
                    bits |= 1 << v
            reach[ci] = bits
        out = {}
        for n, i in idx.items():
            b = reach[comp[i]] & ~(1 << i)
            out[n] = bin(b).count('1')
        self._reach = out
        return out

    def af_target(self, x):
        """If x is an add-factor series attached to an equation (not written inside it), its target."""
        if not hasattr(self, '_af_rev'):
            self._af_rev = {v: k for k, v in self.af.items() if v not in self.eqs[k].drivers}
        return self._af_rev.get(x)

    def shock_mode(self, x, shock):
        """How to apply 'X changes as in scenario B': an add-factor series or an equation driven mostly by
        its add-factor is shocked through its add-factor (the equation is still re-solved)."""
        if shock is None:
            t = self.af_target(x)
            if t:
                return 'addfactor', t
            if x in self.eqs and self.eqs[x].evaluable and self.kind(x) in ('A', 'Q'):
                return 'addfactor', x
        return 'series', x

    def impact_addfactor(self, t):
        nodes = self.descendants(t) | {t}
        ev = [n for n in nodes if n in self.eqs and self.eqs[n].evaluable]
        blocked = sorted(n for n in nodes if n not in self.eqs or not self.eqs[n].evaluable)
        base = self.simulate(ev)
        sim = self.simulate(ev, resid_b={t})
        rng = range(self.ws, self.we + 1)
        return {n: [sim[n][i] - base[n][i] for i in rng] for n in ev}, {'nodes': len(ev), 'blocked': blocked}

    def influence(self, x, shock=None, thr=0.01):
        """Change X (to its scenario-B path, or by a shock), re-solve everything downstream, and
        summarise how many variables move and by how much (% of their scenario-A level)."""
        mode, tgt = self.shock_mode(x, shock)
        if mode == 'addfactor':
            eff, info = self.impact_addfactor(tgt)
        else:
            eff, info = self.impact(x, None, shock)
        rng = list(range(self.ws, self.we + 1))
        rows = []
        for n, e in eff.items():
            if n == x and mode != 'addfactor':
                continue
            a = self.series(n, 'a')
            mx, mxp, tm = 0.0, 0.0, None
            for i, t in enumerate(rng):
                v = e[i]
                if v != v or a[t] != a[t]:
                    continue
                p = abs(v) / abs(a[t]) * 100 if a[t] else (0.0 if v == 0 else float('inf'))
                if p > mxp:
                    mx, mxp, tm = v, p, t
            if tm is not None:
                rows.append((n, mx, mxp if mxp != float('inf') else None, tm, e[-1]))
        moved = [r for r in rows if r[2] is None or r[2] >= thr]
        b1 = sum(1 for r in moved if r[2] is None or r[2] >= 1)
        b2 = sum(1 for r in moved if r[2] is not None and 0.1 <= r[2] < 1)
        b3 = len(moved) - b1 - b2
        moved.sort(key=lambda r: -(r[2] if r[2] is not None else 1e18))
        reach = len(self.descendants(tgt)) + 1 if mode == 'addfactor' else len(self.descendants(x))
        return {'x': x, 'mode': mode, 'target': tgt, 'reach': reach, 'resolved': info['nodes'], 'blocked': len(info['blocked']),
                'moved': len(moved), 'ge1': b1, 'b01': b2, 'b001': b3, 'thr': thr,
                'top': [(n, mx, p, self.labels[t]) for n, mx, p, t, _ in moved[:40]]}

    # ---- ranking helpers
    def ranked(self):
        return sorted(self.differs, key=lambda n: -self.stats[n]['score'])

    def entry_points(self):
        out = [n for n in self.differs if self.kind(n) in ('X', 'A', 'F', 'Q')]
        return sorted(out, key=lambda n: (self.stats[n]['first'], -self.stats[n]['score']))

    def shock_list(self):
        """Scenario shocks, one row each: an add-factor series and the variable it adjusts count once."""
        out, seen = [], set()
        for n in self.entry_points():
            t = self.af_target(n) if self.kind(n) == 'F' else None
            key = t or n
            if key in seen:
                continue
            seen.add(key)
            out.append(n)
        return out

    def downstream(self, x, depth=4):
        dist, prev = {x: 0}, {}
        q = deque([x])
        while q:
            u = q.popleft()
            if dist[u] >= depth:
                continue
            for v in self.children.get(u, ()):
                if v not in dist:
                    dist[v] = dist[u] + 1
                    prev[v] = u
                    q.append(v)
        return dist, prev

    def strongest_path(self, x, y, maxdepth=15):
        import heapq
        INF = float('inf')
        best, prev, hops = {x: 0.0}, {}, {x: 0}
        pq = [(0.0, x)]
        while pq:
            c, u = heapq.heappop(pq)
            if c > best.get(u, INF):
                continue
            if u == y:
                break
            if hops[u] >= maxdepth:
                continue
            for v in self.children.get(u, ()):
                if v not in self.differs_any and v != y:
                    continue
                sh = self.share(v, u)
                w = -math.log(min(1.0, abs(sh))) if sh else 7.0
                if c + w < best.get(v, INF):
                    best[v], prev[v], hops[v] = c + w, u, hops[u] + 1
                    heapq.heappush(pq, (c + w, v))
        if y not in best or y == x:
            return None
        p = [y]
        while p[-1] != x:
            p.append(prev[p[-1]])
        return p[::-1]

    def path_product(self, p):
        prod = 1.0
        for a, b in zip(p, p[1:]):
            sh = self.share(b, a)
            if sh is None:
                return None
            prod *= sh
        return prod

    def paths(self, x, y, k=5, maxdepth=12):
        dist = {x: 0}
        q = deque([x])
        while q:
            u = q.popleft()
            if u == y or dist[u] >= maxdepth:
                continue
            for v in self.children.get(u, ()):
                if v not in dist:
                    dist[v] = dist[u] + 1
                    q.append(v)
        if y not in dist:
            return []
        out = []

        def back(node, acc):
            if len(out) >= k:
                return
            if node == x:
                out.append([x] + acc[::-1])
                return
            for p in self.parents.get(node, ()):
                if dist.get(p) == dist[node] - 1:
                    back(p, acc + [node])
        back(y, [])
        return out

    def ranked_paths(self, x, y, k=5):
        ps = self.paths(x, y, k=40)
        sp = self.strongest_path(x, y)
        if sp and sp not in ps:
            ps.append(sp)

        def key(p):
            v = self.path_product(p)
            return -abs(v) if v is not None else 1
        ps.sort(key=key)
        return ps[:k]


# --------------------------------------------------------------------------
# Formatting helpers
# --------------------------------------------------------------------------
def fmt(v, sig=4):
    if v is None or v != v:
        return 'n/a'
    if v == 0:
        return '0'
    if abs(v) >= 1e4:
        return '{:,.0f}'.format(v)
    return ('%.' + str(sig) + 'g') % v


def fpct(v):
    if v is None or v != v:
        return 'n/a'
    return '%+.2f%%' % v


def fsh(v):
    return '   n/a' if v is None or v != v else '%5.0f%%' % (100 * v)


def _r(v, sig=6):
    if v is None or v != v or v in (float('inf'), float('-inf')):
        return None
    if v == 0:
        return 0
    return float('%.*g' % (sig, v))


KIND_TEXT = {'X': 'potential exogenous (no equation)', 'Q': 'equation differs between A and B', 'F': 'add-factor series', 'A': 'equation, mostly add-factor',
             'E': 'explained by its equation', 'N': 'equation not evaluable', '-': ''}


# --------------------------------------------------------------------------
# Terminal commands
# --------------------------------------------------------------------------
def cmd_info(M, args, eqstats=None, fails=None):
    print('Scenario A: %s' % M.A.label)
    for n in M.A.notes:
        print('   ', n)
    print('Scenario B: %s' % M.B.label)
    for n in M.B.notes:
        print('   ', n)
    print('Periods: %s .. %s (%d), %d per year' % (M.labels[0], M.labels[-1], M.T, M.ppy))
    print('Series: %d in both, %d only in A, %d only in B' % (len(M.common), len(M.only_a), len(M.only_b)))
    print('Comparison window: %s .. %s' % (M.labels[M.ws], M.labels[M.we]))
    if M.eqs_b_all is not None:
        c = Counter(M.eq_status.values())
        print('Equations A vs B: %d changed, %d only in A (set directly in B), %d only in B (set directly in A); '
              'the rest identical' % (c.get('changed', 0), c.get('only_a', 0), c.get('only_b', 0)))
    if M.geo_codes:
        cnt = Counter(M.geo(n) for n in M.common if M.geo(n))
        print('Geography: separator "%s", %d codes (%s%s); %d series have no geography'
              % (M.geo_sep, len(M.geo_codes), ', '.join(g for g, _ in cnt.most_common(8)),
                 ', ...' if len(M.geo_codes) > 8 else '', sum(1 for n in M.common if not M.geo(n))))
    else:
        print('Geography: no geography codes found (no name ending that is shared by 3+ concepts)')
    print('Series that differ: %d' % len(M.differs))
    if eqstats:
        print('Equations: %s' % ', '.join('%s=%d' % kv for kv in sorted(eqstats.items())))
        if M.exog or M.exog_missing:
            print('Set exogenously by the modeller (--exog): %d with an equation, treated as assumptions%s' % (
                len(M.exog), ('; %d listed names have no equation: %s' % (len(M.exog_missing), ', '.join(M.exog_missing[:8])))
                if M.exog_missing else ''))
        print('Add-factor series found: %d (suffix %s), e.g. %s' % (
            len(M.af_series), ','.join(M.af_suffixes), ', '.join(sorted(M.af_series)[:5])))
        bad = Counter(e.why for e in M.eqs.values() if not e.evaluable)
        for why, c in bad.most_common():
            print('    not evaluable: %d (%s)' % (c, why))
        errs, fits = Counter(), []
        for n in list(M.differs)[:3000]:
            d = M.dyn(n)
            if d and 'error' in d:
                errs[re.sub(r'series \S+', 'series X', d['error'])] += 1
        for why, c in errs.most_common(8):
            print('    evaluation error: %d (%s)' % (c, why))
        missing = Counter()
        for e in M.eqs.values():
            for d in e.drivers:
                if d not in M.a and d not in M.b and d not in M.eqs:
                    missing[d] += 1
        if missing:
            print('    drivers not found in data: %d, e.g. %s'
                  % (len(missing), ', '.join(n for n, _ in missing.most_common(8))))
    if fails:
        print('Lines with "=" that could not be read (first %d):' % len(fails))
        for f in fails[:10]:
            print('   ', f)
    print('Loaded in %.1fs' % M.load_secs)


def _filter(M, names, args):
    out = names
    if getattr(args, 'geo', None):
        g = args.geo.upper()
        out = [n for n in out if M.geo(n) == g]
    if getattr(args, 'match', None):
        rx = re.compile(args.match, re.I)
        out = [n for n in out if rx.search(n) or rx.search(M.desc.get(n, ''))]
    if getattr(args, 'type', None):
        out = [n for n in out if M.kind(n) == args.type.upper()]
    return out


def _formula_af(M, n):
    d = M.dyn(n)
    if not d or 'error' in d:
        return None, None
    return d['formula_share_end'], d['af_share_end'] if d['af_changed'] else 0.0


def cmd_top(M, args):
    names = _filter(M, M.ranked(), args)
    print('A = %s   B = %s   window %s..%s   %d series differ'
          % (M.A.label, M.B.label, M.labels[M.ws], M.labels[M.we], len(M.differs)))
    print('score = largest gap / typical period-to-period move of the series in A')
    print('formula / add-f. = share of the gap at the window end explained by the drivers / by the add-factor\n')
    w = max([len(n) for n in names[:args.n]] + [6])
    print('%-*s %7s %11s %9s %9s %8s %7s %7s  %s' % (w, 'series', 'score', 'max gap', 'max %', 'end %',
                                                     'starts', 'formula', 'add-f.', 'type'))
    for n in names[:args.n]:
        s = M.stats[n]
        fs, af = _formula_af(M, n)
        print('%-*s %7.1f %11s %9s %9s %8s %7s %7s  %s  %s' % (
            w, n, s['score'], fmt(s['dmax']), fpct(s['pmax']), fpct(s['pend']),
            M.labels[s['first']] if s['first'] is not None else '', fsh(fs) if fs is not None else '',
            fsh(af) if af is not None else '', M.kind(n), M.desc.get(n, '')[:40]))
    if args.csv:
        with open(args.csv, 'w', newline='', encoding='utf-8') as f:
            wr = csv.writer(f)
            wr.writerow(['series', 'description', 'geo', 'type', 'score', 'max_gap', 'max_gap_period',
                         'max_gap_pct', 'end_gap', 'end_gap_pct', 'first_divergence',
                         'share_from_formula_end', 'share_from_addfactor_end', 'addfactor_series'])
            for n in M.ranked():
                s = M.stats[n]
                fs, af = _formula_af(M, n)
                wr.writerow([n, M.desc.get(n, ''), M.geo(n), M.kind(n), _r(s['score']), _r(s['dmax']),
                             M.labels[s['tmax']], _r(s['pmax']), _r(s['dend']), _r(s['pend']),
                             M.labels[s['first']] if s['first'] is not None else '',
                             _r(fs, 4), _r(af, 4), M.af.get(n, '')])
        print('\nAll %d differing series written to %s' % (len(M.differs), args.csv))


def cmd_addfactors(M, args):
    """Formula vs add-factor for every equation whose variable differs."""
    rows = []
    for n in M.ranked():
        if n not in M.eqs or not M.eqs[n].evaluable:
            continue
        d = M.dyn(n)
        if not d or 'error' in d:
            continue
        rows.append((n, d))
    rows.sort(key=lambda r: -abs(r[1]['af_end']) / (abs(r[1]['gap_end']) + 1e-30) if r[1]['af_changed'] else 0)
    changed = [r for r in rows if r[1]['af_changed']]
    print('%d of %d differing equations have a changed add-factor/residual.\n' % (len(changed), len(rows)))
    w = max([len(n) for n, _ in rows[:args.n]] + [6])
    print('%-*s %11s %9s %8s %8s %8s  %s' % (w, 'series', 'gap end', 'gap %', 'formula', 'add-f.', 'other', 'add-factor series'))
    for n, d in (changed if not args.all else rows)[:args.n]:
        s = M.stats[n]
        afn = d['af_name'] or ''
        afc = ''
        st = M.stats.get(afn) if afn else None
        if st and st['differs']:
            afc = ' (changed: largest %s in %s)' % (fmt(st['dmax']), M.labels[st['tmax']])
        elif afn:
            afc = ' (series unchanged; residual moved)'
        ge = d['gap_end']
        oth = d['inter_end'] / ge if abs(ge) > 1e-12 else None
        print('%-*s %11s %9s %8s %8s %8s  %s%s' % (w, n, fmt(ge), fpct(s['pend']), fsh(d['formula_share_end']),
                                                fsh(d['af_share_end']) if d['af_changed'] else '     0%', fsh(oth), afn, afc))


def _norm(M, name):
    n = name.strip().upper()
    if n in M.a or n in M.b or n in M.eqs or n in M.children:
        return n
    cands = [x for x in set(M.a) | set(M.b) if n in x][:10]
    raise SystemExit('%s not found.%s' % (name, (' Did you mean: ' + ', '.join(sorted(cands))) if cands else ''))


def _expand_focus(M, specs):
    """--focus values: names, comma lists, or wildcards (*_ICAN, FGDP*, ?LBR.*), case-insensitive.
    Wildcards keep only series that differ between the scenarios."""
    import fnmatch
    out = []
    for spec in specs or []:
        for f in spec.split(','):
            f = f.strip().upper()
            if not f:
                continue
            if any(c in f for c in '*?['):
                hits = sorted(n for n in M.differs if fnmatch.fnmatchcase(n, f))
                if not hits:
                    print('--focus %s: no changed series match' % f, file=sys.stderr)
                elif M.verbose:
                    print('--focus %s: %d changed series' % (f, len(hits)), file=sys.stderr)
                out += hits
            else:
                out.append(_norm(M, f))
    return list(dict.fromkeys(out))


def _period_index(M, label):
    if label is None:
        return None
    labs = M.labels[M.ws:M.we + 1]
    if label in labs:
        return labs.index(label)
    p = parse_period(label)
    if p:
        for i, k in enumerate(M.keys[M.ws:M.we + 1]):
            if k == p[0]:
                return i
    raise SystemExit('period %s is not in the window %s..%s' % (label, labs[0], labs[-1]))


def print_contrib(M, n, prefix='', at=None):
    """Print each driver's contribution to the gap of n at one period. Returns contributions."""
    d = M.dyn(n, periods=True)
    eq = M.eqs.get(n)
    if eq is None:
        if n in M.exog:
            print(prefix + '(set exogenously by the modeller: its equation is switched off, so its whole gap is a '
                           'scenario choice. Equation: %s)' % M.exog[n].text[:200])
        else:
            print(prefix + '(no equation: an assumption set by the scenario)')
        return None
    if not d or 'error' in d:
        print(prefix + '(cannot evaluate: %s)' % (d['error'] if d else eq.why))
        ch = [p for p in M.parents.get(n, []) if p in M.differs]
        if ch:
            print(prefix + 'drivers that differ: ' + ', '.join(ch[:15]))
        return None
    i = d['iend'] if at is None else at
    g = d['gap'][i]
    lab = M.labels[M.ws + i]
    print('%sgap B-A in %s = %s' % (prefix, lab, fmt(g)))
    items = sorted(((k, v[i]) for k, v in d['contrib'].items()), key=lambda kv: -abs(kv[1]))
    rows = [(k, v) for k, v in items]
    if d['af']:
        nm = d['af_name'] or 'residual'
        rows.append((M.af_label(n) + ' (formula + add-factor)' if n in M.eq_status else
                     ('add-factor %s' % nm if d['af_name'] else 'add-factor / residual change'), d['af'][i]))
    rows.append(('(interaction, non-linear)', d['inter'][i]))
    big = max([abs(v) for _, v in rows if v == v] + [abs(g) if g == g else 0, 1e-300])
    for k, v in rows:
        if k.startswith('(') and abs(v) < 1e-3 * big:
            continue
        sh = v / g if abs(g) > 1e-12 * big else None
        print('%s  %s %-12s %s' % (prefix, fsh(sh), fmt(v), k))
    if not d['contrib'] and not d['af']:
        print(prefix + '  none of its drivers differ and its add-factor did not change')
    return d


def cmd_explain(M, args):
    root = _norm(M, args.var)
    at = _period_index(M, args.at)
    seen = {root}

    def head(n, prefix):
        s = M.stats.get(n)
        info = ('gap at end %s (%s)' % (fmt(s['dend']), fpct(s['pend']))) if s and s['differs'] else 'no difference'
        print('%s%s  %s  [%s]' % (prefix, n, info, 'set exogenously by the modeller' if n in M.exog else KIND_TEXT.get(M.kind(n), '')))

    def rec(n, depth, prefix):
        d = print_contrib(M, n, prefix + '    ', at)
        if depth <= 1 or not d:
            return
        i = d['iend'] if at is None else at
        g = d['gap'][i]
        for k, c in sorted(d['contrib'].items(), key=lambda kv: -abs(kv[1][i])):
            if k in seen or (abs(g) > 1e-12 and abs(c[i] / g) < args.min_share):
                continue
            seen.add(k)
            print()
            head(k, prefix + '  > ')
            rec(k, depth - 1, prefix + '    ')

    head(root, '')
    if M.desc.get(root):
        print('   ', M.desc[root])
    if root in M.eqs:
        print('    equation: %s' % M.eqs[root].text[:300])
    rec(root, args.depth, '')
    if args.by_period:
        d = M.dyn(root, periods=True)
        if d and 'error' not in d:
            labs = M.labels[M.ws:M.we + 1]
            cols = sorted(d['contrib'], key=lambda k: -abs(d['end'][k]))[:7]
            hdr = cols + (['add-factor'] if d['af'] else [])
            print('\nBy period, contribution to the gap of %s:' % root)
            print('%-9s %11s ' % ('period', 'gap') + ' '.join('%11s' % h[:11] for h in hdr))
            for i, lab in enumerate(labs):
                vals = [d['contrib'][k][i] for k in cols] + ([d['af'][i]] if d['af'] else [])
                print('%-9s %11s ' % (lab, fmt(d['gap'][i])) + ' '.join('%11s' % fmt(v) for v in vals))
    print('\nHow to read: each line is how much the gap (B-A) of the variable would be if ONLY that input '
          'moved from A to B, with the equation re-solved through time (so effects via its own lags are '
          'credited to the input that caused them).')


def cmd_path(M, args):
    x, y = _norm(M, args.x), _norm(M, args.y)
    ps = M.ranked_paths(x, y, k=args.k)
    if not ps:
        back = M.paths(y, x, k=1)
        print('%s does not affect %s through the loaded equations.' % (x, y))
        if back:
            print('(The reverse is true: %s -> %s)' % (y, ' -> '.join(back[0][1:])))
        if not M.children.get(x):
            print('%s is an end point: no other equation uses it.' % x)
        return
    _print_impact(M, x, y, None)
    print('\nChannels (ordered by how much of the change they carry):')
    for i, p in enumerate(ps, 1):
        print('Path %d:' % i)
        prod = 1.0
        ok = True
        for j, n in enumerate(p):
            s = M.stats.get(n)
            gap = ('%s (%s)' % (fmt(s['dend']), fpct(s['pend']))) if s and s['differs'] else 'no difference'
            if j == 0:
                print('   %s   gap at end %s' % (n, gap))
                continue
            sh = M.share(n, p[j - 1])
            if sh is None:
                ok = False
                shtxt = 'share n/a'
            else:
                prod *= sh
                shtxt = '%.0f%% of its gap comes from %s' % (100 * sh, p[j - 1])
            print('-> %s   gap at end %s   [%s]' % (n, gap, shtxt))
        if ok:
            print('   approx. share of %s\'s gap along this path: %.0f%%' % (y, 100 * prod))
        print()


def _print_impact(M, x, y, shock):
    eff, info = M.impact(x, y, shock)
    rng = range(M.ws, M.we + 1)
    labs = [M.labels[t] for t in rng]
    ey = eff.get(y)
    what = 'its change from A to B' if shock is None else ('a %s%s shock over the window' % (
        ('+' if shock[1] >= 0 else '') + ('%g' % shock[1]), '%' if shock[0] == 'pct' else ' unit'))
    if ey is None:
        print('%s cannot be re-solved (no evaluable equation), so the impact cannot be simulated.' % y)
        return
    A, Bv = M.series(y, 'a'), M.series(y, 'b')
    gap = [Bv[t] - A[t] for t in rng]
    iend = max(i for i, g in enumerate(ey) if g == g) if any(g == g for g in ey) else len(ey) - 1
    print('Impact of %s (%s) on %s, re-solving %d equations in between:' % (x, what, y, info['nodes']))
    pe = ey[iend] / abs(A[M.ws + iend]) * 100 if A[M.ws + iend] else NAN
    line = '  at %s: %s changes by %s (%s)' % (labs[iend], y, fmt(ey[iend]), fpct(pe))
    if shock is None and abs(gap[iend]) > 1e-12:
        line += ' = %.0f%% of its total gap %s' % (100 * ey[iend] / gap[iend], fmt(gap[iend]))
    print(line)
    if info['blocked']:
        print('  note: %d variables in between have no evaluable equation and were held fixed (e.g. %s)'
              % (len(info['blocked']), ', '.join(info['blocked'][:5])))
    return eff, gap, labs


def cmd_impact(M, args):
    x = _norm(M, args.x)
    shock = parse_shock(args.shock)
    if args.y:
        y = _norm(M, args.y)
        if y != x and y not in M.descendants(x):
            print('%s does not affect %s through the loaded equations (no chain of equations links them).' % (x, y))
            return
        r = _print_impact(M, x, y, shock)
        if r and args.by_period:
            eff, gap, labs = r
            print('\n%-9s %12s %12s %8s' % ('period', 'impact', 'total gap', 'share'))
            for i, lab in enumerate(labs):
                sh = eff[y][i] / gap[i] if shock is None and abs(gap[i]) > 1e-12 else None
                print('%-9s %12s %12s %8s' % (lab, fmt(eff[y][i]), fmt(gap[i]), fsh(sh)))
        return
    eff, info = M.impact(x, None, shock)
    rows = []
    for n, e in eff.items():
        if n == x:
            continue
        v = e[-1] if e[-1] == e[-1] else NAN
        a = M.series(n, 'a')[M.we]
        p = v / abs(a) * 100 if a else NAN
        g = M.stats.get(n, {}).get('dend')
        rows.append((n, v, p, (v / g) if (shock is None and g) else None))
    rows.sort(key=lambda r: -abs(r[2]) if r[2] == r[2] else 0)
    print('Variables affected by %s (%d re-solved). Effect at %s:' % (x, info['nodes'], M.labels[M.we]))
    print('%-24s %12s %9s %8s' % ('variable', 'effect', '%', 'of gap'))
    for n, v, p, s in rows[:args.n]:
        print('%-24s %12s %9s %8s' % (n, fmt(v), fpct(p), fsh(s) if s is not None else ''))


def check_equations(M):
    """One row per equation that cannot be used (or only partly), with the reason."""
    rows = []
    allnames = set(M.a) | set(M.b)
    for n in sorted(M.eqs):
        eq = M.eqs[n]
        differs = 'yes' if n in M.differs else ('no' if n in allnames else 'not in data')
        cat = detail = ''
        if n not in allnames:
            cat, detail = 'variable not in the data', 'the equation defines %s, but no series by that name was loaded' % n
        elif not eq.evaluable:
            w = eq.why or ''
            if 'specification' in w:
                cat, detail = 'specification only', 'estimation spec (e.g. "ls y c x") with no coefficients'
            elif 'coefficient' in w:
                cat, detail = 'coefficients not substituted', 'contains C(i); export the equation with substituted coefficients'
            else:
                cat, detail = 'could not parse', w.replace('could not parse: ', '')
        else:
            d = M.dyn(n)
            err = d.get('error') if d else 'no result'
            if err:
                m = re.search(r'unsupported function (\S+)', err)
                m2 = re.search(r'series (\S+) is not in the data', err)
                if m:
                    cat, detail = 'unsupported function', m.group(1)
                elif m2:
                    cat, detail = 'input series missing from the data', m2.group(1)
                elif 'missing values' in err:
                    cat, detail = 'equation gives missing values', err
                else:
                    cat, detail = 'evaluation error', err
            else:
                # partly missing: the equation gives NA in quarters where the data has a value
                try:
                    ra = M.resid(n, 'a')
                    a = M.series(n, 'a')
                    bad = [M.labels[M.ws + i] for i, v in enumerate(ra)
                           if v != v and a[M.ws + i] == a[M.ws + i]]
                except EvalError:
                    bad = []
                if bad:
                    cat = 'missing values in some quarters (warning; still used)'
                    detail = '%d quarter(s): %s%s' % (len(bad), ', '.join(bad[:4]), ' ...' if len(bad) > 4 else '')
        if cat:
            rows.append([n, M.desc.get(n, ''), differs, cat, detail, eq.source, eq.text[:500]])
    return rows


def write_check(M, path, quiet=False):
    rows = check_equations(M)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['variable', 'description', 'differs_between_scenarios', 'reason', 'detail', 'file', 'equation'])
        w.writerows(rows)
    if not quiet:
        cnt = Counter(r[3] for r in rows)
        print('%d of %d equations cannot be used fully; list written to %s' % (len(rows), len(M.eqs), path),
              file=sys.stderr)
        for k, v in cnt.most_common():
            print('   %6d  %s' % (v, k), file=sys.stderr)
        fn = Counter(r[4] for r in rows if r[3] == 'unsupported function')
        if fn:
            print('   unsupported functions: ' + ', '.join('%s (%d)' % kv for kv in fn.most_common(12)), file=sys.stderr)
        miss = Counter(r[4] for r in rows if r[3] == 'input series missing from the data')
        if miss:
            print('   missing inputs, most used: ' + ', '.join('%s (%d)' % kv for kv in miss.most_common(8)), file=sys.stderr)
    return rows


def cmd_importance(M, args):
    keys = pick_key_vars(M, getattr(args, 'key_list', None), 8)
    lg = lambda m: print(m, file=sys.stderr)
    if args.method == 'exact':
        res, share, targets = compute_importance(M, args.max_shocks, keys, log=lg)
    else:
        res, share, targets = compute_importance_fast(M, keys, log=lg)
    full = (res, share, targets)
    if args.geo or args.match:
        sel = targets
        if args.geo:
            gs = {g.strip().upper() for g in args.geo.split(',')}
            sel = [y for y in sel if (M.geo(y) or '(NONE)') in gs]
        if args.match:
            rx = re.compile(args.match, re.I)
            sel = [y for y in sel if rx.search(y) or rx.search(M.desc.get(y, ''))]
        if not sel:
            raise SystemExit('no variable that moved matches --geo/--match (geographies: %s)'
                             % ', '.join(sorted(M.geo_codes)[:20]))
        targets = sel
        res = aggregate_importance(M, share, targets, keys)
        share = {y: share[y] for y in targets if y in share}
    col = {'all': 'w_all', 'concept': 'w_concept', 'key': 'w_key'}[args.weight]
    rows = sorted(res.values(), key=lambda r: -abs(r[col] or 0))
    print('Importance of each scenario shock = average share of the gaps it explains (whole window, %s weights).'
          % args.weight)
    print('%d variables with an equation moved%s; %d shocks.\n' % (
        len(targets), (' in ' + ', '.join(x for x in (args.geo, args.match) if x)) if (args.geo or args.match) else '', len(rows)))
    print('%-28s %-12s %8s %8s %8s %7s %7s  %s' % ('shock', 'type', 'all vars', 'concept', 'key vars', '>=10%', '>=50%', 'explains most of'))
    for r in rows[:args.n]:
        lab = M.af_label(r['series']) if r['kind'] == 'addfactor' else r['series']
        top = ', '.join('%s %.0f%%' % (y, 100 * v) for y, v in r['top'][:3])
        print('%-28s %-12s %7.1f%% %7.1f%% %8s %7d %7d  %s' % (lab[:28], r['type'], 100 * r['w_all'], 100 * r['w_concept'],
              ('%7.1f%%' % (100 * r['w_key'])) if r['w_key'] is not None else '-', r['n10'], r['n50'], top))
    if args.csv:
        with open(args.csv, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f)
            w.writerow(['shock', 'type', 'description', 'importance_all_vars', 'importance_by_concept',
                        'importance_key_vars', 'vars_explained_10pct', 'vars_explained_50pct', 'variables_it_feeds'])
            for r in rows:
                w.writerow([r['id'], r['type'], r['desc'], _r(r['w_all'], 4), _r(r['w_concept'], 4), _r(r['w_key'], 4),
                            r['n10'], r['n50'], r['reach']])
        base = os.path.splitext(args.csv)[0] + '_matrix.csv'
        with open(base, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f)
            w.writerow(['variable', 'description', 'shock', 'share_of_gap'])
            for y in sorted(share):
                for sid, v in sorted(share[y].items(), key=lambda kv: -abs(kv[1])):
                    if abs(v) >= 0.01:
                        w.writerow([y, M.desc.get(y, ''), sid, _r(v, 4)])
        print('\nRanking written to %s; every variable x shock share (>=1%%) to %s' % (args.csv, base))
    if args.out:
        if not args.out.lower().endswith(('.html', '.htm')):
            args.out += '.html'
        for k, v in (('embed', 'auto'), ('max_decomp', 50000), ('context', 8), ('focus', None), ('summary_n', 8),
                     ('importance', True), ('influence_n', 20)):
            if not hasattr(args, k):
                setattr(args, k, v)
        args.importance_method = args.method
        args.imp_pre = full
        args.open_view = {'tab': 'imp', 'geo': (args.geo or '').split(',')[0].strip().upper() or None,
                          'weight': args.weight, 'match': args.match}
        print('', file=sys.stderr)
        build_report(M, args, getattr(args, 'eqstats', None))


_MASK_KEEP = {'AR', 'MA', 'SAR', 'SMA', 'UNCOND', 'INCOND', 'ESTSMPL', 'ESTSAMPLE', 'LOG', 'DLOG', 'D', 'EXP',
              'ABS', 'SQR', 'SQRT', 'AND', 'OR', 'NOT', 'NA'}


def _mask(t):
    """Equation text with every name and number hidden, so it can be shared: names -> x, numbers -> 9.
    EViews functions (@MOVAV ...) and keywords (AR, UNCOND, DLOG ...) stay visible, as do operators,
    brackets and spaces; any unusual character is shown as <U+XXXX name>."""
    import unicodedata

    def word(m):
        w = m.group(0)
        if w.startswith('@') or w.upper() in _MASK_KEEP:
            return w
        return '9' if w[0].isdigit() else 'x'
    t = re.sub(r'@?[A-Za-z0-9_]+(?:\.[0-9]+)?', word, t)
    return ''.join(c if (c.isascii() and (c.isprintable() or c == ' ')) else
                   '<U+%04X %s>' % (ord(c), unicodedata.name(c, 'unnamed').lower()) for c in t)


def _mask_why(w):
    """Mask only the quoted equation snippets in an error message."""
    return re.sub(r'(["\'])(.*?)\1', lambda m: m.group(1) + _mask(m.group(2)) + m.group(1), w)


def cmd_mask(args):
    """Print equations containing [ or AR/MA terms in a masked form that is safe to share."""
    files = []
    for q in args.eq:
        if os.path.isdir(q):
            for root, _, fs in os.walk(q):
                files += [os.path.join(root, f) for f in sorted(fs) if os.path.splitext(f)[1].lower() in EQ_EXT]
        else:
            files.append(q)
    keep = re.compile(args.match, re.I) if args.match else None
    shown = 0
    for f in files:
        with open(f, 'r', encoding='utf-8-sig', errors='replace') as fh:
            raw = fh.read()
        for ln in _logical_lines(raw):
            if keep and not keep.search(ln):
                continue
            if not keep and not re.search(r'[\[\uff3b]|\b(ar|ma)\s*\(', ln, re.I):
                continue
            r = parse_equation_line(ln, os.path.basename(f))
            ok = r is not None and not isinstance(r, tuple) and r.evaluable
            if args.failed and ok:
                continue
            why = '' if r is None or isinstance(r, tuple) else r.why
            print('%-4s %s' % ('ok' if ok else 'FAIL', _mask(ln)))
            if why:
                print('     reason: %s' % _mask_why(why))
            shown += 1
            if shown >= args.n:
                print('... (first %d shown; -n to change)' % args.n)
                return
    if not shown:
        print('No matching equation lines found.')


def cmd_layers(M, args):
    """Print how many layers of variables each variable affects (longest chain of equations)."""
    L = args.layers
    names, alab = args.layers_area
    keep = [n for n in L if n in M.parents or n in M.children]
    if args.match:
        rx = re.compile(args.match, re.I)
        keep = [n for n in keep if rx.search(n)]
    k = 1 if (args.by_area and names) else 0
    hist = Counter(L[n][k] for n in keep if L[n][k] is not None)
    print('Layers%s: longest chain of equations from each variable downstream; a loop counts as one layer.'
          % ((' into %s' % alab) if k else ' (overall)'))
    print('  layers  variables')
    for d in sorted(hist):
        print('  %6d  %d' % (d, hist[d]))
    if k:
        print('  (cannot reach %s: %d)' % (alab, sum(1 for n in keep if L[n][1] is None)))
    top = sorted((n for n in keep if L[n][k] is not None), key=lambda n: (-L[n][k], n))[:args.n]
    print()
    for n in top:
        ch = layer_chain(L, n, area=bool(k))
        print('%-28s %3d  %s%s' % (n, L[n][k], ' -> '.join(ch[1:8]), ' -> ...' if len(ch) > 8 else ''))
    if names:
        aset = set(names)
        feeds = defaultdict(list)
        for t in names:
            for p in M.parents.get(t, ()):
                if p not in aset:
                    feeds[p].append(t)
        print('\n%d variables outside %s feed it directly (%d of its variables use at least one):'
              % (len(feeds), alab, len({t for ts in feeds.values() for t in ts})))
        for p in sorted(feeds, key=lambda p: (-len(feeds[p]), p))[:args.n]:
            ts = sorted(feeds[p])
            print('  %-26s -> %3d  %s%s' % (p, len(ts), ', '.join(ts[:6]), ', ...' if len(ts) > 6 else ''))


def cmd_check(M, args):
    write_check(M, args.out)


def cmd_influence(M, args):
    shock = parse_shock(args.shock)
    reach = M.reach_counts()
    if args.x:
        x = _norm(M, args.x)
        r = M.influence(x, shock, args.thr)
        what = 'its change from A to B' if shock is None else ('%s%s' % (args.shock, '' if args.shock.endswith('%') else ' unit(s)'))
        if r['mode'] == 'addfactor':
            what = 'its add-factor change from A to B (equation of %s re-solved)' % r['target']
        if r['mode'] == 'addfactor':
            print('%s: add-factor of %s; linked to %d variables in total (including %s).' % (
                x, r['target'], r['reach'], r['target']))
        else:
            print('%s: used directly by %d equations; linked to %d variables in total.' % (
                x, len(M.children.get(x, [])), reach.get(x, 0)))
        print('Changing %s by %s, re-solving %d equations:' % (x, what, r['resolved']))
        print('  %d variables move by >= %g%% of their level at some point: %d by >= 1%%, %d by 0.1-1%%, %d by %g-0.1%%'
              % (r['moved'], args.thr, r['ge1'], r['b01'], r['b001'], args.thr))
        if r['blocked']:
            print('  (%d linked variables have no evaluable equation and were held fixed)' % r['blocked'])
        print('\n%-26s %12s %9s  %s' % ('variable', 'largest effect', '%', 'when'))
        for n, v, p, lab in r['top'][:args.n]:
            print('%-26s %12s %9s  %s' % (n, fmt(v), fpct(math.copysign(p, v)) if p is not None else 'n/a', lab))
        return
    shocks = M.shock_list()[:args.n]
    print('Influence of each scenario shock (its own A->B change only), largest effect over %s..%s:\n'
          % (M.labels[M.ws], M.labels[M.we]))
    print('%-24s %4s %8s %8s %6s %7s %8s  %s' % ('shock', 'type', 'linked', 'moved', '>=1%', '0.1-1%', '<0.1%', 'most affected'))
    for x in shocks:
        try:
            r = M.influence(x, None, args.thr)
        except EvalError as e:
            print('%-24s  (%s)' % (x, e))
            continue
        top = ', '.join('%s %s' % (n, fpct(math.copysign(p, v)) if p is not None else '') for n, v, p, _ in r['top'][:3])
        lab = x + (' (add-factor)' if r['mode'] == 'addfactor' and M.kind(x) != 'F' else '')
        print('%-24s %4s %8d %8d %6d %7d %8d  %s' % (lab, M.kind(x), r['reach'], r['moved'], r['ge1'], r['b01'], r['b001'], top))


def cmd_attribute(M, args):
    y = _norm(M, args.y)
    res, joint, gap = M.attribute(y)
    i = max(j for j, g in enumerate(gap) if g == g)
    g = gap[i]
    print('%s: gap B-A at %s = %s. Split across the scenario shocks upstream of it:' % (y, M.labels[M.ws + i], fmt(g)))
    res.sort(key=lambda r: -abs(r[2][i]) if r[2][i] == r[2][i] else 0)
    shown = 0
    for kind, s, e in res:
        if abs(e[i]) < 1e-9 * max(1.0, abs(g)):
            continue
        lab = s if kind == 'data' else '%s%s' % (M.af_label(s), (' (%s)' % M.af[s]) if s in M.af else '')
        print('  %s %-12s %s  [%s]' % (fsh(e[i] / g if abs(g) > 1e-12 else None), fmt(e[i]), lab,
                                       ('set exogenously by the modeller' if s in M.exog else KIND_TEXT.get(M.kind(s), ''))
                                       if kind == 'data' else ('equation differs' if s in M.eq_status else 'add-factor / residual')))
        shown += 1
        if shown >= args.n:
            break
    other = g - sum(e[i] for _, _, e in res if e[i] == e[i])
    print('  %s %-12s (interaction between shocks)' % (fsh(other / g if abs(g) > 1e-12 else None), fmt(other)))
    if abs(joint[i] - g) > 1e-6 * max(1.0, abs(g)):
        print('  note: re-solving with all shocks gives %s vs actual gap %s (model fit difference)' % (fmt(joint[i]), fmt(g)))


# --------------------------------------------------------------------------
# HTML report (static file) and local server mode
# --------------------------------------------------------------------------
def _vals_payload(M, n, es):
    a, b = M.series(n, 'a'), M.series(n, 'b')
    if a is None:
        return None
    return [[_r(v) for v in a[es:M.we + 1]], [_r(v) for v in b[es:M.we + 1]]]


def _dec_payload(M, n):
    d = M.dyn(n, periods=True)
    if not d:
        return None
    if 'error' in d:
        return {'err': d['error']}
    cols = [[k, [_r(v, 5) for v in d['contrib'][k]]]
            for k in sorted(d['contrib'], key=lambda k: -abs(d['end'][k]))]
    return {'gap': [_r(v, 6) for v in d['gap']], 'cols': cols,
            'af': [_r(v, 5) for v in d['af']] if d['af'] else None,
            'afn': d['af_name'], 'int': [_r(v, 5) for v in d['inter']], 'iend': d['iend'],
            'afs': _r(d['af_share'], 3)}


def _af_rows(M):
    rows = []
    for n in M.ranked():
        if n not in M.eqs or not M.eqs[n].evaluable:
            continue
        d = M.dyn(n)
        if not d or 'error' in d:
            continue
        afn = d['af_name']
        afc = None
        st = M.stats.get(afn) if afn else None
        if st and st['differs']:
            afc = [_r(st['dmax']), M.labels[st['tmax']]]
        rows.append([n, M.desc.get(n, ''), afn or '', afc, _r(d['af_share'], 3),
                     [_r(v, 5) for v in d['pgap']], [_r(v, 5) for v in d['pf']],
                     [_r(v, 5) for v in d['paf']] if d['paf'] else None])
    return rows


def _layers_data(args, nodes):
    """Layers per report node (aligned with D.nodes), plus the next step of each longest chain."""
    L = getattr(args, 'layers', None)
    if not L:
        return None
    names, label = getattr(args, 'layers_area', (None, None))
    z = (0, None, None, None, 1)
    nx, nxa = {}, {}
    for n in nodes:
        for k, dct in ((2, nx), (3, nxa)):
            cur, guard = n, 0
            while cur not in dct and guard < 80:
                nxt = L.get(cur, z)[k]
                if not nxt:
                    break
                dct[cur] = nxt
                cur, guard = nxt, guard + 1
    extra = sorted({v for v in list(nx.values()) + list(nxa.values())} - set(nodes))
    return {'d': [L.get(n, z)[0] for n in nodes], 'da': [L.get(n, z)[1] for n in nodes],
            'lp': [L.get(n, z)[4] for n in nodes], 'nx': nx, 'nxa': nxa,
            'xd': {n: [L[n][0], L[n][1]] for n in extra if n in L},
            'area': names, 'label': label,
            'max': max((v[0] for v in L.values()), default=0),
            'nall': len(L), 'pruned': bool(getattr(args, 'target', None))}


def build_data(M, args, eqstats, serve=False):
    t0 = time.time()
    ranked = M.ranked()
    labels = M.labels
    es = max(0, M.ws - args.context)
    we = M.we
    per = labels[es:we + 1]

    def log(msg):
        if M.verbose:
            print(msg, file=sys.stderr)

    focus = _expand_focus(M, args.focus)
    todo = [n for n in ranked if n in M.eqs and M.eqs[n].evaluable][:args.max_decomp]
    for i, n in enumerate(todo):
        M.dyn(n)
        if M.verbose and i and i % 2000 == 0:
            log('  analysed %d / %d equations ...' % (i, len(todo)))
    log('  driver contributions: %d equations in %.1fs' % (len(todo), time.time() - t0))

    nodes = sorted(set(M.a) | set(M.b) | set(M.eqs) | set(M.children))
    nidx = {n: i for i, n in enumerate(nodes)}
    par = [[nidx[p] for p in M.parents.get(n, []) if p in nidx] for n in nodes]

    rows = []
    for n in ranked:
        s = M.stats[n]
        rows.append([n, M.desc.get(n, ''), M.geo(n), _r(s['score'], 4), _r(s['dmax']), _r(s['pmax'], 4),
                     _r(s['dend']), _r(s['pend'], 4), labels[s['first']] if s['first'] is not None else '',
                     M.kind(n), M.change_class(n)])
    entry = [[n, M.kind(n)] for n in M.entry_points()]

    emb, seen = [], set()

    def add(n):
        if n not in seen and (n in M.a or n in M.b):
            seen.add(n)
            emb.append(n)
    emb_n = str(args.embed).lower()
    if emb_n == 'all' or (emb_n == 'auto' and len(ranked) <= 5000):
        args.embed = 10 ** 9
    elif emb_n == 'auto':
        args.embed = 1500
    else:
        args.embed = int(emb_n)
    for n in focus:
        add(n)
    for n, _ in entry[:args.embed]:
        add(n)
    for n in ranked[:args.embed]:
        add(n)
    for n in list(emb)[:args.embed]:
        d = M.dyn(n)
        if d and 'error' not in d:
            for k, _ in sorted(d['end'].items(), key=lambda kv: -abs(kv[1]))[:8]:
                add(k)
            if d['af_name']:
                add(d['af_name'])
    vals = {n: _vals_payload(M, n, es) for n in emb}
    dec = {}
    for n in emb:
        p = _dec_payload(M, n)
        if p:
            dec[n] = p
    esh, afsh, eshp, afshp = {}, {}, {}, {}
    pg = {}
    for n in M.differs:  # gap in % of A, every quarter of the window
        a, b = M.series(n, 'a'), M.series(n, 'b')
        pg[n] = [_r((b[t] - a[t]) / abs(a[t]) * 100, 3) if (a[t] == a[t] and b[t] == b[t] and a[t]) else None
                 for t in range(M.ws, M.we + 1)]
    for n, d in M._dd.items():
        if not d or 'error' in d or n not in nidx:
            continue
        esh[str(nidx[n])] = [[nidx[k], _r(v, 3)] for k, v in d['shares'].items() if k in nidx and v is not None]
        ps = d.get('pshare')
        if ps:
            eshp[str(nidx[n])] = [[nidx[k], v] for k, v in ps.items() if k in nidx and any(x is not None for x in v)]
        if d['af_changed'] and d['af_share_end'] is not None:
            afsh[str(nidx[n])] = _r(d['af_share_end'], 3)
        if d.get('pafshare'):
            afshp[str(nidx[n])] = d['pafshare']
    reach = M.reach_counts() if M.has_eqs else {}
    infl = {}
    if M.has_eqs and args.influence_n > 0:
        t1 = time.time()
        for x in M.shock_list()[:args.influence_n]:
            try:
                r = M.influence(x, None)
            except EvalError:
                continue
            r['top'] = [[n, _r(v), _r(p, 4), lab] for n, v, p, lab in r['top']]
            infl[x] = r
        log('  influence of %d scenario shocks in %.1fs' % (len(infl), time.time() - t1))
    summ = build_summary(M, args, log) if M.has_eqs and getattr(args, 'summary_n', 8) > 0 else None
    imp_payload = None
    if M.has_eqs and getattr(args, 'importance', True):
        if getattr(args, 'imp_pre', None):
            res, share, targets = args.imp_pre
        elif getattr(args, 'importance_method', 'fast') == 'exact':
            res, share, targets = compute_importance(M, getattr(args, 'max_shocks', 0), summ['keys'] if summ else None, log)
        else:
            res, share, targets = compute_importance_fast(M, summ['keys'] if summ else None, log)
        imp_payload = {
            'n_targets': len(targets),
            'shocks': [[r['id'], r['type'], r['desc'], _r(r['w_all'], 4), _r(r['w_concept'], 4),
                        _r(r['w_key'], 4) if r['w_key'] is not None else None, r['n10'], r['n50'], r['reach'],
                        [[y, _r(v, 3)] for y, v in r['top']]] for r in res.values()],
            'by_var': {y: [[sid, _r(v, 3)] for sid, v in sorted(d.items(), key=lambda kv: -abs(kv[1]))[:6]]
                       for y, d in share.items()},
            'geo': {}}
        if len(M.geo_codes) >= 2:
            groups = defaultdict(list)
            for y in targets:
                groups[M.geo(y) or '(none)'].append(y)
            for g, ys in groups.items():
                rg = aggregate_importance(M, share, ys, summ['keys'] if summ else None)
                top = sorted(rg.values(), key=lambda r: -abs(r['w_concept']))[:100]
                imp_payload['geo'][g] = {'n': len(ys), 'rows': [[r['id'], _r(r['w_all'], 4), _r(r['w_concept'], 4), r['n10'], r['n50'],
                                                                 [[y, _r(v, 3)] for y, v in r['top'][:3]]] for r in top]}
    eqt = {n: M.eqs[n].text[:700] for n in nodes if n in M.eqs and (n in seen or n in M.differs)}
    eqdiff = []
    if M.eqs_b_all is not None:
        import difflib
        for n, st in sorted(M.eq_status.items()):
            ta = M.eqs[n].text if (st != 'only_b' and n in M.eqs) else ''
            if st == 'only_b':
                ta = ''
            tb = M.eqs_b_all[n].text if n in M.eqs_b_all else ''
            if st == 'only_a' and n in M.exog:
                ta = M.exog[n].text
            def split(t):   # main formula / error-term options (not compared, shown in grey)
                main, brl = _drop_error_terms(t)
                br = ' '.join(brl)
                return re.findall(r'\s+|[\w.$@%#]+|\S', main), br
            ra, bra = split(ta)
            rb, brb = split(tb)
            ka = [x.upper() for x in ra if not x.isspace()]
            kb = [x.upper() for x in rb if not x.isspace()]
            ia = [i for i, x in enumerate(ra) if not x.isspace()]
            ib = [i for i, x in enumerate(rb) if not x.isspace()]
            fa, fb = [0] * len(ra), [0] * len(rb)
            sm = difflib.SequenceMatcher(a=ka, b=kb, autojunk=False)
            for op, i1, i2, j1, j2 in sm.get_opcodes():
                if op != 'equal':
                    for k in range(i1, i2):
                        fa[ia[k]] = 1
                    for k in range(j1, j2):
                        fb[ib[k]] = 1

            def pack(tokens, flags, br):
                out = []
                for t, f in zip(tokens, flags):
                    if out and out[-1][1] == f:
                        out[-1][0] += t
                    else:
                        out.append([t, f])
                if br:
                    out.append(['   ' + br, 2])
                return out
            da, db = pack(ra, fa, bra), pack(rb, fb, brb)
            stt = M.stats.get(n)
            eqdiff.append([n, st, da, db, _r(stt['pend'], 4) if stt and stt['differs'] else None, M.desc.get(n, '')])
    for n, e in M.exog.items():
        eqt[n] = e.text[:700]
    why = {n: M.eqs[n].why for n in eqt if n in M.eqs and M.eqs[n].why}

    geo_rows, stem_rows = [], []
    if len(M.geo_codes) >= 2:
        def grp_ok(n):  # add-factor series are counted through the variable they adjust
            return not (M.has_eqs and M.kind(n) == 'F')
        g_all, s_all = Counter(), defaultdict(set)
        g_cls, s_cls = defaultdict(lambda: defaultdict(list)), defaultdict(lambda: defaultdict(list))
        for n in M.common:
            if not grp_ok(n):
                continue
            g_all[M.geo(n)] += 1
            s_all[M.stem(n)].add(M.geo(n))
        for n in M.differs:
            if not grp_ok(n):
                continue
            c = M.change_class(n) or 'D'
            g_cls[M.geo(n)][c].append(n)
            s_cls[M.stem(n)][c].append(n)

        def med(xs):
            xs = sorted(xs)
            return xs[len(xs) // 2] if xs else None

        def top(cl):
            ns = cl['D'] or cl['U'] or cl['M']
            return sorted(ns, key=lambda n: -M.stats[n]['score'])[:5]
        for g, cl in g_cls.items():
            geo_rows.append([g, len(cl['D']), len(cl['M']), len(cl['U']), g_all[g],
                             _r(med([M.stats[n]['score'] for n in cl['D']]), 3), top(cl)])
        geo_rows.sort(key=lambda r: (-r[1], -r[2]))
        for st, cl in s_cls.items():
            pm = [abs(M.stats[n]['pend']) for n in cl['D'] if M.stats[n]['pend'] == M.stats[n]['pend']]
            tp = top(cl)
            stem_rows.append([st, len(cl['D']), len(cl['M']), len(cl['U']), len(s_all[st]), _r(med(pm), 3),
                              _r(max([M.stats[n]['score'] for n in cl['D']] or [0]), 3), tp, M.desc.get(tp[0], '')])
        stem_rows.sort(key=lambda r: (-r[1], -r[2]))

    data = {
        'meta': {'a': M.A.label, 'b': M.B.label, 'ws': labels[M.ws], 'we': labels[M.we],
                 'first': labels[0], 'last': labels[-1], 'ppy': M.ppy,
                 'n_common': len(M.common), 'n_diff': len(M.differs),
                 'n_only_a': len(M.only_a), 'n_only_b': len(M.only_b),
                 'eqs': dict(eqstats) if eqstats else None, 'has_eqs': M.has_eqs,
                 'generated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
                 'open': getattr(args, 'open_view', None),
                 'adj_thr': M.adj_threshold, 'embedded': len(emb), 'serve': serve,
                 'af_suffix': ','.join(M.af_suffixes), 'n_af': len(M.af_series), 'geo_sep': M.geo_sep,
                 'exog': sorted(M.exog)},
        'per': per, 'wsi': M.ws - es,
        'rows': rows, 'entry': entry, 'nodes': nodes, 'par': par,
        'layers': _layers_data(args, nodes),
        'vals': vals, 'dec': dec, 'esh': esh, 'afsh': afsh, 'eshp': eshp, 'afshp': afshp, 'pg': pg,
        'reach': [reach.get(n, 0) for n in nodes], 'infl': infl,
        'summary': summ,
        'imp': imp_payload, 'eqt': eqt, 'why': why, 'eqdiff': eqdiff,
        'eqstat': dict(M.eq_status),
        'afmap': {k: v for k, v in M.af.items()},
        'afrows': _af_rows(M) if M.has_eqs else [],
        'geo': geo_rows, 'stem': stem_rows,
        'only_a': M.only_a[:500], 'only_b': M.only_b[:500],
        'desc': {n: M.desc[n] for n in nodes if n in M.desc and n not in M.differs},
    }
    log('  report data built in %.1fs (%d series with full detail)' % (time.time() - t0, len(emb)))
    return data, es


def _pct_word(p):
    if p is None or p != p:
        return 'n/a'
    return '%+.2f%%' % p


def compute_importance(M, max_shocks=None, keys=None, log=lambda m: None):
    """Relative importance of every scenario shock.

    Shocks = changed exogenous variables (no equation, or set by the modeller) and changed add-factors.
    For each shock alone: re-solve every equation it feeds, with everything else at scenario A, and
    take its effect e_t on each variable. Its share of variable y's gap g_t over the whole window is
        share = sum_t e_t * g_t / sum_t g_t^2
    (the part of y's gap path that the shock reproduces). Shares of all shocks add up to ~1 for each
    variable (exactly 1 if the model is linear). Importance = weighted average of the shares over the
    variables that moved."""
    t0 = time.time()
    rng = range(M.ws, M.we + 1)
    shocks = []   # (id, kind, series)
    seen = set()
    for n in sorted(M.differs_any):
        if n in M.eqs and M.eqs[n].evaluable:
            continue
        if n in M.af_series and M.af_target(n):
            continue                      # counted through its equation's add-factor below
        if n in M.children or n in M.exog:
            shocks.append((n, 'data', n))
            seen.add(n)
    for n in sorted(M.eqs):
        if M.eqs[n].evaluable and n in M.differs_any:
            d = M.dyn(n)
            if d and 'error' not in d and d['af_changed']:
                shocks.append(('AF:' + n, 'addfactor', n))
    # cheap ones first, so a cap keeps the broad ones... rank by size of what they feed
    reach = M.reach_counts()
    if max_shocks and len(shocks) > max_shocks:
        shocks.sort(key=lambda x: -(reach.get(x[2], 0) + 1) * (M.stats.get(x[2], {}).get('score', 0) or 1))
        shocks = shocks[:max_shocks]
    targets = [n for n in M.differs if n in M.eqs and M.eqs[n].evaluable]
    tset = set(targets)
    gap = {}
    for y in targets:
        a, b = M.series(y, 'a'), M.series(y, 'b')
        gap[y] = [(b[t] - a[t]) if (a[t] == a[t] and b[t] == b[t]) else 0.0 for t in rng]
    g2 = {y: sum(v * v for v in g) for y, g in gap.items()}
    share = defaultdict(dict)    # y -> {shock: share}
    for k, (sid, kind, sname) in enumerate(shocks):
        if kind == 'data':
            nodes = M.descendants(sname)
            ev = [n for n in nodes if n in M.eqs and M.eqs[n].evaluable]
            try:
                sim = M.simulate(ev, override={sname: M.series(sname, 'b')})
            except EvalError:
                continue
        else:
            nodes = M.descendants(sname) | {sname}
            ev = [n for n in nodes if n in M.eqs and M.eqs[n].evaluable]
            try:
                sim = M.simulate(ev, resid_b={sname})
            except EvalError:
                continue
        for y in ev:
            if y not in tset or not g2[y]:
                continue
            a = M.series(y, 'a')
            s_ = sim[y]
            num = 0.0
            for i, t in enumerate(rng):
                e = s_[t] - a[t]
                if e == e:
                    num += e * gap[y][i]
            sh = num / g2[y]
            if abs(sh) >= 1e-4:
                share[y][sid] = sh
        if (k + 1) % 50 == 0:
            log('  importance: %d / %d shocks (%.0fs)' % (k + 1, len(shocks), time.time() - t0))
    # weights
    conc = Counter(M.stem(y) for y in targets)
    keyset = set(keys or ())
    res = {}
    for sid, kind, sname in shocks:
        sh = {y: d[sid] for y, d in share.items() if sid in d}
        ys = [y for y in targets]
        eq_w = sum(sh.get(y, 0.0) for y in ys) / max(1, len(ys))
        cw = sum(sh.get(y, 0.0) / conc[M.stem(y)] for y in ys) / max(1, len(conc))
        kw = (sum(sh.get(y, 0.0) for y in keyset if y in tset) / max(1, len(keyset & tset))) if keyset else None
        res[sid] = {'id': sid, 'kind': kind, 'series': sname, 'type': 'set by modeller' if sname in M.exog else
                    ('add-factor' if kind == 'addfactor' else 'exogenous'),
                    'desc': M.desc.get(sname, ''), 'reach': len(M.descendants(sname)) + (1 if kind == 'addfactor' else 0),
                    'w_all': eq_w, 'w_concept': cw, 'w_key': kw,
                    'n10': sum(1 for v in sh.values() if abs(v) >= 0.10),
                    'n50': sum(1 for v in sh.values() if abs(v) >= 0.50),
                    'top': sorted(sh.items(), key=lambda kv: -abs(kv[1]))[:25]}
    log('  importance of %d shocks over %d variables in %.1fs' % (len(shocks), len(targets), time.time() - t0))
    return res, share, targets


def _scc_order(nodes, preds):
    """Strongly connected components of the graph (edges pred -> node), in dependency order."""
    idx, low, onst, st, comps, cnt = {}, {}, set(), [], [], [0]
    succ = defaultdict(list)
    for v in nodes:
        for u in preds.get(v, ()):
            if u in preds or u in nodes:
                succ[u].append(v)
    nodeset = set(nodes)
    for root in nodes:
        if root in idx:
            continue
        work = [(root, iter([w for w in preds.get(root, ()) if w in nodeset]))]
        idx[root] = low[root] = cnt[0]; cnt[0] += 1; st.append(root); onst.add(root)
        while work:
            v, it = work[-1]
            adv = False
            for w in it:
                if w not in idx:
                    idx[w] = low[w] = cnt[0]; cnt[0] += 1; st.append(w); onst.add(w)
                    work.append((w, iter([z for z in preds.get(w, ()) if z in nodeset])))
                    adv = True
                    break
                elif w in onst:
                    low[v] = min(low[v], idx[w])
            if adv:
                continue
            work.pop()
            if work:
                low[work[-1][0]] = min(low[work[-1][0]], low[v])
            if low[v] == idx[v]:
                c = []
                while True:
                    w = st.pop(); onst.discard(w); c.append(w)
                    if w == v:
                        break
                comps.append(c)
    return comps   # predecessors' components come first


def _reach_of(M, n):
    c = M.__dict__.setdefault('_reach_cache', {})
    if n not in c:
        c[n] = M.reach_counts().get(n, 0) if n in M.reach_counts() else len(M.descendants(n))
    return c[n]


def aggregate_importance(M, share, targets, keys=None):
    """Importance of each shock averaged over the given variables only (e.g. one geography)."""
    tset = set(targets)
    conc = Counter(M.stem(y) for y in targets)
    keyset = set(keys or ()) & tset
    by = defaultdict(dict)
    for y in targets:
        for sid, v in share.get(y, {}).items():
            by[sid][y] = v
    res = {}
    for sid, sh in by.items():
        sname = sid[3:] if sid.startswith('AF:') else sid
        kind = 'addfactor' if sid.startswith('AF:') else 'data'
        res[sid] = {'id': sid, 'kind': kind, 'series': sname,
                    'type': 'set by modeller' if sname in M.exog else (
                        ('equation changed' if sname in M.eq_status else 'add-factor') if kind == 'addfactor' else
                        ('exogenous' if sname not in M.eqs else 'not evaluable')),
                    'desc': M.desc.get(sname, ''), 'reach': _reach_of(M, sname) + (1 if kind == 'addfactor' else 0),
                    'w_all': sum(sh.values()) / max(1, len(targets)),
                    'w_concept': sum(v / conc[M.stem(y)] for y, v in sh.items()) / max(1, len(conc)),
                    'w_key': (sum(sh.get(y, 0.0) for y in keyset) / len(keyset)) if keyset else None,
                    'n10': sum(1 for v in sh.values() if abs(v) >= 0.10),
                    'n50': sum(1 for v in sh.values() if abs(v) >= 0.50),
                    'top': sorted(sh.items(), key=lambda kv: -abs(kv[1]))[:25]}
    return res


def compute_importance_fast(M, keys=None, log=lambda m: None, prune=5e-4, keep=300):
    """Fast importance: trace every shock through the model using each equation's own split of its gap.

    For each variable y with an equation, its gap over the window is split into its inputs k and its
    add-factor (the same split as 'Where the gap comes from'):  w_k(y) = sum_t C_k,t g_t / sum_t g_t^2.
    Then the share of y's gap from a root shock s is
        S(y, s) = sum_k w_k(y) * S(k, s)   (+ w_AF(y) if s is y's own add-factor),
    with S(s, s) = 1 for an exogenous variable s. Loops in the model (GDP -> income -> consumption -> GDP)
    are solved by iterating. One pass over the model instead of one simulation per shock. It is exact for
    a linear model and close otherwise; 'importance --method exact' re-solves the model per shock."""
    t0 = time.time()
    rng = range(M.ws, M.we + 1)
    targets = [n for n in M.differs if n in M.eqs and M.eqs[n].evaluable]
    w, waf, preds = {}, {}, {}
    for k, y in enumerate(targets):
        d = M.dyn(y, periods=True)
        if not d or 'error' in d:
            continue
        g = [v if v == v else 0.0 for v in d['gap']]
        g2 = sum(v * v for v in g)
        if not g2:
            continue
        w[y] = {i: sum((c if c == c else 0.0) * gg for c, gg in zip(cs, g)) / g2 for i, cs in d['contrib'].items()}
        if d.get('af'):
            waf[y] = sum((c if c == c else 0.0) * gg for c, gg in zip(d['af'], g)) / g2
        preds[y] = list(w[y])
        if k and k % 5000 == 0:
            log('  importance (fast): equation splits %d / %d' % (k, len(targets)))
    S = {}
    roots = set()

    def src(k):
        if k in S:
            return S[k]
        if k in w:          # endogenous, filled in its own turn
            return {}
        roots.add(k)        # exogenous / set by modeller / not evaluable: a root shock
        S[k] = {k: 1.0}
        return S[k]

    def combine(y):
        acc = defaultdict(float)
        for k, wk in w[y].items():
            if abs(wk) < 1e-9:
                continue
            for sid, v in src(k).items():
                acc[sid] += wk * v
        if y in waf and abs(waf[y]) > 1e-9:
            acc['AF:' + y] += waf[y]
        out = {a: b for a, b in acc.items() if abs(b) >= prune}
        if len(out) > keep:
            out = dict(sorted(out.items(), key=lambda kv: -abs(kv[1]))[:keep])
        return out
    comps = _scc_order(list(w), preds)
    for c in comps:
        if len(c) == 1 and c[0] not in preds.get(c[0], ()):
            S[c[0]] = combine(c[0])
            continue
        for y in c:
            S[y] = {}
        for it in range(200):
            mx = 0.0
            for y in c:
                new = combine(y)
                old = S[y]
                for kk in set(new) | set(old):
                    mx = max(mx, abs(new.get(kk, 0.0) - old.get(kk, 0.0)))
                S[y] = new
            if mx < 1e-6:
                break
    share = {y: S[y] for y in w}
    res = aggregate_importance(M, share, targets, keys)
    log('  importance (fast): %d shocks traced through %d equations in %.1fs' % (len(res), len(w), time.time() - t0))
    return res, share, targets


def pick_key_vars(M, user=None, n=8):
    """Key variables = the hubs that carry the scenario: changed variables with an equation that the
    largest number of other variables depend on (one per concept), plus any the user names."""
    out = []
    for u in user or ():
        u = u.upper()
        if u in M.a or u in M.b:
            out.append(u)
    reach = M.reach_counts()
    cands = [v for v in M.differs if v in M.eqs and M.eqs[v].evaluable and v not in M.af_series]
    cands.sort(key=lambda v: (-reach.get(v, 0), -M.stats[v]['score']))
    stems = {M.stem(v) for v in out}
    for v in cands:
        if len(out) >= n + len(user or ()):
            break
        if M.stem(v) in stems or v in out:
            continue
        stems.add(M.stem(v))
        out.append(v)
    return out


def build_summary(M, args, log=lambda m: None):
    t0 = time.time()
    rng = list(range(M.ws, M.we + 1))
    per = [M.labels[t] for t in rng]
    keys = pick_key_vars(M, getattr(args, 'key_list', None), getattr(args, 'summary_n', 8))

    def sid(kind, sname):
        return sname if kind == 'data' else 'AF:' + sname

    # 1. attribution of each key variable's gap to the scenario shocks, every quarter
    attr = {}
    for y in keys:
        if y not in M.eqs or not M.eqs[y].evaluable:
            continue
        try:
            res, joint, gap = M.attribute(y)
        except (EvalError, ValueError, ZeroDivisionError, OverflowError):
            continue
        attr[y] = {'gap': [_r(v) for v in gap],
                   'rows': [[sid(k, sn), [_r(v, 5) for v in e]] for k, sn, e in res
                            if max([abs(v) for v in e if v == v] + [0]) > 1e-6 * max([abs(g) for g in gap if g == g] + [1e-30])]}
    log('  summary: attribution of %d key variables in %.1fs' % (len(attr), time.time() - t0))

    # 2. scenario shocks, ranked by how much of the key variables they explain
    weight = defaultdict(float)
    for y, a in attr.items():   # share of the key variable's movement over the whole window
        den = sum(abs(v) for v in a['gap'] if v is not None)
        for sname, e in a['rows']:
            if den:
                weight[sname] += min(1.5, sum(abs(v) for v in e if v is not None) / den)
    shocks = []
    for n in M.shock_list():
        k = M.kind(n)
        t = M.af_target(n) if k == 'F' else None
        if t is None and k in ('A', 'Q'):
            t = n
        node = 'AF:' + t if t else n
        st = M.stats[n]
        shocks.append({'id': node, 'series': n, 'target': t, 'kind': k, 'desc': M.desc.get(t or n, ''),
                       'first': M.labels[st['first']] if st['first'] is not None else '',
                       'dmax': _r(st['dmax']), 'pmax': _r(st['pmax'], 4), 'tmax': M.labels[st['tmax']],
                       'pend': _r(st['pend'], 4), 'dend': _r(st['dend']), 'w': round(weight.get(node, 0.0), 4)})
    known = {x['id'] for x in shocks}
    for sname in weight:  # sources that appear in attribution but not in the shock list (e.g. not-evaluable data)
        if sname not in known:
            base = sname[3:] if sname.startswith('AF:') else sname
            st = M.stats.get(base)
            shocks.append({'id': sname, 'series': base, 'target': base if sname.startswith('AF:') else None,
                           'kind': M.kind(base), 'desc': M.desc.get(base, ''),
                           'first': M.labels[st['first']] if st and st['first'] is not None else '',
                           'dmax': _r(st['dmax']) if st else None, 'pmax': _r(st['pmax'], 4) if st else None,
                           'tmax': M.labels[st['tmax']] if st else '', 'pend': _r(st['pend'], 4) if st else None,
                           'dend': _r(st['dend']) if st else None, 'w': round(weight[sname], 4)})
    shocks.sort(key=lambda x: (-x['w'], x['first']))

    # 3. transmission map: shocks + key variables, linked through the equations
    mshocks = [x for x in shocks if x['w'] >= 0.02][:10] or shocks[:6]
    nodes = [x['id'] for x in mshocks] + [k for k in keys if k not in [x['id'] for x in mshocks]]
    inset = set(k for k in keys)
    edges = {}

    def per_share(child, parent):
        d = M.dyn(child, periods=True)
        if not d or 'error' in d or parent not in d.get('contrib', {}):
            return None
        return [(c / g) if (g == g and c == c and abs(g) > 1e-12 * max(1.0, abs(M.series(child, 'a')[t])))
                else None for c, g, t in zip(d['contrib'][parent], d['gap'], rng)]

    def af_share(child):
        d = M.dyn(child, periods=True)
        if not d or 'error' in d or not d.get('af'):
            return None
        return [(c / g) if (g == g and c == c and abs(g) > 1e-12) else None for c, g in zip(d['af'], d['gap'])]

    def chain(path):
        out = None
        for a, b in zip(path, path[1:]):
            sh = per_share(b, a)
            if sh is None:
                return None
            out = sh if out is None else [(x * y) if (x is not None and y is not None) else None for x, y in zip(out, sh)]
        return out

    for src in nodes:
        if src.startswith('AF:'):
            tgt = src[3:]
            if tgt in inset:
                edges[(src, tgt)] = (1, af_share(tgt))
            start = tgt
            if tgt not in inset:   # add-factor on a non-key variable: continue from it
                pass
            else:
                continue
        else:
            start = src
        # BFS from start, stopping at key variables
        prev, q, seen = {start: None}, deque([start]), {start}
        while q:
            u = q.popleft()
            for v in M.children.get(u, ()):
                if v in seen or v not in M.differs_any:
                    continue
                seen.add(v)
                prev[v] = u
                if v in inset and v != src:
                    path = [v]
                    while path[-1] != start:
                        path.append(prev[path[-1]])
                    path = path[::-1]
                    if src.startswith('AF:'):
                        sh = chain(path)
                        a0 = af_share(start)
                        sh = [(x * y) if (x is not None and y is not None) else None for x, y in zip(a0, sh)] \
                            if (sh is not None and a0 is not None) else None
                    else:
                        sh = chain(path)
                    edges[(src, v)] = (len(path) - 1 + (1 if src.startswith('AF:') else 0), sh)
                    continue
                q.append(v)
    nodemeta = {}
    for nid in nodes:
        base = nid[3:] if nid.startswith('AF:') else nid
        a, b = M.series(base, 'a'), M.series(base, 'b')
        gp = []
        for t in rng:
            x, y = a[t] if a else NAN, b[t] if b else NAN
            gp.append(_r((y - x) / abs(x) * 100, 4) if (x == x and y == y and x) else None)
        nodemeta[nid] = {'label': ('add-factor of ' + base) if nid.startswith('AF:') else base,
                         'desc': M.desc.get(base, ''), 'shock': nid in [x['id'] for x in mshocks],
                         'pgap': gp, 'kind': M.kind(base)}
    mp = {'nodes': [dict(id=n, **nodemeta[n]) for n in nodes],
          'edges': [[a, b, steps, [_r(v, 3) for v in sh] if sh else None] for (a, b), (steps, sh) in edges.items()]}

    # 4. written summary
    cc = Counter(M.change_class(n) for n in M.differs)
    first = min((M.stats[n]['first'] for n in M.differs if M.stats[n]['first'] is not None), default=None)
    lines = []
    lines.append('B (%s) starts to differ from A (%s) in %s. %d series differ: %d were changed directly by the '
                 'scenario (exogenous variables, add-factors%s) and %d moved only because the model\'s equations passed '
                 'those changes on.' % (M.B.label, M.A.label, M.labels[first] if first is not None else '?',
                                        len(M.differs), cc.get('D', 0),
                                        ' or equation changes' if M.eq_status else '', cc.get('M', 0))
                 if M.has_eqs else
                 'B (%s) starts to differ from A (%s) in %s; %d series differ.' % (
                     M.B.label, M.A.label, M.labels[first] if first is not None else '?', len(M.differs)))
    sh_lines = []
    for x in shocks[:6]:
        nm = x['series'] if not x['id'].startswith('AF:') else M.af_label(x['target'])
        dsc = (' (%s)' % x['desc']) if x['desc'] else ''
        if x['id'].startswith('AF:') and x['series'] == x['target']:
            mag = ('its equation differs between A and B' if x['target'] in M.eq_status
                   else 'its equation\'s residual (add-factor) was changed')
        else:
            mag = 'largest change %s (%s) in %s, %s by %s' % (fmt(x['dmax']), _pct_word(x['pmax']), x['tmax'],
                                                           _pct_word(x['pend']), M.labels[M.we])
        sh_lines.append('%s%s: from %s; %s.' % (nm, dsc, x['first'], mag))
    key_lines = []
    lab = {x['id']: (x['series'] if not x['id'].startswith('AF:') else M.af_label(x['target'])) for x in shocks}
    for y in keys:
        st = M.stats.get(y)
        if not st:
            continue
        txt = '%s%s: %s vs A at %s (largest %s in %s).' % (
            y, (' (%s)' % M.desc[y]) if M.desc.get(y) else '', _pct_word(st['pend']), M.labels[st['tend']],
            _pct_word(st['pmax']), M.labels[st['tmax']])
        a = attr.get(y)
        if a:
            g = a['gap']
            i = max(j for j, v in enumerate(g) if v is not None)
            parts = sorted(((r[1][i] or 0) / g[i], r[0]) for r in a['rows'] if g[i]) if g[i] else []
            parts = sorted(parts, key=lambda z: -abs(z[0]))[:3]
            if parts:
                txt += ' Its gap comes from ' + ', '.join('%s %.0f%%' % (lab.get(sn, sn), 100 * v) for v, sn in parts) + '.'
        key_lines.append(txt)
    spread = []
    if M.has_eqs:
        sc = Counter(M.stem(n) for n in M.differs if M.change_class(n) == 'D' and M.kind(n) != 'F')
        if sc:
            spread.append('Directly changed concepts: ' + ', '.join('%s (%d)' % kv for kv in sc.most_common(6)) + '.')
    log('  summary built in %.1fs' % (time.time() - t0))
    return {'per': per, 'keys': keys, 'attr': attr, 'shocks': shocks[:30], 'map': mp,
            'text': {'intro': lines, 'shocks': sh_lines, 'keys': key_lines, 'spread': spread}}


def render_html(M, data):
    js = json.dumps(data, separators=(',', ':'), allow_nan=False).replace('</', '<\\/')
    return HTML_TEMPLATE.replace('__TITLE__', _esc('%s vs %s' % (M.A.label, M.B.label))).replace('__DATA__', js)


def build_report(M, args, eqstats):
    if M.has_eqs:
        write_check(M, os.path.splitext(args.out)[0] + '_not_evaluable.csv', quiet=not M.verbose)
    data, _ = build_data(M, args, eqstats)
    out = render_html(M, data)
    with open(args.out, 'w', encoding='utf-8') as f:
        f.write(out)
    if M.verbose:
        print('Report written: %s (%.1f MB)' % (args.out, os.path.getsize(args.out) / 1e6), file=sys.stderr)


def _impact_payload(M, x, y, shock_s):
    shock = parse_shock(shock_s)
    if y and y != x and y not in M.descendants(x):
        return {'none': True}
    eff, info = M.impact(x, y or None, shock)
    rng = range(M.ws, M.we + 1)
    out = {'per': [M.labels[t] for t in rng], 'nodes': info['nodes'], 'blocked': info['blocked'][:20],
           'shock': shock_s or ''}
    if y:
        A, Bv = M.series(y, 'a'), M.series(y, 'b')
        out['y'] = [_r(v) for v in eff.get(y, [])]
        out['gap'] = [_r(Bv[t] - A[t]) for t in rng]
        out['ya'] = [_r(A[t]) for t in rng]
    top = []
    for n, e in eff.items():
        if n == x or n == y:
            continue
        a = M.series(n, 'a')
        pc = [(v / abs(a[t]) * 100) if (v == v and a[t] == a[t] and a[t]) else None for v, t in zip(e, rng)]
        m = max([abs(p) for p in pc if p is not None] + [0])
        if m > 0:
            top.append((m, n, e, pc))
    top.sort(key=lambda z: -z[0])
    out['top'] = [[n, [_r(v) for v in e], [_r(p, 4) for p in pc]] for _, n, e, pc in top[:60]]
    return out


def _attr_payload(M, y):
    res, joint, gap = M.attribute(y)
    i = max(j for j, g in enumerate(gap) if g == g)
    rows = [[s, kind, M.af.get(s) if kind == 'addfactor' else None, [_r(v, 5) for v in e]]
            for kind, s, e in sorted(res, key=lambda r: -abs(r[2][i]) if r[2][i] == r[2][i] else 0)
            if abs(e[i]) > 1e-9 * max(1.0, abs(gap[i]))]
    return {'per': [M.labels[t] for t in range(M.ws, M.we + 1)], 'iend': i, 'gap': [_r(v) for v in gap],
            'joint': [_r(v) for v in joint], 'rows': rows[:60]}


def cmd_serve(M, args, eqstats):
    from http.server import HTTPServer, BaseHTTPRequestHandler
    from urllib.parse import urlparse, parse_qs
    import webbrowser
    data, es = build_data(M, args, eqstats, serve=True)
    page = render_html(M, data).encode('utf-8')

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, body, ctype='application/json'):
            self.send_response(200)
            self.send_header('Content-Type', ctype + '; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            u = urlparse(self.path)
            q = {k: v[0] for k, v in parse_qs(u.query).items()}
            try:
                if u.path in ('/', '/index.html'):
                    return self._send(page, 'text/html')
                if u.path == '/api/detail':
                    n = _norm(M, q['n'])
                    out = {'n': n, 'vals': _vals_payload(M, n, es), 'dec': _dec_payload(M, n),
                           'eqt': M.eqs[n].text[:700] if n in M.eqs else (M.exog[n].text[:700] if n in M.exog else None),
                           'why': M.eqs[n].why if n in M.eqs else None, 'desc': M.desc.get(n, '')}
                elif u.path == '/api/impact':
                    out = _impact_payload(M, _norm(M, q['x']), _norm(M, q['y']) if q.get('y') else None,
                                          q.get('shock'))
                elif u.path == '/api/influence':
                    r = M.influence(_norm(M, q['x']), parse_shock(q.get('shock')))
                    r['top'] = [[n, _r(v), _r(p, 4), lab] for n, v, p, lab in r['top']]
                    out = r
                elif u.path == '/api/attr':
                    out = _attr_payload(M, _norm(M, q['y']))
                else:
                    self.send_error(404)
                    return
            except SystemExit as e:
                out = {'error': str(e)}
            except Exception as e:  # keep the server alive
                out = {'error': '%s: %s' % (type(e).__name__, e)}
            self._send(json.dumps(out, allow_nan=False, default=str).encode('utf-8'))

    srv = HTTPServer(('127.0.0.1', args.port), H)
    url = 'http://127.0.0.1:%d/' % args.port
    print('Serving %s  (Ctrl+C to stop)' % url, file=sys.stderr)
    if not args.no_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


def _esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;'))


HTML_TEMPLATE = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Scenario explain: __TITLE__</title>
<style>
:root{--bg:#fff;--fg:#1d2433;--mut:#667085;--line:#e4e7ec;--card:#f7f8fa;--a:#2563eb;--b:#d9480f;--pos:#0f9d58;--neg:#d92d20;--hl:#fff5d6}
@media (prefers-color-scheme:dark){:root{--bg:#0f1218;--fg:#e6e9ef;--mut:#98a2b3;--line:#2a2f3a;--card:#171b23;--a:#6c9cff;--b:#ff8a4c;--pos:#3ccb7f;--neg:#ff6b5e;--hl:#3a3216}}
*{box-sizing:border-box}
body{margin:0;font:14px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;background:var(--bg);color:var(--fg)}
header{padding:16px 20px 8px;border-bottom:1px solid var(--line)}
h1{font-size:19px;margin:0 0 4px}
h2{font-size:16px;margin:18px 0 8px}
.mut{color:var(--mut)}
.small{font-size:12px}
.wrap{padding:12px 20px 40px;max-width:1300px}
nav{display:flex;gap:4px;flex-wrap:wrap;padding:8px 20px;border-bottom:1px solid var(--line);position:sticky;top:0;background:var(--bg);z-index:5}
nav button{border:1px solid var(--line);background:var(--card);color:var(--fg);padding:6px 12px;border-radius:16px;cursor:pointer;font:inherit}
nav button.on{background:var(--fg);color:var(--bg);border-color:var(--fg)}
.kpis{display:flex;gap:10px;flex-wrap:wrap;margin:10px 0}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:8px 12px;min-width:120px}
.kpi b{display:block;font-size:20px}
.ctl{display:flex;gap:8px;flex-wrap:wrap;margin:8px 0}
input,select{font:inherit;padding:6px 8px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--fg)}
input[type=text]{min-width:220px}
table{border-collapse:collapse;width:100%}
th,td{padding:5px 8px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{background:var(--bg);cursor:pointer;white-space:nowrap;font-weight:600;font-size:12px;color:var(--mut)}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
tr.cl{cursor:pointer}
tr.cl:hover{background:var(--card)}
.tw{overflow-x:auto}
a.v{color:var(--a);cursor:pointer;text-decoration:none;font-weight:500}
a.v:hover{text-decoration:underline}
.pos{color:var(--pos)}.neg{color:var(--neg)}
.badge{display:inline-block;font-size:11px;padding:1px 7px;border-radius:10px;border:1px solid var(--line);background:var(--card);white-space:nowrap}
.bX{border-color:#c4320a;color:#c4320a}.bF{border-color:#7a5af8;color:#7a5af8}.bA{border-color:#b54708;color:#b54708}.bE{color:var(--mut)}.bN{color:var(--mut);border-style:dashed}
#det{position:fixed;top:0;right:0;bottom:0;width:min(760px,100%);background:var(--bg);border-left:1px solid var(--line);box-shadow:-8px 0 24px rgba(0,0,0,.12);overflow-y:auto;transform:translateX(105%);transition:transform .2s;z-index:20;padding:14px 18px 60px}
#det.open{transform:none}
.dh{display:flex;justify-content:space-between;align-items:center;gap:8px}
.btn{border:1px solid var(--line);background:var(--card);color:var(--fg);border-radius:8px;padding:5px 10px;cursor:pointer;font:inherit}
pre.eq{white-space:pre-wrap;word-break:break-word;background:var(--card);border:1px solid var(--line);padding:8px;border-radius:8px;font-size:12px;margin:6px 0}
svg text{fill:var(--mut);font-size:11px}
.bars td{padding:3px 6px}
.bar{height:12px;border-radius:2px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:10px 12px;margin:10px 0}
.hint{font-size:12px;color:var(--mut);margin:4px 0 8px}
.path{margin:8px 0;padding:8px 10px;border:1px solid var(--line);border-radius:10px}
.step{display:inline-block;margin:2px 0}
.arrow{color:var(--mut);margin:0 6px}
.hide{display:none}
</style></head><body>
<header>
<h1>Scenario explain</h1>
<div id="sub" class="mut"></div>
</header>
<nav id="tabs"></nav>
<div class="wrap">
<section id="t-story"></section>
<section id="t-sum" class="hide"></section>
<section id="t-top" class="hide"></section>
<section id="t-entry" class="hide"></section>
<section id="t-afx" class="hide"></section>
<section id="t-eqs" class="hide"></section>
<section id="t-imp" class="hide"></section>
<section id="t-inf" class="hide"></section>
<section id="t-lay" class="hide"></section>
<section id="t-stem" class="hide"></section>
<section id="t-geo" class="hide"></section>
<section id="t-path" class="hide"></section>
</div>
<div id="det"></div>
<script id="data" type="application/json">__DATA__</script>
<script>
"use strict";
const D=JSON.parse(document.getElementById('data').textContent);
const M=D.meta, NI={}; D.nodes.forEach((n,i)=>NI[n]=i);
const CH=D.nodes.map(()=>[]); D.par.forEach((ps,i)=>ps.forEach(p=>CH[p].push(i)));
const ROW={}; D.rows.forEach(r=>ROW[r[0]]=r);
const KT={Q:'equation changed',X:'potential exogenous',F:'add-factor',A:'mostly add-factor',E:'formula',N:'not evaluable','-':''};
const KL={Q:'Its equation is different in A and B (changed, added or dropped); that switch is reported as one effect, together with its add-factor',X:'No equation computes this variable in the files given, so it is treated as exogenous: set directly by the scenario',F:'An add-factor series'+(D.meta.af_suffix?' (name ends in '+D.meta.af_suffix+')':''),A:'Has an equation, but most of its change comes from its add-factor / residual, not from its drivers',E:'Change is explained by its equation (formula) drivers',N:'Has an equation that could not be evaluated'};
const SERVE=D.meta.serve&&location.protocol.startsWith('http');
const EXOG=new Set(D.meta.exog||[]);
const ES=D.eqstat||{};
function afWord(n){const st=ES[n];return st==='changed'?'equation changed':st==='only_a'?'equation dropped in B':st==='only_b'?'equation added in B':'add-factor'}
function afText(n){return afWord(n)+(ES[n]?': ':' on ')+n}
function afLink(n){return `<i>${afWord(n)}${ES[n]?':':' on'}</i> ${vlink(n)}`}
const $=(s,e=document)=>e.querySelector(s);
const esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function fmt(v){if(v==null||isNaN(v))return'–';if(v===0||Math.abs(v)<1e-9)return'0';const a=Math.abs(v);if(a>=1e4)return v.toLocaleString(undefined,{maximumFractionDigits:0});return Number(v.toPrecision(4)).toString()}
function pct(v){if(v==null||isNaN(v))return'–';return(v>0?'+':'')+v.toFixed(2)+'%'}
function cls(v){return v>0?'pos':v<0?'neg':''}
function badge(k,n){if(n&&EXOG.has(n))return `<span class="badge bX" title="Has an equation, but the modeller set it exogenously (--exog)">set by modeller</span>`;return k&&k!=='-'?`<span class="badge b${k}" title="${esc(KL[k]||'')}">${KT[k]}</span>`:''}
function vlink(n){return `<a class="v" data-v="${esc(n)}">${esc(n)}</a>`}
function desc(n){const r=ROW[n];return r?r[1]:(D.desc[n]||'')}
document.addEventListener('click',e=>{const a=e.target.closest('[data-v]');if(a){e.preventDefault();show(a.dataset.v)}});

// ---------- tabs
const TABS=[['story','Summary'],['sum','Overview'],['top','Key changes'],['entry','Where B starts'],['afx','Formula vs add-factor'],['stem','By concept'],['geo','By geography'],['eqs','Equations A vs B'],['imp','Importance'],['inf','Influence'],['lay','Layers'],['path','Impact X → Y']];
$('#tabs').innerHTML=TABS.filter(t=>t[0]==='story'?D.summary:t[0]==='imp'?D.imp:t[0]==='eqs'?(D.eqdiff&&D.eqdiff.length):t[0]==='afx'?D.afrows.length:t[0]==='inf'?M.has_eqs:t[0]==='lay'?!!D.layers:(t[0]!=='stem'&&t[0]!=='geo')||D[t[0]].length).map(t=>`<button data-t="${t[0]}">${t[1]}</button>`).join('');
function tab(t){document.querySelectorAll('nav button').forEach(b=>b.classList.toggle('on',b.dataset.t===t));TABS.forEach(x=>$('#t-'+x[0]).classList.toggle('hide',x[0]!==t))}
$('#tabs').addEventListener('click',e=>{if(e.target.dataset.t)tab(e.target.dataset.t)});
$('#sub').innerHTML=`<b>A</b> = ${esc(M.a)} &nbsp; vs &nbsp; <b>B</b> = ${esc(M.b)} &nbsp;·&nbsp; window ${M.ws}–${M.we} &nbsp;·&nbsp; generated ${M.generated}`;

// ---------- summary + transmission map
(function(){const S=D.summary;if(!S)return;
 const NP=S.per.length,lab={};S.shocks.forEach(x=>lab[x.id]=x.id.startsWith('AF:')?afText(x.target):x.series);
 const nm=id=>lab[id]||(id.startsWith('AF:')?afText(id.slice(3)):id);
 const lnk=id=>id.startsWith('AF:')?afLink(id.slice(3)):vlink(id);
 let h=`<div class="card">${S.text.intro.map(esc).join(' ')} ${S.text.spread.map(esc).join(' ')}</div>
 <h2>What the scenario changed</h2><div class="hint">The shocks B is built on, ordered by how much of the key variables they explain.</div>
 <div class="tw"><table><thead><tr><th>Shock</th><th>Description</th><th>From</th><th class="n">Largest change</th><th class="n">At ${esc(M.we)}</th><th class="n" title="Sum over key variables of the share of their movement (whole window) that this shock explains">Weight in key variables</th></tr></thead><tbody>${S.shocks.slice(0,15).map(x=>`<tr><td>${x.id.startsWith('AF:')&&x.series!==x.target?`${vlink(x.series)} <span class="small mut">→ ${esc(x.target)}</span>`:x.id.startsWith('AF:')?afLink(x.target):vlink(x.series)}</td><td class="small">${esc(x.desc).slice(0,60)}</td><td>${esc(x.first)}</td><td class="n ${cls(x.dmax)}">${fmt(x.dmax)} <span class="small">(${pct(x.pmax)}, ${esc(x.tmax)})</span></td><td class="n ${cls(x.pend)}">${pct(x.pend)}</td><td class="n">${x.w?x.w.toFixed(2):'–'}</td></tr>`).join('')}</tbody></table></div>
 <h2>Key variables and where their gaps come from</h2><div class="hint">Key variables are picked automatically as the changed variables that the most other variables depend on (one per concept). Use <code>--key</code> to choose your own. Each share re-solves every equation upstream with only that shock switched from A to B.</div>
 <div class="ctl" style="align-items:center"><span class="small mut">Quarter</span><input type="range" id="ssl" min="0" max="${NP-1}" value="${NP-1}" style="flex:1;min-width:200px"><b id="slab"></b></div>
 <div id="smap"></div><div id="skeys"></div>
 <h2>In words</h2><ul>${S.text.shocks.length?'<li><b>Shocks:</b><ul>'+S.text.shocks.map(t=>`<li>${esc(t)}</li>`).join('')+'</ul></li>':''}${S.text.keys.length?'<li><b>Key variables:</b><ul>'+S.text.keys.map(t=>`<li>${esc(t)}</li>`).join('')+'</ul></li>':''}</ul>`;
 $('#t-story').innerHTML=h;
 const upd=()=>{const i=+$('#ssl').value;$('#slab').textContent=S.per[i];$('#smap').innerHTML=mapSVG(S,i,nm);mapZoom();
  $('#skeys').innerHTML=`<div class="tw"><table><thead><tr><th>Key variable</th><th>Description</th><th class="n">Gap %</th><th>Main sources of the gap at ${esc(S.per[i])}</th></tr></thead><tbody>${S.keys.map(k=>{const a=S.attr[k],nd=S.map.nodes.find(n=>n.id===k),gp=nd?nd.pgap[i]:null;
   let src='<span class="mut small">not evaluable</span>';if(a){const g=a.gap[i];src=g?a.rows.map(r=>[r[0],(r[1][i]||0)/g]).filter(r=>Math.abs(r[1])>=.005).sort((x,y)=>Math.abs(y[1])-Math.abs(x[1])).slice(0,4).map(r=>`${lnk(r[0])} <b>${Math.round(100*r[1])}%</b>`).join(' · ')||'<span class="mut small">none above 0.5%</span>':'<span class="mut small">no gap this quarter</span>'}
   return `<tr><td>${vlink(k)}</td><td class="small">${esc(desc(k)).slice(0,50)}</td><td class="n ${cls(gp)}">${pct(gp)}</td><td class="small">${src}</td></tr>`}).join('')}</tbody></table></div>`};
 $('#ssl').addEventListener('input',upd);upd();
})();
var MV;
function mapZoom(){const sv=$('#msvg');if(!sv)return;const set=()=>sv.setAttribute('viewBox',MV.v.join(' '));
 const zoomAt=(f,cx,cy)=>{const [x,y,w,h]=MV.v,nw=Math.min(MV.full[2]*2,Math.max(MV.full[2]/12,w*f)),k=nw/w,nh=h*k;MV.v=[cx-(cx-x)*k,cy-(cy-y)*k,nw,nh];set()};
 const toSvg=(ex,ey)=>{const r=sv.getBoundingClientRect(),[x,y,w,h]=MV.v,sc=Math.max(w/r.width,h/r.height),ox=(r.width*sc-w)/2,oy=(r.height*sc-h)/2;return[x-ox+(ex-r.left)*sc,y-oy+(ey-r.top)*sc,sc]};
 sv.addEventListener('wheel',e=>{e.preventDefault();const [cx,cy]=toSvg(e.clientX,e.clientY);zoomAt(e.deltaY>0?1.15:1/1.15,cx,cy)},{passive:false});
 let drag=null,moved=false;
 sv.addEventListener('pointerdown',e=>{drag=[e.clientX,e.clientY,MV.v.slice()];moved=false;sv.style.cursor='grabbing'});
 window.addEventListener('pointermove',e=>{if(!drag)return;const sc=toSvg(0,0)[2],dx=(e.clientX-drag[0])*sc,dy=(e.clientY-drag[1])*sc;if(Math.abs(dx)+Math.abs(dy)>2)moved=true;MV.v=[drag[2][0]-dx,drag[2][1]-dy,drag[2][2],drag[2][3]];set()});
 window.addEventListener('pointerup',()=>{if(drag){drag=null;sv.style.cursor='grab'}});
 sv.addEventListener('click',e=>{if(moved){e.stopPropagation();e.preventDefault();moved=false}},true);
 document.querySelectorAll('[data-z]').forEach(b=>b.onclick=()=>{const [x,y,w,h]=MV.v;if(b.dataset.z==='reset'){MV.v=MV.full.slice();set()}else zoomAt(b.dataset.z==='in'?1/1.3:1.3,x+w/2,y+h/2)})}
function mapSVG(S,i,nm){
 const N=S.map.nodes,E=S.map.edges,id2=new Map(N.map((n,k)=>[n.id,k]));
 // layers: shocks at 0, others = shortest number of map links from any shock
 const L=new Array(N.length).fill(Infinity),q=[];N.forEach((n,k)=>{if(n.shock){L[k]=0;q.push(k)}});
 const out=N.map(()=>[]);E.forEach(e=>{const a=id2.get(e[0]),b=id2.get(e[1]);if(a!=null&&b!=null)out[a].push(b)});
 while(q.length){const u=q.shift();for(const v of out[u])if(L[v]>L[u]+1){L[v]=L[u]+1;q.push(v)}}
 const mx=Math.max(0,...L.filter(isFinite));N.forEach((n,k)=>{if(!isFinite(L[k]))L[k]=mx+1});
 const nl=Math.max(...L)+1,cols=[...Array(nl)].map(()=>[]);N.forEach((n,k)=>cols[L[k]].push(k));
 const ins=N.map(()=>[]);E.forEach(e=>{const a=id2.get(e[0]),b=id2.get(e[1]);if(a!=null&&b!=null)ins[b].push(a)});
 const pos=new Array(N.length);cols.forEach((c,ci)=>{if(ci>0)c.sort((x,y)=>{const m=z=>ins[z].length?ins[z].reduce((s,p)=>s+(pos[p]||0),0)/ins[z].length:99;return m(x)-m(y)});c.forEach((k,r)=>pos[k]=r)});
 const W=1000,BW=150,BH=40,gapY=18,maxR=Math.max(...cols.map(c=>c.length)),H=Math.max(160,maxR*(BH+gapY)+30);
 const cx=ci=>14+ci*((W-BW-28)/Math.max(1,nl-1)),cy=(ci,r)=>{const n=cols[ci].length,tot=n*BH+(n-1)*gapY;return (H-tot)/2+r*(BH+gapY)};
 const P=N.map((n,k)=>({x:cx(L[k]),y:cy(L[k],pos[k])}));
 const col=v=>{if(v==null)return'var(--card)';const a=Math.min(1,Math.abs(v)/3);return v<0?`rgba(217,45,32,${.12+.55*a})`:`rgba(15,157,88,${.12+.55*a})`};
 let ed='';E.forEach(e=>{const a=id2.get(e[0]),b=id2.get(e[1]);if(a==null||b==null)return;const sh=e[3]?e[3][i]:null,w=sh==null?1:1+9*Math.min(1,Math.abs(sh));
  const x1=P[a].x+BW,y1=P[a].y+BH/2,x2=P[b].x,y2=P[b].y+BH/2;let d,fwd=x2>x1+4,lx,ly;
  if(fwd){const m=(x1+x2)/2;d=`M${x1},${y1} C${m},${y1} ${m},${y2} ${x2},${y2}`;lx=x2-18;ly=y2-5}
  else if(Math.abs(P[a].x-P[b].x)<2){const xr=x1+34;d=`M${x1},${y1} C${xr},${y1} ${xr},${y2} ${x1},${y2}`;lx=xr-4;ly=(y1+y2)/2}
  else{const yy=H+18;d=`M${P[a].x+BW/2},${P[a].y+BH} C${P[a].x+BW/2},${yy} ${P[b].x+BW/2},${yy} ${P[b].x+BW/2},${P[b].y+BH}`;lx=null}
  ed+=`<path d="${d}" fill="none" stroke="${sh!=null&&sh<0?'var(--neg)':'var(--mut)'}" stroke-opacity="${fwd?(sh==null?.35:.35+.5*Math.min(1,Math.abs(sh))):.28}" stroke-width="${fwd?w:Math.min(w,4)}" ${e[2]>1||!fwd?'stroke-dasharray="6,4"':''}><title>${esc(nm(e[0]))} → ${esc(e[1])}: ${sh==null?'share n/a':Math.round(100*sh)+'% of its gap'}${e[2]>1?' (via '+(e[2]-1)+' other variable(s))':''}</title></path>`;
  if(lx!=null&&sh!=null&&Math.abs(sh)>=.05)ed+=`<text x="${lx}" y="${ly}" text-anchor="end" style="font-size:10px;font-weight:600">${Math.round(100*sh)}%</text>`});
 let nd=N.map((n,k)=>{const v=n.pgap[i],lb=n.id.startsWith('AF:')?afWord(n.id.slice(3)):n.id;return `<g class="mn" data-v="${esc(n.id.startsWith('AF:')?n.id.slice(3):n.id)}" style="cursor:pointer"><rect x="${P[k].x}" y="${P[k].y}" width="${BW}" height="${BH}" rx="8" fill="${col(v)}" stroke="${n.shock?'#c4320a':'var(--line)'}" stroke-width="${n.shock?2:1}"/><text x="${P[k].x+8}" y="${P[k].y+16}" style="font-size:12px;fill:var(--fg);font-weight:600">${esc(lb.length>22?lb.slice(0,21)+'…':lb)}</text><text x="${P[k].x+8}" y="${P[k].y+31}" style="font-size:11px">${n.id.startsWith('AF:')?esc(n.id.slice(3)):(v==null?'':pct(v))}${n.shock&&!n.id.startsWith('AF:')?' · shock':''}</text><title>${esc(n.label)}${n.desc?' — '+esc(n.desc):''}: ${pct(v)} vs A</title></g>`}).join('');
 const full=[0,0,W,H+30];if(!MV||MV.full[2]!==W||MV.full[3]!==H+30)MV={full,v:full.slice()};
 return `<div class="ctl" style="margin:4px 0"><button class="btn" data-z="in" title="Zoom in">＋</button><button class="btn" data-z="out" title="Zoom out">－</button><button class="btn" data-z="reset">Reset view</button><span class="small mut">Scroll to zoom, drag to move.</span></div><div style="border:1px solid var(--line);border-radius:10px;overflow:hidden"><svg id="msvg" viewBox="${MV.v.join(' ')}" width="100%" style="height:${Math.min(640,Math.max(260,(H+30)*0.9))}px;cursor:grab;touch-action:none;display:block" preserveAspectRatio="xMidYMid meet">${ed}${nd}</svg></div><div class="small mut">Left: the scenario's shocks (red border). Moving right: key variables in the order the shocks reach them, coloured by their gap vs A (green above, red below). Solid lines are equation links; thickness and % = share of the receiving variable's gap that comes from that link. Dashed lines either pass through variables not shown or feed back to the left (model feedback loops). Hover for details; click a box to open it.</div>`}

// ---------- overview
(function(){
 const kinds={},ccs={};D.rows.forEach(r=>{kinds[r[9]]=(kinds[r[9]]||0)+1;ccs[r[10]]=(ccs[r[10]]||0)+1});
 let h=`<div class="kpis"><div class="kpi"><span class="mut small">series in both</span><b>${M.n_common.toLocaleString()}</b></div>
 <div class="kpi"><span class="mut small">differ between A and B</span><b>${M.n_diff.toLocaleString()}</b></div>`;
 if(M.has_eqs){h+=`<div class="kpi"><span class="mut small">changed directly (assumption / add-factor)</span><b>${(ccs.D||0).toLocaleString()}</b></div>
 <div class="kpi"><span class="mut small">moved only through their formula</span><b>${(ccs.M||0).toLocaleString()}</b></div>
 <div class="kpi"><span class="mut small">changed assumptions (no equation)</span><b>${(kinds.X||0).toLocaleString()}</b></div>
 <div class="kpi"><span class="mut small">add-factor series changed</span><b>${(kinds.F||0).toLocaleString()}</b></div>
 <div class="kpi"><span class="mut small">variables driven mostly by add-factors</span><b>${(kinds.A||0).toLocaleString()}</b></div>`}
 h+=`</div><p class="mut small">Comparison window ${M.ws}–${M.we} (data ${M.first}–${M.last}). ${M.n_only_a?M.n_only_a+' series only in A. ':''}${M.n_only_b?M.n_only_b+' series only in B. ':''}</p>`;
 h+=`<h2>Biggest changes</h2><div class="hint">Score = largest gap between B and A, divided by the series' typical period-to-period move in A, so series in different units can be ranked together.</div>`;
 h+=tableHTML(D.rows.slice(0,15));
 if(D.entry.length){h+=`<h2>Where scenario B departs from A first</h2><div class="hint">Assumptions (no equation), add-factor series that changed, and variables whose gap comes mostly from their add-factor rather than their formula. These are the shocks the scenario is built on — everything else follows through the equations.</div>`;
  h+=tableHTML(D.entry.slice(0,15).map(e=>ROW[e[0]]).filter(Boolean),'first')}
 if(M.eqs){h+=`<p class="mut small">Equations loaded: ${esc(Object.entries(M.eqs).map(([k,v])=>k+' '+v).join(', '))}</p>`}
 $('#t-sum').innerHTML=h;
})();

function tableHTML(rows,sortHint){
 return `<div class="tw"><table><thead><tr><th>Series</th><th>Description</th><th class="n">Score</th><th class="n">Max gap</th><th class="n">Max gap %</th><th class="n">Gap at end %</th><th>Diverges</th><th>Type</th></tr></thead><tbody>${rows.map(r=>`<tr class="cl" data-v="${esc(r[0])}"><td><a class="v">${esc(r[0])}</a></td><td class="small">${esc(r[1]).slice(0,80)}</td><td class="n">${fmt(r[3])}</td><td class="n ${cls(r[4])}">${fmt(r[4])}</td><td class="n ${cls(r[5])}">${pct(r[5])}</td><td class="n ${cls(r[7])}">${pct(r[7])}</td><td>${esc(r[8])}</td><td>${badge(r[9],r[0])}</td></tr>`).join('')}</tbody></table></div>`}

// ---------- key changes (filterable, complete list)
const CC={D:'changed directly',M:'via formula only',U:'not evaluable','':''};
const CCL={D:'Changed directly: an assumption, an add-factor series, or its equation\'s add-factor changed',M:'Same equation and same add-factor in both scenarios; it moved only because its inputs moved',U:'Its equation could not be evaluated, so this cannot be told'};
function ccBadge(c){return c?`<span class="small ${c==='D'?'':'mut'}" title="${esc(CCL[c]||'')}">${CC[c]}</span>`:''}
let KC={};
(function(){
 const geos=[...new Set(D.rows.map(r=>r[2]).filter(Boolean))].sort();
 let h=`<div class="hint">Every series that differs between A and B. ${M.has_eqs?'<b>Changed directly</b> = an assumption or add-factor changed. <b>Via formula only</b> = same equation and same add-factor; it moved only because its inputs moved.':''}</div>
 <div class="ctl"><input type="text" id="q" placeholder="Search name or description (regex ok)">
 ${M.has_eqs?`<select id="cc"><option value="">All changed series</option><option value="D">Changed directly</option><option value="M">Moved via formula only</option><option value="U">Not evaluable</option></select>`:''}
 ${geos.length>1?`<select id="g"><option value="">All geographies</option>${D.rows.some(r=>!r[2])?'<option value="__none__">(no geography)</option>':''}${geos.map(g=>`<option>${esc(g)}</option>`).join('')}</select>`:''}
 ${M.has_eqs?`<select id="k"><option value="">All types</option><option value="X">Potential exogenous</option><option value="Q">Equation changed</option><option value="F">Add-factor series</option><option value="A">Mostly add-factor</option><option value="E">Formula</option><option value="N">Not evaluable</option></select>`:''}
 <select id="dir"><option value="">Up or down</option><option value="u">B above A</option><option value="d">B below A</option></select>
 <button class="btn" id="csv">Download CSV</button></div>
 <div id="cnt" class="hint"></div><div class="tw"><table id="tt"><thead><tr>
 <th data-c="0">Series</th><th data-c="1">Description</th><th data-c="3" class="n">Score ▾</th><th data-c="4" class="n">Max gap</th><th data-c="5" class="n">Max gap %</th><th data-c="7" class="n">Gap at end %</th><th data-c="8">Diverges</th><th data-c="9">Type</th>${M.has_eqs?'<th data-c="10">Change</th>':''}</tr></thead><tbody></tbody></table></div><div class="ctl" id="more"></div>`;
 $('#t-top').innerHTML=h;
 let sc=3,sd=-1,lim=500,cur=[];
 function filt(){
  const q=$('#q').value.trim();let rx=null;try{rx=q?new RegExp(q,'i'):null}catch(e){rx=new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'i')}
  const g=$('#g')?$('#g').value:'',k=$('#k')?$('#k').value:'',c=$('#cc')?$('#cc').value:'',dr=$('#dir').value;
  let rs=D.rows.filter(r=>(!rx||rx.test(r[0])||rx.test(r[1]))&&(!g||(g==='__none__'?!r[2]:r[2]===g))&&(!k||r[9]===k)&&(!c||r[10]===c)&&(!dr||(dr==='u'?r[4]>0:r[4]<0)));
  rs.sort((x,y)=>{let a=x[sc],b=y[sc];if(typeof a==='number'||typeof b==='number'){a=a==null?-Infinity:Math.abs(a);b=b==null?-Infinity:Math.abs(b)}return a<b?-sd:a>b?sd:0});
  return rs}
 function render(){
  cur=filt();
  $('#cnt').textContent=cur.length.toLocaleString()+' series'+(cur.length>lim?` (showing ${lim.toLocaleString()})`:'')+' · numeric columns sort by size, ignoring sign';
  $('#tt tbody').innerHTML=cur.slice(0,lim).map(r=>`<tr class="cl" data-v="${esc(r[0])}"><td><a class="v">${esc(r[0])}</a></td><td class="small">${esc(r[1]).slice(0,80)}</td><td class="n">${fmt(r[3])}</td><td class="n ${cls(r[4])}">${fmt(r[4])}</td><td class="n ${cls(r[5])}">${pct(r[5])}</td><td class="n ${cls(r[7])}">${pct(r[7])}</td><td>${esc(r[8])}</td><td>${badge(r[9],r[0])}</td>${M.has_eqs?`<td>${ccBadge(r[10])}</td>`:''}</tr>`).join('');
  $('#more').innerHTML=cur.length>lim?`<button class="btn" id="m1">Show ${Math.min(2000,cur.length-lim).toLocaleString()} more</button><button class="btn" id="m2">Show all ${cur.length.toLocaleString()}</button>`:'';
  if(cur.length>lim){$('#m1').onclick=()=>{lim+=2000;render()};$('#m2').onclick=()=>{lim=cur.length;render()}}}
 ['q','g','k','cc','dir'].forEach(id=>{const e=$('#'+id);if(e)e.addEventListener('input',()=>{lim=500;render()})});
 $('#tt thead').addEventListener('click',e=>{const c=e.target.dataset.c;if(c==null)return;const ci=+c;if(ci===sc)sd=-sd;else{sc=ci;sd=(ci<=2||ci>=8)?1:-1}document.querySelectorAll('#tt th').forEach(t=>t.textContent=t.textContent.replace(/ [▾▴]$/,''));e.target.textContent+=sd<0?' ▾':' ▴';render()});
 $('#csv').onclick=()=>{const q=v=>{v=v==null?'':String(v);return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v};
  const hdr=['series','description','geo','score','max_gap','max_gap_pct','end_gap','end_gap_pct','first_divergence','type','change'];
  const lines=[hdr.join(',')].concat(cur.map(r=>[r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7],r[8],KT[r[9]]||r[9],CC[r[10]]||''].map(q).join(',')));
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([lines.join('\n')],{type:'text/csv'}));a.download='changed_series.csv';document.body.appendChild(a);a.click();a.remove()};
 KC={set:(o)=>{['q','g','k','cc','dir'].forEach(id=>{const e=$('#'+id);if(e)e.value=o[id]||''});lim=500;render();tab('top')}};
 render();
})();

// ---------- entry points
(function(){
 let h=`<div class="hint">${M.has_eqs?'Ordered by when each first diverges. <b>Assumptions</b> have no equation in the model, so they are set by the scenario directly. <b>Add-factor</b> series that changed, and variables whose gap comes mostly (≥'+Math.round(M.adj_thr*100)+'%) from their add-factor rather than their formula. Start here to see what the scenario actually changed.':'No equations were loaded, so assumptions cannot be told apart from model outcomes. Listed instead: series ordered by when they first diverge.'}</div>`;
 const rows=M.has_eqs?D.entry.map(e=>ROW[e[0]]).filter(Boolean):D.rows.slice().sort((a,b)=>a[8]<b[8]?-1:a[8]>b[8]?1:b[3]-a[3]).slice(0,300);
 h+=`<div class="hint">${rows.length.toLocaleString()} series</div>`+tableHTML(rows.slice(0,500));
 $('#t-entry').innerHTML=h;
})();

// ---------- formula vs add-factor (time slider)
(function(){if(!D.afrows.length)return;
 const NP=D.afrows[0][5].length,P=i=>D.per[i+D.wsi];
 $('#t-afx').innerHTML=`<div class="hint">For every variable with an equation: how much of its gap (B−A) comes from its <b>formula</b> (its inputs changed) versus its <b>add-factor</b> (${M.af_suffix?esc(M.af_suffix)+' series or ':''}residual changed). The equation is the same in both scenarios, so whatever the formula does not explain comes from the add-factor. Move the slider to check any quarter, or use the whole window.</div>
 <div class="ctl" style="align-items:center"><select id="amode"><option value="p">At one quarter</option><option value="w">Whole window (sum of all quarters)</option></select>
 <input type="range" id="asl" min="0" max="${NP-1}" value="${NP-1}" style="flex:1;min-width:200px"><b id="alab">${esc(P(NP-1))}</b></div>
 <div class="ctl"><input type="text" id="aq" placeholder="Search"><select id="am"><option value="">All</option><option value="af" selected>Add-factor changed</option><option value="maf">Mostly add-factor (≥50%)</option><option value="f">Formula only</option></select><select id="as"><option value="s">Sort by size of change</option><option value="a">Sort by add-factor share</option></select></div><div id="acnt" class="hint"></div>
 <div class="tw"><table><thead><tr><th>Series</th><th>Description</th><th class="n">Gap (B−A)</th><th class="n">From formula</th><th class="n">From add-factor</th><th class="n">Other</th><th>Add-factor series</th></tr></thead><tbody id="atb"></tbody></table></div><div class="ctl" id="amore"></div>`;
 const sh=v=>v==null?'–':Math.round(v*100)+'%';let lim=500;
 const val=(x,i,w)=>{const sum=a=>a?a.reduce((s,v)=>s+(v||0),0):0;if(w)return[sum(x[5]),sum(x[6]),sum(x[7])];return[x[5][i],x[6][i],x[7]?x[7][i]:0]};
 const r=()=>{const q=$('#aq').value.trim().toUpperCase(),m=$('#am').value,w=$('#amode').value==='w',i=+$('#asl').value;
  $('#asl').disabled=w;$('#alab').textContent=w?`${M.ws}–${M.we}`:P(i);
  let rs=D.afrows.map(x=>{const [g,f,a]=val(x,i,w),ok=g!=null&&Math.abs(g)>1e-12;return [x,g,ok?f/g:null,ok?(a||0)/g:null,ok?(g-(f||0)-(a||0))/g:null]})
   .filter(([x,g,fs,as])=>(!q||x[0].includes(q)||x[1].toUpperCase().includes(q))&&(!m||(m==='af'?x[7]!==null:m==='maf'?Math.abs(as||0)>=.5:x[7]===null)));
  if($('#as').value==='a')rs.sort((u,v)=>Math.abs(v[3]||0)-Math.abs(u[3]||0));
  $('#acnt').textContent=rs.length.toLocaleString()+' variables'+(rs.length>lim?` (showing ${lim.toLocaleString()})`:'');
  $('#atb').innerHTML=rs.slice(0,lim).map(([x,g,fs,as,os])=>`<tr class="cl" data-v="${esc(x[0])}"><td><a class="v">${esc(x[0])}</a></td><td class="small">${esc(x[1]).slice(0,60)}</td><td class="n ${cls(g)}">${fmt(g)}</td><td class="n">${sh(fs)}</td><td class="n"><b>${sh(as)}</b></td><td class="n mut">${os!=null&&Math.abs(os)>=.005?sh(os):''}</td><td>${x[2]?vlink(x[2])+(x[3]?` <span class="small ${cls(x[3][0])}">(${fmt(x[3][0])} in ${esc(x[3][1])})</span>`:' <span class="small mut">(unchanged)</span>'):(x[7]?'<span class="mut small">residual (no '+esc(M.af_suffix)+' series)</span>':'')}</td></tr>`).join('');
  $('#amore').innerHTML=rs.length>lim?`<button class="btn" id="am2">Show all ${rs.length.toLocaleString()}</button>`:'';if(rs.length>lim)$('#am2').onclick=()=>{lim=rs.length;r()}};
 ['aq','am','as','amode','asl'].forEach(id=>$('#'+id).addEventListener('input',()=>{if(id!=='asl')lim=500;r()}));r()})();

// ---------- equations A vs B
(function(){const E=D.eqdiff;if(!E||!E.length)return;const c={changed:0,only_a:0,only_b:0};E.forEach(e=>c[e[1]]++);
 const L={changed:'changed',only_a:'only in A (B sets it directly)',only_b:'only in B (A sets it directly)'};
 $('#t-eqs').innerHTML=`<div class="kpis"><div class="kpi"><span class="mut small">equations changed</span><b>${c.changed}</b></div><div class="kpi"><span class="mut small">only in A</span><b>${c.only_a}</b></div><div class="kpi"><span class="mut small">only in B</span><b>${c.only_b}</b></div></div>
 <div class="hint">Variables whose equation is not the same in the two scenarios. Differences are highlighted. Spacing, upper/lower case and error-term options like [ar(1)=…] are ignored when comparing. Click a name for its full breakdown, where the switch appears as one effect.</div>
 <div class="ctl"><select id="eqf"><option value="">All</option><option value="changed">Changed</option><option value="only_a">Only in A</option><option value="only_b">Only in B</option></select><input type="text" id="eqq" placeholder="Search"></div><div id="eqlist"></div>`;
 const r=()=>{const f=$('#eqf').value,q=$('#eqq').value.trim().toUpperCase();$('#eqlist').innerHTML=E.filter(e=>(!f||e[1]===f)&&(!q||e[0].includes(q)||e[5].toUpperCase().includes(q))).slice(0,500).map(e=>`<div class="card"><b>${vlink(e[0])}</b> <span class="badge">${L[e[1]]}</span> ${e[4]!=null?`<span class="small ${cls(e[4])}">gap at end ${pct(e[4])}</span>`:'<span class="small mut">no difference in its values</span>'} <span class="small mut">${esc(e[5])}</span>${eqPair(e)}</div>`).join('')};
 $('#eqf').addEventListener('input',r);$('#eqq').addEventListener('input',r);r()})();

// ---------- importance
(function(){const I=D.imp;if(!I)return;
 const lab=id=>id.startsWith('AF:')?afText(id.slice(3)):id,lk=id=>id.startsWith('AF:')?afLink(id.slice(3)):vlink(id);
 $('#t-imp').innerHTML=`<div class="hint">Which shocks matter most? Each <b>shock</b> is a changed exogenous variable (no equation, or set by the modeller) or a changed add-factor. Each one is switched from A to B <i>on its own</i>, everything it feeds is traced through the equations, and for each of the ${I.n_targets.toLocaleString()} variables that moved we measure the share of its gap the shock explains (whole window: Σ effect·gap / Σ gap²). For each variable, the shares of all shocks add up to about 100%. <b>Importance</b> = the average of those shares across variables.</div>
 <div class="ctl"><label class="small mut">Weight variables <select id="iw"><option value="4">by concept (each concept counts once, so 50 state copies don't dominate)</option><option value="3">all equally</option>${I.shocks.some(r=>r[5]!=null)?'<option value="5">key variables only (Summary tab)</option>':''}</select></label><select id="it"><option value="">All shock types</option><option>exogenous</option><option>add-factor</option><option>equation changed</option><option>set by modeller</option></select><input type="text" id="iq" placeholder="Search"></div>
 ${Object.keys(I.geo||{}).length>1?`<div class="ctl"><label class="small mut">Variables to explain <select id="ig"><option value="">All geographies (${I.n_targets.toLocaleString()} variables)</option>${Object.entries(I.geo).sort((a,b)=>a[0]<b[0]?-1:1).map(([g,o])=>`<option value="${esc(g)}">${esc(g)} only (${o.n.toLocaleString()} variables)</option>`).join('')}</select></label><span class="small mut" id="ign"></span></div>`:''}
 <div id="ichart"></div><div class="tw"><table><thead><tr><th>#</th><th>Shock</th><th>Type</th><th>Description</th><th class="n">Importance</th><th class="n" title="Number of variables where this shock explains at least 10% of the gap">Explains ≥10% of</th><th class="n">≥50% of</th><th class="n">Variables it feeds</th><th>Explains most of</th></tr></thead><tbody id="itb"></tbody></table></div>
 <h2>Look up a variable</h2><div class="ctl"><input type="text" id="iv" placeholder="Variable, e.g. FGDP$_ICAN"><button class="btn" id="ivgo">Show its main shocks</button></div><div id="ivout"></div>`;
 const r=()=>{let c=+$('#iw').value;const ty=$('#it').value,q=$('#iq').value.trim().toUpperCase(),g=$('#ig')?$('#ig').value:'';
  let base=I.shocks;
  if(g){const G=I.geo[g],meta={};I.shocks.forEach(x=>meta[x[0]]=x);
   // geography rows: [id, w_all, w_concept, n10, n50, top]  ->  same layout as the full table
   base=G.rows.map(z=>{const m=meta[z[0]]||[z[0],z[0].startsWith('AF:')?'add-factor':'exogenous','',null,null,null,0,0,0,[]];return [z[0],m[1],m[2],z[1],z[2],null,z[3],z[4],m[8],z[5]]});
   if(c===5)c=4;$('#ign').textContent=`Importance = average share of the gaps of ${G.n.toLocaleString()} ${g} variables. National and other-country shocks count if they move ${g}.`}else if($('#ign'))$('#ign').textContent='';
  const rs=base.filter(x=>(!ty||x[1]===ty)&&(!q||x[0].includes(q)||x[2].toUpperCase().includes(q))).slice().sort((a,b)=>Math.abs(b[c]||0)-Math.abs(a[c]||0));
  const top=rs.slice(0,20),mx=Math.max(1e-9,...top.map(x=>Math.abs(x[c]||0)));
  $('#ichart').innerHTML=`<table class="bars">${top.map(x=>`<tr><td style="width:30%">${lk(x[0])}</td><td style="width:50%">${bar(x[c]||0,mx)}</td><td class="n"><b>${x[c]==null?'–':(100*x[c]).toFixed(1)+'%'}</b></td></tr>`).join('')}</table>`;
  $('#itb').innerHTML=rs.slice(0,1000).map((x,i)=>`<tr><td class="mut">${i+1}</td><td>${lk(x[0])}</td><td class="small">${esc(x[1])}</td><td class="small">${esc(x[2]).slice(0,50)}</td><td class="n"><b>${x[c]==null?'–':(100*x[c]).toFixed(1)+'%'}</b></td><td class="n">${x[6].toLocaleString()}</td><td class="n">${x[7].toLocaleString()}</td><td class="n mut">${x[8].toLocaleString()}</td><td class="small">${x[9].slice(0,3).map(t=>vlink(t[0])+' '+Math.round(100*t[1])+'%').join(', ')}</td></tr>`).join('')};
 ['iw','it','iq','ig'].forEach(id=>{const e=$('#'+id);if(e)e.addEventListener('input',r)});r();
 const lv=()=>{const v=$('#iv').value.trim().toUpperCase(),l=I.by_var[v];
  $('#ivout').innerHTML=!l?`<div class="card">${esc(v)}: no shares (it did not move, has no evaluable equation, or is itself a shock). ${sugg(v)}</div>`:`<div class="hint">Share of ${esc(v)}'s gap (whole window) explained by each shock on its own:</div><table class="bars">${l.map(t=>`<tr><td style="width:40%">${lk(t[0])}</td><td style="width:40%">${bar(t[1],Math.max(...l.map(z=>Math.abs(z[1]))))}</td><td class="n"><b>${Math.round(100*t[1])}%</b></td></tr>`).join('')}</table>`};
 $('#ivgo').onclick=lv;$('#iv').addEventListener('keydown',e=>{if(e.key==='Enter')lv()});
})();

// ---------- layers
(function(){const Y=D.layers;if(!Y)return;
 const N=D.nodes.length,AREA=Y.area?new Set(Y.area):null,AL=Y.label||'';
 const val=(i,a)=>a?Y.da[i]:Y.d[i];
 const chain=(n,a)=>{const nx=a?Y.nxa:Y.nx,out=[n];while(out.length<80){const v=nx[out[out.length-1]];if(!v||out.includes(v))break;out.push(v)}return out};
 const lpOf=n=>{const i=NI[n];return i==null?1:Y.lp[i]};
 const gapOf=n=>{const r=ROW[n];return r?r[7]:null};
 const inA=n=>AREA&&AREA.has(n);
 const chainHTML=(n,a)=>chain(n,a).map((v,k)=>`${k?'<span class="arrow">→</span>':''}<span class="step">${vlink(v)}${lpOf(v)>1?` <span class="badge" title="part of a loop of ${lpOf(v)} variables; the loop counts as one layer">loop ${lpOf(v)}</span>`:''}${inA(v)?' <span class="badge bE">area</span>':''}</span>`).join('');
 let h=`<div class="hint"><b>Layers</b> = the longest chain of equations from a variable downstream: 0 = no equation uses it; 3 = it feeds a variable that feeds a variable that feeds a variable. A loop (e.g. GDP → income → consumption → GDP) counts as <b>one</b> layer, since inside a loop the chain never ends. Lags count as links. This is the model's structure, not the size of an effect.${Y.pruned?` Counted on the full model before <code>--target</code> dropped variables.`:''}</div>`;
 h+=`<div class="ctl">${AREA?`<label>Measure <select id="lm"><option value="0">longest chain, anywhere</option><option value="1">longest chain ending in ${esc(AL)}</option></select></label>`:''} <input type="text" id="lq" placeholder="filter names, e.g. _ICAN or FGDP"> <label><input type="checkbox" id="lc"> changed only</label></div>
 <div id="lh"></div><div id="lch"></div><div class="tw"><table><thead><tr><th>Variable</th><th>Type</th><th class="n">Layers</th>${AREA?`<th class="n">Into ${esc(AL)}</th>`:''}<th class="n">Used directly by</th><th class="n">Gap at end</th><th>Longest chain</th></tr></thead><tbody id="lt"></tbody></table></div><div id="lmore"></div>`;
 if(AREA){h+=`<h2>Inputs into ${esc(AL)}</h2><div id="lin"></div>`}
 h+=`<h2>Layer maps</h2><div class="hint">The model in three maps by complexity. <b>Map 1</b> is the core: variables whose longest chain has the most layers. <b>Map 2</b> shows the next tier that Map 1 feeds directly, <b>Map 3</b> the tier after that, fed by Map 2. Dashed boxes on the left of a map are the variables from the map above that feed it. Click a box to narrow the next map to what that variable reaches (click again, or ✕, to clear); double-click for its details. In each map, columns are layers (most layers on the left), arrows are equations, colour is the gap at the end (green up, red down, grey unchanged). Scroll to zoom, drag to move.</div>
 <div class="ctl">Map 1: layers ≥ <input type="number" id="t1" min="2" style="width:60px"> &nbsp; Map 2: layers ≥ <input type="number" id="t2" min="1" style="width:60px"> &nbsp; Map 3: the rest (≥ 1) &nbsp; <label>max boxes per map <input type="number" id="mn" value="200" min="20" max="3000" step="50" style="width:80px"></label> <label><input type="checkbox" id="mlink" checked> only variables linked to the map above</label></div>
 ${[1,2,3].map(k=>`<h3 id="mh${k}"></h3><div class="ctl"><span id="mf${k}"></span> <span class="zb"><button class="btn" data-lz${k}="in">+</button><button class="btn" data-lz${k}="out">−</button><button class="btn" data-lz${k}="reset">reset</button></span></div><div id="mnote${k}" class="hint"></div><div id="lmap${k}"></div>`).join('')}`;
 $('#t-lay').innerHTML=h;
 let LIM=300;
 const meas=()=>AREA&&$('#lm')&&$('#lm').value==='1';
 function hist(a){const c={};let none=0;for(let i=0;i<N;i++){const v=val(i,a);if(v==null){none++;continue}c[v]=(c[v]||0)+1}
  const ks=Object.keys(c).map(Number).sort((x,y)=>x-y),mx=Math.max(1,...Object.values(c));
  return `<div class="card"><b>How many variables reach each number of layers</b>${a?` (chains ending in ${esc(AL)})`:''}<table class="bars">${ks.map(k=>`<tr class="cl" data-lk="${k}"><td class="n">${k}</td><td style="width:70%"><div class="bar" style="width:${Math.max(1,100*c[k]/mx)}%;background:var(--a)"></div></td><td class="n">${c[k].toLocaleString()}</td></tr>`).join('')}</table>${a&&none?`<div class="hint">${none.toLocaleString()} variables cannot reach ${esc(AL)}.</div>`:''}<div class="hint">Click a row to list those variables.</div></div>`}
 let only=null;
 function table(){const a=meas(),q=$('#lq').value.trim().toUpperCase(),ch=$('#lc').checked;
  let ids=[];for(let i=0;i<N;i++){const v=val(i,a);if(v==null)continue;const n=D.nodes[i];if(q&&!n.includes(q))continue;if(ch&&!ROW[n])continue;if(only!=null&&v!==only)continue;ids.push(i)}
  ids.sort((x,y)=>val(y,a)-val(x,a)||CH[y].length-CH[x].length||(D.nodes[x]<D.nodes[y]?-1:1));
  $('#lt').innerHTML=ids.slice(0,LIM).map(i=>{const n=D.nodes[i],c=chain(n,a),g=gapOf(n),r=ROW[n];return `<tr class="cl" data-lr="${esc(n)}"><td>${vlink(n)}</td><td>${badge(r?r[9]:(M.has_eqs&&D.par[i].length?'E':'X'),n)}</td><td class="n"><b>${Y.d[i]}</b></td>${AREA?`<td class="n">${Y.da[i]==null?'–':Y.da[i]}</td>`:''}<td class="n">${CH[i].length}</td><td class="n ${cls(g)}">${g==null?'–':pct(g)}</td><td class="small">${c.slice(1,5).map(esc).join(' → ')}${c.length>5?' → …':''}</td></tr>`}).join('');
  $('#lmore').innerHTML=`<div class="hint">${only!=null?`Only variables with ${only} layers. <a href="#" id="lall">show all</a>. `:''}Showing ${Math.min(LIM,ids.length).toLocaleString()} of ${ids.length.toLocaleString()}${ids.length>LIM?` <button class="btn" id="lmb">show 500 more</button>`:''}. Click a row for its full chain.</div>`;
  const b=$('#lmb');if(b)b.onclick=()=>{LIM+=500;table()};const al=$('#lall');if(al)al.onclick=e=>{e.preventDefault();only=null;table()}}
 function refresh(){$('#lh').innerHTML=hist(meas());$('#lch').innerHTML='';table()}
 $('#t-lay').addEventListener('click',e=>{if(e.target.closest('a'))return;const r=e.target.closest('[data-lr]');if(r){const n=r.dataset.lr,a=meas(),i=NI[n];
   $('#lch').innerHTML=`<div class="card"><b>${esc(n)}</b>: ${val(i,a)} layers${a?` into ${esc(AL)}`:''}. Longest chain:<div style="margin-top:6px">${chainHTML(n,a)}</div>${!a&&AREA&&Y.da[i]!=null&&Y.da[i]!==Y.d[i]?`<div class="hint" style="margin-top:6px">Longest chain ending in ${esc(AL)} (${Y.da[i]} layers): ${chainHTML(n,true)}</div>`:''}</div>`;$('#lch').scrollIntoView({block:'nearest'});return}
  const k=e.target.closest('[data-lk]');if(k){only=+k.dataset.lk;LIM=300;table()}});
 ['#lq','#lc'].forEach(s=>$(s).addEventListener('input',()=>{LIM=300;table()}));if($('#lm'))$('#lm').addEventListener('input',()=>{only=null;LIM=300;refresh();FOC[1]=FOC[2]=null;drawMaps()});
 refresh();
 // inputs into the area: variables outside it that its equations use directly
 if(AREA){const feeds={};Y.area.forEach(t=>{const i=NI[t];if(i==null)return;D.par[i].forEach(p=>{const pn=D.nodes[p];if(!AREA.has(pn))(feeds[pn]=feeds[pn]||[]).push(t)})});
  const ks=Object.keys(feeds).sort((x,y)=>feeds[y].length-feeds[x].length||(x<y?-1:1)),fed=new Set(Object.values(feeds).flat());
  $('#lin').innerHTML=`<div class="hint"><b>${ks.length.toLocaleString()}</b> variables outside ${esc(AL)} are used directly in its equations; ${fed.size.toLocaleString()} of its ${Y.area.length.toLocaleString()} variables use at least one. Everything else outside reaches ${esc(AL)} only through these.</div>
  <div class="tw"><table><thead><tr><th>Input</th><th>Type</th><th class="n">Gap at end</th><th class="n">${esc(AL)} variables using it</th><th>Which</th></tr></thead><tbody>${ks.map(p=>{const r=ROW[p],g=gapOf(p),ts=feeds[p].sort();return `<tr><td>${vlink(p)}</td><td>${badge(r?r[9]:(D.par[NI[p]].length?'E':'X'),p)}</td><td class="n ${cls(g)}">${g==null?'unchanged':pct(g)}</td><td class="n">${ts.length}</td><td class="small">${ts.slice(0,12).map(vlink).join(', ')}${ts.length>12?` … and ${ts.length-12} more`:''}</td></tr>`}).join('')}</tbody></table></div>`}
 // layer maps in three tiers
 const LVS={1:{},2:{},3:{}},FOC={1:null,2:null};
 const mxL=Math.max(1,Y.max);$('#t1').value=mxL>=6?5:Math.max(2,Math.ceil(mxL*0.6));$('#t2').value=+$('#t1').value>=5?3:Math.max(1,+$('#t1').value-1);
 const col=g=>{if(g==null)return'var(--card)';const t=Math.min(1,Math.abs(g)/3);return g<0?`rgba(217,45,32,${.12+.55*t})`:`rgba(15,157,88,${.12+.55*t})`};
 const rank=a=>(x,y)=>val(y,a)-val(x,a)||Math.abs(gapOf(D.nodes[y])||0)-Math.abs(gapOf(D.nodes[x])||0)||CH[y].length-CH[x].length;
 const RC=new Map();function reach(start){if(RC.has(start))return RC.get(start);const seen=new Set(),q=[start];while(q.length){const u=q.pop();for(const v of CH[u])if(!seen.has(v)){seen.add(v);q.push(v)}}RC.set(start,seen);return seen}
 function draw(k,ids,ghosts,a,note){const el=$('#lmap'+k);$('#mnote'+k).textContent=note;if(!ids.length){el.innerHTML='';return}
  const sel=new Set(ids),gs=new Set(ghosts),all=new Set([...ids,...ghosts]);
  const lay=i=>gs.has(i)?Infinity:val(i,a),cols=new Map();[...ghosts,...ids].forEach(i=>{const L=lay(i);if(!cols.has(L))cols.set(L,[]);cols.get(L).push(i)});
  const ks=[...cols.keys()].sort((x,y)=>y-x),pos={};ks.forEach(L=>cols.get(L).forEach((i,r)=>pos[i]=r));
  const inS=i=>D.par[i].filter(p=>all.has(p));
  for(let sw=0;sw<3;sw++)ks.forEach((L,j)=>{if(!j)return;const c=cols.get(L);c.sort((x,y)=>{const m=z=>{const ps=inS(z);return ps.length?ps.reduce((s,p)=>s+pos[p],0)/ps.length:1e9};return m(x)-m(y)});c.forEach((i,r)=>pos[i]=r)});
  const BW=150,BH=26,GX=90,SX=14,GY=8,R=Math.max(12,Math.min(40,Math.ceil(Math.sqrt(all.size)*1.6)));
  const nsub={},x0={};let xc=10;ks.forEach(L=>{nsub[L]=Math.max(1,Math.ceil(cols.get(L).length/R));x0[L]=xc;xc+=nsub[L]*(BW+SX)-SX+GX});
  const rows=L=>Math.ceil(cols.get(L).length/nsub[L]),mr=Math.max(...ks.map(rows)),W=xc-GX+20,H=mr*(BH+GY)+20;
  const P={};ks.forEach(L=>{const c=cols.get(L),rk=rows(L),off=(H-rk*(BH+GY))/2+6;c.forEach((i,r)=>P[i]={x:x0[L]+Math.floor(r/rk)*(BW+SX),y:off+(r%rk)*(BH+GY)})});
  let ed='';ids.forEach(v=>inS(v).forEach(u=>{const A=P[u],B=P[v];if(!A||!B)return;let d;
   if(A.x<B.x){const x1=A.x+BW,y1=A.y+BH/2,x2=B.x,y2=B.y+BH/2,m=(x1+x2)/2;d=`M${x1},${y1} C${m},${y1} ${m},${y2} ${x2},${y2}`}
   else{const x1=A.x+BW,y1=A.y+BH/2,x2=B.x+BW,y2=B.y+BH/2,bx=Math.max(x1,x2)+30;d=`M${x1},${y1} C${bx},${y1} ${bx},${y2} ${x2},${y2}`}
   ed+=`<path d="${d}" fill="none" stroke="var(--mut)" stroke-opacity=".35" stroke-width="1" marker-end="url(#la${k})" data-e="${u},${v}"${gs.has(u)?' stroke-dasharray="4 3"':''}/>`}));
  const foc=FOC[k];let nd='';[...all].forEach(i=>{const n=D.nodes[i],p=P[i],g=gapOf(n),gh=gs.has(i),f=foc===i;
   nd+=`<g class="ln" data-n="${esc(n)}" data-i="${i}" style="cursor:pointer"><title>${esc(n)}: ${Y.d[i]} layers${AREA&&Y.da[i]!=null?`, ${Y.da[i]} into ${AL}`:''}${g!=null?`, gap ${pct(g)}`:', unchanged'}${Y.lp[i]>1?`, in a loop of ${Y.lp[i]}`:''}${gh?` (from Map ${k-1})`:''}</title><rect x="${p.x}" y="${p.y}" width="${BW}" height="${BH}" rx="5" fill="${col(g)}" fill-opacity="${gh?.45:1}" stroke="${f?'var(--b)':inA(n)?'var(--a)':'var(--mut)'}" stroke-opacity="${f||inA(n)?1:.4}" stroke-width="${f?3:inA(n)?2:1}"${gh?' stroke-dasharray="4 3"':''}/><text x="${p.x+6}" y="${p.y+17}" font-size="11" fill="var(--fg)">${esc(n.length>22?n.slice(0,21)+'…':n)}</text></g>`});
  const hd=ks.map(L=>`<text x="${x0[L]}" y="12" font-size="11" fill="var(--mut)">${L===Infinity?`from Map ${k-1} (${cols.get(L).length})`:`${L} layer${L===1?'':'s'} (${cols.get(L).length})`}</text>`).join('');
  const S=LVS[k];S.full=[0,-6,W,H+20];S.v=S.full.slice();
  el.innerHTML=`<svg id="lsvg${k}" viewBox="${S.v.join(' ')}" style="width:${Math.min(($('#t-lay').clientWidth||1100),Math.round(W*1.15))}px;max-width:100%;height:${Math.round(Math.min(680,Math.max(120,Math.min(($('#t-lay').clientWidth||1100),W*1.15)*(H+20)/W)))}px;border:1px solid var(--line);border-radius:10px;cursor:grab;touch-action:none"><defs><marker id="la${k}" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="6" markerHeight="6" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="var(--mut)"/></marker></defs>${hd}${ed}${nd}</svg>`;
  const sv=$('#lsvg'+k);panZoom(sv,S,'[data-lz'+k+']','lz'+k);
  sv.addEventListener('mouseover',e=>{const g=e.target.closest('.ln');if(!g)return;const i=g.dataset.i;sv.querySelectorAll('path[data-e]').forEach(pth=>{const [u,v]=pth.dataset.e.split(',');const on=u===i||v===i;pth.setAttribute('stroke',on?'var(--a)':'var(--mut)');pth.setAttribute('stroke-opacity',on?'.9':'.35');pth.setAttribute('stroke-width',on?'2':'1')})});
  sv.addEventListener('click',e=>{const g=e.target.closest('.ln');if(!g||k===3||gs.has(+g.dataset.i))return;const i=+g.dataset.i;FOC[k]=FOC[k]===i?null:i;if(k===1)FOC[2]=null;drawMaps()});
  sv.addEventListener('dblclick',e=>{const g=e.target.closest('.ln');if(g)show(g.dataset.n)})}
 function drawMaps(){const a=meas(),cap=Math.max(20,+$('#mn').value||200),link=$('#mlink').checked;
  let t1=Math.max(2,+$('#t1').value||2),t2=Math.max(1,Math.min(t1-1,+$('#t2').value||1));
  const tier=(lo,hi)=>{const out=[];for(let i=0;i<N;i++){const v=val(i,a);if(v!=null&&v>=lo&&v<hi)out.push(i)}return out};
  const T=[null,tier(t1,Infinity),tier(t2,t1),tier(1,t2)],shown={};
  const label=['',`Map 1 — core: ${t1}+ layers`,`Map 2 — ${t2===t1-1?t2:`${t2}–${t1-1}`} layers`,`Map 3 — ${t2-1<=1?1:`1–${t2-1}`} layer${t2-1===1?'':'s'}`];
  for(const k of [1,2,3]){$('#mh'+k).textContent=label[k];let ids=T[k],ghosts=[],extra='';
   if(k>1){const up=new Set(T[k-1]),f=FOC[k-1]!=null?FOC[k-1]:(k===3?FOC[1]:null);
    if(f!=null){const r=reach(f);ids=ids.filter(i=>r.has(i));extra=` reached from ${D.nodes[f]}`}
    else if(link){ids=ids.filter(i=>D.par[i].some(p=>up.has(p)))}
    const cand=ids.slice().sort(rank(a)).slice(0,cap),gc=new Map();
    cand.forEach(i=>D.par[i].forEach(p=>{if(up.has(p))gc.set(p,(gc.get(p)||0)+1)}));
    ghosts=[...gc.keys()].sort((x,y)=>gc.get(y)-gc.get(x)).slice(0,40);
    if(FOC[k-1]!=null&&!ghosts.includes(FOC[k-1]))ghosts.unshift(FOC[k-1])}
   const total=ids.length;ids=ids.slice().sort(rank(a)).slice(0,cap);shown[k]=ids;
   {const fk=k>1?(FOC[k-1]!=null?k-1:(k===3&&FOC[1]!=null?1:null)):null;$('#mf'+k).innerHTML=fk!=null?`<span class="badge" style="cursor:pointer" data-cf="${fk}">narrowed to ${esc(D.nodes[FOC[fk]])} ✕</span>`:''}
   draw(k,ids,ghosts,a,total?`${total.toLocaleString()} variables${extra}${k>1&&!extra&&link?` linked to Map ${k-1}`:''}${total>cap?`; showing the ${cap} with the most layers and largest gaps`:''}.`:`No variables${extra} in this tier.`)}}
 $('#t-lay').addEventListener('click',e=>{const c=e.target.closest('[data-cf]');if(c){FOC[+c.dataset.cf]=null;if(c.dataset.cf==='1')FOC[2]=null;drawMaps()}});
 ['#t1','#t2','#mn','#mlink'].forEach(s=>$(s).addEventListener('input',()=>{FOC[1]=FOC[2]=null;drawMaps()}));
 if($('#lm'))$('#lm').addEventListener('input',()=>RC.clear());
 drawMaps();
})();
function panZoom(sv,S,btnSel,attr){attr=attr||'lz';if(!sv)return;const set=()=>sv.setAttribute('viewBox',S.v.join(' '));
 const zoomAt=(f,cx,cy)=>{const [x,y,w,h]=S.v,nw=Math.min(S.full[2]*2,Math.max(S.full[2]/40,w*f)),k=nw/w;S.v=[cx-(cx-x)*k,cy-(cy-y)*k,nw,h*k];set()};
 const toSvg=(ex,ey)=>{const r=sv.getBoundingClientRect(),[x,y,w,h]=S.v,sc=Math.max(w/r.width,h/r.height),ox=(r.width*sc-w)/2,oy=(r.height*sc-h)/2;return[x-ox+(ex-r.left)*sc,y-oy+(ey-r.top)*sc,sc]};
 sv.addEventListener('wheel',e=>{e.preventDefault();const [cx,cy]=toSvg(e.clientX,e.clientY);zoomAt(e.deltaY>0?1.15:1/1.15,cx,cy)},{passive:false});
 let drag=null,moved=false;
 sv.addEventListener('pointerdown',e=>{drag=[e.clientX,e.clientY,S.v.slice()];moved=false});
 window.addEventListener('pointermove',e=>{if(!drag)return;const sc=toSvg(0,0)[2],dx=(e.clientX-drag[0])*sc,dy=(e.clientY-drag[1])*sc;if(Math.abs(dx)+Math.abs(dy)>2)moved=true;S.v=[drag[2][0]-dx,drag[2][1]-dy,drag[2][2],drag[2][3]];set()});
 window.addEventListener('pointerup',()=>{drag=null});
 sv.addEventListener('click',e=>{if(moved){e.stopPropagation();e.preventDefault();moved=false}},true);
 document.querySelectorAll(btnSel).forEach(b=>b.onclick=()=>{const [x,y,w,h]=S.v,z=b.dataset[attr];if(z==='reset'){S.v=S.full.slice();set()}else zoomAt(z==='in'?1/1.3:1.3,x+w/2,y+h/2)})}
// ---------- influence
function infSummary(r,what){return `<div class="kpis"><div class="kpi"><span class="mut small">linked through equations</span><b>${r.reach.toLocaleString()}</b></div><div class="kpi"><span class="mut small">actually move (≥${r.thr}%)</span><b>${r.moved.toLocaleString()}</b></div><div class="kpi"><span class="mut small">move ≥ 1%</span><b>${r.ge1.toLocaleString()}</b></div><div class="kpi"><span class="mut small">0.1–1%</span><b>${r.b01.toLocaleString()}</b></div><div class="kpi"><span class="mut small">${r.thr}–0.1%</span><b>${r.b001.toLocaleString()}</b></div></div>
 <div class="hint">${what}. ${r.resolved.toLocaleString()} equations re-solved${r.blocked?`; ${r.blocked} linked variables have no evaluable equation and were held fixed`:''}. Largest effect over ${esc(M.ws)}–${esc(M.we)}, as % of the scenario-A level:</div>
 <div class="tw"><table><thead><tr><th>Variable</th><th>Description</th><th class="n">Largest effect</th><th class="n">%</th><th>When</th></tr></thead><tbody>${r.top.map(t=>`<tr><td>${vlink(t[0])}</td><td class="small">${esc(desc(t[0])).slice(0,60)}</td><td class="n ${cls(t[1])}">${fmt(t[1])}</td><td class="n ${cls(t[1])}">${t[2]==null?'–':(t[1]<0?'−':'+')+t[2].toFixed(2)+'%'}</td><td>${esc(t[3])}</td></tr>`).join('')}</tbody></table></div>`}
(function(){if(!M.has_eqs)return;
 const sh=Object.values(D.infl);
 let h=`<div class="hint">If one variable changes, how many others move, and by how much? <b>Linked</b> = variables that use it directly or indirectly in their equations. <b>Move</b> = the change is re-run through the equations with everything else held at scenario A, and a variable counts if it moves by at least 0.01% of its level.</div>
 <div class="ctl"><input type="text" id="ix" placeholder="Any variable, e.g. FOIL">${SERVE?'<select id="is"><option value="">its change from A to B</option><option value="1%">+1%</option><option value="-1%">−1%</option><option value="1">+1 unit</option><option value="0.25">+0.25 unit</option></select>':''}<button class="btn" id="igo">Show</button></div><div id="iout"></div>`;
 if(sh.length){h+=`<h2>Each scenario shock on its own</h2><div class="tw"><table><thead><tr><th>Shock</th><th>Type</th><th class="n">Linked</th><th class="n">Move</th><th class="n">≥1%</th><th class="n">0.1–1%</th><th class="n">&lt;0.1%</th><th>Most affected</th></tr></thead><tbody>${sh.map(r=>`<tr class="cl" data-inf="${esc(r.x)}"><td>${vlink(r.x)}${r.mode==='addfactor'&&r.target!==r.x?' <span class="small mut">→ '+esc(r.target)+'</span>':r.mode==='addfactor'?' <span class="small mut">(its add-factor)</span>':''}</td><td>${badge(ROW[r.x]?ROW[r.x][9]:'')}</td><td class="n">${r.reach.toLocaleString()}</td><td class="n"><b>${r.moved.toLocaleString()}</b></td><td class="n">${r.ge1}</td><td class="n">${r.b01}</td><td class="n">${r.b001}</td><td class="small">${r.top.slice(0,3).map(t=>esc(t[0])+' '+(t[2]==null?'':(t[1]<0?'−':'+')+t[2].toFixed(2)+'%')).join(', ')}</td></tr>`).join('')}</tbody></table></div><div class="hint">Click a row for the full list.</div>`}
 $('#t-inf').innerHTML=h;
 $('#t-inf').addEventListener('click',e=>{const r=e.target.closest('[data-inf]');if(r&&!e.target.closest('a')){$('#ix').value=r.dataset.inf;go()}});
 const go=()=>{const x=$('#ix').value.trim().toUpperCase(),i=NI[x];if(i==null){$('#iout').innerHTML=`<div class="card">${esc(x)} not found. ${sugg(x)}</div>`;return}
  const shock=SERVE?$('#is').value:'',afT=D.infl[x]&&D.infl[x].mode==='addfactor'?D.infl[x].target:null,direct=CH[i].length;
  let h=`<div class="card"><b>${esc(x)}</b> ${afT&&afT!==x?'is the add-factor of '+vlink(afT)+'; it':afT?'is moved by its own add-factor; it is used directly in <b>'+direct+'</b> other equation(s) and':'is used directly in <b>'+direct+'</b> equation(s) and'} is linked to <b>${(afT?D.infl[x].reach:D.reach[i]).toLocaleString()}</b> variables in total${afT?' (through its add-factor, including '+esc(afT)+' itself)':''}.${direct||afT?'':' It is an end point: changing it moves nothing else.'}</div>`;
  if(!direct&&!afT){$('#iout').innerHTML=h;return}
  if(!shock&&D.infl[x]){const r=D.infl[x];$('#iout').innerHTML=h+infSummary(r,r.mode==='addfactor'?`Effect of the add-factor change of ${esc(r.target)} (from A to B)`:`Effect of ${esc(x)}'s own change from A to B`);return}
  if(!SERVE){$('#iout').innerHTML=h+`<div class="hint">To see how much the linked variables move, run <code>python scenario_explain.py serve …</code> (any variable, any shock) or <code>influence ${esc(x)}</code> in the terminal.</div>`;return}
  $('#iout').innerHTML=h+'<div class="hint">Re-solving the equations…</div>';
  fetch(`/api/influence?x=${encodeURIComponent(x)}&shock=${encodeURIComponent(shock)}`).then(r=>r.json()).then(o=>{
   $('#iout').innerHTML=h+(o.error?`<div class="card">${esc(o.error)}</div>`:infSummary(o,shock?`Effect of a ${esc(shock)}${shock.endsWith('%')?'':' unit'} change in ${esc(x)}`:`Effect of ${esc(x)}'s own change from A to B`))})};
 $('#igo').onclick=go;$('#ix').addEventListener('keydown',e=>{if(e.key==='Enter')go()});
})();

// ---------- by concept / geography
(function(){
 const note=M.has_eqs?`<div class="hint">A series counts as <b>changed</b> only if the scenario changed it directly: an assumption, or its add-factor changed. Series with the same equation and the same add-factor that moved only because their inputs moved are counted under <b>via formula only</b>. Click a row for the full list.</div>`:'<div class="hint">Click a row for the full list.</div>';
 const esc2=t=>t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
 if(D.stem.length){$('#t-stem').innerHTML=note+`<div class="tw"><table><thead><tr><th>Concept</th><th>Description</th><th class="n">Geos changed</th>${M.has_eqs?'<th class="n">Via formula only</th><th class="n">Not evaluable</th>':''}<th class="n">of</th><th class="n">Median |gap at end %| (changed)</th><th>Largest</th></tr></thead><tbody>${D.stem.slice(0,3000).map(r=>`<tr class="cl" data-stem="${esc(r[0])}"><td><b>${esc(r[0])}</b></td><td class="small">${esc(r[8]).slice(0,70)}</td><td class="n"><b>${r[1]}</b></td>${M.has_eqs?`<td class="n mut">${r[2]}</td><td class="n mut">${r[3]||''}</td>`:''}<td class="n mut">${r[4]}</td><td class="n">${r[5]==null?'–':r[5].toFixed(2)+'%'}</td><td class="small">${r[7].map(vlink).join(', ')}</td></tr>`).join('')}</tbody></table></div>`;
  $('#t-stem').addEventListener('click',e=>{const t=e.target.closest('[data-stem]');if(t&&!e.target.closest('a'))KC.set({q:'^'+esc2(t.dataset.stem)+esc2(D.meta.geo_sep||'.')})})}
 if(D.geo.length){$('#t-geo').innerHTML=note+`<div class="tw"><table><thead><tr><th>Geography</th><th class="n">Series changed</th>${M.has_eqs?'<th class="n">Via formula only</th><th class="n">Not evaluable</th>':''}<th class="n">of</th><th class="n">Median score (changed)</th><th>Largest changes</th></tr></thead><tbody>${D.geo.map(r=>`<tr class="cl" data-geo="${esc(r[0])}"><td><b>${esc(r[0]||'(none)')}</b></td><td class="n"><b>${r[1]}</b></td>${M.has_eqs?`<td class="n mut">${r[2]}</td><td class="n mut">${r[3]||''}</td>`:''}<td class="n mut">${r[4]}</td><td class="n">${fmt(r[5])}</td><td class="small">${r[6].map(vlink).join(', ')}</td></tr>`).join('')}</tbody></table></div>`;
  $('#t-geo').addEventListener('click',e=>{const t=e.target.closest('[data-geo]');if(t&&!e.target.closest('a'))KC.set({g:t.dataset.geo||'__none__'})})}
})();

// ---------- path tracing
function esh(child,parent){const l=D.esh[NI[child]];if(!l)return null;const p=NI[parent];for(const e of l)if(e[0]===p)return e[1];return null}
function findPaths(x,y,k){const xi=NI[x],yi=NI[y];if(xi==null||yi==null)return[];const dist=new Map([[xi,0]]),q=[xi];
 while(q.length){const u=q.shift();if(u===yi)continue;for(const v of CH[u])if(!dist.has(v)){dist.set(v,dist.get(u)+1);q.push(v)}}
 if(!dist.has(yi))return[];const out=[];(function back(n,acc){if(out.length>=k)return;if(n===xi){out.push([xi,...acc.slice().reverse()]);return}for(const p of D.par[n])if(dist.get(p)===dist.get(n)-1)back(p,[...acc,n])})(yi,[]);return out.map(p=>p.map(i=>D.nodes[i]))}
function strongest(x,y){const xi=NI[x],yi=NI[y];if(xi==null||yi==null||xi===yi)return null;
 const best=new Map([[xi,0]]),prev=new Map(),hops=new Map([[xi,0]]),h=[[0,xi]];
 const push=e=>{h.push(e);let i=h.length-1;while(i>0){const p=(i-1)>>1;if(h[p][0]<=h[i][0])break;[h[p],h[i]]=[h[i],h[p]];i=p}};
 const pop=()=>{const t=h[0],l=h.pop();if(h.length){h[0]=l;let i=0;for(;;){let a=2*i+1,b=a+1,m=i;if(a<h.length&&h[a][0]<h[m][0])m=a;if(b<h.length&&h[b][0]<h[m][0])m=b;if(m===i)break;[h[m],h[i]]=[h[i],h[m]];i=m}}return t};
 while(h.length){const [c,u]=pop();if(c>(best.get(u)??Infinity))continue;if(u===yi)break;if(hops.get(u)>=15)continue;
  for(const v of CH[u]){if(v!==yi&&!ROW[D.nodes[v]])continue;const sh=esh(D.nodes[v],D.nodes[u]);const w=sh?-Math.log(Math.min(1,Math.abs(sh))):7;
   if(c+w<(best.get(v)??Infinity)){best.set(v,c+w);prev.set(v,u);hops.set(v,hops.get(u)+1);push([c+w,v])}}}
 if(!best.has(yi))return null;const p=[yi];while(p[p.length-1]!==xi)p.push(prev.get(p[p.length-1]));return p.reverse().map(i=>D.nodes[i])}
function prodOf(p){let pr=1;for(let j=1;j<p.length;j++){const sh=esh(p[j],p[j-1]);if(sh==null)return null;pr*=sh}return pr}
function rankedPaths(x,y,k){const ps=findPaths(x,y,40),sp=strongest(x,y);if(sp&&!ps.some(p=>p.join()===sp.join()))ps.push(sp);
 ps.sort((a,b)=>{const pa=prodOf(a),pb=prodOf(b);return (pb==null?-1:Math.abs(pb))-(pa==null?-1:Math.abs(pa))});return ps.slice(0,k)}
function gapTxt(n){const r=ROW[n];if(!r)return'<span class="mut">no change</span>';if(r[7]==null)return `<span class="${cls(r[4])} small">max gap ${fmt(r[4])} (${esc(r[8])})</span>`;return `<span class="${cls(r[7])}">${pct(r[7])}</span>`}
function eshq(child,parent,i){const l=D.eshp[NI[child]];if(!l)return null;const p=NI[parent];for(const e of l)if(e[0]===p)return e[1][i]==null?null:e[1][i]/1000;return null}
function gapQ(n,i){const g=D.pg[n];if(!g)return'<span class="mut">no change</span>';const v=g[i];return v==null?'<span class="mut">–</span>':`<span class="${cls(v)}">${pct(v)}</span>`}
let PATHS=null;
function pathsQ(i){const {x,y,ps}=PATHS,WP=D.per.length-D.wsi;
 const prodQ=(p,k)=>{let pr=1;for(let j=1;j<p.length;j++){const sh=eshq(p[j],p[j-1],k);if(sh==null)return null;pr*=sh}return pr};
 const ord=ps.map(p=>[p,prodQ(p,i)]).sort((a,b)=>(b[1]==null?-1:Math.abs(b[1]))-(a[1]==null?-1:Math.abs(a[1])));
 return ord.map(([p,prod],k)=>{const st=p.map((n,j)=>{let s='';if(j>0){const sh=eshq(n,p[j-1],i);s=sh==null?'<span class="mut small">share n/a (no gap this quarter)</span>':`<span class="small">${Math.round(sh*100)}% of its gap from ${esc(p[j-1])}</span>`}return `<span class="step">${j?'<span class="arrow">→</span>':''}${vlink(n)} ${gapQ(n,i)} ${s}</span>`}).join('<br>');
  const series=[...Array(WP).keys()].map(q=>{const v=prodQ(p,q);return v==null?null:100*v});
  return `<div class="path"><b>Path ${k+1}</b>${k===0&&ord.length>1?' <span class="small mut">(ordered by how much of the change they carry in this quarter)</span>':''}<br>${st}${p.length>1?`<div class="hint">${prod==null?'Share along this channel not available in this quarter.':`≈ <b>${Math.round(prod*100)}%</b> of ${esc(y)}'s change at ${esc(D.per[D.wsi+i])} travels along this channel (product of the link shares, approximate).`}</div>`+lineChart2(series,null,D.per.slice(D.wsi),'share of '+y+'\'s gap along this channel, %','',i):''}</div>`}).join('')}
function pathHTML(x,y){
 const ps=rankedPaths(x,y,6);
 if(!ps.length){const rev=findPaths(y,x,1);return `<div class="card">${esc(x)} does not feed into ${esc(y)} through the loaded equations.${rev.length?` The reverse holds: ${rev[0].map(vlink).join(' → ')}.`:''}</div>`}
 if(D.eshp&&Object.keys(D.eshp).length){const WP=D.per.length-D.wsi;PATHS={x,y,ps};
  setTimeout(()=>{const sl=$('#csl');if(!sl)return;const f=()=>{$('#clab').textContent=D.per[D.wsi+(+sl.value)];$('#cbody').innerHTML=pathsQ(+sl.value)};sl.addEventListener('input',f);f()},0);
  return `<div class="ctl" style="align-items:center"><span class="small mut">Quarter</span><input type="range" id="csl" min="0" max="${WP-1}" value="${WP-1}" style="flex:1;min-width:200px"><b id="clab"></b></div><div id="cbody"></div>`}
 return ps.map((p,i)=>{let prod=1,ok=true;const st=p.map((n,j)=>{let s='';if(j>0){const sh=esh(n,p[j-1]);if(sh==null){ok=false;s='<span class="mut small">share n/a</span>'}else{prod*=sh;s=`<span class="small">${Math.round(sh*100)}% of its gap from ${esc(p[j-1])}</span>`}}return `<span class="step">${j?'<span class="arrow">→</span>':''}${vlink(n)} ${gapTxt(n)} ${s}</span>`}).join('<br>');
  return `<div class="path"><b>Path ${i+1}</b>${i===0&&ps.length>1?' <span class="small mut">(paths ordered by how much of the change they carry)</span>':''}<br>${st}${ok&&p.length>1?`<div class="hint">≈ ${Math.round(prod*100)}% of ${esc(y)}'s change travels along this channel (product of the link shares, approximate).</div>`:''}</div>`}).join('')}
function upHTML(x,depth){const xi=NI[x];const dist=new Map([[xi,0]]),nxt=new Map(),q=[xi];
 while(q.length){const u=q.shift();if(dist.get(u)>=depth)continue;for(const p of D.par[u])if(!dist.has(p)){dist.set(p,dist.get(u)+1);nxt.set(p,u);q.push(p)}}
 const rs=[...dist.keys()].filter(i=>i!==xi&&ROW[D.nodes[i]]).map(i=>[D.nodes[i],dist.get(i),D.nodes[nxt.get(i)]]);
 [...dist.keys()].forEach(i=>{const a=D.afsh[i];if(a!=null&&Math.abs(a)>=.005)rs.push(['__AF__'+D.nodes[i],dist.get(i)+1,D.nodes[i],a])});
 rs.sort((a,b)=>a[1]-b[1]||Math.abs(b[3]??esh(b[2],b[0])??0)-Math.abs(a[3]??esh(a[2],a[0])??0));
 if(!D.par[xi].length)return `<div class="hint">${esc(x)} has no equation either — it is an assumption set directly by the scenario.</div>`;
 if(!rs.length)return `<div class="hint">None of the variables driving ${esc(x)} (within ${depth} steps) differ between A and B.</div>`;
 return `<div class="tw"><table><thead><tr><th>Driver</th><th class="n">Steps back</th><th>Feeds</th><th class="n">Gap at end %</th><th class="n">Share of that variable's gap</th></tr></thead><tbody>${rs.slice(0,300).map(r=>{if(r[0].startsWith('__AF__')){const t=r[2],af=D.afmap[t];return `<tr><td><i>add-factor</i> ${af?vlink(af):'<span class="small mut">(residual)</span>'}</td><td class="n">${r[1]}</td><td>${vlink(t)}</td><td class="n">${af?gapTxt(af):''}</td><td class="n"><b>${Math.round(r[3]*100)}%</b></td></tr>`}const sh=esh(r[2],r[0]);return `<tr><td>${vlink(r[0])}</td><td class="n">${r[1]}</td><td>${vlink(r[2])}</td><td class="n">${gapTxt(r[0])}</td><td class="n">${sh==null?'–':Math.round(sh*100)+'%'}</td></tr>`}).join('')}</tbody></table></div>`}
function downHTML(x,depth){const xi=NI[x];if(xi==null)return'';
 if(!CH[xi].length)return `<div class="card"><b>${esc(x)} is an end point of the model:</b> no other equation uses it, so a change in it does not flow anywhere. (A lag of itself in its own equation doesn't count.)</div><h2>What drives ${esc(x)} instead</h2>`+upHTML(x,depth);const dist=new Map([[xi,0]]),prev=new Map(),q=[xi];
 while(q.length){const u=q.shift();if(dist.get(u)>=depth)continue;for(const v of CH[u])if(!dist.has(v)){dist.set(v,dist.get(u)+1);prev.set(v,u);q.push(v)}}
 const rs=[...dist.keys()].filter(i=>i!==xi&&ROW[D.nodes[i]]).map(i=>[D.nodes[i],dist.get(i),D.nodes[prev.get(i)]]);
 rs.sort((a,b)=>a[1]-b[1]||ROW[b[0]][3]-ROW[a[0]][3]);
 if(!rs.length)return `<div class="hint">Nothing that ${esc(x)} feeds within ${depth} steps differs between scenarios.</div>`;
 return `<div class="hint">${rs.length} variables within ${depth} steps of ${esc(x)} differ between A and B.</div><div class="tw"><table><thead><tr><th>Variable</th><th class="n">Steps</th><th>Via</th><th class="n">Gap at end %</th><th class="n">Share from via</th></tr></thead><tbody>${rs.slice(0,300).map(r=>{const sh=esh(r[0],r[2]);return `<tr><td>${vlink(r[0])}</td><td class="n">${r[1]}</td><td>${vlink(r[2])}</td><td class="n">${gapTxt(r[0])}</td><td class="n">${sh==null?'–':Math.round(sh*100)+'%'}</td></tr>`}).join('')}</tbody></table></div>`}
(function(){
 $('#t-path').innerHTML=`<div class="hint">Can X affect Y, and by how much? ${SERVE?'The equations between X and Y are re-solved with <b>only X changed</b>, which gives the exact effect.':'The channels below show, at each link, what share of the next variable\'s gap comes from the previous one. <b>For the exact effect</b>, run <code>python scenario_explain.py serve …</code> (or the <code>impact X Y</code> command), which re-solves the equations with only X changed.'}</div>
 <div class="ctl"><input type="text" id="px" placeholder="X (e.g. FOIL)"><input type="text" id="py" placeholder="Y (optional)">${SERVE?'<select id="ps"><option value="">X moves as in scenario B</option><option value="1%">X +1%</option><option value="-1%">X −1%</option><option value="1">X +1 unit</option><option value="0.25">X +0.25 unit</option></select>':''}<select id="pd"><option value="2">2 steps</option><option value="3" selected>3 steps</option><option value="5">5 steps</option></select><button class="btn" id="pgo">Go</button></div><div id="pimp"></div><div id="pout"></div>`;
 const go=()=>{const x=$('#px').value.trim().toUpperCase(),y=$('#py').value.trim().toUpperCase();$('#pimp').innerHTML='';
  if(NI[x]==null){$('#pout').innerHTML=`<div class="card">${esc(x)} not found. ${sugg(x)}</div>`;return}
  if(y&&NI[y]==null){$('#pout').innerHTML=`<div class="card">${esc(y)} not found. ${sugg(y)}</div>`;return}
  $('#pout').innerHTML=y?'<h2>Channels</h2>'+pathHTML(x,y):`<h2>What ${esc(x)} feeds</h2>`+downHTML(x,+$('#pd').value);
  if(SERVE)runImpact(x,y,$('#ps').value);else if(y&&staticImpact(x,y))$('#pimp').insertAdjacentHTML('afterbegin','<div class="hint">Exact effect from the Summary tab (precomputed for scenario shocks → key variables). For any other pair, use <code>serve</code>.</div>')};
 $('#pgo').onclick=go;['px','py'].forEach(i=>$('#'+i).addEventListener('keydown',e=>{if(e.key==='Enter')go()}));
})();
function runImpact(x,y,shock){$('#pimp').innerHTML='<div class="hint">Re-solving the equations…</div>';
 fetch(`/api/impact?x=${encodeURIComponent(x)}&y=${encodeURIComponent(y)}&shock=${encodeURIComponent(shock)}`).then(r=>r.json()).then(o=>renderImpact(o,x,y,shock))}
function renderImpact(o,x,y,shock){
  if(o.error){$('#pimp').innerHTML=`<div class="card">${esc(o.error)}</div>`;return}
  if(o.none){$('#pimp').innerHTML=`<div class="card"><b>No.</b> ${esc(x)} cannot affect ${esc(y)}: no chain of equations links them.</div>`;return}
  const what=shock?`a ${esc(shock)}${shock.endsWith('%')?'':' unit'} change in ${esc(x)}`:`${esc(x.startsWith('AF:')?'the '+afText(x.slice(3)):x)}'s change from A to B`,NP=o.per.length;
  $('#pimp').innerHTML=`<div class="ctl" style="align-items:center"><span class="small mut">Quarter</span><input type="range" id="isl" min="0" max="${NP-1}" value="${NP-1}" style="flex:1;min-width:200px"><b id="ilab"></b></div><div id="ibody"></div>`;
  const draw=()=>{const i=+$('#isl').value;$('#ilab').textContent=o.per[i];let h='';
   if(y&&o.y){const v=o.y[i],g=o.gap?o.gap[i]:null,a=o.ya?o.ya[i]:null;
    h+=`<div class="card"><b>Effect of ${what} on ${esc(y)}</b> at <b>${esc(o.per[i])}</b>: <b class="${cls(v)}">${fmt(v)}</b>${a?' ('+pct(100*v/Math.abs(a))+')':''}${!shock&&g?` = <b>${Math.round(100*v/g)}%</b> of ${esc(y)}'s gap (${fmt(g)})`:''}.${o.nodes!=null?` ${o.nodes} equations re-solved.`:''}${o.blocked&&o.blocked.length?` <span class="mut small">${o.blocked.length} variables in between have no evaluable equation and were held fixed.</span>`:''}</div>`;
    h+=lineChart2(o.y,shock?null:o.gap,o.per,'Effect of '+(x.startsWith('AF:')?afText(x.slice(3)):x),'Total gap B−A',i);
    if(!shock&&o.gap){const sh=o.y.map((v,k)=>o.gap[k]&&Math.abs(o.gap[k])>1e-9*Math.max(...o.gap.map(z=>Math.abs(z||0)))?100*v/o.gap[k]:null);h+=`<div class="small mut" style="margin-top:6px">Share of ${esc(y)}'s gap from ${esc(x)}, by quarter (%):</div>`+lineChart2(sh,null,o.per,'share of gap, %','',i)}}
   else if(y&&!o.y)h+=`<div class="card">${esc(y)} has no evaluable equation, so the effect cannot be simulated.</div>`;
   if(o.top&&o.top.length){const rs=o.top.slice().sort((a,b)=>Math.abs(b[2][i]||0)-Math.abs(a[2][i]||0)).slice(0,40);
    h+=`<h2>${y?'Variables in between':'Most affected variables'} at ${esc(o.per[i])}</h2><div class="tw"><table><thead><tr><th>Variable</th><th class="n">Effect</th><th class="n">%</th><th class="n">Largest % (any quarter)</th></tr></thead><tbody>${rs.map(r=>{const m=r[2].reduce((s,p)=>p!=null&&Math.abs(p)>Math.abs(s)?p:s,0);return `<tr><td>${vlink(r[0])}</td><td class="n ${cls(r[1][i])}">${fmt(r[1][i])}</td><td class="n ${cls(r[2][i])}">${pct(r[2][i])}</td><td class="n mut">${pct(m)}</td></tr>`}).join('')}</tbody></table></div>`}
   $('#ibody').innerHTML=h};
  $('#isl').addEventListener('input',draw);draw()}
function staticImpact(x,y){const S=D.summary;if(!S||!y||!S.attr[y])return false;
 const ids=[x,'AF:'+x];let r=null;for(const id of ids){r=S.attr[y].rows.find(z=>z[0]===id);if(r){x=id;break}}if(!r)return false;
 renderImpact({per:S.per,y:r[1],gap:S.attr[y].gap,ya:null,nodes:null,blocked:[],top:[]},x,y,'');return true}
function lineChart2(a,b,per,la,lb,sel){const W=640,H=180,L=56,R=8,T=10,B=24,n=per.length;const vs=a.concat(b||[]).concat([0]).filter(v=>v!=null);let lo=Math.min(...vs),hi=Math.max(...vs);if(lo===hi)hi=lo+1;
 const X=i=>L+(W-L-R)*(n>1?i/(n-1):.5),Y=v=>T+(H-T-B)*(1-(v-lo)/(hi-lo));const pl=arr=>arr.map((v,i)=>v==null?'':(i&&arr[i-1]!=null?'L':'M')+X(i).toFixed(1)+','+Y(v).toFixed(1)).join('');
 return `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:780px">${sel!=null?`<line x1="${X(sel)}" x2="${X(sel)}" y1="${T}" y2="${H-B}" stroke="var(--mut)" stroke-dasharray="3,3"/>${a[sel]!=null?`<circle cx="${X(sel)}" cy="${Y(a[sel])}" r="4" fill="var(--b)"/>`:''}`:''}<line x1="${L}" x2="${W-R}" y1="${Y(0)}" y2="${Y(0)}" stroke="var(--line)"/><text x="${L-6}" y="${Y(hi)+8}" text-anchor="end">${fmt(hi)}</text><text x="${L-6}" y="${Y(lo)}" text-anchor="end">${fmt(lo)}</text><text x="${L}" y="${H-6}">${esc(per[0])}</text><text x="${W-R}" y="${H-6}" text-anchor="end">${esc(per[n-1])}</text>
 ${b?`<path d="${pl(b)}" fill="none" stroke="var(--mut)" stroke-width="2" stroke-dasharray="4,3"/>`:''}<path d="${pl(a)}" fill="none" stroke="var(--b)" stroke-width="2.5"/></svg><div class="small"><span style="color:var(--b)">━ ${esc(la)}</span>${b?` &nbsp; <span class="mut">╍ ${esc(lb)}</span>`:''}</div>`}
document.addEventListener('click',e=>{const g=e.target.closest('g.mn');if(g)show(g.dataset.v)});
function sugg(x){if(!x)return'';const m=D.nodes.filter(n=>n.includes(x)).slice(0,12);return m.length?'Did you mean: '+m.map(vlink).join(', '):''}

// ---------- charts
function lineChart(a,b,per,wsi){
 const W=640,H=200,L=56,R=8,T=10,B=24,n=per.length;const vs=a.concat(b).filter(v=>v!=null);if(!vs.length)return'';
 let lo=Math.min(...vs),hi=Math.max(...vs);if(lo===hi){lo-=1;hi+=1}const pad=(hi-lo)*.06;lo-=pad;hi+=pad;
 const X=i=>L+(W-L-R)*(n>1?i/(n-1):.5),Y=v=>T+(H-T-B)*(1-(v-lo)/(hi-lo));
 const pl=arr=>{let d='',pen=false;arr.forEach((v,i)=>{if(v==null){pen=false;return}d+=(pen?'L':'M')+X(i).toFixed(1)+','+Y(v).toFixed(1);pen=true});return d};
 return `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img"><line x1="${L}" x2="${W-R}" y1="${H-B}" y2="${H-B}" stroke="var(--line)"/>
 ${wsi>0?`<line x1="${X(wsi)}" x2="${X(wsi)}" y1="${T}" y2="${H-B}" stroke="var(--line)" stroke-dasharray="3,3"/>`:''}
 <text x="${L-6}" y="${Y(hi-pad)+4}" text-anchor="end">${fmt(hi-pad)}</text><text x="${L-6}" y="${Y(lo+pad)+4}" text-anchor="end">${fmt(lo+pad)}</text>
 <text x="${L}" y="${H-6}">${esc(per[0])}</text>${wsi>0?`<text x="${X(wsi)}" y="${H-6}" text-anchor="middle">${esc(per[wsi])}</text>`:''}<text x="${W-R}" y="${H-6}" text-anchor="end">${esc(per[n-1])}</text>
 <path d="${pl(a)}" fill="none" stroke="var(--a)" stroke-width="2"/><path d="${pl(b)}" fill="none" stroke="var(--b)" stroke-width="2" stroke-dasharray="5,3"/></svg>
 <div class="small"><span style="color:var(--a)">━ A: ${esc(M.a)}</span> &nbsp; <span style="color:var(--b)">╍ B: ${esc(M.b)}</span></div>`}
function barChart(vals,per,label){
 const W=640,H=120,L=56,R=8,T=8,B=20,n=per.length;const vs=vals.filter(v=>v!=null);if(!vs.length)return'';
 let lo=Math.min(0,...vs),hi=Math.max(0,...vs);if(lo===hi)hi=1;const X=i=>L+(W-L-R)*i/n,bw=Math.max(1,(W-L-R)/n-1),Y=v=>T+(H-T-B)*(1-(v-lo)/(hi-lo));
 return `<svg viewBox="0 0 ${W} ${H}" width="100%"><line x1="${L}" x2="${W-R}" y1="${Y(0)}" y2="${Y(0)}" stroke="var(--line)"/>
 ${vals.map((v,i)=>v==null?'':`<rect x="${X(i)}" width="${bw}" y="${Math.min(Y(v),Y(0))}" height="${Math.max(.5,Math.abs(Y(v)-Y(0)))}" fill="${v>=0?'var(--pos)':'var(--neg)'}" opacity=".8"><title>${esc(per[i])}: ${fmt(v)}</title></rect>`).join('')}
 <text x="${L-6}" y="${Y(hi)+8}" text-anchor="end">${fmt(hi)}</text><text x="${L-6}" y="${Y(lo)}" text-anchor="end">${fmt(lo)}</text>
 <text x="${L}" y="${H-4}">${esc(per[0])}</text><text x="${W-R}" y="${H-4}" text-anchor="end">${esc(per[n-1])}</text></svg><div class="small mut">${label}</div>`}

// ---------- detail panel
const hist=[];
function show(n,noPush){
 n=String(n).toUpperCase();
 if(SERVE&&!(n in D.vals)){fetch('/api/detail?n='+encodeURIComponent(n)).then(r=>r.json()).then(o=>{D.vals[n]=o.error?null:o.vals;if(o.dec)D.dec[n]=o.dec;if(o.eqt)D.eqt[n]=o.eqt;if(o.why)D.why[n]=o.why;if(o.desc&&!ROW[n])D.desc[n]=o.desc;show(n,noPush)});return}
 if(!noPush){hist.push(n)}
 const r=ROW[n],v=D.vals[n],d=D.dec[n],eq=D.eqt[n],kind=r?r[9]:(M.has_eqs?(eq?'':'X'):'-');
 let h=`<div class="dh"><div>${hist.length>1?'<button class="btn" id="bk">← Back</button> ':''}</div><button class="btn" id="cx">Close ✕</button></div>
 <h2 style="margin-top:8px">${esc(n)} ${badge(r?r[9]:'',n)}</h2><div class="mut">${esc(desc(n))}</div>`;
 if(r)h+=`<div class="kpis"><div class="kpi"><span class="mut small">gap at end (B−A)</span><b class="${cls(r[6])}">${fmt(r[6])}</b><span class="small ${cls(r[7])}">${pct(r[7])}</span></div><div class="kpi"><span class="mut small">largest gap</span><b class="${cls(r[4])}">${fmt(r[4])}</b><span class="small ${cls(r[5])}">${pct(r[5])}</span></div><div class="kpi"><span class="mut small">first differs</span><b>${esc(r[8])}</b></div><div class="kpi"><span class="mut small">score</span><b>${fmt(r[3])}</b></div></div>`;
 else h+=`<div class="card">No difference between A and B in the window${(n in D.vals)?'':' (or series not in the data)'}.</div>`;
 if(v){h+=lineChart(v[0],v[1],D.per,D.wsi);const df=v[0].map((x,i)=>x==null||v[1][i]==null?null:v[1][i]-x);h+=barChart(df,D.per,'Gap B − A by period')}
 else if(r){const g=D.pg[n];if(g)h+=lineChart2(g,null,D.per.slice(D.wsi),'gap B − A, % of A','',null);
  h+=`<div class="hint">The A and B paths are not embedded for this variable, only its gap (% of A) by quarter${D.eshp&&D.eshp[NI[n]]?' and, below, where it comes from':''}. For the full charts, build the report with <code>--embed all</code> or use <code>serve</code>.</div>`}
 if(eq&&EXOG.has(n))h+=`<h2>Equation (switched off)</h2><pre class="eq">${esc(eq)}</pre><div class="card">The modeller set <b>${esc(n)}</b> exogenously, so this equation is not used: its whole gap is a scenario choice, and shocks to its inputs stop here.</div>`;
 else if(ES[n]){const e=(D.eqdiff||[]).find(z=>z[0]===n);h+=`<h2>Equations</h2>`+eqPair(e)+`<div class="hint">${ES[n]==='changed'?'The formula differs between the scenarios. The effect of switching from A\'s equation to B\'s (together with its add-factor) is shown below as one line, <i>equation changed</i>. Each effect is averaged over A\'s and B\'s equation.':ES[n]==='only_a'?'B has no equation for it: B sets it directly. That switch is shown below as <i>equation dropped in B</i>.':'A has no equation for it: A sets it directly. B\'s formula is used for both, and the switch is shown as <i>equation added in B</i>.'}</div>`}
 else if(eq){h+=`<h2>Equation</h2><pre class="eq">${esc(eq)}</pre>`;if(/\[[^\]]*\b(ar|ma|sar|sma)\s*\(/i.test(eq))h+=`<div class="hint">The [AR/MA …] part describes how the equation's error evolves. It is not dropped: the tool backs the error out of the data in each scenario, so the AR/MA process is already inside it. A change in an input moves ${esc(n)} through the formula, the error path stays as in scenario A, and a changed add-factor carries forward through the AR term exactly as in the model.</div>`;if(D.why[n])h+=`<div class="hint">${esc(D.why[n])}</div>`;if(D.afmap[n])h+=`<div class="hint">Add-factor series: ${vlink(D.afmap[n])} ${gapTxt(D.afmap[n])}</div>`}
 else if(M.has_eqs)h+=`<div class="card">No equation for ${esc(n)} — it is an <b>assumption</b> (exogenous). Its path is set directly by the scenario.</div>`;
 if(d&&d.err)h+=`<div class="card">Could not decompose: ${esc(d.err)}</div>`;
 else if(d)h+=decHTML(n,d);
 if(SERVE&&eq&&r)h+=`<p><button class="btn" id="atb2">Split ${esc(n)}'s gap across the scenario shocks</button></p><div id="atout"></div>`;
 else if(r&&eq&&!D.dec[n]&&D.eshp[NI[n]]){h+=shareOnlyHTML(n)}
 else if(r&&eq&&!D.dec[n]){const l=D.esh[NI[n]];if(l&&l.length){h+=`<h2>Why it changed</h2><div class="hint">Share of the change in its left-hand side coming from each driver:</div>`+l.slice().sort((a,b)=>Math.abs(b[1])-Math.abs(a[1])).map(e=>`${vlink(D.nodes[e[0]])} ${Math.round(e[1]*100)}%`).join(' · ')+(D.afsh[NI[n]]!=null?` · <i>add-factor</i> ${D.afmap[n]?vlink(D.afmap[n]):''} <b>${Math.round(D.afsh[NI[n]]*100)}%</b>`:'')}}
 const ps=(D.par[NI[n]]||[]).map(i=>D.nodes[i]),cs=(CH[NI[n]]||[]).map(i=>D.nodes[i]);
 if(ps.length)h+=`<h2>Drivers in its equation</h2><div>${ps.map(p=>`${vlink(p)} ${gapTxt(p)}`).join(' &nbsp;·&nbsp; ')}</div>`;
 if(M.has_eqs&&NI[n]!=null&&cs.length)h+=`<div class="hint">Linked to ${D.reach[NI[n]].toLocaleString()} variables in total (directly or indirectly). See the Influence tab for how much they move.</div>`;
 if(!cs.length&&M.has_eqs)h+=`<h2>Feeds into</h2><div class="hint">Nothing — no other equation uses ${esc(n)}. It is an end point of the model.</div>`;
 if(cs.length){const cd=cs.filter(c=>ROW[c]),cn=cs.length-cd.length;h+=`<h2>Feeds into</h2><div class="hint">${cs.length} variable(s) use ${esc(n)} in their equations${cn?`; ${cn} of them do not differ`:''}.</div><div>${cd.sort((a,b)=>ROW[b][3]-ROW[a][3]).slice(0,60).map(c=>{const sh=esh(c,n);return `${vlink(c)} ${gapTxt(c)}${sh!=null?` <span class="small mut">(${Math.round(sh*100)}% from ${esc(n)})</span>`:''}`}).join(' &nbsp;·&nbsp; ')}</div>`}
 const det=$('#det');det.innerHTML=h;det.classList.add('open');det.scrollTop=0;
 $('#cx').onclick=()=>{det.classList.remove('open');hist.length=0};
 const bk=$('#bk');if(bk)bk.onclick=()=>{hist.pop();show(hist[hist.length-1],true)};
 const pb=$('#pb');if(pb)pb.onclick=()=>{$('#ptab').classList.toggle('hide')};
 const sos=$('#sosl');if(sos){shareOnlyUpdate(n,+sos.value);sos.oninput=()=>shareOnlyUpdate(n,+sos.value)}
 const dsl=$('#dsl');if(dsl){const dd=D.dec[n];decUpdate(n,dd,+dsl.value);dsl.oninput=()=>decUpdate(n,dd,+dsl.value)}
 const at=$('#atb2');if(at)at.onclick=()=>{$('#atout').innerHTML='<div class="hint">Re-solving…</div>';fetch('/api/attr?y='+encodeURIComponent(n)).then(r=>r.json()).then(o=>{if(o.error){$('#atout').innerHTML=`<div class="card">${esc(o.error)}</div>`;return}
  const i=o.iend,g=o.gap[i],mx=Math.max(1e-30,...o.rows.map(r=>Math.abs(r[3][i]||0)));let s=0;o.rows.forEach(r=>s+=r[3][i]||0);
  $('#atout').innerHTML=`<div class="hint">Gap at ${esc(o.per[i])}: <b>${fmt(g)}</b>. Each line re-solves the model with only that shock switched from A to B.</div><table class="bars">${o.rows.map(r=>`<tr><td style="width:42%">${r[1]==='addfactor'?'add-factor of '+vlink(r[0])+(r[2]?' ('+esc(r[2])+')':''):vlink(r[0])}</td><td style="width:34%">${bar(r[3][i],mx)}</td><td class="n">${fmt(r[3][i])}</td><td class="n mut">${g?Math.round(100*r[3][i]/g)+'%':''}</td></tr>`).join('')}<tr><td><i>Interaction between shocks</i></td><td></td><td class="n">${fmt(g-s)}</td><td class="n mut">${g?Math.round(100*(g-s)/g)+'%':''}</td></tr></table>`})};
}
document.addEventListener('keydown',e=>{if(e.key==='Escape'){$('#det').classList.remove('open');hist.length=0}});
function bar(v,mx){v=v||0;const w=Math.min(100,100*Math.abs(v)/mx);return `<div style="display:flex"><div style="width:50%;display:flex;justify-content:flex-end">${v<0?`<div class="bar" style="width:${w}%;background:var(--neg)"></div>`:''}</div><div style="width:50%">${v>0?`<div class="bar" style="width:${w}%;background:var(--pos)"></div>`:''}</div></div>`}
function decBars(n,d,i){
 const g=d.gap[i];const rows=d.cols.map(c=>[c[0],c[1][i]]);
 if(d.af)rows.push(['__AF__',d.af[i]]);
 const mx=Math.max(1e-30,Math.abs(g||0),...rows.map(r=>Math.abs(r[1]||0)));
 if(d.int&&Math.abs(d.int[i]||0)>1e-3*mx)rows.push(['__INT__',d.int[i]]);
 const sh=v=>g&&Math.abs(g)>1e-12*mx?Math.round(100*(v||0)/g)+'%':'';
 let h=`<table class="bars"><tr><td style="width:40%"><b>Gap B−A</b></td><td style="width:34%">${bar(g,mx)}</td><td class="n"><b>${fmt(g)}</b></td><td class="n mut">100%</td></tr>`;
 h+=rows.map(r=>`<tr><td>${r[0]==='__AF__'?(ES[n]?`<i>${afWord(n)}</i> <span class="small mut">(formula + add-factor)</span>`:`<i>Add-factor</i> ${d.afn?vlink(d.afn):'<span class="small mut">(residual)</span>'}`):r[0]==='__INT__'?'<i>Interaction (non-linear)</i>':vlink(r[0])+' '+badge(ROW[r[0]]?ROW[r[0]][9]:'')}</td><td>${bar(r[1],mx)}</td><td class="n">${fmt(r[1])}</td><td class="n mut">${sh(r[1])}</td></tr>`).join('');
 if(!d.cols.length&&!d.af)h+=`<tr><td colspan="4" class="mut">None of its inputs differ and its add-factor did not change${g?'':' — that is why it did not move'}.</td></tr>`;
 return h+'</table>'}
const PAL=['#2563eb','#0f9d58','#7a5af8','#c2185b','#00838f','#8d6e63','#5c6bc0'],AFC='#f59e0b';
function decSeries(d){const top=d.cols.slice(0,6),rest=d.cols.slice(6),out=top.map((c,k)=>({name:c[0],v:c[1],c:PAL[k]}));
 if(rest.length)out.push({name:`${rest.length} other input(s)`,v:d.gap.map((_,i)=>rest.reduce((s,c)=>s+(c[1][i]||0),0)),c:'#9aa4b2'});
 if(d.af)out.push({name:(ES[d.n]?afWord(d.n):'Add-factor')+(d.afn?' '+d.afn:''),v:d.af,c:AFC});return out}
function stackChart(d,sel){
 const S=decSeries(d),n=d.gap.length,W=640,H=210,L=56,R=8,T=10,B=24;let lo=0,hi=0;
 for(let i=0;i<n;i++){let p=0,m=0;S.forEach(s=>{const v=s.v[i]||0;if(v>0)p+=v;else m+=v});lo=Math.min(lo,m,d.gap[i]||0);hi=Math.max(hi,p,d.gap[i]||0)}if(lo===hi)hi=1;
 const bw=(W-L-R)/n,X=i=>L+bw*i,Y=v=>T+(H-T-B)*(1-(v-lo)/(hi-lo));let r='';
 for(let i=0;i<n;i++){let p=0,m=0;S.forEach(s=>{const v=s.v[i]||0;if(!v)return;const y0=v>0?p:m,y1=y0+v;if(v>0)p=y1;else m=y1;r+=`<rect x="${X(i)+1}" width="${Math.max(1,bw-2)}" y="${Math.min(Y(y0),Y(y1))}" height="${Math.max(.5,Math.abs(Y(y1)-Y(y0)))}" fill="${s.c}"><title>${esc(D.per[i+D.wsi])} · ${esc(s.name)}: ${fmt(v)}</title></rect>`})}
 const gl=d.gap.map((v,i)=>v==null?'':(i&&d.gap[i-1]!=null?'L':'M')+(X(i)+bw/2).toFixed(1)+','+Y(v).toFixed(1)).join('');
 return `<svg viewBox="0 0 ${W} ${H}" width="100%"><rect x="${X(sel)}" width="${bw}" y="${T}" height="${H-T-B}" fill="var(--hl)"/><line x1="${L}" x2="${W-R}" y1="${Y(0)}" y2="${Y(0)}" stroke="var(--line)"/>${r}<path d="${gl}" fill="none" stroke="var(--fg)" stroke-width="1.5"/>${d.gap.map((v,i)=>v==null?'':`<circle cx="${X(i)+bw/2}" cy="${Y(v)}" r="2" fill="var(--fg)"/>`).join('')}
 <text x="${L-6}" y="${Y(hi)+8}" text-anchor="end">${fmt(hi)}</text><text x="${L-6}" y="${Y(lo)}" text-anchor="end">${fmt(lo)}</text><text x="${L}" y="${H-6}">${esc(D.per[D.wsi])}</text><text x="${W-R}" y="${H-6}" text-anchor="end">${esc(D.per[D.wsi+n-1])}</text></svg>
 <div class="small">${S.map(s=>`<span style="white-space:nowrap;margin-right:10px"><span style="display:inline-block;width:10px;height:10px;background:${s.c};margin-right:4px"></span>${esc(s.name)}</span>`).join('')}<span style="white-space:nowrap">● ━ gap B−A</span></div>`}
function shareChart(d,sel){
 const n=d.gap.length,W=640,H=170,L=56,R=8,T=10,B=24,big=Math.max(1e-30,...d.gap.map(v=>Math.abs(v||0)));
 const f=d.gap.map((g,i)=>g==null||Math.abs(g)<1e-3*big?null:Math.max(-1.5,Math.min(2.5,d.cols.reduce((s,c)=>s+(c[1][i]||0),0)/g)));
 const a=d.af?d.gap.map((g,i)=>g==null||Math.abs(g)<1e-3*big?null:Math.max(-1.5,Math.min(2.5,(d.af[i]||0)/g))):null;
 const vs=f.concat(a||[]).filter(v=>v!=null).concat([0,1]);const lo=Math.min(...vs),hi=Math.max(...vs);
 const X=i=>L+(W-L-R)*(n>1?i/(n-1):.5),Y=v=>T+(H-T-B)*(1-(v-lo)/(hi-lo||1));const pl=arr=>arr.map((v,i)=>v==null?'':(i&&arr[i-1]!=null?'L':'M')+X(i).toFixed(1)+','+Y(v).toFixed(1)).join('');
 return `<svg viewBox="0 0 ${W} ${H}" width="100%"><line x1="${X(sel)}" x2="${X(sel)}" y1="${T}" y2="${H-B}" stroke="var(--mut)" stroke-dasharray="3,3"/>
 <line x1="${L}" x2="${W-R}" y1="${Y(0)}" y2="${Y(0)}" stroke="var(--line)"/><line x1="${L}" x2="${W-R}" y1="${Y(1)}" y2="${Y(1)}" stroke="var(--line)" stroke-dasharray="2,3"/>
 <text x="${L-6}" y="${Y(1)+4}" text-anchor="end">100%</text><text x="${L-6}" y="${Y(0)+4}" text-anchor="end">0%</text><text x="${L}" y="${H-6}">${esc(D.per[D.wsi])}</text><text x="${W-R}" y="${H-6}" text-anchor="end">${esc(D.per[D.wsi+n-1])}</text>
 <path d="${pl(f)}" fill="none" stroke="var(--a)" stroke-width="2"/>${a?`<path d="${pl(a)}" fill="none" stroke="${AFC}" stroke-width="2"/>`:''}</svg>
 <div class="small"><span style="color:var(--a)">━ from ${ES[d.n]?'inputs':'formula'}</span>${a?` &nbsp; <span style="color:${AFC}">━ from ${ES[d.n]?afWord(d.n):'add-factor'}</span>`:''} &nbsp; <span class="mut">(share of the gap in each quarter; hidden where the gap is ~0)</span></div>`}
function decHTML(n,d){
 const i=d.iend,NP=d.gap.length;
 let h=`<h2>Where the gap comes from</h2><div class="hint">Each input's part is the gap ${esc(n)} would have if <i>only</i> that input moved from A to B, with the equation re-solved through time, so effects that build up through its own lags are credited to the input that caused them. ${ES[n]?'Its equation differs between the scenarios: the effect of switching to B\'s equation (with its add-factor) is the <b>equation change</b> line; every effect is averaged over A\'s and B\'s equation, so inputs that only one of them uses still get their share.':'The equation is the same in both scenarios, so the rest comes from the <b>add-factor</b>.'}</div>
 <div id="dcharts"></div>
 <div class="ctl" style="align-items:center"><span class="small mut">Quarter</span><input type="range" id="dsl" min="0" max="${NP-1}" value="${i}" style="flex:1;min-width:180px"><b id="dlab"></b></div>
 <div id="dsum" class="hint"></div><div id="dbars"></div>`;
 if(d.cols.length||d.af){const ks=d.cols.slice(0,8);h+=`<p><button class="btn" id="pb">Show every quarter as a table</button></p><div id="ptab" class="hide tw"><table><thead><tr><th>Period</th><th class="n">Gap</th>${ks.map(k=>`<th class="n">${esc(k[0])}</th>`).join('')}${d.af?'<th class="n">Add-factor</th>':''}</tr></thead><tbody>${d.gap.map((v,k)=>`<tr><td>${esc(D.per[k+D.wsi])}</td><td class="n">${fmt(v)}</td>${ks.map(c=>`<td class="n ${cls(c[1][k])}">${fmt(c[1][k])}</td>`).join('')}${d.af?`<td class="n ${cls(d.af[k])}">${fmt(d.af[k])}</td>`:''}</tr>`).join('')}</tbody></table></div>`}
 return h}
function eqPair(e){if(!e)return'';const tk=a=>a.length?a.map(t=>t[1]===1?`<mark style="background:#fde68a;color:#111;border-radius:3px">${esc(t[0])}</mark>`:t[1]===2?`<span class="mut" title="Error-term options: not part of the comparison">${esc(t[0])}</span>`:esc(t[0])).join(''):'<span class="mut">(no equation: set directly by the scenario)</span>';
 return `<div class="small mut">Scenario A</div><pre class="eq">${tk(e[2])}</pre><div class="small mut">Scenario B</div><pre class="eq">${tk(e[3])}</pre>`}
function shareOnlyHTML(n){const NP=D.per.length-D.wsi;
 return `<h2>Where the gap comes from</h2><div class="hint">Share of ${esc(n)}'s gap from each input and from its add-factor, quarter by quarter (each input switched from A to B on its own, equation re-solved through time).</div><div id="sochart"></div>
 <div class="ctl" style="align-items:center"><span class="small mut">Quarter</span><input type="range" id="sosl" min="0" max="${NP-1}" value="${NP-1}" style="flex:1;min-width:180px"><b id="solab"></b></div><div id="sobars"></div>`}
function shareOnlyUpdate(n,i){const ix=NI[n],l=D.eshp[ix]||[],af=D.afshp[ix],g=D.pg[n];
 $('#solab').textContent=D.per[D.wsi+i];
 const rows=l.map(e=>[D.nodes[e[0]],e[1][i]==null?null:e[1][i]/1000]);if(af)rows.push(['__AF__',af[i]==null?null:af[i]/1000]);
 rows.sort((a,b)=>Math.abs(b[1]||0)-Math.abs(a[1]||0));const tot=rows.reduce((s,r)=>s+(r[1]||0),0),mx=Math.max(1,...rows.map(r=>Math.abs(r[1]||0)));
 $('#sobars').innerHTML=`<div class="hint">Gap at ${esc(D.per[D.wsi+i])}: <b>${g&&g[i]!=null?pct(g[i]):'–'}</b> vs A.</div><table class="bars">${rows.map(r=>`<tr><td style="width:40%">${r[0]==='__AF__'?(ES[n]?`<i>${afWord(n)}</i> <span class="small mut">(formula + add-factor)</span>`:`<i>Add-factor</i> ${D.afmap[n]?vlink(D.afmap[n]):''}`):vlink(r[0])}</td><td style="width:34%">${bar(r[1]||0,mx)}</td><td class="n">${r[1]==null?'n/a':Math.round(100*r[1])+'%'}</td></tr>`).join('')}${Math.abs(1-tot)>=.005&&rows.some(r=>r[1]!=null)?`<tr><td><i>Interaction (non-linear)</i></td><td></td><td class="n">${Math.round(100*(1-tot))}%</td></tr>`:''}</table>`;
 const NP=D.per.length-D.wsi,f=[...Array(NP).keys()].map(k=>{let s=0,ok=false;l.forEach(e=>{if(e[1][k]!=null){s+=e[1][k]/10;ok=true}});return ok?s:null}),a=af?af.map(v=>v==null?null:v/10):null;
 $('#sochart').innerHTML=lineChart2(f,a,D.per.slice(D.wsi),'share from formula (inputs), %','share from add-factor, %',i)}
function decUpdate(n,d,i){d.n=n;
 $('#dlab').textContent=D.per[i+D.wsi];
 const g=d.gap[i],fs=d.cols.reduce((s,c)=>s+(c[1][i]||0),0);
 $('#dsum').innerHTML=g&&Math.abs(g)>1e-12?`At <b>${esc(D.per[i+D.wsi])}</b> the gap is <b>${fmt(g)}</b>: <b>${Math.round(100*fs/g)}%</b> from ${ES[n]?'the inputs':'the formula'}${d.af?`, <b>${Math.round(100*(d.af[i]||0)/g)}%</b> from the ${ES[n]?afWord(n):'add-factor'}`:''}.`:`At ${esc(D.per[i+D.wsi])} there is no gap.`;
 $('#dbars').innerHTML=decBars(n,d,i);
 if(d.cols.length||d.af)$('#dcharts').innerHTML=`<div class="small mut" style="margin-top:6px">Contribution of each input to the gap, by quarter (bars), and the gap itself (line):</div>`+stackChart(d,i)+`<div class="small mut" style="margin-top:6px">Share of the gap from ${ES[n]?'the inputs vs the equation change':'the formula vs the add-factor'}, by quarter:</div>`+shareChart(d,i)}
tab(D.summary?'story':'sum');
(function(){const o=M.open;if(!o||!o.tab||!$('#t-'+o.tab)||!document.querySelector('nav button[data-t="'+o.tab+'"]'))return;tab(o.tab);
 if(o.tab==='imp'){const w={all:'3',concept:'4',key:'5'}[o.weight],iw=$('#iw'),ig=$('#ig'),iq=$('#iq');
  if(iw&&w&&[...iw.options].some(x=>x.value===w))iw.value=w;
  if(ig&&o.geo&&[...ig.options].some(x=>x.value===o.geo))ig.value=o.geo;
  (ig||iw).dispatchEvent(new Event('input'))}})();
</script></body></html>
'''


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def build_model(args):
    t0 = time.time()
    say = (lambda m: print(m, file=sys.stderr)) if not getattr(args, 'quiet', False) else (lambda m: None)
    say('Loading scenario A ...')
    A = load_scenario(args.a, args.name_a or _label(args.a), args.sheet)
    say('Loading scenario B ...')
    B = load_scenario(args.b, args.name_b or _label(args.b), args.sheet)
    eqs, addl, eqstats, fails = {}, {}, None, []
    eqs_b = None
    src_a = getattr(args, 'eq_a', None) or args.eq
    src_b = getattr(args, 'eq_b', None)
    if src_a:
        say('Reading equations%s ...' % (' for A' if src_b else ''))
        eqs, addl, eqstats, fails = load_equations(src_a)
    if src_b:
        say('Reading equations for B ...')
        eqs_b, addl_b, st_b, fails_b = load_equations(src_b)
        addl = dict(addl_b, **addl)
        eqstats = Counter(eqstats or {})
        for k, v in st_b.items():
            eqstats['B ' + k] = v
        fails = list(fails) + ['(B) ' + f for f in fails_b]
        if not src_a:
            eqs = {}
    meta = load_meta(args.meta) if args.meta else None
    exog_names = set(load_exog(args.exog))
    lp = defaultdict(set)
    for E in (eqs, eqs_b or {}):
        for n, e in E.items():
            if n not in exog_names:
                lp[n].update(d for d in e.drivers if d != n)
    allnames = set(A.series) | set(B.series) | set(eqs) | set(eqs_b or {})
    area, area_label = None, None
    if getattr(args, 'area', None):
        area, area_label = _match_spec(allnames, args.area, '--area', say), ', '.join(args.area)
    elif getattr(args, 'target', None):
        area, area_label = _match_spec(allnames, args.target, '--target', lambda m: None), ', '.join(args.target)
    t1 = time.time()
    args.layers = compute_layers(lp, area)
    args.layers_area = (sorted(area) if area else None, area_label)
    say('Layers: longest chain %d%s, %d variables in %.1fs'
        % (max((v[0] for v in args.layers.values()), default=0),
           (', into %s %s' % (area_label, max((v[1] for v in args.layers.values() if v[1] is not None), default=0)))
           if area else '', len(args.layers), time.time() - t1))
    if getattr(args, 'target', None):
        args.target_set = _apply_target(args.target, A, B, eqs, eqs_b, addl, load_exog(args.exog),
                                        args.af_suffix, say, outside_exog=getattr(args, 'target_exog', False))
    elif getattr(args, 'target_exog', False):
        raise SystemExit('--target-exog needs --target')
    M = Model(A, B, eqs, addl, start=args.start, end=args.end, tol=args.tol, geo_sep=args.geo_sep,
              adj_threshold=args.adj_threshold, meta=meta, verbose=not getattr(args, 'quiet', False),
              af_suffix=args.af_suffix, exog=load_exog(args.exog), eqs_b=eqs_b)
    say('Loaded %d + %d series, %d periods, %d equations in %.1fs'
        % (len(A.series), len(B.series), M.T, len(eqs), time.time() - t0))
    return M, eqstats, fails


def _match_spec(names, spec_list, opt, say):
    """Names matching --target / --area specs: exact names, wildcards (*_ICAN), or a bare geography code
    (ICAN = *.ICAN and *_ICAN); comma lists allowed."""
    import fnmatch
    out = set()
    for spec in spec_list or ():
        for f in re.split(r'[,;\s]+', spec):
            f = f.strip().strip('"\'').upper()
            if not f:
                continue
            if any(c in f for c in '*?['):
                hit = {n for n in names if fnmatch.fnmatchcase(n, f)}
            elif f in names:
                hit = {f}
            else:
                code = f.lstrip('._')
                hit = {n for n in names if n.endswith('.' + code) or n.endswith('_' + code)}
            if not hit:
                say('%s %s: nothing matches' % (opt, f))
            out |= hit
    return out


def compute_layers(parents, area=None):
    """Longest chain of equations from each variable downstream ("layers it affects").
    parents: variable -> variables its equation uses (any lag). A loop counts as one layer, since inside
    a loop the longest chain is endless. Returns per variable:
      d    longest chain to any variable (0 = feeds nothing)
      da   longest chain that ends in `area` (None = cannot reach the area; 0 = in the area, feeds no other)
      nx / nxa   next variable on that chain;  lp  size of the loop it sits in (1 = no loop)."""
    nodes = set(parents)
    for ps in parents.values():
        nodes.update(ps)
    nodes = sorted(nodes)
    comps = _scc_order(nodes, parents)          # upstream components first
    comp = {}
    for ci, c in enumerate(comps):
        for n in c:
            comp[n] = ci
    out_edges = defaultdict(list)               # component -> [(u, v)] edges leaving it
    for v, ps in parents.items():
        cv = comp[v]
        for u in ps:
            if comp[u] != cv:
                out_edges[comp[u]].append((u, v))
    area = set(area or ())
    nc = len(comps)
    d, da, nx, nxa = [0] * nc, [None] * nc, [None] * nc, [None] * nc
    for ci in range(nc - 1, -1, -1):            # downstream components first
        if any(n in area for n in comps[ci]):
            da[ci] = 0
        for u, v in out_edges.get(ci, ()):
            cv = comp[v]
            if d[cv] + 1 > d[ci]:
                d[ci], nx[ci] = d[cv] + 1, (u, v)
            if da[cv] is not None and (da[ci] is None or da[cv] + 1 > da[ci]):
                da[ci], nxa[ci] = da[cv] + 1, (u, v)
    res = {}
    for n in nodes:
        ci = comp[n]
        res[n] = (d[ci], da[ci], nx[ci][1] if nx[ci] else None, nxa[ci][1] if nxa[ci] else None, len(comps[ci]))
    return res


def layer_chain(L, n, area=False, limit=60):
    """The longest chain from n as a list of names (loops shown by one member each)."""
    out, k = [n], 3 if area else 2
    while len(out) < limit:
        nxt = L.get(out[-1], (0, None, None, None, 1))[k]
        if not nxt or nxt in out:
            break
        out.append(nxt)
    return out


def _apply_target(spec_list, A, B, eqs, eqs_b, addl, exog, af_suffix, say, outside_exog=False):
    """--target: keep only the target variables and every series that feeds them through the equations
    (directly or via any chain, lags included), plus their add-factor series. Everything else is dropped
    before the model is built, so all results are unchanged for the kept variables, only faster.
    A spec is an exact name, a wildcard (*_ICAN, FGDP*), or a bare geography code (ICAN = *.ICAN and *_ICAN).
    outside_exog (--target-exog): only the targets keep their equations; every other variable their equations
    use is taken as given from each scenario's data (exogenous), and nothing upstream of it is kept."""
    names = set(A.series) | set(B.series) | set(eqs) | set(eqs_b or {})
    targets = _match_spec(names, spec_list, '--target', say)
    if not targets:
        raise SystemExit('--target matched no series')
    exog = set(exog or ())
    drivers = defaultdict(set)
    for E in (eqs, eqs_b or {}):
        for n, e in E.items():
            if n not in exog:
                drivers[n].update(d for d in e.drivers if d != n)
    for y, af in (addl or {}).items():
        drivers[y].add(af)
    keep, stack = set(targets), list(targets)
    while stack:
        n = stack.pop()
        for d in drivers.get(n, ()):
            if d not in keep:
                keep.add(d)
                if not outside_exog:           # with --target-exog, stop at the first variable outside
                    stack.append(d)
    # add-factor series named by suffix (GDP.US_A or GDP_A.US) of kept variables
    sufs = [x.strip().upper() for x in (af_suffix or '').split(',') if x.strip()]
    for n in set(A.series) | set(B.series):
        if n in keep:
            continue
        for sf in sufs:
            base = [n[:-len(sf)]] if n.endswith(sf) else []
            for sep in ('.', '_'):
                if sf + sep in n:
                    base.append(n.replace(sf + sep, sep, 1))
            if any(b in (targets if outside_exog else keep) for b in base):
                keep.add(n)
                break
    total = len(set(A.series) | set(B.series))
    for sc in (A, B):
        sc.series = {n: v for n, v in sc.series.items() if n in keep}
    for E in (eqs, eqs_b):
        if E is not None:
            for n in [n for n in E if n not in keep]:
                del E[n]
    for y in [y for y in (addl or {}) if y not in keep]:
        del addl[y]
    if outside_exog:
        n_off = 0
        for E in (eqs, eqs_b):
            if E is not None:
                for n in [n for n in E if n not in targets]:
                    del E[n]
                    n_off += 1
        for y in [y for y in (addl or {}) if y not in targets]:
            del addl[y]
        say('--target-exog: %d variables outside the target feed it directly; they are taken as given '
            '(exogenous) from each scenario, %d of their equations switched off'
            % (len(keep - targets), n_off))
    say('--target: %d target variables; keeping %d of %d series (the targets and everything that feeds them)'
        % (len(targets), len(set(A.series) | set(B.series)), total))
    return targets


def load_exog(items):
    names = []
    for it in items or ():
        if os.path.isfile(it):
            if os.path.splitext(it)[1].lower() in ('.xlsx', '.xlsm'):
                text = ' '.join(str(c) for r in read_xlsx(it) for c in r)
            else:
                with open(it, 'r', encoding='utf-8-sig', errors='replace') as f:
                    text = f.read()
        else:
            text = it
        for tok in re.split(r'[\s,;]+', text):
            tok = tok.strip().strip('"\'').upper()
            if tok and not tok.startswith(("'", '#')):
                names.append(tok)
    return names


def _label(paths):
    return os.path.splitext(os.path.basename(paths[0]))[0]


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest='cmd')

    def common(sp):
        sp.add_argument('--a', action='append', required=True, metavar='FILE',
                        help='scenario A data (csv/tsv/xlsx). Repeat to combine several files.')
        sp.add_argument('--b', action='append', required=True, metavar='FILE', help='scenario B data')
        sp.add_argument('--eq-a', action='append', metavar='PATH',
                        help='equations of scenario A only, when A and B use different equations (repeatable)')
        sp.add_argument('--eq-b', action='append', metavar='PATH',
                        help='equations of scenario B only; differences from A are reported as "equation changed"')
        sp.add_argument('--eq', action='append', metavar='PATH',
                        help='equation text file(s) or a folder of them (EViews syntax). Repeatable.')
        sp.add_argument('--name-a', help='label for A (default: file name)')
        sp.add_argument('--name-b', help='label for B')
        sp.add_argument('--sheet', help='xlsx sheet name or number (default: first)')
        sp.add_argument('--meta', help='optional CSV with series name + description columns')
        sp.add_argument('--start', help='first period to compare (default: where scenarios first diverge)')
        sp.add_argument('--end', help='last period to compare (default: end of data)')
        sp.add_argument('--target', action='append', metavar='SPEC',
                        help='keep only these variables and everything that feeds them: a geography code (ICAN), '
                             'a wildcard ("*_ICAN"), names or a comma list; repeatable')
        sp.add_argument('--target-exog', action='store_true',
                        help='with --target: treat every variable outside the target as exogenous (taken from '
                             'each scenario), so only the target\'s own equations are solved')
        sp.add_argument('--area', action='append', metavar='SPEC',
                        help='area for the Layers tab: longest chains that end in these variables (geography code, '
                             'wildcard or names; default: --target)')
        sp.add_argument('--tol', type=float, default=1e-6, help='relative tolerance for "differs" (default 1e-6)')
        sp.add_argument('--geo-sep', default='auto',
                        help='separator before the geography code: "." (FGDP.ICAN), "_" (FGDP_ICAN) or auto (default)')
        sp.add_argument('--adj-threshold', type=float, default=0.25,
                        help='flag a variable as "mostly add-factor" above this share of its change (0.25)')
        sp.add_argument('--exog', action='append', metavar='LIST',
                        help='variables the modeller set exogenously although they have an equation: a file '
                             '(names separated by spaces, commas or new lines) or a comma-separated list. Repeatable.')
        sp.add_argument('--af-suffix', default='',
                        help='suffix of add-factor series, e.g. GDP.US_A (default _A; comma-separate several)')
        sp.add_argument('--quiet', action='store_true')

    def report_opts(sp):
        sp.add_argument('--embed', default='auto',
                        help='series with charts + full detail in the report: a number, "all", or auto '
                             '(all if at most 5000 series changed, else the top 1500). Other series still get '
                             'their per-quarter shares.')
        sp.add_argument('--max-decomp', type=int, default=50000,
                        help='max equations analysed for driver contributions (default 50000)')
        sp.add_argument('--context', type=int, default=8, help='periods of history shown before the window (8)')
        sp.add_argument('--focus', action='append', help='variables to always include in full detail: a name, a comma list, or a wildcard '
                             'like "*_ICAN" (quote it); repeatable')
        sp.add_argument('--key', action='append', metavar='LIST',
                        help='key variables for the Summary tab (file or comma list); default: picked automatically')
        sp.add_argument('--summary-n', type=int, default=8,
                        help='how many key variables to pick automatically for the Summary tab (default 8; 0 = off)')
        sp.add_argument('--no-importance', dest='importance', action='store_false',
                        help='skip the Importance tab (it re-solves the model once per scenario shock)')
        sp.add_argument('--max-shocks', type=int, default=0,
                        help='Importance tab, exact method: only the N broadest shocks (0 = all)')
        sp.add_argument('--importance-method', choices=['fast', 'exact'], default='fast',
                        help='Importance tab: fast (one pass, default) or exact (re-solve per shock)')
        sp.add_argument('--influence-n', type=int, default=20,
                        help='scenario shocks whose influence is simulated for the Influence tab (default 20)')

    sp = sub.add_parser('report', help='build the interactive HTML report (a single file)')
    common(sp)
    report_opts(sp)
    sp.add_argument('-o', '--out', default='scenario_explain.html')

    sp = sub.add_parser('serve', help='same report, served locally, with exact impact X->Y and shock attribution')
    common(sp)
    report_opts(sp)
    sp.add_argument('--port', type=int, default=8765)
    sp.add_argument('--no-browser', action='store_true')

    sp = sub.add_parser('top', help='print the biggest changes')
    common(sp)
    sp.add_argument('-n', type=int, default=30)
    sp.add_argument('--geo', help='only this geography code, e.g. IUSA')
    sp.add_argument('--match', help='regex on name or description')
    sp.add_argument('--type', help='X=assumption, F=add-factor series, A=mostly add-factor, E=formula, N=not evaluable')
    sp.add_argument('--csv', help='write stats for ALL differing series to this CSV')

    sp = sub.add_parser('explain', help='contribution of each input to the gap of a variable')
    sp.add_argument('var')
    common(sp)
    sp.add_argument('--at', help='period to explain (default: end of window)')
    sp.add_argument('--depth', type=int, default=1, help='levels of inputs to follow (default 1)')
    sp.add_argument('--min-share', type=float, default=0.1, help='only follow inputs above this share (0.1)')
    sp.add_argument('--by-period', action='store_true', help='also print contributions period by period')

    sp = sub.add_parser('addfactors', help='formula vs add-factor share of the gap, for every equation')
    common(sp)
    sp.add_argument('-n', type=int, default=40)
    sp.add_argument('--all', action='store_true', help='include equations whose add-factor did not change')

    sp = sub.add_parser('impact', help='can X affect Y, and by how much? (re-solves the equations)')
    sp.add_argument('x')
    sp.add_argument('y', nargs='?')
    common(sp)
    sp.add_argument('--shock', help='hypothetical change in X instead of its A->B change: "1%%" or "0.25" (units)')
    sp.add_argument('--by-period', action='store_true')
    sp.add_argument('-n', type=int, default=30, help='rows when Y is omitted')

    sp = sub.add_parser('influence', help='if X changes, how many variables move and by how much')
    sp.add_argument('x', nargs='?', help='variable (omit: table of all scenario shocks)')
    common(sp)
    sp.add_argument('--shock', help='hypothetical change instead of the A->B change: "1%%" or "0.25" (units)')
    sp.add_argument('--thr', type=float, default=0.01, help='count a variable as moved above this %% of its level (0.01)')
    sp.add_argument('-n', type=int, default=25)

    sp = sub.add_parser('attribute', help='split the gap of Y across all scenario shocks upstream of it')
    sp.add_argument('y')
    common(sp)
    sp.add_argument('-n', type=int, default=25)

    sp = sub.add_parser('path', help='the channels through which X reaches Y')
    sp.add_argument('x')
    sp.add_argument('y')
    common(sp)
    sp.add_argument('-k', type=int, default=5, help='max number of paths')

    sp = sub.add_parser('importance', help='rank the scenario shocks by how much of all variables they explain')
    common(sp)
    sp.add_argument('--weight', choices=['all', 'concept', 'key'], default='concept',
                    help='average over all variables, per concept (each concept counts once), or key variables (default concept)')
    sp.add_argument('--max-shocks', type=int, default=0, help='exact method: only the N broadest shocks (0 = all)')
    sp.add_argument('--method', choices=['fast', 'exact'], default='fast',
                    help='fast: trace shocks through each equation\'s own gap split (one pass, default); '
                         'exact: re-solve the model once per shock (slow with many shocks)')
    sp.add_argument('--csv', help='write the ranking, and every variable x shock share, to CSV')
    sp.add_argument('-o', '--out', help='also write the full HTML report, opened on the Importance tab '
                                        '(with --geo / --weight already selected)')
    sp.add_argument('--embed', default='auto', help='with -o: series with full charts in the report (as in report)')
    sp.add_argument('--geo', help='only average over variables of these geographies, e.g. ICAN (comma list)')
    sp.add_argument('--match', help='only average over variables whose name or description matches (regex)')
    sp.add_argument('-n', type=int, default=30)
    sp.add_argument('--key', action='append', metavar='LIST', help='key variables for the "key" weighting')

    sp = sub.add_parser('check', help='list equations that cannot be used, with the reason, in a CSV')
    common(sp)
    sp.add_argument('-o', '--out', default='not_evaluable.csv')

    sp = sub.add_parser('layers', help='how many layers of variables each variable affects (longest chain)')
    common(sp)
    sp.add_argument('-n', type=int, default=25)
    sp.add_argument('--match', help='only variables matching this regex')
    sp.add_argument('--by-area', action='store_true', help='rank by the longest chain into --area / --target')

    sp = sub.add_parser('mask', help='print equations with error-term brackets with names/numbers hidden (safe to share)')
    sp.add_argument('--eq', '--eq-a', '--eq-b', action='append', required=True, help='equation file(s) or folder')
    sp.add_argument('--failed', action='store_true', help='only lines that cannot be used')
    sp.add_argument('--match', help='only lines matching this regex (default: lines with [ or ar(/ma()')
    sp.add_argument('-n', type=int, default=20)

    sp = sub.add_parser('info', help='show what was loaded and parsed')
    common(sp)

    args = p.parse_args(argv)
    if not args.cmd:
        p.print_help()
        return 1
    if args.cmd == 'mask':
        cmd_mask(args)
        return 0
    M, eqstats, fails = build_model(args)
    if hasattr(args, 'key'):
        args.key_list = load_exog(args.key)
    if args.cmd == 'info':
        cmd_info(M, args, eqstats, fails)
    elif args.cmd == 'top':
        cmd_top(M, args)
    elif args.cmd == 'explain':
        cmd_explain(M, args)
    elif args.cmd == 'path':
        cmd_path(M, args)
    elif args.cmd == 'report':
        build_report(M, args, eqstats)
    elif args.cmd == 'layers':
        cmd_layers(M, args)
    elif args.cmd == 'importance':
        args.eqstats = eqstats
        cmd_importance(M, args)
    elif args.cmd == 'serve':
        cmd_serve(M, args, eqstats)
    elif args.cmd == 'addfactors':
        cmd_addfactors(M, args)
    elif args.cmd == 'impact':
        cmd_impact(M, args)
    elif args.cmd == 'attribute':
        cmd_attribute(M, args)
    elif args.cmd == 'influence':
        cmd_influence(M, args)
    elif args.cmd == 'check':
        cmd_check(M, args)
    return 0


if __name__ == '__main__':
    sys.exit(main())
