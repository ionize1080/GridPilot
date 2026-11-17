"""Bounded Excel 2010+ analysis on managed copies, using an isolated hidden Excel.

The existing workspace_analyze engine remains usable without Excel. This engine
creates native PivotTables/charts and uses Excel for workbook calculation. It never
attaches to the user's Excel instance or saves an input file in place.
"""
from contextlib import contextmanager
from datetime import datetime, timezone
import gc
import hashlib
import json
import math
import operator
from pathlib import Path
import re
import shutil
import statistics
import uuid
import xml.etree.ElementTree as ET

import workspace_files as files

MAX_CELLS, MAX_ROWS, MAX_COLUMNS = 2000000, 100001, 200
LEGACY_FUNCTIONS = set('SUM SUMIF SUMIFS SUMPRODUCT COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS AVERAGE AVERAGEIF AVERAGEIFS MIN MAX MEDIAN MODE STDEV STDEVP VAR VARP STDEV.S STDEV.P VAR.S VAR.P PERCENTILE QUARTILE PERCENTILE.INC QUARTILE.INC CORREL COVAR SLOPE INTERCEPT RSQ RANK RANK.EQ LARGE SMALL IF IFERROR AND OR NOT TRUE FALSE ABS ROUND ROUNDUP ROUNDDOWN INT MOD POWER SQRT LOG LN EXP SIGN SUBTOTAL AGGREGATE VLOOKUP HLOOKUP INDEX MATCH CHOOSE OFFSET ROW ROWS COLUMN COLUMNS ADDRESS INDIRECT LEFT RIGHT MID LEN TRIM CLEAN UPPER LOWER PROPER CONCATENATE TEXT VALUE SUBSTITUTE REPLACE FIND SEARCH EXACT DATE YEAR MONTH DAY TODAY NOW TIME HOUR MINUTE SECOND WEEKDAY WEEKNUM EDATE EOMONTH DATEDIF DAYS360 NETWORKDAYS WORKDAY ISBLANK ISNUMBER ISTEXT ISERROR ISNA NA N PV FV PMT NPV IRR SUMSQ PRODUCT'.split())
MODERN_PROBES = {
    'XLOOKUP': 'XLOOKUP(1,{1},{2})', 'XMATCH': 'XMATCH(1,{1})',
    'FILTER': 'SUM(FILTER({1;2},{1;0}))', 'SORT': 'SUM(SORT({2;1}))',
    'UNIQUE': 'SUM(UNIQUE({1;1}))', 'LET': 'LET(x,1,x+1)',
    'TEXTJOIN': 'TEXTJOIN(",",TRUE,"a","b")', 'IFS': 'IFS(TRUE,1)',
    'SWITCH': 'SWITCH(1,1,2)',
    'SEQUENCE': 'SUM(SEQUENCE(2))', 'SORTBY': 'SUM(SORTBY({1;2},{2;1}))',
    'TAKE': 'SUM(TAKE({1;2},1))', 'DROP': 'SUM(DROP({1;2},1))',
    'CHOOSECOLS': 'SUM(CHOOSECOLS({1,2},1))', 'CHOOSEROWS': 'SUM(CHOOSEROWS({1;2},1))',
    'VSTACK': 'SUM(VSTACK(1,2))', 'HSTACK': 'SUM(HSTACK(1,2))',
    'TOCOL': 'SUM(TOCOL({1,2}))', 'TOROW': 'SUM(TOROW({1;2}))',
    'TEXTSPLIT': 'COUNTA(TEXTSPLIT("a,b",","))', 'TEXTBEFORE': 'TEXTBEFORE("a,b",",")',
    'TEXTAFTER': 'TEXTAFTER("a,b",",")', 'LAMBDA': 'LAMBDA(x,x+1)(1)',
    'BYROW': 'SUM(BYROW({1;2},LAMBDA(x,SUM(x))))',
    'BYCOL': 'SUM(BYCOL({1,2},LAMBDA(x,SUM(x))))',
    'MAP': 'SUM(MAP({1;2},LAMBDA(x,x*2)))',
    'REDUCE': 'REDUCE(0,{1;2},LAMBDA(a,b,a+b))',
    'SCAN': 'SUM(SCAN(0,{1;2},LAMBDA(a,b,a+b)))',
}
AGGREGATES = {'sum': -4157, 'count': -4112, 'mean': -4106, 'average': -4106,
              'min': -4139, 'max': -4136, 'product': -4149, 'count_numbers': -4113,
              'stdev': -4155, 'stdevp': -4156, 'var': -4164, 'varp': -4165}
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def bounds(address):
    from openpyxl.utils.cell import range_boundaries
    if not isinstance(address, str) or not re.fullmatch(r'\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?', address):
        raise ValueError('请使用单个有限 A1 区域，例如 A1:F100；不接受整列、整行或外部引用。')
    c1, r1, c2, r2 = range_boundaries(address)
    if c2 < c1 or r2 < r1 or c2 > 16384 or r2 > 1048576 or r2-r1+1 > MAX_ROWS or c2-c1+1 > MAX_COLUMNS or (c2-c1+1)*(r2-r1+1) > MAX_CELLS:
        raise ValueError('区域超过 100001 行、200 列或 200 万单元格上限。')
    return c1, r1, c2, r2


def sheet_name(name):
    if not isinstance(name, str) or not name or len(name) > 31 or re.search(r'[\\/*?:\[\]\x00-\x1f]', name) or name.startswith("'") or name.endswith("'"):
        raise ValueError('工作表名称须为 1–31 个字符，不能包含路径或 Excel 保留字符。')
    return name


def validate_formula(expression, compatibility='excel2010', available=()):
    from openpyxl.formula import Tokenizer
    if not isinstance(expression, str) or not expression.startswith('=') or len(expression) > 2048 or any(c in expression for c in ('[', ']', '|', '\x00')):
        raise ValueError('公式须以 = 开头；不接受外部/结构化引用、DDE 或超过 2048 字符的公式。')
    allowed = LEGACY_FUNCTIONS | (set(available) if compatibility == 'installed' else set())
    try:
        for token in Tokenizer(expression).items:
            if token.type == 'OPERAND' and token.subtype == 'RANGE' and ('\\' in token.value or '://' in token.value or '!' in token.value and any(c in token.value.split('!')[0] for c in ('/', ':'))):
                raise ValueError('公式只能引用本工作簿的数据。')
            if token.type == 'FUNC' and token.subtype == 'OPEN':
                name = token.value[:-1].upper().removeprefix('_XLFN.').removeprefix('_XLWS.')
                if name not in allowed:
                    raise ValueError(f'当前兼容模式不支持公式函数 {name}。请用 transform 的查找/筛选/排序/去重，或先检查 Excel 能力。')
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError('公式无法解析。') from exc
    return expression


def check_input(path, compatibility='excel2010', available=()):
    if path.suffix.lower() != '.xlsx':
        raise ValueError('原生 Excel 操作接受 XLSX；旧 XLS、CSV 或 XLSM 可先用 workspace_analyze 生成 XLSX 底稿。')
    with files.checked_zip(path) as archive:
        for name in archive.namelist():
            low = name.lower()
            if any(x in low for x in ('vbaproject', 'macrosheets/', 'externalLinks/'.lower(), 'connections.xml', 'querytables/', 'embeddings/')):
                raise ValueError('此工作簿包含宏、外部连接或嵌入对象；请先另存为仅含本机数据的 XLSX 副本。')
            if low == '[content_types].xml':
                for part in ET.fromstring(archive.read(name)):
                    content_type = part.get('ContentType','').lower()
                    if any(marker in content_type for marker in ('macroenabled', 'vbaproject', 'macrosheet', 'oleobject', 'externallink', 'connections', 'querytable')):
                        raise ValueError('此工作簿包含宏、外部连接或嵌入对象。')
            if low.endswith('.rels'):
                for relationship in ET.fromstring(archive.read(name)):
                    if relationship.get('TargetMode','').lower() == 'external':
                        raise ValueError('此工作簿含外部链接，请使用仅含本机数据的 XLSX 副本。')
            if low.startswith('xl/') and low.endswith('.xml'):
                root = ET.fromstring(archive.read(name))
                for element in root.iter():
                    tag = element.tag.rsplit('}', 1)[-1]
                    if tag in ('f', 'definedName') and element.text:
                        validate_formula('=' + element.text.lstrip('='), compatibility, available)
                    if tag == 'cacheSource' and element.get('type') != 'worksheet':
                        raise ValueError('仅支持工作簿内数据源的普通数据透视表。')
                    if tag == 'dimension' and element.get('ref'):
                        bounds(element.get('ref'))


@contextmanager
def application():
    import pythoncom
    import win32api
    import win32com.client
    import win32job
    import win32process
    pythoncom.CoInitialize()
    excel, job, process = None, None, None
    prior = set(win32process.EnumProcesses())
    try:
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
        except Exception as exc:
            raise ValueError('原生透视表和完整公式计算需要本机安装 Excel 2010 或更新版本；离线分析仍可使用 workspace_analyze。') from exc
        _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        if pid in prior:
            excel = None  # Never close or reconfigure an existing user process.
            raise ValueError('无法创建独立 Excel 进程，已停止以保留现有工作簿。')
        job = win32job.CreateJobObject(None, 'GridPilot-Excel-'+uuid.uuid4().hex)
        limits = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
        limits['BasicLimitInformation']['LimitFlags'] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, limits)
        process = win32api.OpenProcess(0x0100 | 0x0001, False, pid)
        win32job.AssignProcessToJobObject(job, process)
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3
        if float(excel.Version) < 14:
            raise ValueError('原生 Excel 操作最低需要 Excel 2010。')
        yield excel
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        excel = None
        gc.collect()
        if job is not None:
            job.Close()  # Also terminates only this dedicated Excel on worker timeout.
        if process is not None:
            process.Close()
        pythoncom.CoUninitialize()


def capabilities(excel):
    supported = []
    for name, expression in MODERN_PROBES.items():
        try:
            value = excel.Evaluate('=' + expression)
            if not (isinstance(value, int) and value < -2000000000):
                supported.append(name)
        except Exception:
            pass
    return dict(engine='Microsoft Excel', version=str(excel.Version), visible=bool(excel.Visible),
                minimum_version='Excel 2010 Standard', pivot_version=4,
                modern_functions=supported, compatibility_default='excel2010',
                compatible_alternatives=['exact lookup/join', 'filter', 'sort', 'unique', 'fill missing', 'descriptive statistics', 'correlation'],
                note='Modern functions are probed, not inferred from version 16. Default output uses Excel 2010 formulas and classic native PivotTables. Compatible transforms write values; they do not emulate dynamic spill formulas, Power Query, Power Pivot, Python in Excel or Copilot.')


def info(workspace_id):
    files.workspace(workspace_id)
    with application() as excel:
        book = excel.Workbooks.Add(-4167)
        try:
            return capabilities(excel)
        finally:
            book.Close(False)


def safe_value(value):
    if isinstance(value, (datetime,)):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    return value


def read(workspace_id, file_id, sheet='', cell_range='', max_rows=50, max_columns=20, expected_revision=None):
    import openpyxl
    path, stamp = files.validated(workspace_id, file_id, expected_revision)
    if path.suffix.lower() != '.xlsx':
        raise ValueError('请先用 workspace_analyze 将数据保存为 XLSX 底稿。')
    if type(max_rows) is not int or not 1 <= max_rows <= 100 or type(max_columns) is not int or not 1 <= max_columns <= 50:
        raise ValueError('读取预览限制为 100 行、50 列。')
    with files.checked_zip(path):
        pass
    with __import__('contextlib').ExitStack() as stack:
        book = openpyxl.load_workbook(path, read_only=False, data_only=False, keep_links=False)
        stack.callback(book.close)
        cached = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_links=False)
        stack.callback(cached.close)
        sheets = [dict(name=ws.title, rows=ws.max_row, columns=ws.max_column, tables=list(ws.tables),
                       pivots=[dict(name=p.name, range=p.location.ref, cache_source=p.cache.cacheSource.worksheetSource.ref if p.cache and p.cache.cacheSource.worksheetSource else None) for p in ws._pivots], charts=len(ws._charts)) for ws in book]
        result = dict(file_id=file_id, revision=stamp, sheets=sheets)
        if sheet or cell_range:
            selected = sheet or book.sheetnames[0]
            if selected not in book.sheetnames:
                raise ValueError('工作表不存在。')
            ws, values = book[selected], cached[selected]
            requested = cell_range or f'A1:{openpyxl.utils.get_column_letter(min(ws.max_column, MAX_COLUMNS))}{min(ws.max_row, MAX_ROWS)}'
            c1, r1, c2, r2 = bounds(requested)
            end_row, end_col = min(r2, r1+max_rows-1), min(c2, c1+max_columns-1)
            rows, formulas, errors = [], [], []
            for r in range(r1, end_row+1):
                row = []
                for c in range(c1, end_col+1):
                    cell, cached_cell = ws.cell(r, c), values.cell(r, c)
                    v = safe_value(cached_cell.value if cell.data_type == 'f' else cell.value)
                    row.append(v[:500] if isinstance(v, str) else v)
                    if cell.data_type == 'f':
                        formulas.append(dict(cell=cell.coordinate, formula=getattr(cell.value,'text',cell.value)))
                    if cached_cell.data_type == 'e':
                        errors.append(dict(cell=cell.coordinate, error=cached_cell.value))
                rows.append(row)
            result.update(sheet=selected, requested_range=requested, returned_range=f'{openpyxl.utils.get_column_letter(c1)}{r1}:{openpyxl.utils.get_column_letter(end_col)}{end_row}',
                          values=rows, formulas=formulas, errors=errors, truncated=end_row<r2 or end_col<c2,
                          next_row=end_row+1 if end_row<r2 else None, next_column=end_col+1 if end_col<c2 else None)
        files.validated(workspace_id, file_id, stamp)
        return result


def worksheet(book, name):
    sheet_name(name)
    try:
        return book.Worksheets(name)
    except Exception as exc:
        raise ValueError('工作表不存在：' + name) from exc


def matrix(rng):
    value = rng.Value2
    return [[value]] if rng.Count == 1 else [list(row) for row in value]


def write_values(ws, address, values):
    from openpyxl.utils import get_column_letter
    c, r, c2, r2 = bounds(address)
    if c != c2 or r != r2:
        raise ValueError('写入位置须为单个起始单元格。')
    if not isinstance(values, list) or not values or not isinstance(values[0], list) or not values[0]:
        raise ValueError('values 须为非空二维数组。')
    width = len(values[0])
    if any(not isinstance(row, list) or len(row) != width for row in values):
        raise ValueError('写入数据的每一行须有相同列数。')
    target = f'{get_column_letter(c)}{r}:{get_column_letter(c+width-1)}{r+len(values)-1}'
    bounds(target)
    def literal(v):
        if v is None or type(v) in (bool, int):
            return v
        if type(v) is float and math.isfinite(v):
            return v
        if isinstance(v, str) and len(v) <= 32000:
            return "'" + v  # Value2 must never execute model/source text as a formula.
        raise ValueError('单元格仅支持有限数字、布尔值、文本和空值；文本最多 32000 字符。')
    ws.Range(target).Value2 = tuple(tuple(literal(v) for v in row) for row in values)
    return target


def table_data(book, item, prefix=''):
    ws = worksheet(book, item[prefix+'sheet'])
    address = item[prefix+'range']
    bounds(address)
    if ws.Evaluate('SUMPRODUCT(--ISERROR('+address+'))'):
        raise ValueError('数据区域含 Excel 错误值，请先修复公式或缩小范围。')
    rows = matrix(ws.Range(address))
    headers = rows[0]
    if not headers or any(not isinstance(h, str) or not h.strip() for h in headers) or len({h.casefold() for h in headers}) != len(headers):
        raise ValueError('表格首行须为不重复的非空文本列名。')
    return headers, [dict(zip(headers, row)) for row in rows[1:]]


def key(value, case_sensitive=False):
    if isinstance(value, str):
        return ('text', value if case_sensitive else value.casefold())
    if value is None or value == '':
        return ('blank', '')
    return ('bool' if isinstance(value, bool) else 'number', value)


def transform(headers, records, item, lookup_data=None):
    records = [dict(r) for r in records]
    headers = list(headers)
    def field(name):
        if name not in headers:
            raise ValueError('未知列：' + str(name))
        return name
    for name in item.get('trim', []):
        field(name)
        for row in records:
            if isinstance(row[name], str):
                row[name] = row[name].strip()
    for name in item.get('numeric', []):
        field(name)
        for row in records:
            if row[name] not in (None, ''):
                try:
                    v = float(row[name])
                    if isinstance(row[name], bool) or not math.isfinite(v):
                        raise ValueError()
                    row[name] = v
                except (ValueError, TypeError) as exc:
                    raise ValueError('数值转换失败：' + name) from exc
    for fill in item.get('fill', []):
        field(fill['column'])
        for row in records:
            if row[fill['column']] in (None, ''):
                row[fill['column']] = fill.get('value')
    lookup = item.get('lookup')
    if lookup:
        left = field(lookup['left_key'])
        right_headers, right_records = lookup_data
        right_key, selected = lookup['right_key'], lookup['columns']
        if right_key not in right_headers or not selected or any(c not in right_headers or c in headers for c in selected):
            raise ValueError('查找表的键/返回列无效，或返回列与现有列重名。')
        policy = lookup.get('duplicates', 'first')
        if policy not in ('first', 'last', 'error'):
            raise ValueError('duplicates 须为 first、last 或 error。')
        mapping = {}
        for row in right_records:
            k = key(row[right_key], lookup.get('case_sensitive', False))
            if k in mapping and policy == 'error':
                raise ValueError('查找表含重复键。')
            if k not in mapping or policy == 'last':
                mapping[k] = row
        for row in records:
            match = mapping.get(key(row[left], lookup.get('case_sensitive', False)), {})
            row.update({c: match.get(c, lookup.get('if_not_found')) for c in selected})
        headers.extend(selected)
    ops = {'eq': operator.eq, 'ne': operator.ne, 'gt': operator.gt, 'ge': operator.ge, 'lt': operator.lt, 'le': operator.le,
           'contains': lambda a, b: str(b).casefold() in str(a or '').casefold(), 'in': lambda a, b: a in b,
           'blank': lambda a, b: a is None or a == '', 'not_blank': lambda a, b: a is not None and a != ''}
    for f in item.get('filters', []):
        name = field(f['column'])
        if f.get('op') not in ops:
            raise ValueError('未知筛选操作。')
        try:
            records = [r for r in records if ops[f['op']](r[name], f.get('value'))]
        except TypeError as exc:
            raise ValueError('筛选值类型不一致，请先 numeric 转换或填补缺失值。') from exc
    unique = item.get('unique_by')
    if unique:
        for name in unique:
            field(name)
        seen, kept = set(), []
        for row in records:
            k = tuple(key(row[c], item.get('case_sensitive', False)) for c in unique)
            if k not in seen:
                seen.add(k)
                kept.append(row)
        records = kept
    for order in reversed(item.get('sort_by', [])):
        name = field(order['column'])
        present, missing = [r for r in records if r[name] not in (None, '')], [r for r in records if r[name] in (None, '')]
        present.sort(key=lambda r: key(r[name]), reverse=bool(order.get('descending', False)))
        records = present + missing
    selected = item.get('select', headers)
    if not selected or len(set(selected)) != len(selected):
        raise ValueError('select 须包含不重复的现有列名。')
    for name in selected:
        field(name)
    return selected, [[row[c] for c in selected] for row in records]


def target_sheet(book, name):
    from pythoncom import Empty as Missing
    sheet_name(name)
    for ws in book.Worksheets:
        if ws.Name.casefold() == name.casefold():
            return ws
    ws = book.Worksheets.Add(Missing, book.Worksheets(book.Worksheets.Count))
    ws.Name = name
    return ws


def execute(book, excel, item, compatibility, available):
    from pythoncom import Empty as Missing
    op = item.get('op')
    if op == 'add_sheet':
        name = sheet_name(item['name'])
        if any(ws.Name.casefold() == name.casefold() for ws in book.Worksheets):
            raise ValueError('工作表已存在：' + name)
        target_sheet(book, name)
    elif op in ('rename_sheet', 'copy_sheet', 'delete_sheet'):
        ws = worksheet(book, item['sheet'])
        if op == 'rename_sheet':
            ws.Name = sheet_name(item['name'])
        elif op == 'copy_sheet':
            name = sheet_name(item['name'])
            if any(s.Name.casefold() == name.casefold() for s in book.Worksheets):
                raise ValueError('目标工作表已存在。')
            ws.Copy(Missing, book.Worksheets(book.Worksheets.Count))
            book.Worksheets(book.Worksheets.Count).Name = name
        else:
            if book.Worksheets.Count <= 1:
                raise ValueError('必须保留至少一张工作表。')
            ws.Delete()
    elif op == 'write_range':
        return dict(op=op, range=write_values(worksheet(book, item['sheet']), item.get('cell', 'A1'), item['values']))
    elif op in ('formula', 'clear_range', 'format', 'table', 'filter', 'sort', 'deduplicate'):
        ws = worksheet(book, item['sheet'])
        c1, r1, c2, r2 = bounds(item['range'])
        rng = ws.Range(item['range'])
        if op == 'formula':
            expression = validate_formula(item['formula'], compatibility, available)
            if compatibility == 'installed' and item.get('dynamic', False):
                rng.Formula2 = expression
            else:
                rng.Cells(1, 1).Formula = expression
                if rng.Columns.Count > 1:
                    rng.Rows(1).FillRight()
                if rng.Rows.Count > 1:
                    rng.FillDown()
        elif op == 'clear_range':
            rng.ClearContents()
        elif op == 'format':
            if 'number_format' in item:
                rng.NumberFormat = str(item['number_format'])[:200]
            if 'bold' in item:
                rng.Font.Bold = bool(item['bold'])
            if item.get('autofit'):
                rng.Columns.AutoFit()
        elif op == 'table':
            table_data(book, item)
            name = item['name']
            if not re.fullmatch(r'[A-Za-z_][A-Za-z_0-9]{0,79}', name) or re.fullmatch(r'[A-Za-z]{1,3}\d+', name):
                raise ValueError('表名称须为字母/下划线开头的标识符，不能是单元格地址。')
            table = ws.ListObjects.Add(1, rng, Missing, 1)
            table.Name = name
            table.TableStyle = 'TableStyleMedium2'
        elif op == 'filter':
            headers, _ = table_data(book, item)
            if ws.AutoFilterMode:
                ws.AutoFilterMode = False
            rng.AutoFilter()
            for f in item.get('filters', []):
                if f['column'] not in headers or f['op'] not in ('eq', 'ne', 'gt', 'ge', 'lt', 'le', 'contains', 'blank', 'not_blank'):
                    raise ValueError('筛选列或条件无效。')
                escaped = str(f.get('value', '')).replace('~', '~~').replace('*', '~*').replace('?', '~?')
                criteria = {'eq': '=', 'ne': '<>', 'gt': '>', 'ge': '>=', 'lt': '<', 'le': '<='}.get(f['op'], '') + escaped
                if f['op'] == 'contains': criteria = '*' + escaped + '*'
                if f['op'] == 'blank': criteria = '='
                if f['op'] == 'not_blank': criteria = '<>'
                rng.AutoFilter(headers.index(f['column'])+1, criteria)
        elif op == 'sort':
            headers, _ = table_data(book, item)
            sorter = ws.Sort
            sorter.SortFields.Clear()
            for order in item.get('sort_by', []):
                col = headers.index(order['column'])+1
                sorter.SortFields.Add(rng.Columns(col), 0, 2 if order.get('descending') else 1, Missing, 0)
            sorter.SetRange(rng)
            sorter.Header = 1
            sorter.Apply()
        elif op == 'deduplicate':
            headers, _ = table_data(book, item)
            columns = item.get('columns', headers)
            if not columns or any(c not in headers for c in columns): raise ValueError('去重列无效。')
            rng.RemoveDuplicates(tuple(headers.index(c)+1 for c in columns), 1)
    elif op in ('transform', 'describe', 'correlation'):
        excel.Calculate()
        headers, records = table_data(book, item)
        if op == 'transform':
            lookup = item.get('lookup')
            out_headers, rows = transform(headers, records, item, table_data(book, lookup) if lookup else None)
        elif op == 'describe':
            out_headers, rows = ['Column', 'Count', 'Missing', 'Sum', 'Mean', 'Median', 'Min', 'Max', 'StdDev', 'Q1', 'Q3'], []
            for name in item.get('columns', headers):
                if name not in headers: raise ValueError('未知统计列。')
                values = [r[name] for r in records if r[name] not in (None, '')]
                if any(type(v) not in (int, float) or not math.isfinite(v) for v in values): raise ValueError('描述统计须选择数值列；非数值不能静默忽略。')
                quartiles = statistics.quantiles(values, n=4, method='inclusive') if len(values)>1 else values*3
                rows.append([name, len(values), len(records)-len(values), sum(values), statistics.mean(values) if values else None, statistics.median(values) if values else None,
                             min(values) if values else None, max(values) if values else None, statistics.stdev(values) if len(values)>1 else None, quartiles[0] if values else None, quartiles[2] if values else None])
        else:
            names = item['columns']
            if len(names)!=2 or any(c not in headers for c in names): raise ValueError('相关分析须指定两列。')
            pairs = [[r[c] for c in names] for r in records if all(r[c] not in (None, '') for c in names)]
            if any(type(v) not in (int, float) or not math.isfinite(v) for row in pairs for v in row): raise ValueError('相关分析仅接受数值。')
            if len(pairs)<2: raise ValueError('相关分析至少需要两对有效数据。')
            try: correlation = statistics.correlation([r[0] for r in pairs], [r[1] for r in pairs])
            except statistics.StatisticsError as exc: raise ValueError('常数列无法计算 Pearson 相关系数。') from exc
            out_headers, rows = ['X', 'Y', 'Pearson', 'Pairs', 'MissingPairs'], [[*names, correlation, len(pairs), len(records)-len(pairs)]]
        target = target_sheet(book, item['target_sheet'])
        address = write_values(target, item.get('target_cell', 'A1'), [out_headers]+rows)
        return dict(op=op, sheet=target.Name, range=address, rows=len(rows), columns=out_headers, preview=rows[:10], preview_complete=len(rows)<=10,
                    note='Compatible computed values; refresh by rerunning this action when source data changes.')
    elif op == 'pivot':
        headers, _ = table_data(book, dict(sheet=item['source_sheet'], range=item['source_range']))
        row_fields, column_fields, page_fields = item.get('rows', []), item.get('columns', []), item.get('filters', [])
        fields = row_fields+column_fields+[f['field'] for f in page_fields]
        values = item.get('values', [])
        if len(set(fields))!=len(fields) or any(f not in headers for f in fields) or not values or len(values)>20:
            raise ValueError('透视表字段须存在且不重复，values 至少包含一个聚合指标。')
        if any(v.get('field') not in headers or v.get('aggregate') not in AGGREGATES for v in values): raise ValueError('透视表值字段或聚合方式无效；经典透视表不支持 distinct count。')
        destination = target_sheet(book, item['target_sheet'])
        location = item.get('target_cell', 'A3')
        c1,r1,c2,r2 = bounds(location)
        if c1!=c2 or r1!=r2: raise ValueError('透视表目标须为单个单元格。')
        if destination.Name == item['source_sheet']: raise ValueError('请把透视表放在独立工作表，避免覆盖源数据。')
        name = item.get('name', 'Pivot_'+uuid.uuid4().hex[:8])
        if not re.fullmatch(r'[A-Za-z_][A-Za-z_0-9]{0,79}', name): raise ValueError('透视表名称须为字母/下划线开头的标识符。')
        excel.Calculate()
        c1,r1,c2,r2 = bounds(item['source_range'])
        source_address = "'"+item['source_sheet'].replace("'","''")+f"'!R{r1}C{c1}:R{r2}C{c2}"
        cache = book.PivotCaches().Create(1, source_address, 4)
        pivot = cache.CreatePivotTable(destination.Range(location), name, True, 4)
        pivot.ManualUpdate = True
        for orientation, names in [(1,row_fields),(2,column_fields),(3,[f['field'] for f in page_fields])]:
            for position, field in enumerate(names,1):
                p = pivot.PivotFields(field)
                p.Orientation, p.Position = orientation, position
        for metric in values:
            field = pivot.AddDataField(pivot.PivotFields(metric['field']), metric.get('name') or metric['aggregate']+'_'+metric['field'], AGGREGATES[metric['aggregate']])
            field.NumberFormat = metric.get('number_format', '#,##0.00')
        pivot.RowAxisLayout(1)
        pivot.RowGrand, pivot.ColumnGrand = bool(item.get('grand_totals', True)), bool(item.get('grand_totals', True))
        pivot.ManualUpdate = False
        for f in page_fields:
            if 'value' in f: pivot.PivotFields(f['field']).CurrentPage = str(f['value'])
        cache.RefreshOnFileOpen = False
        pivot.RefreshTable()
        return dict(op=op, name=name, sheet=destination.Name, range=pivot.TableRange2.Address.replace('$',''), native=True, pivot_version=4,
                    note='Native Excel 2010 PivotTable with saved cache; adjustable and refreshable in Excel. Use refresh_pivot with a new source range after adding rows.')
    elif op == 'refresh_pivot':
        ws = worksheet(book, item['sheet'])
        pivot = ws.PivotTables(item['name'])
        excel.Calculate()
        if item.get('source_range'):
            table_data(book, dict(sheet=item['source_sheet'], range=item['source_range']))
            c1,r1,c2,r2 = bounds(item['source_range'])
            source_address = "'"+item['source_sheet'].replace("'","''")+f"'!R{r1}C{c1}:R{r2}C{c2}"
            pivot.ChangePivotCache(book.PivotCaches().Create(1, source_address, 4))
        pivot.PivotCache().Refresh()
        pivot.RefreshTable()
    elif op == 'chart':
        ws = worksheet(book, item['sheet'])
        bounds(item['range'])
        types = {'column':51, 'bar':57, 'line':4, 'pie':5, 'scatter':-4169, 'area':1}
        if item.get('type','column') not in types: raise ValueError('图表支持 column/bar/line/pie/scatter/area。')
        target = target_sheet(book, item.get('target_sheet', item['sheet']))
        chart = target.ChartObjects().Add(320, 20, 600, 340).Chart
        chart.SetSourceData(ws.Range(item['range']))
        chart.ChartType = types[item.get('type','column')]
        chart.HasTitle = True
        chart.ChartTitle.Text = str(item.get('title','分析图表'))[:100]
    else:
        raise ValueError('未知 Excel 操作：' + str(op))
    return dict(op=op, completed=True)


def edit(workspace_id, actions, file_id='', output_name='Excel分析底稿', expected_revision=None, compatibility='excel2010'):
    from pythoncom import Empty as Missing
    files.workspace(workspace_id)
    if not isinstance(actions,list) or not 1<=len(actions)<=40 or any(not isinstance(a,dict) for a in actions):
        raise ValueError('每次接受 1–40 项声明式操作。')
    if compatibility not in ('excel2010','installed'):
        raise ValueError('compatibility 须为 excel2010 或 installed。')
    folder = files.home()/'results'/workspace_id
    for part in [folder,*folder.parents]:
        if part.is_symlink() or (hasattr(part,'is_junction') and part.is_junction()): raise ValueError('底稿目录不能是链接。')
    folder.mkdir(parents=True,exist_ok=True)
    label = re.sub(r'[<>:"/\\|?*\x00-\x1f]','_',str(output_name)).strip(' .')[:70] or 'Excel分析底稿'
    target = folder/(label+'-'+uuid.uuid4().hex[:10]+'.xlsx')
    temporary = folder/('.'+uuid.uuid4().hex+'.xlsx')
    source, stamp, provenance = None, None, None
    if file_id:
        source, stamp = files.validated(workspace_id,file_id,expected_revision)
        provenance = dict(file_id=file_id,name=source.name,revision=stamp,sha256=hashlib.sha256(source.read_bytes()).hexdigest())
    try:
        with application() as excel:
            probe = excel.Workbooks.Add(-4167)
            caps = capabilities(excel)
            probe.Close(False)
            if source:
                check_input(source,compatibility,caps['modern_functions'])
                shutil.copyfile(source,temporary)
                files.validated(workspace_id,file_id,stamp)
                book = excel.Workbooks.Open(str(temporary),0,False,Missing,'','',True,Missing,Missing,False,False,Missing,False)
            else:
                book = excel.Workbooks.Add(-4167)
                book.Worksheets(1).Name = 'Sheet1'
            try:
                results = []
                for i,item in enumerate(actions):
                    try:
                        results.append(execute(book,excel,item,compatibility,caps['modern_functions']))
                    except Exception as exc:
                        detail = str(exc) if isinstance(exc,(ValueError,KeyError,IndexError)) else 'Excel 拒绝此操作，请检查字段、名称、区域及版本支持。'
                        raise ValueError(f'第 {i+1} 项 {item.get("op")} 失败：{detail}；本批次未保存。') from exc
                excel.CalculateFull()
                for item in actions:
                    if item.get('op') == 'formula':
                        ws = worksheet(book,item['sheet'])
                        if ws.Evaluate('SUMPRODUCT(--ISERROR('+item['range']+'))'):
                            raise ValueError('新写入公式产生 Excel 错误值；本批次未保存，请检查公式、数据或 Excel 版本。')
                audit_name = 'GridPilot_Audit'
                existing = {ws.Name.casefold() for ws in book.Worksheets}
                count = 1
                while audit_name.casefold() in existing:
                    count += 1
                    audit_name = 'GridPilot_Audit_'+str(count)
                audit = target_sheet(book,audit_name)
                write_values(audit,'A1',[['Item','Value'],['Created UTC',datetime.now(timezone.utc).isoformat()],['Source',json.dumps(provenance,ensure_ascii=False)],['Compatibility',compatibility],['Excel version',caps['version']]])
                write_values(audit,'A7',[['Action','Parameters']]+[[a['op'],json.dumps(a,ensure_ascii=False)[:30000]] for a in actions])
                # In-place source files are never opened or saved, even on failure.
                book.CheckCompatibility = False
                book.SaveAs(str(temporary),51,Missing,Missing,False,False,1,Missing,False)
            finally:
                book.Close(False)
        if source:
            files.validated(workspace_id,file_id,stamp)
        temporary.replace(target)
        identifier = files.file_id({'id':'results'},target.name)
        try:
            result = read(workspace_id,identifier)
        except Exception:
            target.unlink(missing_ok=True)
            raise
        result.update(saved=True,name=target.name,source=provenance,actions=results,engine=caps,compatibility=compatibility,
                      note='Saved as a new managed copy. Original and previous versions remain unchanged. Default: Excel 2010-compatible classic native pivots, charts and formulas; installed mode may require newer Excel.')
        return result
    finally:
        temporary.unlink(missing_ok=True)
