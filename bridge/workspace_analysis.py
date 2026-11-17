"""Offline, declarative spreadsheet analysis with retained, auditable XLSX papers.

No Excel process, eval, shell, macros, external links, or source-file writes.
"""
import ast
import csv
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
import hashlib
import json
import operator
from pathlib import Path
import re
import uuid
import workspace_files as files

MAX_ROWS, MAX_CELLS = 100000, 2000000


def number(value):
    if isinstance(value, bool) or value is None or str(value).strip() == '':
        raise ValueError('数值列含空值或非数值。请清洗数据，或显式使用 missing=skip。')
    try:
        result = Decimal(str(value))
    except InvalidOperation as exc:
        raise ValueError(f'不能将值 {str(value)[:40]} 作为数值。') from exc
    if not result.is_finite():
        raise ValueError('计算不接受 NaN 或无穷值。')
    return result


def scalar(value):
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (datetime,)):
        return value.isoformat()
    return value


def formula(expression, columns):
    if not isinstance(expression, str) or len(expression) > 500:
        raise ValueError('计算表达式长度须不超过 500。')
    tree = ast.parse(expression, mode='eval')
    allowed = (ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Name, ast.Load, ast.Add, ast.Sub, ast.Mult, ast.Div, ast.USub, ast.UAdd, ast.Call)
    if len(list(ast.walk(tree))) > 100 or any(not isinstance(n, allowed) for n in ast.walk(tree)):
        raise ValueError('计算仅支持列名、col("列名")、+ - * /、abs 和 round。')
    def evaluate(node, row, excel=False, row_number=2):
        from openpyxl.utils import get_column_letter
        def column(name):
            if name not in columns:
                raise ValueError('未知计算列：' + str(name))
            return f'{get_column_letter(columns.index(name) + 1)}{row_number}' if excel else number(row[name])
        if isinstance(node, ast.Expression):
            return evaluate(node.body, row, excel, row_number)
        if isinstance(node, ast.Constant) and type(node.value) in (int, float):
            return str(node.value) if excel else number(node.value)
        if isinstance(node, ast.Name):
            return column(node.id)
        if isinstance(node, ast.UnaryOp):
            value = evaluate(node.operand, row, excel, row_number)
            return ('-' if isinstance(node.op, ast.USub) else '+') + '(' + value + ')' if excel else -value if isinstance(node.op, ast.USub) else value
        if isinstance(node, ast.BinOp):
            left, right = evaluate(node.left, row, excel, row_number), evaluate(node.right, row, excel, row_number)
            operations = {ast.Add: (operator.add, '+'), ast.Sub: (operator.sub, '-'), ast.Mult: (operator.mul, '*'), ast.Div: (operator.truediv, '/')}
            op, symbol = operations[type(node.op)]
            if excel:
                return f'({left}{symbol}{right})'
            if isinstance(node.op, ast.Div) and right == 0:
                raise ValueError('计算出现除以零。')
            return op(left, right)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and not node.keywords:
            name = node.func.id
            if name == 'col' and len(node.args) == 1 and isinstance(node.args[0], ast.Constant) and isinstance(node.args[0].value, str):
                return column(node.args[0].value)
            if name == 'abs' and len(node.args) == 1:
                value = evaluate(node.args[0], row, excel, row_number)
                return f'ABS({value})' if excel else abs(value)
            if name == 'round' and len(node.args) == 2 and isinstance(node.args[1], ast.Constant) and type(node.args[1].value) is int and -6 <= node.args[1].value <= 12:
                value = evaluate(node.args[0], row, excel, row_number)
                digits = node.args[1].value
                return f'ROUND({value},{digits})' if excel else value.quantize(Decimal(10) ** -digits, rounding=ROUND_HALF_UP)
        raise ValueError('不支持的计算表达式。只支持 col、abs、round 和基础四则运算。')
    return lambda row, excel=False, row_number=2: evaluate(tree, row, excel, row_number)


def load_table(path, sheet, header_row):
    import openpyxl
    source_formulas = []
    ext = path.suffix.lower()
    if ext in {'.xlsx', '.xlsm'}:
        with files.checked_zip(path):
            pass
        book = openpyxl.load_workbook(path, read_only=True, data_only=False, keep_links=False)
        try:
            if sheet and sheet not in book.sheetnames:
                raise ValueError('工作表不存在。')
            ws = book[sheet] if sheet else book.worksheets[0]
            selected = ws.title
            rows, formula_cells = [], set()
            for index, cells in enumerate(ws.iter_rows(), 1):
                if index > MAX_ROWS + 1 or index * len(cells) > MAX_CELLS:
                    raise ValueError('分析范围超过 10 万行或 200 万单元格，请先拆分。')
                for cell in cells:
                    if cell.data_type == 'e': raise ValueError('源工作表含 Excel 错误值，请先修复或选择其他工作表。')
                    if cell.data_type == 'f':
                        if not isinstance(cell.value, str): raise ValueError('暂不支持数组或数据表公式。')
                        formula_cells.add(cell.coordinate)
                rows.append([scalar(c.value) for c in cells])
            from spreadsheet_calc import FormulaGrid
            grid = FormulaGrid(rows, formula_cells)
            from openpyxl.utils import get_column_letter
            for r, row in enumerate(rows):
                for c, value in enumerate(row):
                    coordinate = f'{get_column_letter(c + 1)}{r + 1}'
                    if coordinate in formula_cells:
                        source_formulas.append([coordinate, value, scalar(grid.cell(coordinate))])
            rows = [[scalar(grid.cell(f'{get_column_letter(c + 1)}{r + 1}')) for c in range(len(row))] for r, row in enumerate(rows)]
            rows = rows[header_row - 1:]
        finally:
            book.close()
    elif ext in {'.csv', '.tsv'}:
        with path.open(encoding='utf-8-sig', newline='') as f:
            allrows = csv.reader(f, delimiter='\t' if ext == '.tsv' else ',')
            rows = []
            for index, row in enumerate(allrows, 1):
                if index < header_row:
                    continue
                if len(rows) > MAX_ROWS or (len(rows) + 1) * len(row) > MAX_CELLS:
                    raise ValueError('分析范围过大，请先拆分。')
                rows.append(row)
        selected = 'CSV' if ext == '.csv' else 'TSV'
    elif ext == '.xls':
        import xlrd
        book = xlrd.open_workbook(str(path), on_demand=True)
        try:
            ws = book.sheet_by_name(sheet) if sheet else book.sheet_by_index(0)
            if ws.nrows > MAX_ROWS + 1 or ws.nrows * ws.ncols > MAX_CELLS:
                raise ValueError('分析范围过大，请先拆分。')
            rows = [ws.row_values(i) for i in range(header_row - 1, ws.nrows)]
            selected = ws.name
        finally:
            book.release_resources()
        # xlrd exposes cached values. Make the provenance explicit, never imply recalc.
    else:
        raise ValueError('数据分析输入支持 XLSX / XLSM / XLS / CSV / TSV。')
    if not rows:
        raise ValueError('工作表没有数据。')
    return [str(c).strip() if c is not None else '' for c in rows[0]], rows[1:], selected, source_formulas


def analyze(workspace_id, file_id='', sheet='', header_row=1, columns=None, rows=None, calculations=None,
            filters=None, group_by=None, pivot_column='', metrics=None, missing='error', output_name='分析底稿'):
    workspace = files.workspace(workspace_id)
    if type(header_row) is not int or header_row < 1 or missing not in ('error', 'skip'):
        raise ValueError('表头行号或空值处理方式无效。')
    calculations, filters, group_by, metrics = calculations or [], filters or [], group_by or [], metrics or []
    if len(calculations) > 20 or len(filters) > 20 or len(group_by) > 6 or len(metrics) > 20:
        raise ValueError('分析操作过多，请拆分为多个底稿。')
    source, path, stamp, source_formulas = {'kind': 'model_supplied_data'}, None, None, []
    if file_id:
        path, stamp = files.validated(workspace_id, file_id)
        columns, rows, selected_sheet, source_formulas = load_table(path, sheet, header_row)
        with path.open('rb') as f:
            digest = hashlib.file_digest(f, 'sha256').hexdigest()
        source = dict(name=path.name, file_id=file_id, sheet=selected_sheet, header_row=header_row, revision=stamp, sha256=digest, recalculatedFormulas=len(source_formulas),
                      values='cached legacy XLS values; not recalculated' if path.suffix.lower() == '.xls' else 'literal values and supported same-sheet formulas recalculated; stale caches ignored')
    if not isinstance(columns, list) or not 1 <= len(columns) <= 200 or not all(isinstance(c, str) and 0 < len(c) <= 100 for c in columns) or len(set(columns)) != len(columns):
        raise ValueError('表头须为 1–200 个不重复的非空列名。')
    if not isinstance(rows, list) or not 1 <= len(rows) <= MAX_ROWS or len(rows) * len(columns) > MAX_CELLS:
        raise ValueError('数据须为 1–100000 行，总计不超过 200 万单元格。')
    if any(not isinstance(row, list) or len(row) > len(columns) for row in rows):
        raise ValueError('数据行列数超过表头。')
    if any(not isinstance(v, (str, int, float, bool, type(None))) or isinstance(v, str) and len(v) > 32000 for row in rows for v in row):
        raise ValueError('单元格仅支持文本、数字、布尔值或空值，文本不超过 32000 字符。')
    original_columns = list(columns)
    original = [row + [None] * (len(columns) - len(row)) for row in rows]
    data, excluded, skipped = [], 0, 0
    operators = {'eq': operator.eq, 'ne': operator.ne, 'gt': operator.gt, 'ge': operator.ge, 'lt': operator.lt, 'le': operator.le, 'contains': lambda a, b: str(b) in str(a)}
    for f in filters:
        if f.get('column') not in columns or f.get('op') not in operators:
            raise ValueError('筛选条件列名或运算符无效。')
    compiled = []
    for item in calculations:
        name = item.get('name')
        if not isinstance(name, str) or not name or len(name) > 100 or name in columns:
            raise ValueError('计算列须使用新的非空列名。')
        compiled.append((name, formula(item.get('expression'), list(columns))))
        columns.append(name)
    with localcontext() as context:
        context.prec = 28
        for row in original:
            record = dict(zip(original_columns, row))
            include = True
            for f in filters:
                a, b = record[f['column']], f.get('value')
                if f['op'] in ('gt', 'ge', 'lt', 'le'):
                    a, b = number(a), number(b)
                if not operators[f['op']](a, b):
                    include = False
                    break
            if not include:
                excluded += 1
                continue
            for name, compute in compiled:
                try:
                    record[name] = compute(record)
                except (ValueError, InvalidOperation, ZeroDivisionError):
                    if missing == 'error':
                        raise
                    record[name] = None
                    skipped += 1
            data.append(record)
        if any(c not in columns for c in group_by) or pivot_column and pivot_column not in columns:
            raise ValueError('分组或透视列不存在。')
        if pivot_column and not metrics:
            raise ValueError('数据透视需要至少一个指标。')
        valid_ops = {'sum', 'mean', 'min', 'max', 'count', 'count_distinct'}
        metric_names = []
        for m in metrics:
            if m.get('op') not in valid_ops or m.get('column') not in columns:
                raise ValueError('指标只支持 sum / mean / min / max / count / count_distinct，并须指定现有列。')
            name = m.get('name') or m['op'] + '_' + m['column']
            if name in metric_names or name in group_by:
                raise ValueError('指标名称重复。')
            metric_names.append(name)
        buckets, categories = {}, []
        for record in data:
            key = tuple(str(record[c]) if record[c] is not None else '(空)' for c in group_by)
            category = str(record[pivot_column]) if pivot_column else ''
            if category not in categories:
                categories.append(category)
            if len(categories) > 50 or len(buckets) > 2000:
                raise ValueError('分组或透视类别过多，请增加筛选条件。')
            buckets.setdefault(key, {}).setdefault(category, []).append(record)
        summary_columns = group_by + [f'{category} · {name}' if pivot_column else name for category in categories for name in metric_names]
        summary = []
        if metrics:
            for key, category_rows in buckets.items():
                result = list(key)
                for category in categories:
                    records = category_rows.get(category, [])
                    for metric in metrics:
                        values = [r[metric['column']] for r in records]
                        op = metric['op']
                        if op == 'count':
                            value = len([v for v in values if v is not None and v != ''])
                        elif op == 'count_distinct':
                            value = len(set(str(v) for v in values if v is not None and v != ''))
                        else:
                            numeric = []
                            for v in values:
                                try:
                                    numeric.append(number(v))
                                except ValueError:
                                    if missing == 'error':
                                        raise
                                    skipped += 1
                            value = (sum(numeric, Decimal(0)) if op == 'sum' else sum(numeric, Decimal(0)) / len(numeric) if op == 'mean' else min(numeric) if op == 'min' else max(numeric)) if numeric else None
                        result.append(scalar(value))
                summary.append(result)
    if path:
        files.validated(workspace_id, file_id, stamp)
    # Output goes only to a managed, per-workspace directory. Names cannot escape.
    folder = files.home() / 'results' / workspace_id
    for ancestor in [folder, *folder.parents]:
        if ancestor.is_symlink() or (hasattr(ancestor, 'is_junction') and ancestor.is_junction()):
            raise ValueError('底稿目录不能是链接。')
    folder.mkdir(parents=True, exist_ok=True)
    if folder.is_symlink() or (hasattr(folder, 'is_junction') and folder.is_junction()):
        raise ValueError('底稿目录不能是链接。')
    label = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(output_name)).strip(' .')[:70] or '分析底稿'
    created = datetime.now(timezone.utc).isoformat()
    target = folder / (label + '-' + datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:6] + '.xlsx')
    audit = dict(createdUTC=created, source=source, inputRows=len(original), includedRows=len(data), filteredRows=excluded,
                 skippedValues=skipped, missing=missing, calculations=calculations, filters=filters, group_by=group_by,
                 pivot_column=pivot_column, metrics=metrics, engine='GridPilot Decimal(28), ROUND_HALF_UP; no Excel process',
                 limitations='Pivot is a static computed cross-tab, not a native Excel PivotTable. Supported same-sheet XLSX formulas are recalculated; unsupported formulas fail. Legacy XLS uses cached values. Excel numeric storage uses up to 15 significant digits. Source formula text is preserved separately and is not executed in the paper. Review source provenance before writing conclusions.')
    import xlsxwriter
    temporary = target.with_suffix('.tmp')
    try:
        with xlsxwriter.Workbook(str(temporary), {'strings_to_formulas': False, 'strings_to_urls': False, 'nan_inf_to_errors': False}) as book:
            header = book.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#275D80', 'text_wrap': True, 'valign': 'vcenter'})
            numeric_format = book.add_format({'num_format': '#,##0.00;[Red](#,##0.00)', 'font_color': '#176A49'})
            def table(name, headers, records, calculations_sheet=False):
                ws = book.add_worksheet(name)
                ws.freeze_panes(1, 0)
                ws.set_row(0, 28)
                ws.set_column(0, max(0, len(headers) - 1), 19)
                for c, h in enumerate(headers):
                    ws.write_string(0, c, h, header)
                for r, record in enumerate(records, 1):
                    values = [record.get(c) for c in headers] if isinstance(record, dict) else record
                    for c, value in enumerate(values):
                        if calculations_sheet and headers[c] in dict(compiled) and value is not None:
                            expression = dict(compiled)[headers[c]](record, True, r + 1)
                            ws.write_formula(r, c, '=' + expression, numeric_format, scalar(value))
                        elif isinstance(value, bool):
                            ws.write_boolean(r, c, value)
                        elif isinstance(value, (int, float, Decimal)):
                            ws.write_number(r, c, scalar(value), numeric_format)
                        elif value is not None:
                            ws.write_string(r, c, str(value))
                if records:
                    ws.autofilter(0, 0, len(records), len(headers) - 1)
                return ws
            table('源数据', original_columns, original)
            if source_formulas:
                table('源公式复核', ['单元格', '原公式（文本）', '重新计算值'], source_formulas)
            table('计算明细', columns, data, True)
            if metrics and summary_columns:
                table('透视汇总' if pivot_column else '分组汇总', summary_columns, summary)
            ws = book.add_worksheet('说明与来源')
            ws.set_column(0, 0, 24)
            ws.set_column(1, 1, 100)
            ws.write_row(0, 0, ['项目', '记录'], header)
            wrap = book.add_format({'text_wrap': True, 'valign': 'top'})
            for i, (key, value) in enumerate(audit.items(), 1):
                ws.write_string(i, 0, key)
                ws.write_string(i, 1, json.dumps(value, ensure_ascii=False) if not isinstance(value, str) else value, wrap)
            book.set_properties({'title': label, 'comments': 'Generated analysis working paper. Source data is preserved; no source files modified.'})
        temporary.replace(target)
    finally:
        temporary.unlink(missing_ok=True)
    return dict(file_id='results:' + __import__('base64').urlsafe_b64encode(target.name.encode()).decode().rstrip('='), name=target.name,
                saved=True, source=source, input_rows=len(original), included_rows=len(data), filtered_rows=excluded, skipped_values=skipped,
                result_columns=summary_columns if metrics else columns,
                result_preview=summary[:10] if metrics else [[scalar(record[c]) for c in columns] for record in data[:10]],
                result_rows=len(summary) if metrics else len(data), preview_complete=(len(summary) if metrics else len(data)) <= 10,
                note='Analysis paper retained locally. User can open it in the current Excel instance, export it or delete it in the workspace. Read additional rows before drawing conclusions beyond the preview. Source data, formulas with computed cached values, and provenance are included. Edits to the current Excel workbook require the existing GridPilot approval workflow.')
