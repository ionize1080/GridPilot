"""Small, explicit Excel formula subset. Unsupported functions fail, never use caches."""
import ast
from decimal import Decimal, ROUND_HALF_UP
import json
import operator
import re

TOKEN = re.compile(r'\s+|"(?:[^"]|"")*"|(?:\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|[A-Za-z_][A-Za-z_0-9.]*|<>|>=|<=|[+*/^(),<>=%-]')
CELL = re.compile(r'^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?$')


class FormulaGrid:
    def __init__(self, rows, formula_cells=None):
        self.rows, self.cache, self.visiting = rows, {}, set()
        self.formula_cells = formula_cells
        self.calculated = 0
    def cell(self, reference):
        from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
        reference = reference.replace('$', '').upper()
        if reference in self.cache: return self.cache[reference]
        if reference in self.visiting or len(self.visiting) > 100: raise ValueError('公式存在循环或依赖层级过深。')
        column, row = coordinate_from_string(reference); col = column_index_from_string(column)
        if row > 1048576 or col > 16384: raise ValueError('单元格引用超出 Excel 工作表边界。')
        value = self.rows[row - 1][col - 1] if row <= len(self.rows) and col <= len(self.rows[row - 1]) else None
        if isinstance(value, str) and value.startswith('=') and (self.formula_cells is None or reference in self.formula_cells):
            self.visiting.add(reference)
            try: value = self.expression(value[1:]); self.calculated += 1
            finally: self.visiting.remove(reference)
        self.cache[reference] = value
        return value
    def range(self, reference):
        from openpyxl.utils.cell import range_boundaries, get_column_letter
        a, b, c, d = range_boundaries(reference.replace('$','').upper())
        a,c = sorted((a,c)); b,d = sorted((b,d))
        if d > 1048576 or c > 16384: raise ValueError('公式范围超出 Excel 工作表边界。')
        if (c - a + 1) * (d - b + 1) > 100000: raise ValueError('公式范围超过 10 万单元格。')
        return [self.cell(f'{get_column_letter(column)}{row}') for row in range(b, d + 1) for column in range(a, c + 1)]
    @staticmethod
    def numeric(value):
        if value is None: return Decimal(0)
        if isinstance(value, bool): return Decimal(int(value))
        if not isinstance(value, (int, float, Decimal)): raise ValueError('公式需要数值，但引用了文本。')
        result = Decimal(str(value))
        if not result.is_finite(): raise ValueError('公式结果不是有限数值。')
        return result
    def expression(self, source):
        if len(source) > 2000: raise ValueError('公式过长。')
        parts, refs, position = [], {}, 0
        for match in TOKEN.finditer(source):
            if match.start() != position: raise ValueError('暂不支持跨表、外部链接、结构化引用或此公式语法。')
            token = match.group(); position = match.end()
            if token.isspace(): continue
            if token.startswith('"'): parts.append(json.dumps(token[1:-1].replace('""','"')))
            elif CELL.fullmatch(token):
                label = '_ref' + str(len(refs)); refs[label] = token; parts.append(label)
            elif token == '^': raise ValueError('暂不支持幂运算，请先展开为基础乘除。')
            elif token == '=': parts.append('==')
            elif token == '<>': parts.append('!=')
            elif token == '%': parts.append('/100')
            else: parts.append(token.upper())
        if position != len(source): raise ValueError('不支持的公式语法。')
        try: tree = ast.parse(''.join(parts), mode='eval')
        except SyntaxError as exc: raise ValueError('公式语法无效。') from exc
        if len(list(ast.walk(tree))) > 300: raise ValueError('公式过于复杂。')
        def evaluate(node):
            if isinstance(node, ast.Expression): return evaluate(node.body)
            if isinstance(node, ast.Constant) and type(node.value) in (int, float, str): return Decimal(str(node.value)) if not isinstance(node.value,str) else node.value
            if isinstance(node, ast.Name):
                if node.id in refs:
                    ref = refs[node.id]; return self.range(ref) if ':' in ref else self.cell(ref)
                if node.id in ('TRUE','FALSE'): return node.id == 'TRUE'
            if isinstance(node, ast.UnaryOp) and isinstance(node.op,(ast.UAdd,ast.USub)):
                value = self.numeric(evaluate(node.operand));return -value if isinstance(node.op,ast.USub) else value
            if isinstance(node, ast.BinOp):
                a,b=self.numeric(evaluate(node.left)),self.numeric(evaluate(node.right))
                ops={ast.Add:operator.add,ast.Sub:operator.sub,ast.Mult:operator.mul,ast.Div:operator.truediv}
                if type(node.op) not in ops: raise ValueError('不支持的公式运算符。')
                if isinstance(node.op,ast.Div) and b==0: raise ValueError('公式出现除以零。')
                return ops[type(node.op)](a,b)
            if isinstance(node,ast.Compare) and len(node.ops)==1:
                a,b=evaluate(node.left),evaluate(node.comparators[0]);ops={ast.Eq:operator.eq,ast.NotEq:operator.ne,ast.Lt:operator.lt,ast.LtE:operator.le,ast.Gt:operator.gt,ast.GtE:operator.ge}
                if isinstance(a,str) or isinstance(b,str): raise ValueError('暂不支持公式中的文本比较。')
                if type(node.ops[0]) in ops: return ops[type(node.ops[0])](self.numeric(a),self.numeric(b))
            if isinstance(node,ast.Call) and isinstance(node.func,ast.Name) and not node.keywords:
                name=node.func.id
                if name=='IF' and len(node.args)==3: return evaluate(node.args[1] if self.numeric(evaluate(node.args[0])) else node.args[2])
                arguments=[evaluate(n) for n in node.args]
                if not arguments: raise ValueError('函数缺少参数。')
                if any(isinstance(a,ast.Constant) and isinstance(a.value,(str,bool)) for a in node.args):
                    raise ValueError('函数参数不支持隐式文本或布尔值转换。')
                flat=[v for a in arguments for v in (a if isinstance(a,list) else [a])]
                if name in ('SUM','COUNT','AVERAGE','MIN','MAX'):
                    for argument, value in zip(node.args, arguments):
                        is_reference = isinstance(argument, ast.Name) and argument.id in refs
                        if not is_reference and isinstance(value,(str,bool)):
                            raise ValueError('聚合公式不支持非引用参数的隐式文本或布尔值转换。')
                numeric=[self.numeric(v) for v in flat if isinstance(v,(int,float,Decimal)) and not isinstance(v,bool)]
                if name=='SUM': return sum(numeric,Decimal(0))
                if name=='COUNT': return len(numeric)
                if name in ('AVERAGE','MIN','MAX'):
                    if not numeric: raise ValueError('公式聚合范围没有数值。')
                    return sum(numeric,Decimal(0))/len(numeric) if name=='AVERAGE' else min(numeric) if name=='MIN' else max(numeric)
                if name=='ABS' and len(arguments)==1:return abs(self.numeric(arguments[0]))
                if name=='ROUND' and len(arguments)==2:
                    digits=self.numeric(arguments[1])
                    if digits!=int(digits) or not -6<=digits<=12:raise ValueError('ROUND 位数须为 -6..12 整数。')
                    return self.numeric(arguments[0]).quantize(Decimal(10)**-int(digits),rounding=ROUND_HALF_UP)
                if name=='AND':return all(self.numeric(v) for v in flat)
                if name=='OR':return any(self.numeric(v) for v in flat)
                if name=='NOT' and len(arguments)==1:return not self.numeric(arguments[0])
                raise ValueError('不支持 Excel 函数 '+name+'；未使用缓存结果。请改用支持的计算。')
            raise ValueError('不支持的公式表达式。')
        return evaluate(tree)
