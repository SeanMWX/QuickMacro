import ast
p='QuickMacro.py'
with open(p,'r',encoding='utf-8',errors='ignore') as f:
    src=f.read()
try:
    ast.parse(src)
    print('AST OK')
except SyntaxError as e:
    print('SyntaxError', e.lineno, e.offset, e.msg)
    print('LINE:', src.splitlines()[e.lineno-1])
