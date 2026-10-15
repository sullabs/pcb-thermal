import glob
import os
import re

SRC = r"C:\Users\Sultan\Desktop\thermal_similation\src"
RESERVED = set(x.lower() for x in """Imp Eqv Mid Len Name Date Error Left Right Val Str Type Get Set Open Close Time Now Format Round Abs
Array Dir Input Line Mod Not And Or Xor Is Like New Nothing Null Empty True False Print Put Seek Stop End Select Case
If Then Else For Each Next Do Loop While Wend With Exit Call Const Dim ReDim Static Public Private Friend Option Explicit
String Integer Long Double Single Boolean Byte Variant Object Currency Decimal Collection Rem Let Lset Rset Resume Return
Event Error Erl Err Fix Int Sgn Sqr Tab Spc Timer Trim Split Join Replace Space String Hex Oct Asc Chr Cos Sin Tan Atn Exp Log
InStr StrComp UCase LCase CStr CLng CDbl CInt CBool Choose Switch IIf Partition Filter Min Max Sum Count Item Items Keys
Application Declare Default Property Optional ParamArray Preserve Alias Lib Base Compare Binary Text Database Enum
Implements Attribute Step To Until WithEvents Sub Function Lock Unlock Width Name Reset Kill Load Unload Date Day Month Year Hour Minute Second
Mid Left Right Randomize Rnd Shell Environ Loc Lof Eof Input Output Append Random Access Read Write Shared Seek Me""".split())
# names that are only dangerous as variables; we ignore ones that are fine as members
IGNORE = set(x.lower() for x in "Count Item Items Keys Min Max Sum Text Compare Base Default Output Read Write Access Append Random Binary Database Reset Load Width Filter Count".split())
CHECK = RESERVED - IGNORE

problems = []
for f in sorted(glob.glob(os.path.join(SRC, "*.bas"))):
    b = open(f, 'rb').read()
    base = os.path.basename(f)
    if any(c > 127 for c in b):
        problems.append((base, 0, "non-ascii"))
    if b.count(b'\r\n') != b.count(b'\n'):
        problems.append((base, 0, "LF without CR"))
    lines = b.decode('ascii', 'replace').split('\r\n')
    if not lines[0].startswith('Attribute VB_Name'):
        problems.append((base, 1, "first line is not Attribute VB_Name"))
    if 'Option Explicit' not in lines[1]:
        problems.append((base, 2, "no Option Explicit"))
    # procedures
    proc = None
    names = {}
    for n, raw in enumerate(lines, 1):
        ln = raw
        # strip comments and strings
        code = re.sub(r'"[^"]*"', '""', ln)
        code = code.split("'")[0].strip()
        m = re.match(r'(?:Public |Private |Friend )?(?:Static )?(Sub|Function)\s+(\w+)\s*\((.*)', code)
        if m and not code.startswith("End "):
            proc = m.group(2)
            names = {}
            sig = code
            # continuation lines for signature
            k = n
            while sig.rstrip().endswith('_'):
                sig = sig.rstrip()[:-1] + ' ' + lines[k].split("'")[0].strip()
                k += 1
            inner = sig[sig.index('(') + 1: sig.rindex(')')] if ')' in sig else ''
            depth = 0
            parts = []
            cur = ''
            for ch in inner:
                if ch == '(':
                    depth += 1
                if ch == ')':
                    depth -= 1
                if ch == ',' and depth == 0:
                    parts.append(cur)
                    cur = ''
                else:
                    cur += ch
            if cur.strip():
                parts.append(cur)
            for p in parts:
                p = p.strip()
                if not p:
                    continue
                pm = re.match(r'(?:Optional\s+)?(ByVal|ByRef)?\s*(\w+)(\(\))?', p)
                if not pm:
                    continue
                if pm.group(1) is None:
                    problems.append((base, n, "param without ByVal/ByRef: " + p))
                nm = pm.group(2).lower()
                is_arr = pm.group(3) is not None
                if pm.group(1) == 'ByRef' and not is_arr and not re.search(r'As\s+(Long|Double|String|Variant|Boolean|Object)', p) is None:
                    pass
                if nm in CHECK:
                    problems.append((base, n, "reserved/builtin parameter name: " + nm))
                names[nm] = n
            continue
        if re.match(r'End (Sub|Function)', code):
            proc = None
            continue
        if proc is None:
            # module level Dim
            md = re.match(r'(?:Dim|Private|Public)\s+(.*)', code)
            if md and not code.startswith('Public Function') and not code.startswith('Private Function') and not code.startswith('Public Sub') and not code.startswith('Private Sub'):
                body = re.sub(r'\([^)]*\)', '', md.group(1))
                for part in body.split(','):
                    pm = re.match(r'\s*(\w+)', part)
                    if pm and pm.group(1).lower() in CHECK and pm.group(1).lower() not in ('const',):
                        problems.append((base, n, "reserved module-level name: " + pm.group(1)))
            continue
        md = re.match(r'Dim\s+(.*)', code)
        if md:
            body = re.sub(r'\([^)]*\)', '', md.group(1))
            for part in body.split(':')[0].split(','):
                pm = re.match(r'\s*(\w+)', part)
                if not pm:
                    continue
                nm = pm.group(1).lower()
                if nm in CHECK:
                    problems.append((base, n, "reserved/builtin variable name: " + nm + " in " + proc))
                if nm in names:
                    problems.append((base, n, "duplicate Dim/param: " + nm + " in " + proc + " (first line %d)" % names[nm]))
                names[nm] = n
        # Dim on a line after ':'  e.g. "x = 1: Dim y"
        for mm in re.finditer(r':\s*Dim\s+(\w+)', code):
            nm = mm.group(1).lower()
            if nm in names:
                problems.append((base, n, "duplicate Dim: " + nm))
            names[nm] = n

for p in problems:
    print(p)
print(len(problems), "findings")
