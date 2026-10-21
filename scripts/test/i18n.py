"""Check the task pane's language files (pptcom/UI/web/locales/<language>.json) against the code that uses them.

The add-in and the page refer to every sentence the user sees by a message code (see pptcom/UI/web/i18n.js).
This checks that
  - each file is a JSON object of "UPPER_SNAKE_CODE": "text", and every other language has exactly the codes of
    ja.json (the base) with the same {n} placeholders;
  - every code the code uses is in ja.json, and every code in ja.json is used;
  - a call passes as many values as its text has placeholders ({0}..{n-1}, each used);
  - no Japanese text is written in the C# or JavaScript code itself (outside comments), except the few strings
    listed in ALLOWED below, which are not shown in the task pane.

usage: python -X utf8 scripts/test/i18n.py      exit code 0 = no problems
"""
import json, pathlib, re, sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
WEB = ROOT / 'pptcom' / 'UI' / 'web'
LOCALES = WEB / 'locales'
CODE_RE = re.compile(r'^[A-Z][A-Z0-9_]*$')
PLACEHOLDER = re.compile(r'\{(\d+)\}')
JAPANESE = re.compile(r'[぀-ヿ㐀-鿿＀-￯〜]')

# C# calls whose first argument is a code and whose other arguments are its values (not their declarations).
CS_CALLS = re.compile(r'(?<!void )\b(?:Log|SetStatus|PaneText\.Of|Texts\.Get)\(')
# C# calls whose first argument is a code with no values of its own (the separator of a Join).
CS_CODE_ONLY = re.compile(r'PaneText\.Join\(')
JS_CALLS = re.compile(r'\bi18n\.(?:t|fill)\(')

# Japanese strings in the code that are not shown in the task pane: errors returned to the program that drives
# the add-in (AddinAutomation, ThisAddIn.AttachFromOutside), and characters the code reads (the wave dash a slide
# range may be written with).
ALLOWED = {
    'pptcom/AddinAutomation.cs': ['外部からの操作は無効です', '実行中の依頼があります', '実行中は新しい会話を始められません', '実行中は資料を添付できません',
                                  '資料を読み取っています', 'プレゼンテーションのウィンドウがありません', '続けられる依頼がありません'],
    'pptcom/ThisAddIn.cs': ['ファイルが見つかりません: ', '50MB を超えるファイルは添付できません: '],
    'pptcom/ComProxy/DeckReader.cs': ['〜'],
}

problems = []


def problem(text):
    problems.append(text)


def strip_comments(src, js):
    """Source with comments blanked out (strings kept), so that line numbers stay."""
    out, i, n = [], 0, len(src)
    while i < n:
        c = src[i]
        if src.startswith('//', i):
            j = src.find('\n', i)
            j = n if j < 0 else j
            out.append(' ' * (j - i)); i = j
        elif src.startswith('/*', i):
            j = src.find('*/', i + 2)
            j = n if j < 0 else j + 2
            out.append(re.sub(r'[^\n]', ' ', src[i:j])); i = j
        elif not js and src.startswith('@"', i):
            j = i + 2
            while j < n and not (src[j] == '"' and not src.startswith('""', j)):
                j += 2 if src.startswith('""', j) else 1
            out.append(src[i:j + 1]); i = j + 1
        elif c in '"\'' or (js and c == '`'):
            if js and c == '/':
                pass
            j = i + 1
            while j < n and src[j] != c and src[j] != '\n':
                j += 2 if src[j] == '\\' else 1
            out.append(src[i:j + 1]); i = j + 1
        elif js and c == '/' and re.match(r'[=(,:!&|?{};]\s*$', ''.join(out[-1:])[-20:] or '(') and not src.startswith('/=', i):
            # a regular expression literal: copy it up to its closing slash
            j = i + 1
            in_class = False
            while j < n and src[j] != '\n':
                if src[j] == '\\': j += 2; continue
                if src[j] == '[': in_class = True
                elif src[j] == ']': in_class = False
                elif src[j] == '/' and not in_class: break
                j += 1
            out.append(src[i:j + 1]); i = j + 1
        else:
            out.append(c); i += 1
    return ''.join(out)


def string_literals(text, js):
    """(value, start) of each string literal in text (comments already removed)."""
    found, i, n = [], 0, len(text)
    while i < n:
        if not js and text.startswith('@"', i):
            j = i + 2; buf = []
            while j < n:
                if text.startswith('""', j): buf.append('"'); j += 2; continue
                if text[j] == '"': break
                buf.append(text[j]); j += 1
            found.append((''.join(buf), i)); i = j + 1
        elif text[i] in '"\'' or (js and text[i] == '`'):
            q = text[i]; j = i + 1; buf = []
            while j < n and text[j] != q and text[j] != '\n':
                if text[j] == '\\' and j + 1 < n: buf.append(text[j:j + 2]); j += 2; continue
                buf.append(text[j]); j += 1
            if not js and q == "'":   # a C# char literal
                i = j + 1; continue
            found.append((''.join(buf), i)); i = j + 1
        else:
            i += 1
    return found


def call_arguments(text, open_paren):
    """The top-level arguments of the call whose '(' is at open_paren, as source text."""
    args, depth, i, start, n = [], 0, open_paren + 1, open_paren + 1, len(text)
    while i < n:
        c = text[i]
        if c in '"\'`':
            j = i + 1
            while j < n and text[j] != c:
                j += 2 if text[j] == '\\' else 1
            i = j + 1; continue
        if text.startswith('@"', i):
            j = i + 2
            while j < n and not (text[j] == '"' and not text.startswith('""', j)):
                j += 2 if text.startswith('""', j) else 1
            i = j + 1; continue
        if c in '([{':
            depth += 1
        elif c in ')]}':
            if depth == 0:
                if text[start:i].strip(): args.append(text[start:i])
                return args
            depth -= 1
        elif c == ',' and depth == 0:
            args.append(text[start:i]); start = i + 1
        i += 1
    return args


def line_of(text, pos):
    return text.count('\n', 0, pos) + 1


def placeholders(text):
    return sorted({int(m) for m in PLACEHOLDER.findall(text)})


# ---- the language files ---------------------------------------------------------------------------------

tables = {}
for path in sorted(LOCALES.glob('*.json')):
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
    except ValueError as e:
        problem(f'{path.relative_to(ROOT)}: not JSON: {e}'); continue
    if not isinstance(data, dict):
        problem(f'{path.relative_to(ROOT)}: not a JSON object'); continue
    for code, text in data.items():
        if not CODE_RE.match(code): problem(f'{path.name}: {code}: a code is UPPER_SNAKE_CASE')
        if not isinstance(text, str): problem(f'{path.name}: {code}: the text is not a string')
    tables[path.stem] = data

if 'ja' not in tables:
    sys.exit('pptcom/UI/web/locales/ja.json is missing')
base = tables['ja']
for lang, table in tables.items():
    if lang == 'ja': continue
    for code in sorted(set(base) - set(table)): problem(f'{lang}.json: missing {code}')
    for code in sorted(set(table) - set(base)): problem(f'{lang}.json: {code} is not in ja.json')
    for code in sorted(set(base) & set(table)):
        if placeholders(base[code]) != placeholders(table[code]):
            problem(f'{lang}.json: {code} has placeholders {placeholders(table[code])}, ja.json {placeholders(base[code])}')

# ---- the codes the code uses ----------------------------------------------------------------------------

used = {}   # code -> first place


def use(code, where, values):
    used.setdefault(code, where)
    if code not in base:
        problem(f'{where}: {code} is not in ja.json'); return
    if values is None: return
    want = placeholders(base[code])
    if want != list(range(values)):
        problem(f'{where}: {code} gets {values} value(s), its text has {want or "no placeholders"}')


def codes_in(arg):
    """The codes an argument names: "CODE", or cond ? "A" : "B"."""
    lits = [v for v, _ in string_literals(arg, False)]
    return [v for v in lits if CODE_RE.match(v)] if lits and all(CODE_RE.match(v) for v in lits) else []


for path in sorted((ROOT / 'pptcom').rglob('*.cs')):
    if 'obj' in path.parts or 'bin' in path.parts: continue
    rel = path.relative_to(ROOT).as_posix()
    src = strip_comments(path.read_text(encoding='utf-8-sig'), js=False)
    for m in CS_CALLS.finditer(src):
        args = call_arguments(src, m.end() - 1)
        codes = codes_in(args[0]) if args else []
        if args and args[0].strip() == 'code':
            continue   # Log and SetStatus handing their code on to PaneText.Of
        if not codes:
            problem(f'{rel}:{line_of(src, m.start())}: {m.group(0)} without a code written out'); continue
        for code in codes: use(code, f'{rel}:{line_of(src, m.start())}', len(args) - 1)
    for m in CS_CODE_ONLY.finditer(src):
        args = call_arguments(src, m.end() - 1)
        for code in (codes_in(args[0]) if args else []): use(code, f'{rel}:{line_of(src, m.start())}', 0)
    for value, pos in string_literals(src, False):
        if JAPANESE.search(value) and not any(a in value for a in ALLOWED.get(rel, [])):
            problem(f'{rel}:{line_of(src, pos)}: Japanese text in the code: "{value[:60]}" (use a message code)')

for path in sorted(WEB.glob('*.js')):
    rel = path.relative_to(ROOT).as_posix()
    src = strip_comments(path.read_text(encoding='utf-8'), js=True)
    for m in JS_CALLS.finditer(src):
        args = call_arguments(src, m.end() - 1)
        lits = [v for v, _ in string_literals(args[0], True)] if args else []
        codes = [v for v in lits if CODE_RE.match(v)]
        fill = m.group(0).startswith('i18n.fill')
        for code in codes: use(code, f'{rel}:{line_of(src, m.start())}', len(args) - 1 if fill else 0)
    # the codes taskpane.js routes by (logKinds)
    block = re.search(r'const logKinds = \{([^}]*)\}', src)
    if block:
        for code in re.findall(r'\b([A-Z][A-Z0-9_]+)\s*:', block.group(1)): use(code, f'{rel} (logKinds)', None)
    for value, pos in string_literals(src, True):
        if JAPANESE.search(value): problem(f'{rel}:{line_of(src, pos)}: Japanese text in the code: "{value[:60]}" (use a message code)')

html = (WEB / 'index.html').read_text(encoding='utf-8')
for m in re.finditer(r'data-i18n(?:-placeholder|-title|-label)?="([^"]*)"', html):
    use(m.group(1), f'index.html:{line_of(html, m.start())}', 0)
visible = re.sub(r'<!--.*?-->|<script.*?</script>|<style.*?</style>', ' ', html, flags=re.S)
for m in re.finditer(r'>([^<]*)<|(?:title|placeholder|aria-label)="([^"]*)"', visible):
    text = m.group(1) or m.group(2) or ''
    if JAPANESE.search(text): problem(f'index.html: Japanese text in the page itself: "{text.strip()[:60]}" (use data-i18n)')

for code in sorted(set(base) - set(used)):
    problem(f'ja.json: {code} is not used')

if problems:
    print('\n'.join(problems))
    print(f'{len(problems)} problem(s)')
    sys.exit(1)
print(f'i18n: {len(base)} codes in ja.json, all used; {len(tables) - 1} other language(s) match')
