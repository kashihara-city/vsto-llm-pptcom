"""Check the warnings of C# builds (MSBuild logs written with -flp:logfile=<log>;warningsonly).

Directory.Build.props makes every warning an error but one: CS8032, an analyzer that cannot be created, so that
Meziantou's RemoveUnnecessaryClosedModifierAnalyzer (built against an experimental Roslyn API) may fail to load
in this compiler. That exception covers every analyzer, though. This check lets that one through and nothing
else: another warning is an analyzer that stopped loading (its rules then silently do not run), or a warning
that is not a compiler one.

usage: python -X utf8 scripts/test/build-warnings.py <log> ...      exit code 0 = only the known warning
"""
import pathlib
import re
import sys

KNOWN = ('CS8032', 'RemoveUnnecessaryClosedModifierAnalyzer')
WARNING = re.compile(r': warning ([A-Z]+[0-9]+):')


def main(paths):
    if not paths:
        print(__doc__)
        return 2
    known, other = set(), set()
    for path in paths:
        p = pathlib.Path(path)
        if not p.exists():
            print('build-warnings: no log ' + path)
            return 2
        for line in p.read_text(encoding='utf-8-sig', errors='replace').splitlines():
            m = WARNING.search(line)
            if not m:
                continue
            text = re.sub(r'\s*\[[^\]]*\]$', '', line.strip())   # the project MSBuild appends
            if m.group(1) == KNOWN[0] and KNOWN[1] in line:
                known.add(text)
            else:
                other.add(text)
    for text in sorted(other):
        print('UNEXPECTED ' + text)
    print(f'build-warnings: {len(paths)} log(s), {len(known)} known CS8032, {len(other)} other warning(s)')
    return 1 if other else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
