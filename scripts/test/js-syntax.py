"""Check that the task pane's own JavaScript parses (pptcom/UI/web/*.js; the vendor/ copies are not ours).

Node's --check reads a file without running it. index.html loads some files as classic scripts; every other one
is a module (a <script type="module">, a module worker or an import). A module is checked as a copy named .mjs,
so that the result does not depend on how a Node version guesses what a .js file is.

usage: python -X utf8 scripts/test/js-syntax.py      exit code 0 = every file parses
"""
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile

WEB = pathlib.Path(__file__).resolve().parents[2] / 'pptcom' / 'UI' / 'web'


def main():
    node = shutil.which('node')
    if node is None:
        print('js-syntax: node is not on PATH')
        return 2
    version = subprocess.run([node, '--version'], capture_output=True, text=True, check=False).stdout.strip()

    html = (WEB / 'index.html').read_text(encoding='utf-8')
    classic = set()
    for tag in re.findall(r'<script\b[^>]*>', html):
        src = re.search(r'\bsrc="([^"]+)"', tag)
        if src and 'type="module"' not in tag:
            classic.add(src.group(1))

    files = sorted(p for p in WEB.glob('*.js'))
    failed = []
    with tempfile.TemporaryDirectory() as tmp:
        for f in files:
            target = f
            if f.name not in classic:
                target = pathlib.Path(tmp) / (f.stem + '.mjs')
                shutil.copyfile(f, target)
            r = subprocess.run([node, '--check', str(target)], capture_output=True, text=True, check=False)
            if r.returncode != 0:
                failed.append(f.name + ':\n' + (r.stderr or r.stdout).strip().replace(str(target), f.name))

    for f in failed:
        print('FAIL ' + f)
    kinds = ', '.join(f.name + (' (script)' if f.name in classic else ' (module)') for f in files)
    print(f'js-syntax: {len(files) - len(failed)} of {len(files)} files parse with node {version}: {kinds}')
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
