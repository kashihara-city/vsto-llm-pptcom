"""Check the CycloneDX SBOM scripts/build/sbom.py makes for an MSI: the same content each time it is made,
valid against the official 1.6 schema (deps/cyclonedx), and what it lists matches the records it came from -
the MSI, deps/web-sources.lock.json, NuGet's assets file, anydoc's Cargo.lock and what cargo builds.

  python -X utf8 scripts/test/sbom.py --msi release/vsto-llm-pptcom-<version>.msi
"""
import argparse, base64, copy, hashlib, json, pathlib, re, subprocess, sys, tempfile

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.dont_write_bytecode = True  # no __pycache__ next to the scripts
sys.path.insert(0, str(ROOT / 'scripts/test'))
from json_schema import Schemas  # noqa: E402

failures = []


def check(ok, what, detail=''):
    print(('PASS: ' if ok else 'FAIL: ') + what + (f' ({detail})' if detail and not ok else ''), flush=True)
    if not ok:
        failures.append(what)


def generate(msi, out):
    subprocess.run([sys.executable, '-X', 'utf8', str(ROOT / 'scripts/build/sbom.py'), '--msi', str(msi), '--out', str(out)],
                   check=True, capture_output=True)
    return out.read_bytes()


def sha256(path):
    return hashlib.sha256(pathlib.Path(path).read_bytes()).hexdigest()


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument('--msi', required=True)
    msi = pathlib.Path(ap.parse_args().msi)

    with tempfile.TemporaryDirectory() as tmp:
        first = generate(msi, pathlib.Path(tmp) / 'a.cdx.json')
        second = generate(msi, pathlib.Path(tmp) / 'b.cdx.json')
    check(first == second, 'the same MSI and commit give the same SBOM, byte for byte')
    bom = json.loads(first)

    pins = json.loads((ROOT / 'deps/cyclonedx/schema.lock.json').read_text(encoding='utf-8'))
    changed = [n for n, f in pins['files'].items() if sha256(ROOT / 'deps/cyclonedx' / n) != f['sha256']]
    check(not changed, f"the schemas are the pinned CycloneDX {pins['version']} ones", ', '.join(changed))
    schemas = Schemas(ROOT / 'deps/cyclonedx')
    schema = schemas.docs['bom-1.6.schema.json']
    errors = schemas.errors(bom, schema)
    check(not errors, 'valid against the CycloneDX 1.6 schema', '; '.join(errors[:5]))
    broken = copy.deepcopy(bom)
    broken['components'][0]['hashes'][0]['alg'] = 'SHA-0'
    check(bool(schemas.errors(broken, schema)), 'and the schema check does catch a component that breaks it')

    top = bom['metadata']['component']
    check(top['hashes'][0]['content'] == sha256(msi), 'the SBOM describes this MSI (SHA-256)')
    commit = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip()
    check(top['bom-ref'].endswith('@' + commit), 'and the commit it was built from')

    by_ref = {c['bom-ref']: c for c in bom['components']}
    lock = json.loads((ROOT / 'deps/web-sources.lock.json').read_text(encoding='utf-8'))
    for name in ('anydoc', 'pdfjs'):
        entry = lock[name]
        c = by_ref.get(f"pkg:github/{entry['repository']}@{entry['version']}")
        check(c is not None and c['hashes'][0]['content'] == entry['sha256'], f'{name}: the archive pinned in the lock')

    assets = json.loads((ROOT / 'temp/build/pptcom/obj/project.assets.json').read_text(encoding='utf-8'))
    # PrivateAssets="all" (suppressParent All): only builds it (ILRepack, the analyzers), listed under metadata.tools
    build_only = {n for fw in assets['project']['frameworks'].values()
                  for n, d in fw.get('dependencies', {}).items() if d.get('suppressParent') == 'All'}
    tools = {t.get('purl'): t for t in bom['metadata'].get('tools', {}).get('components', [])}
    for key, lib in assets['libraries'].items():
        name, version = key.split('/')
        c = (tools if name in build_only else by_ref).get(f'pkg:nuget/{name}@{version}')
        check(c is not None and c['hashes'][0]['content'] == base64.b64decode(lib['sha512']).hex(),
              f'{name} {version}: the package NuGet restored' + (' (a build tool)' if name in build_only else ''))

    # The crates cargo builds into anydoc's .wasm, read independently of sbom.py's parsing
    source = ROOT / 'temp/build/pptcom/obj/sbom/anydoc'
    tree = subprocess.check_output(['rustup', 'run', lock['rust']['toolchain'], 'cargo', 'tree', '--locked', '--offline',
                                    '--manifest-path', str(source / 'wasm/Cargo.toml'), '--target', lock['rust']['target'],
                                    '-e', 'normal,no-proc-macro', '--prefix', 'none', '--format', '{p}'], cwd=ROOT, text=True)
    built = {(m[1], m[2]) for m in re.finditer(r'^([\w-]+) v(\S+)', tree, re.M)}
    built = {k for k in built if not k[0].startswith('anydoc')}
    listed = {(c['name'], c['version']) for c in bom['components'] if c['purl'].startswith('pkg:cargo/')}
    check(built == listed, f'the {len(built)} crates cargo builds into the .wasm are the ones listed',
          f'missing {sorted(built - listed)[:5]}, extra {sorted(listed - built)[:5]}')
    cargo_lock = (source / 'Cargo.lock').read_text(encoding='utf-8')
    checksums = {(m[1], m[2]): m[3] for m in re.finditer(
        r'\[\[package\]\]\nname = "([^"]+)"\nversion = "([^"]+)"\nsource = "[^"]+"\nchecksum = "([0-9a-f]+)"', cargo_lock)}
    wrong = [c['name'] for c in bom['components'] if c['purl'].startswith('pkg:cargo/')
             and c['hashes'][0]['content'] != checksums.get((c['name'], c['version']))]
    check(not wrong, "each crate's SHA-256 is the checksum in anydoc's Cargo.lock", ', '.join(wrong[:5]))

    unlicensed = [c['bom-ref'] for c in bom['components'] if not c.get('licenses')]
    check(not unlicensed, 'every component names its license', ', '.join(unlicensed[:5]))
    refs = set(by_ref) | {top['bom-ref']} | {p['bom-ref'] for c in bom['components'] for p in c.get('components', [])}
    dangling = [r for d in bom['dependencies'] for r in [d['ref'], *d.get('dependsOn', [])] if r not in refs]
    check(not dangling, 'every dependency names a component of the SBOM', ', '.join(dangling[:5]))

    if failures:
        raise SystemExit(f'{len(failures)} SBOM check(s) failed')
    print('SBOM checks passed')


if __name__ == '__main__':
    main()
