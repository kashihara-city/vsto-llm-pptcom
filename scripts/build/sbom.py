"""Write the CycloneDX 1.6 SBOM of the MSI that scripts/build/build.ps1 made: what pptcom.dll and the MSI carry and
what built them, from the pinned records only - no network access.

  python -X utf8 scripts/build/sbom.py --msi release/vsto-llm-pptcom-<version>.msi --out <file>.cdx.json

  components        Newtonsoft.Json and the WebView2 SDK (NuGet, with the SHA-512 NuGet recorded in
                    temp/build/pptcom/obj/project.assets.json), the VSTO Utilities DLL shipped next to pptcom.dll, PDF.js (the
                    release archive pinned in deps/web-sources.lock.json, with the parts it brings), anydoc (the
                    pinned source) and the Rust crates compiled into its WebAssembly build (cargo tree for the
                    wasm32 target, normal dependencies, as anydoc's Cargo.lock pins them, with their checksums)
  metadata.tools    what built it: Rust, wasm-bindgen, wasm-opt (binaryen), ILRepack, the analyzers, the WiX SDK,
                    this script

The same inputs give the same file: the serial number is derived from the MSI's hash and the timestamp is the
commit's. scripts/test/sbom.py checks it against the official schema and against the records.
"""
import argparse, base64, datetime, hashlib, json, pathlib, re, shutil, subprocess, tarfile, uuid

ROOT = pathlib.Path(__file__).resolve().parents[2]
REPOSITORY = 'kashihara-city/vsto-llm-pptcom'
LOCK = ROOT / 'deps/web-sources.lock.json'
BUILD = ROOT / 'temp/build'
ASSETS = BUILD / 'pptcom/obj/project.assets.json'
CSPROJ = ROOT / 'pptcom/pptcom.csproj'
WIXPROJ = ROOT / 'scripts/build/pptcom.Installer.wixproj'
UTILITIES = BUILD / 'pptcom/bin/Release/Microsoft.Office.Tools.Common.v4.0.Utilities.dll'
WORK = BUILD / 'pptcom/obj/sbom'
SPEC = '1.6'

# What ships, and how (shown as a property of each component).
NUGET = {
    'Newtonsoft.Json': ('James Newton-King', [{'license': {'id': 'MIT'}}], 'merged into pptcom.dll'),
    'Microsoft.Web.WebView2': ('Microsoft Corporation',
                               [{'license': {'name': 'Microsoft WebView2 SDK license (BSD-style; text in THIRD-PARTY-NOTICES.md)'}}],
                               'Core and WinForms merged into pptcom.dll; WebView2Loader.dll for x86 and x64 next to it'),
}
# NuGet packages that only build it (metadata.tools): nothing of them is in pptcom.dll or the MSI.
BUILD_NUGET = {
    'ILRepack': 'merges Newtonsoft.Json and the WebView2 SDK into pptcom.dll',
    'Microsoft.CodeAnalysis.NetAnalyzers': 'code analysis in the build (Directory.Build.props)',
    'Microsoft.CodeAnalysis.CSharp.CodeStyle': 'code style analysis in the build (Directory.Build.props)',
    'Meziantou.Analyzer': 'code analysis in the build (Directory.Build.props)',
}
# The parts PDF.js's release archive brings that pptcom.dll embeds, with their own licenses.
PDFJS_PARTS = [
    ('pdfjs-cmaps', 'Adobe CMap resources', 'BSD-3-Clause', 'vendor/pdfjs/cmaps'),
    ('openjpeg', 'OpenJPEG decoder as WebAssembly', 'BSD-2-Clause', 'vendor/pdfjs/wasm/openjpeg.wasm'),
    ('pdfium-jbig2', 'JBIG2 decoder from PDFium as WebAssembly', 'BSD-3-Clause', 'vendor/pdfjs/wasm/jbig2.wasm'),
    ('qcms', 'qcms colour management as WebAssembly', 'MIT', 'vendor/pdfjs/wasm/qcms_bg.wasm'),
    ('cgats001compat-v2-micro-icc', 'ICC profile CGATS001Compat-v2-micro', 'CC0-1.0', 'vendor/pdfjs/iccs'),
]


def run(*args, cwd=ROOT):
    return subprocess.check_output(list(args), cwd=cwd, encoding='utf-8').strip()


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for block in iter(lambda: f.read(1 << 20), b''):
            h.update(block)
    return h.hexdigest()


def license_expression(text):
    """A crate's declared license as an SPDX expression (the old `A/B` form read as `A OR B`)."""
    return re.sub(r'\s*/\s*', ' OR ', text.strip())


def property_list(pairs):
    return [{'name': k, 'value': v} for k, v in pairs if v is not None]


def anydoc_source():
    """deps/anydoc.tar.gz unpacked into temp/build/pptcom/obj/sbom/anydoc (fresh each time)."""
    target = WORK / 'anydoc'
    if target.exists():
        shutil.rmtree(target)
    target.mkdir(parents=True)
    with tarfile.open(ROOT / 'deps/anydoc.tar.gz') as tar:
        for member in tar.getmembers():
            parts = pathlib.PurePosixPath(member.name).parts
            if len(parts) < 2:
                continue
            member.name = str(pathlib.PurePosixPath(*parts[1:]))
            tar.extract(member, target, filter='data')
    return target


def crates(lock):
    """The crates compiled into anydoc's .wasm: [(name, version, license, depth)] in cargo tree order, anydoc's
    own crates included (depth 0 = anydoc-wasm). Reads cargo's local cache only (--offline): the build has
    fetched them already."""
    source = anydoc_source()
    out = run('rustup', 'run', lock['rust']['toolchain'], 'cargo', 'tree', '--locked', '--offline',
              '--manifest-path', str(source / 'wasm/Cargo.toml'), '--target', lock['rust']['target'],
              '-e', 'normal,no-proc-macro', '--prefix', 'depth', '--format', '{p}|{l}')
    rows = []
    for line in out.splitlines():
        m = re.match(r'(\d+)([\w-]+) v([\w.+-]+)(?: \([^)]*\))?\|(.*)$', line.replace(' (*)', ''))
        if m:
            rows.append((m[2], m[3], m[4], int(m[1])))
    cargo_lock = (source / 'Cargo.lock').read_text(encoding='utf-8')
    checksums = {(m[1], m[2]): m[3] for m in re.finditer(
        r'\[\[package\]\]\nname = "([^"]+)"\nversion = "([^"]+)"\nsource = "[^"]+"\nchecksum = "([0-9a-f]+)"', cargo_lock)}
    return rows, checksums


def crate_graph(rows):
    """dependsOn per crate from cargo tree's depths (a crate printed again with (*) has the same edges)."""
    edges, stack = {}, []
    for name, version, _, depth in rows:
        key = (name, version)
        del stack[depth:]
        if stack:
            edges.setdefault(stack[-1], set()).add(key)
        edges.setdefault(key, set())
        stack.append(key)
    return edges


def build(msi):
    lock = json.loads(LOCK.read_text(encoding='utf-8'))
    assets = json.loads(ASSETS.read_text(encoding='utf-8'))
    commit = run('git', 'rev-parse', 'HEAD')
    commit_time = run('git', 'log', '-1', '--format=%cI', 'HEAD')
    # the commit's time, so that the same commit and MSI give the same SBOM
    timestamp = datetime.datetime.fromtimestamp(int(run('git', 'log', '-1', '--format=%ct', 'HEAD')),
                                                datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    app_version = re.search(r'<ApplicationVersion>([^<]+)</ApplicationVersion>', CSPROJ.read_text(encoding='utf-8'))[1]
    version = '.'.join(app_version.split('.')[:3])
    msi_hash = sha256(msi)
    app_ref = f'pkg:github/{REPOSITORY}@{commit}'

    components, dependencies, top = [], [], []

    # NuGet packages, with the SHA-512 NuGet checked when it restored them
    for name, (supplier, licenses, shipped) in NUGET.items():
        key = next(k for k in assets['libraries'] if k.split('/')[0] == name)
        ver = key.split('/')[1]
        purl = f'pkg:nuget/{name}@{ver}'
        components.append({
            'type': 'library', 'bom-ref': purl, 'supplier': {'name': supplier}, 'name': name, 'version': ver,
            'hashes': [{'alg': 'SHA-512', 'content': base64.b64decode(assets['libraries'][key]['sha512']).hex()}],
            'licenses': licenses, 'purl': purl,
            'externalReferences': [{'type': 'distribution', 'url': f'https://www.nuget.org/packages/{name}/{ver}'}],
            'properties': property_list([('pptcom:shipped', shipped)]),
        })
        top.append(purl)

    # The VSTO runtime's Utilities DLL, shipped unmodified next to pptcom.dll
    util_version = run('powershell', '-NoProfile', '-Command', f"(Get-Item '{UTILITIES}').VersionInfo.FileVersion")
    util_ref = f'pkg:generic/microsoft/Microsoft.Office.Tools.Common.v4.0.Utilities@{util_version}'
    components.append({
        'type': 'library', 'bom-ref': util_ref, 'supplier': {'name': 'Microsoft Corporation'},
        'name': 'Microsoft.Office.Tools.Common.v4.0.Utilities', 'version': util_version,
        'hashes': [{'alg': 'SHA-256', 'content': sha256(UTILITIES)}],
        'licenses': [{'license': {'name': 'Visual Studio 2022 Distributable Code'}}],
        'purl': util_ref,
        'properties': property_list([('pptcom:shipped', 'next to pptcom.dll, unmodified (from Visual Studio\'s Office workload)')]),
    })
    top.append(util_ref)

    # PDF.js: Mozilla's release archive, and the parts of it pptcom.dll embeds
    p = lock['pdfjs']
    pdfjs_ref = f"pkg:github/{p['repository']}@{p['version']}"
    parts = []
    for slug, title, spdx, where in PDFJS_PARTS:
        ref = f"pkg:generic/{slug}@{p['version'].lstrip('v')}?download_url={p['url']}"
        parts.append({'type': 'library', 'bom-ref': ref, 'name': slug, 'version': p['version'].lstrip('v'), 'description': title,
                      'licenses': [{'license': {'id': spdx}}], 'purl': ref,
                      'properties': property_list([('pptcom:shipped', f'embedded in pptcom.dll ({where})')])})
    components.append({
        'type': 'library', 'bom-ref': pdfjs_ref, 'supplier': {'name': 'Mozilla Foundation'}, 'name': 'pdf.js',
        'version': p['version'], 'hashes': [{'alg': 'SHA-256', 'content': p['sha256']}],
        'licenses': [{'license': {'id': 'Apache-2.0'}}], 'purl': pdfjs_ref,
        'externalReferences': [{'type': 'distribution', 'url': p['url']}],
        'properties': property_list([('pptcom:shipped', 'embedded in pptcom.dll (vendor/pdfjs: pdf.mjs, pdf.worker.mjs)'),
                                     ('pptcom:published', p['published_at'])]),
        'components': parts,
    })
    top.append(pdfjs_ref)

    # anydoc: the pinned source, built here; and the crates in its .wasm
    a = lock['anydoc']
    anydoc_ref = f"pkg:github/{a['repository']}@{a['version']}"
    rows, checksums = crates(lock)
    crate_ref = {}
    crate_components = []
    for name, ver, lic, _ in rows:
        key = (name, ver)
        if key in crate_ref:
            continue
        checksum = checksums.get(key)
        if checksum is None:
            # anydoc's own crates (anydoc, anydoc-wasm) come from the source archive, not crates.io
            crate_ref[key] = anydoc_ref
            continue
        purl = f'pkg:cargo/{name}@{ver}'
        crate_ref[key] = purl
        crate_components.append({
            'type': 'library', 'bom-ref': purl, 'name': name, 'version': ver,
            'hashes': [{'alg': 'SHA-256', 'content': checksum}],
            'licenses': [{'expression': license_expression(lic)}] if lic else [{'license': {'name': 'not declared by the crate'}}],
            'purl': purl,
            'properties': property_list([('pptcom:shipped', "compiled into anydoc's WebAssembly build (embedded in pptcom.dll)")]),
        })
    components.append({
        'type': 'library', 'bom-ref': anydoc_ref, 'supplier': {'name': 'Sideguide Technologies Inc. (Firecrawl)'},
        'name': 'anydoc', 'version': a['version'], 'hashes': [{'alg': 'SHA-256', 'content': a['sha256']}],
        'licenses': [{'license': {'id': 'MIT'}}], 'purl': anydoc_ref,
        'externalReferences': [{'type': 'distribution', 'url': a['url']}, {'type': 'vcs', 'url': f"https://github.com/{a['repository']}"}],
        'properties': property_list([('pptcom:shipped', 'built from this source to WebAssembly (anydoc_wasm.js, anydoc_wasm_bg.wasm) and embedded in pptcom.dll'),
                                     ('pptcom:commit', a['commit'])]),
    })
    top.append(anydoc_ref)
    components.extend(sorted(crate_components, key=lambda c: c['bom-ref']))

    edges = crate_graph(rows)
    anydoc_deps = set()
    for key, targets in edges.items():
        source = crate_ref[key]
        refs = {crate_ref[t] for t in targets} - {source}
        if source == anydoc_ref:
            anydoc_deps |= refs
    dependencies.append({'ref': app_ref, 'dependsOn': sorted(top)})
    dependencies.append({'ref': anydoc_ref, 'dependsOn': sorted(anydoc_deps)})
    crate_deps = {}
    for key, targets in edges.items():
        source = crate_ref[key]
        if source == anydoc_ref:
            continue
        crate_deps.setdefault(source, set()).update(crate_ref[t] for t in targets if crate_ref[t] != anydoc_ref)
    for ref in sorted(crate_deps):
        dependencies.append({'ref': ref, 'dependsOn': sorted(crate_deps[ref] - {ref})})
    dependencies.append({'ref': pdfjs_ref, 'dependsOn': sorted(c['bom-ref'] for c in parts)})

    # What built it
    wix = re.search(r'Sdk="WixToolset\.Sdk/([^"]+)"', WIXPROJ.read_text(encoding='utf-8'))[1]
    build_nuget = []
    for name, description in BUILD_NUGET.items():
        key = next(k for k in assets['libraries'] if k.split('/')[0] == name)
        ver = key.split('/')[1]
        build_nuget.append({'type': 'application', 'name': name, 'version': ver, 'purl': f'pkg:nuget/{name}@{ver}',
                            'hashes': [{'alg': 'SHA-512', 'content': base64.b64decode(assets['libraries'][key]['sha512']).hex()}],
                            'description': description})
    tools = [
        {'type': 'application', 'name': 'rust', 'version': lock['rust']['toolchain'], 'purl': f"pkg:generic/rust@{lock['rust']['toolchain']}",
         'description': f"rustc and cargo, target {lock['rust']['target']}: build anydoc"},
        {'type': 'application', 'name': 'wasm-bindgen', 'version': lock['wasm-bindgen']['version'],
         'purl': f"pkg:github/{lock['wasm-bindgen']['repository']}@{lock['wasm-bindgen']['version']}",
         'hashes': [{'alg': 'SHA-256', 'content': lock['wasm-bindgen']['sha256']}], 'description': "generates anydoc's JavaScript glue"},
        {'type': 'application', 'name': 'binaryen', 'version': lock['binaryen']['version'],
         'purl': f"pkg:github/{lock['binaryen']['repository']}@{lock['binaryen']['version']}",
         'hashes': [{'alg': 'SHA-256', 'content': lock['binaryen']['sha256']}], 'description': "wasm-opt: optimizes anydoc's .wasm"},
        *build_nuget,
        {'type': 'application', 'name': 'WixToolset.Sdk', 'version': wix, 'purl': f'pkg:nuget/WixToolset.Sdk@{wix}', 'description': 'builds the MSI'},
        {'type': 'application', 'name': f'{REPOSITORY} scripts/build/sbom.py', 'version': commit},
    ]

    return {
        'bomFormat': 'CycloneDX',
        'specVersion': SPEC,
        'serialNumber': f'urn:uuid:{uuid.uuid5(uuid.NAMESPACE_URL, f"https://github.com/{REPOSITORY}#{msi_hash}")}',
        'version': 1,
        'metadata': {
            'timestamp': timestamp,
            'tools': {'components': tools},
            'supplier': {'name': 'Kashihara City'},
            'component': {
                'type': 'application', 'bom-ref': app_ref, 'supplier': {'name': 'Kashihara City'},
                'name': 'vsto-llm-pptcom', 'version': version,
                'description': 'PowerPoint VSTO add-in: an LLM edits the open deck through pptcom',
                'hashes': [{'alg': 'SHA-256', 'content': msi_hash}],
                'licenses': [{'license': {'id': 'MIT'}}], 'purl': app_ref,
                'externalReferences': [{'type': 'vcs', 'url': f'https://github.com/{REPOSITORY}'}],
                'properties': property_list([('pptcom:msi', pathlib.Path(msi).name), ('pptcom:commit', commit),
                                             ('pptcom:commit_time', commit_time)]),
            },
        },
        'components': components,
        'dependencies': dependencies,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument('--msi', required=True)
    ap.add_argument('--out', required=True)
    args = ap.parse_args()
    bom = build(args.msi)
    out = pathlib.Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(bom, indent=2, ensure_ascii=False) + '\n', encoding='utf-8', newline='\n')
    crates_count = sum(1 for c in bom['components'] if c['purl'].startswith('pkg:cargo/'))
    print(f"SBOM {out}: {len(bom['components'])} components ({crates_count} Rust crates), MSI SHA-256 {bom['metadata']['component']['hashes'][0]['content']}")


if __name__ == '__main__':
    main()
