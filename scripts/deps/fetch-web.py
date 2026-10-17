"""Fetch the task pane's pinned official sources into deps/ and record them in deps/web-sources.lock.json.

Run it when a version below changes (and only then); the build itself never downloads these. Like ksip's
scripts/deps/fetch-baresip.py: an entry already in the lock at the same version is kept as it is and its
archive is checked against the recorded hash; a new version is read from its official source and refused
if it was published less than seven days ago.

  anydoc  the source of the tagged commit (GitHub codeload), built to WebAssembly by pptcom/WebVendor.ps1
  PDF.js  Mozilla's release archive (pdfjs-<version>-dist.zip): building PDF.js from source needs Node/npm
  tools   the Rust toolchain and the prebuilt wasm-bindgen and wasm-opt (binaryen) that build anydoc; they
          are not kept in deps/, only their URL and hash (WebVendor.ps1 downloads and checks them)

usage: python -X utf8 scripts/deps/fetch-web.py
"""
import datetime, hashlib, io, json, pathlib, re, subprocess, tarfile, urllib.request

ROOT = pathlib.Path(__file__).resolve().parents[2]
LOCK = ROOT / 'deps' / 'web-sources.lock.json'
# Nothing published in the last seven days is accepted, whenever this runs.
CUTOFF = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=7)

# name: (repository, tag); the archive is deps/<name>.tar.gz
SOURCES = {'anydoc': ('firecrawl/anydoc', 'v0.2.4')}
# name: (repository, tag, release asset); the archive is deps/<name>.zip
RELEASES = {'pdfjs': ('mozilla/pdf.js', 'v6.4.299', 'pdfjs-6.4.299-dist.zip')}
# name: (repository, tag, release asset, version); not kept in deps/
TOOLS = {
    # must be the version of the wasm-bindgen crate in anydoc's Cargo.lock (checked below)
    'wasm-bindgen': ('wasm-bindgen/wasm-bindgen', '0.2.126', 'wasm-bindgen-0.2.126-x86_64-pc-windows-msvc.tar.gz', '0.2.126'),
    # the version wasm-pack 0.15.0 (anydoc's own release build) uses
    'binaryen': ('WebAssembly/binaryen', 'version_117', 'binaryen-version_117-x86_64-windows.tar.gz', 'version_117'),
}
# The compiler ksip pins too (its rust-toolchain.toml), with the wasm32-unknown-unknown target.
RUST = '1.99.0'


def get(url, accept=None):
    headers = {'User-Agent': 'pptcom-deps/0.1'}
    if accept:
        headers['Accept'] = accept
    return urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=120).read()


def api(path):
    return json.loads(get('https://api.github.com/repos/' + path, 'application/vnd.github+json'))


def old_enough(name, date):
    if datetime.datetime.fromisoformat(date.replace('Z', '+00:00')) > CUTOFF:
        raise RuntimeError(f'{name}: published {date}, less than seven days ago')


def release_asset(name, repo, tag, asset_name):
    """(url, published_at, sha256 GitHub publishes for it or None, data)"""
    rel = api(f'{repo}/releases/tags/{tag}')
    asset = next((a for a in rel['assets'] if a['name'] == asset_name), None)
    if asset is None:
        raise RuntimeError(f'{name}: no asset {asset_name} in {repo} {tag}')
    published = rel['published_at']
    old_enough(name, published)
    data = get(asset['browser_download_url'])
    digest = hashlib.sha256(data).hexdigest()
    expected = (asset.get('digest') or '').removeprefix('sha256:') or None
    if expected is None:
        # older releases (binaryen) publish the hash as a separate .sha256 asset
        side = next((a for a in rel['assets'] if a['name'] == asset_name + '.sha256'), None)
        if side is not None:
            expected = get(side['browser_download_url']).decode().split()[0]
    if expected is None:
        raise RuntimeError(f'{name}: {repo} publishes no hash for {asset_name}')
    if digest != expected:
        raise RuntimeError(f'{name}: SHA256 of the download does not match the one {repo} publishes')
    return asset['browser_download_url'], published, digest, data


def main():
    old = json.loads(LOCK.read_text(encoding='utf-8')) if LOCK.exists() else {}
    records = {}

    for name, (repo, tag) in SOURCES.items():
        archive = ROOT / 'deps' / f'{name}.tar.gz'
        rec = old.get(name)
        if not (rec and rec['version'] == tag and archive.exists()):
            ref = api(f'{repo}/commits/{tag}')
            date = ref['commit']['committer']['date']
            old_enough(name, date)
            url = f'https://codeload.github.com/{repo}/tar.gz/{ref["sha"]}'
            data = get(url)
            archive.parent.mkdir(parents=True, exist_ok=True)
            archive.write_bytes(data)
            rec = {'repository': repo, 'version': tag, 'commit': ref['sha'], 'source_date': date,
                   'url': url, 'archive': archive.name, 'sha256': hashlib.sha256(data).hexdigest()}
            print('fetched', name, tag, rec['sha256'], flush=True)
        if hashlib.sha256(archive.read_bytes()).hexdigest() != rec['sha256']:
            raise RuntimeError(f'{name}: deps/{archive.name} does not match the lock')
        records[name] = rec

    for name, (repo, tag, asset) in RELEASES.items():
        archive = ROOT / 'deps' / f'{name}.zip'
        rec = old.get(name)
        if not (rec and rec['version'] == tag and archive.exists()):
            url, published, digest, data = release_asset(name, repo, tag, asset)
            archive.write_bytes(data)
            rec = {'repository': repo, 'version': tag, 'published_at': published, 'url': url,
                   'archive': archive.name, 'sha256': digest}
            print('fetched', name, tag, digest, flush=True)
        if hashlib.sha256(archive.read_bytes()).hexdigest() != rec['sha256']:
            raise RuntimeError(f'{name}: deps/{archive.name} does not match the lock')
        records[name] = rec

    for name, (repo, tag, asset, version) in TOOLS.items():
        rec = old.get(name)
        if not (rec and rec['version'] == version):
            url, published, digest, _ = release_asset(name, repo, tag, asset)
            rec = {'repository': repo, 'version': version, 'published_at': published, 'url': url,
                   'sha256': digest, 'source_type': 'tool'}
            print('recorded', name, version, digest, flush=True)
        records[name] = rec

    rec = old.get('rust')
    if not (rec and rec['toolchain'] == RUST):
        # the release date is the commit date rustc reports ("rustc 1.99.0 (b940084d7 2026-09-28)")
        out = subprocess.run(['rustc', f'+{RUST}', '--version'], capture_output=True, text=True, check=True).stdout
        date = re.search(r'(\d{4}-\d{2}-\d{2})\)', out).group(1) + 'T00:00:00Z'
        old_enough('rust', date)
        rec = {'toolchain': RUST, 'target': 'wasm32-unknown-unknown', 'source_date': date, 'source_type': 'tool'}
    records['rust'] = rec

    # wasm-bindgen's generated glue must match the crate the source pins
    with tarfile.open(ROOT / 'deps' / 'anydoc.tar.gz') as tar:
        lock_member = next(m for m in tar.getmembers() if m.name.count('/') == 1 and m.name.endswith('/Cargo.lock'))
        cargo_lock = tar.extractfile(lock_member).read().decode()
    crate = re.search(r'name = "wasm-bindgen"\nversion = "([^"]+)"', cargo_lock).group(1)
    if crate != records['wasm-bindgen']['version']:
        raise RuntimeError(f'wasm-bindgen: anydoc pins the crate at {crate}, the tool is {records["wasm-bindgen"]["version"]}')

    LOCK.write_text(json.dumps(records, indent=2) + '\n', encoding='utf-8')
    print('lock:', ', '.join(f'{k} {v.get("version") or v.get("toolchain")}' for k, v in records.items()))


if __name__ == '__main__':
    main()
