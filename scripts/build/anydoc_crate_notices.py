"""Prints the notices of the Rust crates compiled into anydoc's WebAssembly build, as Markdown for
licenses/THIRD-PARTY-NOTICES.md. Run it again when anydoc's version in deps/web-sources.lock.json changes,
and replace the "Rust crates in anydoc's WebAssembly build" part of that file with its output. It reads the
source the build uses, deps/anydoc.tar.gz, unpacked:

  mkdir <dir> && tar -xzf deps/anydoc.tar.gz -C <dir> --strip-components 1
  python -I scripts/build/anydoc_crate_notices.py <dir> > crates.md

Needs cargo (Rust). Only `cargo tree` runs: it reads the lock file and downloads the crate archives into
cargo's cache, but builds nothing. The crates are those of the wasm32 target, normal dependencies only
(no build scripts, no proc-macros: they run at compile time and are not in the .wasm). Each crate's
license files are read from its archive; a crate that offers a choice is used under MIT where MIT is one
of the choices. Crates whose license text is identical share one copy of it.
"""
import glob
import os
import re
import subprocess
import sys
import tarfile

LICENSE_FILE = re.compile(r"^(LICEN[CS]E|COPYING|COPYRIGHT|NOTICE|UNLICENSE)([-._].*)?$", re.I)

MIT_TEMPLATE = """MIT License

Copyright (c) {authors}

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE."""


def crates(source):
    out = subprocess.run(
        ["cargo", "tree", "--manifest-path", os.path.join(source, "wasm", "Cargo.toml"), "--locked",
         "--target", "wasm32-unknown-unknown", "-e", "normal,no-proc-macro", "--prefix", "none",
         "--format", "{p}|{l}"],
        check=True, capture_output=True, text=True, encoding="utf-8").stdout
    seen = {}
    for line in out.splitlines():
        line = line.replace(" (*)", "").strip()
        if not line or line.startswith("anydoc"):
            continue
        spec, lic = line.split("|", 1)
        name, version = spec.split(" ")[:2]
        seen[(name, version.lstrip("v"))] = lic
    return sorted(seen.items())


def used_under(expr):
    """The licenses the crate is used under: MIT out of a choice, plus whatever is ANDed to it."""
    expr = expr.replace("/", " OR ")
    parts = [p.strip(" ()") for p in re.split(r"\)\s+AND\s+|\s+AND\s+\(", expr)]
    chosen = []
    for p in parts:
        options = [o.strip(" ()") for o in p.split(" OR ")]
        chosen.append("MIT" if "MIT" in options else options[0])
    return chosen


def license_texts(name, version, chosen):
    archive = glob.glob(os.path.join(os.path.expanduser("~"), ".cargo", "registry", "cache", "*", f"{name}-{version}.crate"))
    if not archive:
        raise SystemExit(f"{name} {version}: archive not in cargo's cache")
    texts = {}
    authors = ""
    with tarfile.open(archive[0], "r:gz") as tar:
        for m in tar.getmembers():
            parts = m.name.split("/")
            if len(parts) == 2 and m.isfile() and LICENSE_FILE.match(parts[1]):
                texts[parts[1]] = tar.extractfile(m).read().decode("utf-8", "replace").replace("\r\n", "\n").strip()
            if len(parts) == 2 and parts[1] == "Cargo.toml":
                toml = tar.extractfile(m).read().decode("utf-8", "replace")
                found = re.search(r"^authors\s*=\s*\[(.*?)\]", toml, re.M | re.S)
                authors = ", ".join(re.findall(r'"([^"]*)"', found.group(1))) if found else ""
    if not texts and chosen == ["MIT"]:
        # The crate ships no license file: the MIT text with the authors its Cargo.toml names.
        return [("MIT, from Cargo.toml", MIT_TEMPLATE.format(authors=authors or "the " + name + " authors"))], []
    picked = []
    for lic in chosen:
        key = {"MIT": "MIT", "Apache-2.0": "APACHE", "Zlib": "ZLIB", "BSD-3-Clause": "BSD", "Unicode-3.0": "UNICODE"}.get(lic, "")
        files = [f for f in texts if key and key in f.upper()]
        if not files and lic == "BSD-3-Clause":
            files = [f for f in texts if "WHATWG" in f.upper()]   # encoding_rs: the WHATWG Encoding Standard's BSD license
        if not files and len(texts) >= 1:
            # a single LICENSE / COPYING file, or the extra license of an AND (encoding_rs: COPYRIGHT-style files)
            files = [f for f in texts if f.upper() in ("LICENSE", "LICENSE.MD", "LICENSE.TXT", "COPYING")] or \
                    [f for f in texts if lic.split("-")[0].upper() in f.upper()]
        picked += [f for f in files if f not in picked]
    # NOTICE and COPYRIGHT files always go along
    picked += [f for f in texts if f.upper().startswith(("NOTICE", "COPYRIGHT")) and f not in picked]
    return [(f, texts[f]) for f in picked], sorted(texts)


def main():
    sys.stdout.reconfigure(encoding="utf-8", newline="\n")
    source = sys.argv[1]
    groups = {}   # text -> [crate labels]
    rows = []
    missing = []
    for (name, version), expr in crates(source):
        chosen = used_under(expr)
        picked, available = license_texts(name, version, chosen)
        rows.append((name, version, expr, " AND ".join(chosen)))
        if not picked:
            missing.append(f"{name} {version} ({expr}; files: {', '.join(available) or 'none'})")
        for f, text in picked:
            groups.setdefault(text, []).append(f"{name} {version} ({f})")
    if missing:
        raise SystemExit("no license text found for: " + "; ".join(missing))

    print("| Crate | Version | License | Used under |")
    print("|---|---|---|---|")
    for name, version, expr, chosen in rows:
        print(f"| {name} | {version} | {expr} | {chosen} |")
    print()
    for text, labels in groups.items():
        print("#### " + ", ".join(labels))
        print()
        print("```")
        print(text)
        print("```")
        print()


if __name__ == "__main__":
    main()
