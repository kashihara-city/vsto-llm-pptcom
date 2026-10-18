"""Fail if anything Git tracks names the local network, a host, a machine's paths or a secret.

Like ksip's scripts/test/no-secrets.py, with three additions for this repository: API keys (the add-in keeps
its key in Windows Credential Manager, never in a file), paths of a developer's machine, and an optional deny
list of exact strings - the real endpoint host or model aliases, which no pattern can tell from ordinary
words. The deny list sits in notes/secrets-denylist.txt (git-ignored, one string per line), so the names
themselves are never committed; the check uses it when it exists.

  python -X utf8 scripts/test/no-secrets.py              the files Git tracks now
  python -X utf8 scripts/test/no-secrets.py --history    every file in every commit, and the commit messages
                                                         (before the first push of a history)
  python -X utf8 scripts/test/no-secrets.py --self-test
"""
import pathlib, re, subprocess, sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
DENYLIST = ROOT / "notes" / "secrets-denylist.txt"

# 公開してよいのは手順だけ。接続先・API キー・モデル名はレジストリと資格情報から読む（README）。
PATTERNS = [
    ("私有IPアドレス", re.compile(r"\b(?:10\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[01])\.\d{1,3}|192\.168\.\d{1,3})\.\d{1,3}\b")),
    ("社内ホスト名", re.compile(r"\b[A-Za-z0-9][A-Za-z0-9-]*\.(?:local|lan|internal|intra|corp|home)\b")),
    ("32桁の秘密らしき文字列", re.compile(r"(?<![0-9a-fA-F])[0-9a-f]{32}(?![0-9a-fA-F])")),
    ("APIキーらしき文字列", re.compile(r"\b(?:sk|pk|rk)-[A-Za-z0-9_-]{16,}|\bBearer\s+[A-Za-z0-9._~+/-]{16,}=*")),
    ("開発機のパス", re.compile(r"(?i)\b[a-z]:[\\/]+Users[\\/]+(?!(?:Public|Default|<|%|\{|\$|\.\.\.|…))[^\\/\s\"'<>|:*?]+|/c/Users/(?!Public\b)[^/\s\"']+|/home/(?!runner\b)[a-z][a-z0-9_-]+")),
    # 行頭に限る。本物のPEMは必ず行頭から始まり、ソース中の文字列リテラルは通る。
    ("PEMの中身", re.compile(r"^-----BEGIN (?:RSA |EC )?PRIVATE KEY-----|^-----BEGIN CERTIFICATE-----")),
]
# 依存の固定に使うハッシュと、説明のための例示アドレスは秘密ではない。
ALLOWED = [
    re.compile(r'"(?:sha256|checksum|commit)"\s*:\s*"[0-9a-f]+"'),
    re.compile(r"codeload\.github\.com/[\w.-]+/[\w.-]+/tar\.gz/[0-9a-f]{40}"),
    re.compile(r"\b(?:192\.0\.2|198\.51\.100|203\.0\.113)\.\d{1,3}\b"),  # RFC 5737 の例示用
    re.compile(r"\b(?:127\.0\.0\.1|0\.0\.0\.0)\b"),
    re.compile(r"[0-9a-f]{40}"),  # git の SHA-1
    re.compile(r"\bpptcom\.invalid\b"),  # 作業ウィンドウの仮想ホスト（.invalid は名前解決されない予約名）
    re.compile(r"\bsigning\.local\.props\b|\bsigning\.local\b"),  # 署名の設定ファイルの名前（git 管理外）
    re.compile(r"Version[=>\"]\s*\d+(?:\.\d+){3}\b"),  # アセンブリの版（Version=10.0.0.0）
    re.compile(r"\bPath\.home\(\)"),  # Python の pathlib（ホーム フォルダー）
]
# ファイルを限って許すもの。公式の CycloneDX スキーマ（deps/cyclonedx、SHA-256 を schema.lock.json で
# 固定）にある、ハッシュ値の例示（definitions.hash-content.examples）。値ごとに限るので、同じファイルの別の値は見逃さない。
ALLOWED_IN = {
    "deps/cyclonedx/bom-1.6.schema.json": [re.compile(r'"examples": \["3942447fac867ae5cdb3229b658f4d48"\]')],
}
# 中身が秘密そのものではなく、形式上ハッシュを大量に含むもの。
SKIP = {"deps/web-sources.lock.json", "scripts/test/no-secrets.py"}
BINARY = {".tar", ".gz", ".zip", ".exe", ".dll", ".ico", ".png", ".jpg", ".pdf", ".docx", ".xlsx", ".pptx", ".msi", ".pfx", ".snk"}


def denylist():
    if not DENYLIST.exists():
        return []
    words = [w.strip() for w in DENYLIST.read_text(encoding="utf-8").splitlines()]
    return [w for w in words if w and not w.startswith("#")]


def findings(line, deny=(), name=""):
    """What in one line looks like a secret, minus what an allowed value covers.

    The allow list is matched by position, not by line: a permitted example address or a dependency
    hash on the same line as a real address does not hide the real one. A file may have values of its
    own allowed (ALLOWED_IN). Deny-list words are not subject to the allow list."""
    patterns = ALLOWED + ALLOWED_IN.get(name, [])
    allowed = [match.span() for pattern in patterns for match in pattern.finditer(line)]
    found = []
    for reason, pattern in PATTERNS:
        for match in pattern.finditer(line):
            if any(start <= match.start() and match.end() <= end for start, end in allowed):
                continue
            found.append((reason, match.group(0)))
    lower = line.lower()
    for word in deny:
        if word.lower() in lower:
            found.append(("禁止語リストの語", word))
    return found


def selftest():
    """The scanner on made-up lines: an allowed value beside a forbidden one hides nothing."""
    assert findings("server = 192.0.2.10") == []
    assert findings('"sha256": "' + "ab" * 32 + '"') == []
    assert [r for r, _ in findings("192.0.2.10 and 10.1.2.3 on one line")] == ["私有IPアドレス"]
    assert [r for r, _ in findings('"sha256": "' + "ab" * 32 + '" host llm.intra')] == ["社内ホスト名"]
    assert findings("loopback 127.0.0.1 only") == []
    assert findings("https://pptcom.invalid/index.html") == []
    assert [r for r, _ in findings("key = sk-" + "a1B2" * 6)] == ["APIキーらしき文字列"]
    assert [r for r, _ in findings(r"C:\Users\someone\Desktop\x")] == ["開発機のパス"]
    assert [r for r, _ in findings("/c/Users/someone/AppData")] == ["開発機のパス"]
    assert findings(r"%LOCALAPPDATA%\KashiharaCity\pptcom and C:\Users\<name>\AppData") == []
    assert [r for r, _ in findings("model = Some-Model-Alias", ["some-model-alias"])] == ["禁止語リストの語"]
    assert findings(r'<Reference Include="X, Version=10.0.0.0, Culture=neutral">') == []
    assert [r for r, _ in findings("Version=10.0.0.0 at 10.0.0.5")] == ["私有IPアドレス"]
    assert findings("Microsoft.CodeAnalysis.NetAnalyzers 10.0.401") == [], "a three-part version is not an address"
    assert findings("cache = pathlib.Path.home() / '.nuget'") == []
    assert [r for r, _ in findings("Path.home() and nas.home")] == ["社内ホスト名"]
    assert findings("pptcom/signing.local.props") == []
    assert [r for r, _ in findings("signing.local.props on build.local")] == ["社内ホスト名"]
    example = '"examples": ["3942447fac867ae5cdb3229b658f4d48"],'
    assert findings(example, (), "deps/cyclonedx/bom-1.6.schema.json") == []
    assert [r for r, _ in findings(example)] == ["32桁の秘密らしき文字列"], "the schema's example is allowed in the schema only"
    assert [r for r, _ in findings('"examples": ["' + "cd" * 16 + '"],', (), "deps/cyclonedx/bom-1.6.schema.json")] == ["32桁の秘密らしき文字列"]
    print("self-test: allowed values do not hide a neighbour")


def git(*args, binary=False):
    out = subprocess.check_output(["git", *args], cwd=ROOT)
    return out if binary else out.decode("utf-8", errors="replace")


def scan_text(label, text, deny, problems, name=""):
    for number, line in enumerate(text.splitlines(), 1):
        for reason, found in findings(line, deny, name):
            problems.append("%s:%d %s: %s" % (label, number, reason, found))


def scan_files(deny):
    problems, count = [], 0
    for name in git("ls-files").splitlines():
        if not name or name in SKIP or pathlib.Path(name).suffix.lower() in BINARY:
            continue
        path = ROOT / name
        if not path.is_file():
            continue
        count += 1
        scan_text(name, path.read_text(encoding="utf-8", errors="replace"), deny, problems, name)
    return problems, "tracked files: %d" % count


def scan_history(deny):
    """Every blob reachable from any commit (once per content), and every commit message."""
    problems, blobs = [], 0
    seen = set()
    for commit in git("rev-list", "--all").split():
        scan_text("commit %s (message)" % commit[:7], git("log", "-1", "--format=%B", commit), deny, problems)
        for entry in git("ls-tree", "-r", commit).splitlines():
            meta, name = entry.split("\t", 1)
            _, kind, sha = meta.split()
            if kind != "blob" or sha in seen or name in SKIP or pathlib.Path(name).suffix.lower() in BINARY:
                continue
            seen.add(sha)
            blobs += 1
            scan_text("%s@%s" % (name, commit[:7]), git("cat-file", "blob", sha, binary=True).decode("utf-8", errors="replace"), deny, problems, name)
    return problems, "commits: %d, distinct files: %d" % (len(git("rev-list", "--all").split()), blobs)


def main():
    if "--self-test" in sys.argv:
        selftest()
        return
    deny = denylist()
    problems, summary = scan_history(deny) if "--history" in sys.argv else scan_files(deny)
    if problems:
        print("\n".join(problems))
        raise SystemExit("公開してはいけない情報が追跡対象にあります: %d件" % len(problems))
    print(summary, "：接続先・ホスト名・パス・秘密は見つからない",
          "（禁止語リスト %d 語も照合）" % len(deny) if deny else "（禁止語リストなし）")


if __name__ == "__main__":
    main()
