"""Collect every pptcom script from past runs into one corpus file for parser regression checks.

Usage: python scripts/test/extract_corpus.py <corpus.txt>
Then:  ComScriptRunner.exe parsecheck <corpus.txt> <out.tsv>   (once before and once after a parser
       change; compare the two TSVs - a line that was OK and is now NG is a regression).

Sources: generated-pptx/*/session.log (add-in runs) and generated-pptx/*/transcript.json (rig runs), in
%LOCALAPPDATA%\\KashiharaCity\\pptcom\\generated-pptx.
Old notation is normalized to the current one (pres -> ActivePresentation, app -> Application,
$name -> name) so those lines test the grammar, not the renames. The corpus holds model output
from past runs; keep it out of the repository, like the run folders themselves."""
import io
import json
import os
import re
import sys

ROOT = os.path.expandvars(r"%LOCALAPPDATA%\KashiharaCity\pptcom\generated-pptx")


def normalize(s):
    out = []
    for line in s.replace("\r\n", "\n").split("\n"):
        # outside strings only
        parts = re.split(r'("(?:[^"]|"")*")', line)
        for i in range(0, len(parts), 2):
            p = parts[i]
            p = re.sub(r"\bpres\b", "ActivePresentation", p)
            p = re.sub(r"\bapp\b", "Application", p)
            p = re.sub(r"\$([A-Za-z_])", r"\1", p)
            parts[i] = p
        out.append("".join(parts))
    return "\n".join(out)


def scripts_of(run_dir):
    scripts = []
    log = os.path.join(run_dir, "session.log")
    tr = os.path.join(run_dir, "transcript.json")
    if os.path.isfile(log):
        lines = io.open(log, encoding="utf-8", errors="replace").read().split("\n")
        for i, l in enumerate(lines):
            if l.endswith("--- tool pptcom args:") and i + 1 < len(lines):
                try:
                    scripts.append(json.loads(lines[i + 1])["script"])
                except Exception:
                    pass
    elif os.path.isfile(tr):
        try:
            msgs = json.load(io.open(tr, encoding="utf-8"))
        except Exception:
            msgs = []
        if isinstance(msgs, dict):
            msgs = msgs.get("messages", [])
        for m in msgs:
            for c in (m.get("tool_calls") or []) if isinstance(m, dict) else []:
                f = c.get("function", {})
                if f.get("name") == "pptcom":
                    try:
                        scripts.append(json.loads(f.get("arguments") or "{}").get("script") or "")
                    except Exception:
                        pass
    return scripts


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    runs = []
    for name in sorted(os.listdir(ROOT)):
        d = os.path.join(ROOT, name)
        if os.path.isdir(d):
            scripts = scripts_of(d)
            if scripts:
                runs.append((name, scripts))
    n = 0
    with io.open(sys.argv[1], "w", encoding="utf-8", newline="\n") as f:
        for name, scripts in runs:
            f.write("### RUN " + name + "\n")
            for s in scripts:
                f.write("### CALL\n" + normalize(s).rstrip("\n") + "\n")
                n += 1
    print(len(runs), "runs,", n, "scripts")


if __name__ == "__main__":
    main()
