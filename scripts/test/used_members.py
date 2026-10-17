"""Count the COM members the model has actually used in every logged ComScript
(add-in session.log, rig transcripts, older one-shot scripts)."""
import collections
import glob
import io
import json
import os
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..")
STRING_LITERAL = re.compile(r'"(?:[^"\\]|\\.|"")*"')
MEMBER = re.compile(r"\.([A-Za-z_]\w*)")

scripts = []
log = open(os.path.expandvars(r"%LOCALAPPDATA%\KashiharaCity\pptcom\session.log"), encoding="utf-8").read()
for m in re.finditer(r"--- tool pptcom args:\n(\{.*?\})\nresult:", log, re.S):
    try:
        scripts.append(json.loads(m.group(1))["script"])
    except Exception:
        pass
for f in glob.glob(os.path.expandvars(r"%LOCALAPPDATA%\KashiharaCity\pptcom\generated-pptx\run*\transcript.json")):
    for x in json.load(open(f, encoding="utf-8")):
        for tc in x.get("tool_calls") or []:
            if tc["function"]["name"] == "pptcom":
                try:
                    scripts.append(json.loads(tc["function"]["arguments"])["script"])
                except Exception:
                    pass
for f in glob.glob(os.path.join(ROOT, "old", "experiment", "out", "*", "com_script*.txt")) + \
         glob.glob(os.path.join(ROOT, "old", "experiment", "out", "run8*", "reply.txt")):
    scripts.append(open(f, encoding="utf-8").read())

cnt = collections.Counter()
for s in scripts:
    for line in s.splitlines():
        code = STRING_LITERAL.sub('""', line)
        for mm in MEMBER.finditer(code):
            cnt[mm.group(1)] += 1

print("scripts:", len(scripts), " distinct members used:", len(cnt))
print(", ".join("%s(%d)" % kv for kv in cnt.most_common()))
