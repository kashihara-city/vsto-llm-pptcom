"""Controlled test of the parse-error message: rebuild run13's deck up to the turn where the
model sent a script with an unclosed string, replay the conversation to that point, answer the
broken script with the CURRENT parse-error text (from ComScriptRunner), and watch what the model
sends next.

usage: python replay_parse_error.py [trials=3]

Success = the next pptcom call resends the whole slide-3 script (defines box5 itself) and runs OK.
"""
import base64
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import tool_loop  # noqa: E402

SRC = os.path.join(tool_loop.GENERATED_DIR, "run13_autorender", "transcript.json")
FAILING_CALL = 2   # 0-based index among pptcom calls in run13 (the unclosed-string script)


def load():
    msgs = json.load(open(SRC, encoding="utf-8"))
    pptcom_seen = -1
    cut = None
    for i, m in enumerate(msgs):
        for tc in m.get("tool_calls") or []:
            if tc["function"]["name"] == "pptcom":
                pptcom_seen += 1
                if pptcom_seen == FAILING_CALL:
                    cut = (i, tc)
    assert cut, "failing call not found"
    return msgs, cut


def trial(k):
    msgs, (fail_idx, fail_tc) = load()
    runner = tool_loop.Runner()
    try:
        # 1. rebuild the deck state as it was before the failing call
        for m in msgs[:fail_idx]:
            for tc in m.get("tool_calls") or []:
                args = json.loads(tc["function"]["arguments"])
                if tc["function"]["name"] in ("zimage", "pptcom"):
                    text, _ = tool_loop.run_tool(runner, tc["function"]["name"], args)
                    if not text.startswith(("OK", "Generated")):
                        print("   rebuild step failed:", text.splitlines()[0])

        # 2. conversation up to (and including) the failing assistant turn; image messages are
        #    re-rendered from the rebuilt deck
        convo = [{"role": "system", "content": tool_loop.SYSTEM}]
        for m in msgs[1:fail_idx + 1]:
            if m["role"] == "user" and isinstance(m.get("content"), str) and m["content"].startswith("[text +"):
                slides = [1, 2]
                parts = [{"type": "text", "text": "Slide rendering(s):"}]
                for n in slides:
                    png = runner.call({"export": n}).get("png")
                    if png:
                        parts.append({"type": "text", "text": "Slide %d:" % n})
                        parts.append({"type": "image_url", "image_url": {"url": "data:image/png;base64," + base64.b64encode(open(png, "rb").read()).decode()}})
                convo.append({"role": "user", "content": parts})
            else:
                convo.append(m)

        # 3. answer the broken script with today's parse-error message
        broken = json.loads(fail_tc["function"]["arguments"])
        text, _ = tool_loop.run_tool(runner, "pptcom", broken)
        convo.append({"role": "tool", "tool_call_id": fail_tc["id"], "content": text})
        if k == 0:
            print("---- message given to the model ----\n" + text + "\n------------------------------------")

        # 4. what does the model do next?
        payload = {"model": tool_loop.MODEL, "messages": tool_loop.strip_private(convo), "tools": tool_loop.TOOLS,
                   "tool_choice": "auto", "max_tokens": 32000, "stream": True, "stream_options": {"include_usage": True},
                   "mm_processor_kwargs": {"min_pixels": 3136}}
        r = tool_loop.stream(payload)
        calls = [c for c in r["tool_calls"] if c["function"]["name"] == "pptcom"]
        if not calls:
            print("trial %d: no pptcom call; tools=%s content=%r" % (k + 1, [c["function"]["name"] for c in r["tool_calls"]], r["content"][:80]))
            return False
        script = json.loads(calls[0]["function"]["arguments"])["script"]
        first = script.strip().splitlines()[0]
        res, _ = tool_loop.run_tool(runner, "pptcom", {"script": script})
        ok = res.startswith("OK")
        print("trial %d: %4.1fs reasoning=%d chars | %d lines, starts with %r | %s"
              % (k + 1, r["elapsed_sec"], len(r["reasoning"]), len(script.strip().splitlines()), first[:50], res.splitlines()[0][:90]))
        return ok
    finally:
        runner.close()


if __name__ == "__main__":
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 3
    results = [trial(k) for k in range(n)]
    print("recovered in one turn: %d / %d" % (sum(results), n))
