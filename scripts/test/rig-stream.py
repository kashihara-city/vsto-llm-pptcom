"""How the rig (tool_loop.py) reads a reply and ends a run that failed, without an LLM or PowerPoint.

usage: python scripts/test/rig-stream.py     exit code 0 = all passed

1. stream() against a server on the loopback address: what counts as a whole reply, and which failure each
   broken answer is (StreamFailure: HTTP 5xx, a dropped or refused connection, an error event, a stream that
   stalled or was cut off; RequestRefused: HTTP 4xx). A body or a stream that stops sending fails after
   STREAM_IDLE_SEC (2 s here).
2. The server's certificate and host name are checked (tls_context).
3. main() with a stand-in for ComScriptRunner: a run that a failure ended still saves the deck and
   summary.json, keeps the cause as its outcome, and exits with 1.
"""
import contextlib
import io
import json
import os
import shutil
import socket
import ssl
import sys
import tempfile
import threading
import time

sys.dont_write_bytecode = True  # no __pycache__ next to the scripts
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import tool_loop  # noqa: E402  (reads the add-in's settings; the test sends nothing to the endpoint)

tool_loop.API_KEY = ""  # nothing to send to the loopback server
tool_loop.STREAM_IDLE_SEC = 2

passed = failed = 0


def check(name, ok, detail):
    global passed, failed
    if ok:
        passed += 1
        return
    failed += 1
    print("FAIL " + name)
    print("     " + str(detail))


# ---- the server -------------------------------------------------------------------------------------

def sse(*events):
    return "".join("data: " + e + "\n\n" for e in events)


def content(text):
    return json.dumps({"choices": [{"delta": {"content": text}, "finish_reason": None}]})


def finish(reason):
    return json.dumps({"choices": [{"delta": {}, "finish_reason": reason}]})


def whole(body, status=200):
    """A whole response: with a Content-Length it ends when the body does."""
    b = body.encode("utf-8")
    return (b"HTTP/1.1 %d Status\r\nContent-Type: text/event-stream\r\nContent-Length: %d\r\nConnection: close\r\n\r\n"
            % (status, len(b)) + b)


def started(first, status=200):
    """The start of a response that goes on: chunked, one chunk so far, no end."""
    b = first.encode("utf-8")
    return (b"HTTP/1.1 %d Status\r\nContent-Type: text/event-stream\r\nTransfer-Encoding: chunked\r\n\r\n" % status +
            b"%X\r\n" % len(b) + b + b"\r\n")


def serve(response, hold):
    """One exchange: reads the request, writes `response`, then holds the connection open for `hold` seconds
    (sending nothing more) and closes it."""
    srv = socket.socket()
    srv.bind(("127.0.0.1", 0))
    srv.listen(1)
    port = srv.getsockname()[1]

    def run():
        conn, _ = srv.accept()
        data = b""
        while b"\r\n\r\n" not in data:
            data += conn.recv(65536)
        head, body = data.split(b"\r\n\r\n", 1)
        length = 0
        for line in head.split(b"\r\n"):
            if line.lower().startswith(b"content-length:"):
                length = int(line.split(b":")[1])
        while len(body) < length:
            body += conn.recv(65536)
        conn.sendall(response)
        if hold:
            time.sleep(hold)
        conn.close()
        srv.close()

    threading.Thread(target=run, daemon=True).start()
    return "http://127.0.0.1:%d" % port


def outcome_of(base_url):
    """What stream() made of the answer: ("ok", content, finish_reason) or (exception type, message), and the
    seconds it took."""
    tool_loop.BASE_URL = base_url
    t0 = time.time()
    try:
        r = tool_loop.stream({"model": "m", "messages": [], "stream": True})
        got = ("ok", r["content"], r["finish_reason"])
    except Exception as ex:  # what the rig would see
        got = (type(ex).__name__, str(ex))
    return got, time.time() - t0


def expect(name, response, ok, hold=0, within=None):
    got, secs = outcome_of(serve(response, hold))
    check(name, ok(got) and (within is None or secs < within), "%s after %.1f s" % (got, secs))


# ---- 1. stream() ------------------------------------------------------------------------------------

def failure(kind, text):
    return lambda g: g[0] == kind and text in g[1]


def replies():
    expect("finish_reason and [DONE]", whole(sse(content("Hello"), finish("stop"), "[DONE]")),
           lambda g: g == ("ok", "Hello", "stop"))
    expect("finish_reason without [DONE]", whole(sse(content("Hello"), finish("stop"))),
           lambda g: g == ("ok", "Hello", "stop"))
    expect("[DONE] without finish_reason", whole(sse(content("Hello"), "[DONE]")),
           lambda g: g == ("ok", "Hello", None))
    expect("a line that is not JSON is skipped", whole(sse(content("Hel"), "{not json", content("lo"), finish("stop"))),
           lambda g: g == ("ok", "Hello", "stop"))

    expect("cut off: no finish_reason, no [DONE]", whole(sse(content("partial"))), failure("StreamFailure", "ended before"))
    expect("an error event", whole(sse(json.dumps({"error": {"code": "500", "message": "x"}}))),
           failure("StreamFailure", "stream error"))
    expect("a stream that stops sending", started(sse(content("partial"))), failure("StreamFailure", "stalled"),
           hold=8, within=6)
    expect("the connection drops during the stream", started(sse(content("partial"))), failure("StreamFailure", ""),
           hold=0.5)

    expect("HTTP 503", whole("busy", 503), failure("StreamFailure", "503 busy"))
    expect("HTTP 503 whose body does not finish", started("busy", 503), failure("StreamFailure", "503"),
           hold=8, within=6)
    expect("HTTP 400", whole("bad request", 400), failure("RequestRefused", "400 bad request"))
    expect("HTTP 400 whose body does not finish", started("bad", 400), failure("RequestRefused", "400"),
           hold=8, within=6)

    # nobody listens on the port
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    got, secs = outcome_of("http://127.0.0.1:%d" % port)
    check("a refused connection", failure("StreamFailure", "connection failed")(got), got)


# ---- 2. TLS -----------------------------------------------------------------------------------------

def tls():
    ctx = tool_loop.tls_context()
    check("the certificate is checked", ctx.verify_mode == ssl.CERT_REQUIRED, ctx.verify_mode)
    check("the host name is checked", ctx.check_hostname, ctx.check_hostname)

    # stream() hands urlopen that context
    seen = {}

    def urlopen(req, context=None, timeout=None):
        seen["context"] = context
        raise OSError("not sent")

    real = tool_loop.urllib.request.urlopen
    tool_loop.urllib.request.urlopen = urlopen
    try:
        outcome_of("https://127.0.0.1:1")
    finally:
        tool_loop.urllib.request.urlopen = real
    sent = seen.get("context")
    check("stream() connects with it", isinstance(sent, ssl.SSLContext) and sent.verify_mode == ssl.CERT_REQUIRED and
          sent.check_hostname, sent)


# ---- 3. main() --------------------------------------------------------------------------------------

class FakeRunner:
    """ComScriptRunner serve, without PowerPoint: answers the probe and the save, records what was asked."""
    calls = []
    save_fails = False

    def __init__(self):
        FakeRunner.calls = []

    def call(self, req):
        FakeRunner.calls.append(next(iter(req)))
        if "save" in req:
            if FakeRunner.save_fails:
                raise RuntimeError("runner died")
            return {"saved": req["save"], "slides": 0}
        return {"summary": "", "placeholders": {}}

    def close(self):
        FakeRunner.calls.append("close")


def run_main(name, failing_stream, save_fails=False):
    """main() for one run whose first turn fails as `failing_stream` does. Returns the exit code, the summary's
    outcome (None when summary.json was not written) and what the runner was asked."""
    def stream(payload):
        failing_stream()

    out = tempfile.mkdtemp(prefix="rig-stream-")
    saved = (tool_loop.Runner, tool_loop.stream, tool_loop.GENERATED_DIR, tool_loop.BASE_URL, sys.argv)
    tool_loop.Runner, tool_loop.stream, tool_loop.GENERATED_DIR = FakeRunner, stream, out
    tool_loop.BASE_URL = "http://127.0.0.1:1"
    FakeRunner.save_fails = save_fails
    sys.argv = ["tool_loop.py", name, "request"]
    os.environ.pop("OPEN_PPTX", None)
    try:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            code = tool_loop.main()
    finally:
        tool_loop.Runner, tool_loop.stream, tool_loop.GENERATED_DIR, tool_loop.BASE_URL, sys.argv = saved
    path = os.path.join(out, name, "summary.json")
    summary = None
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            summary = json.load(f)
    shutil.rmtree(out, ignore_errors=True)
    return code, summary and summary["outcome"], FakeRunner.calls


def raises(ex):
    def fail():
        raise ex
    return fail


def main_ends():
    for name, failing, outcome in [
        ("server", raises(tool_loop.StreamFailure("LLM call failed: 503 busy")), "server_error: LLM call failed: 503"),
        ("refused", raises(tool_loop.RequestRefused("LLM call failed: 400 bad")), "error: LLM call failed: 400"),
        ("other", raises(ValueError("unexpected")), "error: ValueError: unexpected"),
    ]:
        code, got, calls = run_main(name, failing)
        check("main, %s: the deck and the summary are saved" % name,
              code == 1 and got is not None and got.startswith(outcome) and calls[-2:] == ["save", "close"],
              "exit %s, outcome %r, runner %s" % (code, got, calls))

    code, got, calls = run_main("nosave", raises(tool_loop.StreamFailure("LLM call failed: 503 busy")), save_fails=True)
    check("main, the save fails too: the summary keeps the first cause",
          code == 1 and got is not None and got.startswith("server_error") and calls[-1] == "close",
          "exit %s, outcome %r, runner %s" % (code, got, calls))


replies()
tls()
main_ends()
print()
print("ALL PASSED (%d)" % passed if failed == 0 else "%d FAILED, %d passed" % (failed, passed))
sys.exit(1 if failed else 0)
