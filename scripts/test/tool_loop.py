"""Tool-calling driver: the model gets two tools, `zimage` (generate an illustration) and
`pptcom` (run ComScript against the live deck), and drives itself - the same shape as the
opencode session whose reasoning stayed short. Execution goes through `ComScriptRunner serve`
(persistent PowerPoint + the add-in's own parser/engine/policy).

usage: python tool_loop.py <run-name> [user request]
env:   LLM_MODEL (default: chat_model in the add-in settings), MAX_TURNS (default 12), PASS_REASONING (1/0, default 1)

Endpoint and models come from the add-in's settings in the registry (HKCU\\Software\\KashiharaCity\\pptcom),
the API key from Windows Credential Manager (generic credential "KashiharaCity/pptcom:api_key"); nothing
environment-specific is built in. Set them once in the add-in's task pane.
"""
import ctypes
import ctypes.wintypes
import datetime
import http.client
import json
import os
import ssl
import subprocess
import sys
import traceback
import urllib.error
import urllib.request


REGISTRY_PATH = r"Software\KashiharaCity\pptcom"   # same as AddinSettings.RegistryPath
CREDENTIAL_TARGET = "KashiharaCity/pptcom:api_key"  # same as AddinSettings.CredentialTarget


class _CREDENTIAL(ctypes.Structure):
    _fields_ = [("Flags", ctypes.wintypes.DWORD), ("Type", ctypes.wintypes.DWORD),
                ("TargetName", ctypes.wintypes.LPWSTR), ("Comment", ctypes.wintypes.LPWSTR),
                ("LastWritten", ctypes.wintypes.FILETIME), ("CredentialBlobSize", ctypes.wintypes.DWORD),
                ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)), ("Persist", ctypes.wintypes.DWORD),
                ("AttributeCount", ctypes.wintypes.DWORD), ("Attributes", ctypes.c_void_p),
                ("TargetAlias", ctypes.wintypes.LPWSTR), ("UserName", ctypes.wintypes.LPWSTR)]


def read_credential(target: str) -> str:
    advapi = ctypes.WinDLL("advapi32", use_last_error=True)
    p = ctypes.POINTER(_CREDENTIAL)()
    if not advapi.CredReadW(target, 1, 0, ctypes.byref(p)):  # CRED_TYPE_GENERIC
        return ""
    try:
        c = p.contents
        return bytes(c.CredentialBlob[:c.CredentialBlobSize]).decode("utf-16-le")
    finally:
        advapi.CredFree(p)


def load_settings() -> dict:
    import winreg
    settings = {"endpoint": "", "chat_model": "", "image_model": ""}
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH) as key:
            for name in settings:
                try:
                    settings[name] = str(winreg.QueryValueEx(key, name)[0]).strip()
                except FileNotFoundError:
                    pass
    except FileNotFoundError:
        pass
    settings["api_key"] = read_credential(CREDENTIAL_TARGET)
    return settings


SETTINGS = load_settings()
BASE_URL = SETTINGS["endpoint"].rstrip("/")
API_KEY = SETTINGS["api_key"]
MODEL = os.environ.get("LLM_MODEL") or SETTINGS["chat_model"]
MAX_TURNS = int(os.environ.get("MAX_TURNS", "12"))
PASS_REASONING = os.environ.get("PASS_REASONING", "1") == "1"
# Same setting as the add-in's AttachGeneratedImages (default off).
ATTACH_GENERATED_IMAGES = os.environ.get("ATTACH_GENERATED_IMAGES", "1") == "1"
HERE = os.path.dirname(os.path.abspath(__file__))
# Every generated deck (pptx + slide PNGs + transcript + reasoning) lands in generated-pptx/<run>, next to
# the add-in's own run folders (its 途中経過を保存 setting): %LOCALAPPDATA%\KashiharaCity\pptcom\generated-pptx.
GENERATED_DIR = os.path.expandvars(r"%LOCALAPPDATA%\KashiharaCity\pptcom\generated-pptx")
RUNNER = os.path.join(os.path.dirname(os.path.dirname(HERE)), "temp", "build", "ComScriptRunner", "bin", "Debug", "ComScriptRunner.exe")
DEFAULT_REQUEST = "litellmが何か説明するパワポを３枚ぐらいで作って。画像も入れて。"

# Deliberately short and in the coding-agent register: the run8 experiment (old/experiment/out) showed that this
# framing (not the model, not any reasoning parameter) is what keeps the thinking model's
# reasoning at opencode's length. No design directions - the model designs better without.
SYSTEM = """You are a PowerPoint editing agent running inside PowerPoint. You build and edit the currently open presentation through the tools below. Use the instructions and tools to complete the user's request.

# Tone and style
You must be concise and direct. Minimize output tokens. Do not explain what you are about to do; just call the tools. When the deck is finished, reply with one short line (no summary of the slides). Write slide text in the user's language.

# Doing tasks
- Decide the slide structure and design once, then write the script. Rough coordinates are fine; do not compute exact text fit in your head.
- Every pptcom result comes with renderings of the slides that call changed (up to 20). Look at them and fix only what looks wrong (overflow, overlap, clipping, unreadable text) in your next pptcom call. When the latest renderings look right, the deck is done: reply immediately. Each rendering is followed by the slide's shape list ([index] = the variable that holds the shape, type, position, size, fill and line with their colors, text, text size/bold/color): refer to a shape by that variable, not by a guessed index.
- If you want illustrations, first decide only which slides get one and its aspect, and call `zimage` right away (one call per image, in parallel is fine). Plan the layout and coordinates after you have seen the images, not before; then use the returned variable name in `pptcom`.
- `pptread` shows slides you have not changed in this request, the same way (rendering + shape list; "3", "4-7", or all): use it first when the user asks to change an existing deck. The slides you changed are already rendered for you. Judge renderings visually (overflow, overlap, readability); do not measure pixels or recompute coordinates from them.
- `pptcom` runs every statement in order. If one fails (or pptcom cannot read it), it and the statements that use a variable it should have set are reported as FAILED / SKIPPED; everything else is applied. Resend only the reported lines, fixed.
- Do not probe the deck with extra pptcom reads, and do not pptread slides you just changed.

# pptcom
Write the script in VBA, as in the body of a PowerPoint VBA macro: one statement per line, starting from ActivePresentation (the presentation this request works on) or Application. Variables keep their values between pptcom calls. When this request started, the slide was %SLIDE_W% x %SLIDE_H% points (each shape list states its slide's size).
pptcom runs plain VBA statements only: no If/For/Do/While/With/Select or other control flow, no Sub/Function, no VBA functions except RGB and Chr, no comparisons or And/Or/Not, no named arguments. Variable names cannot be VBA keywords (Sub, End, Type, Date, Next, ...): call a subtitle subTitle, not sub. Only ActivePresentation may be edited (never Presentations.Add), and never its slide size or orientation (PageSetup). Application.Run / VBE / COMAddIns / AddOLEObject / FollowHyperlink / Quit are blocked.

Example of the style (not a design to copy): one statement per line, Set for objects, one variable per shape.
  Set sld = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
  Set title = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 40, 30, 880, 60)
  title.TextFrame.TextRange.Text = "Title"
  Set card = sld.Shapes.AddShape(msoShapeRoundedRectangle, 40, 120, 420, 300)
  card.Fill.ForeColor.RGB = RGB(240, 240, 240)
  card.Line.Visible = msoFalse
  Set cardText = card.TextFrame.TextRange
  cardText.Text = "First point" & vbCr & "Second point"
  cardText.ParagraphFormat.Alignment = ppAlignLeft
  cardText.ParagraphFormat.SpaceWithin = 1.2   ' line spacing, as a multiple of lines
  cardText.Font.Color.RGB = RGB(40, 40, 40)

A pptcom result looks like this:
  OK: 12 lines executed. Defined: sld, title, card. Values: h = 64.5
or, when something failed:
  10 of 12 lines applied; 1 failed, 1 skipped. ... FAILED line 5: <the line> -> <why> ... SKIPPED ...
To read a value, assign it to a variable (h = title.Height); it comes back under Values.

# Measured in this environment
This deck's defaults: new text %TEXT_SIZE% pt, color %TEXT_COLOR% ; text frame margins left/right %MARGIN_LR%, top/bottom %MARGIN_TB% ; paragraph SpaceBefore / SpaceAfter %SPACE_BEFORE% / %SPACE_AFTER%, SpaceWithin %SPACE_WITHIN% ; line weight %LINE_WEIGHT% ; rounded rectangle Adjustments(1) %ROUND_ADJ% ; text in AddShape shapes is %SHAPE_ANCHOR% and %SHAPE_ALIGN% (set ppAlignLeft for body text), color %SHAPE_TEXT_COLOR% - set Font.Color when that does not stand out on the fill.
Generated images: a 1024 x 1024 one placed by AddPicture with width and height -1 comes out %IMAGE_SIZE% (PowerPoint shrinks a picture larger than the slide to fit); give both to set the size.
Tables: a new table has %TABLE_STYLE% style and %TABLE_TEXT_SIZE% pt text, and a row never gets shorter than its text: set each cell's Font.Size first, then row heights. ApplyStyle "{5940675A-B579-460E-94D1-54222C63F5DA}" = plain black grid, no fills ; ApplyStyle "{2D5ABB26-0587-4C30-8999-92F81FD0307C}" = no grid, no fills.
Bullets: Bullet.Visible = msoTrue alone gives a numbered list (1. 2. 3.). Set Bullet.Type = ppBulletUnnumbered or ppBulletNumbered, then Bullet.UseTextFont = msoTrue (otherwise some themes draw the mark as an arrow). Body placeholders of an existing deck already have the theme's bullets.
Fonts: do not set font names - text then uses the deck's theme fonts, which are installed and consistent on every slide. Only if the user asks for a specific font: set both Font.Name (Latin letters/digits) and Font.NameFarEast (Japanese characters); Font.Name alone does not change Japanese text.
"""

TOOLS = [
    {"type": "function", "function": {"name": "zimage", "description": "Generate an image with the internal Z-Image Turbo model and bind it to a variable usable in pptcom by that name. Returns the variable name and pixel size, and the image itself is attached in the next message: check it for garbled text or a wrong subject BEFORE placing it, and if it is wrong call zimage again with the same name (the older version is then dropped). Calling again with the same name rebinds the name to the new image; pictures already placed are not changed. Prompt rules: (1) Write one concrete, objective English paragraph (60-150 words): subject and count, composition/viewpoint, style/medium (e.g. flat vector illustration, isometric 3D render, photo), color palette (name the deck's colors), lighting, background. (2) Negative prompts do not work; put exclusions in the prompt itself, e.g. 'no text, no letters, no logos, no watermark'. (3) Text renders reliably only in English or Chinese and only when the exact words are in double quotes with position and size (e.g. a sign reading \"LiteLLM\" at the top center). Never request Japanese text in an image; put it on the slide instead. (4) Screens, dashboards, signs and documents get garbled pseudo-text unless you say they show only abstract shapes and bars. (5) No metaphors, no emotional words, no meta tags such as 8K or masterpiece. (6) For slides: use a plain background in the deck's background color and leave empty space where slide text will go. (7) The generated background never matches the slide background exactly. Show each image as a deliberate panel (a framed card, or a large area of the slide); do not try to hide its edge by cropping, recoloring or regenerating. Regenerate an image at most once, and only for a wrong subject or garbled text. (8) Do not use zimage for small icons or badges (anything under about 200 pt). Draw those with shapes instead: an oval or rounded square in an accent color with a number or one short word in it. Use zimage for 1-3 large visuals per deck (hero image, illustration, scene). (9) zimage is for illustrations and photos (objects, scenes, atmosphere). Do not use it for diagrams, flowcharts, comparisons, timelines, charts or other infographics: they come out as decorative pseudo-diagrams that explain nothing. Draw those with shapes, arrows and tables in pptcom.", "parameters": {"type": "object", "properties": {"name": {"type": "string", "description": "Variable name for the image, e.g. img1"}, "prompt": {"type": "string", "description": "English prompt following the rules above."}, "size": {"type": "string", "enum": ["1024x1024", "1024x768", "768x1024", "1024x512", "512x1024"], "description": "Pixel size. 1024x1024 square (icons, side images); 1024x768 or 1024x512 wide (hero, banner, background strip); 768x1024 or 512x1024 tall. Default 1024x1024."}}, "required": ["name", "prompt"]}}},
    {
        "type": "function",
        "function": {
            "name": "pptread",
            "description": "Look at slides of the open presentation without changing them: each slide is rendered at its own size (a 960x540 slide as 1280x720 px; at most 1920 px on the long side) and shown to you in the next message, followed by its shape list - the same as after a pptcom call (per shape: [index] = the script variable that holds it, type, Left/Top/Width/Height in points, fill (RGB, gradient colors and GradientAngle, Transparency), line or outline (RGB, weight in points, DashStyle when not solid), the text as a VBA string, and the text's size, bold and color per paragraph). At most 20 slides per call.",
            "parameters": {
                "type": "object",
                "properties": {"slide": {"type": "string", "description": "One slide (\"3\"), a range (\"4-7\") or a list of them (\"1,3-5\"), 1-based. Omit for all slides."}},
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "pptcom",
            "description": "Execute a pptcom script (see grammar in the system prompt) against the open presentation. Statements run in order; a failing statement does not stop the others - only statements that depend on what it should have set are skipped, and the result lists the failed and skipped lines. Variables persist between calls. The slides the script changed (up to 20) are rendered and attached as images in the next message.",
            "parameters": {
                "type": "object",
                "properties": {"script": {"type": "string", "description": "The script, one statement per line."}},
                "required": ["script"],
            },
        },
    },
]



ZIMAGE_ATTACHED = "Returns the variable name and pixel size, and the image itself is attached in the next message: check it for garbled text or a wrong subject BEFORE placing it, and if it is wrong call zimage again with the same name (the older version is then dropped). "
ZIMAGE_NOT_ATTACHED = "Returns the variable name and pixel size; the image itself is not shown to you here - you see it in the slide renderings once it is placed. If it looks wrong there, call zimage again with the same name and replace the picture. "
if not ATTACH_GENERATED_IMAGES:
    for _t in TOOLS:
        if _t["function"]["name"] == "zimage":
            assert ZIMAGE_ATTACHED in _t["function"]["description"]
            _t["function"]["description"] = _t["function"]["description"].replace(ZIMAGE_ATTACHED, ZIMAGE_NOT_ATTACHED)

class Runner:
    def __init__(self):
        self.p = subprocess.Popen([RUNNER, "serve"], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                  encoding="utf-8", bufsize=1)
        ready = self.p.stdout.readline()
        if '"ready"' not in ready:
            raise SystemExit("runner did not start: " + ready)

    def call(self, req: dict) -> dict:
        self.p.stdin.write(json.dumps(req, ensure_ascii=False) + "\n")
        self.p.stdin.flush()
        line = self.p.stdout.readline()
        if not line:
            raise RuntimeError("runner died")
        return json.loads(line)

    def close(self):
        try:
            self.p.stdin.write(json.dumps({"quit": True}) + "\n")
            self.p.stdin.flush()
        except Exception:
            pass
        try:
            self.p.wait(timeout=30)
        except Exception:
            self.p.kill()


SUPERSEDED_NOTE = "[omitted: an older version of {0}; a newer one appears later in the conversation]"


def strip_private(messages):
    """Drop our bookkeeping keys (e.g. _slide on image parts) before sending."""
    out = []
    for m in messages:
        if isinstance(m.get("content"), list):
            m = {**m, "content": [{k: v for k, v in p.items() if not k.startswith("_")} for p in m["content"]]}
        out.append(m)
    return out


MAX_AUTO_RENDERS = 20

# Same as the add-in (ThisAddIn.ElideOldScripts): once a prompt passes this many tokens, the pptcom
# scripts older than the last KEEP_RECENT_SCRIPTS assistant turns that ran one are replaced, once, by a
# one-line VBA comment, so the conversation shrinks back and the prefix cache keeps working after.
SCRIPT_ELISION_PROMPT_TOKENS = 100000
KEEP_RECENT_SCRIPTS = 2
ELIDED_SCRIPT_MARK = "' ("


def elide_old_scripts(messages, keep):
    with_scripts = [m for m in messages if any(c["function"]["name"] == "pptcom" for c in (m.get("tool_calls") or []))]
    replaced = 0
    for m in with_scripts[:max(0, len(with_scripts) - keep)]:
        for c in m["tool_calls"]:
            if c["function"]["name"] != "pptcom":
                continue
            try:
                script = json.loads(c["function"]["arguments"]).get("script") or ""
            except Exception:
                script = c["function"]["arguments"] or ""
            if script.startswith(ELIDED_SCRIPT_MARK):
                continue
            c["function"]["arguments"] = json.dumps({"script": ELIDED_SCRIPT_MARK + "%d lines, already run - left out of this conversation to keep it short)" % len(script.split("\n"))})
            replaced += 1
    return replaced


# Statements that put an object into a variable without Set (accepted by the engine, only counted;
# same as the add-in's summary.json missing_set).
MISSING_SET = []


# Renderings are (label, png path, shape list or None); the shape list (the runner's "brief", the same
# DeckReader.Brief text the add-in sends) goes right after the image.
def render_slide(runner, n):
    r = runner.call({"export": n})
    if not r.get("png"):
        return None
    return ("slide %d" % n, r["png"], runner.call({"brief": n}).get("text"))


def auto_render(runner, touched):
    out = []
    for n in touched[:MAX_AUTO_RENDERS]:
        rendering = render_slide(runner, n)
        if rendering:
            out.append(rendering)
    return out


def render_note(touched, renders):
    done = [int(label.split()[1]) for label, _, _ in renders]
    note = ""
    if done:
        note += "\nRenderings of changed slide(s) %s are attached in the next message, each with its shape list." % ", ".join(map(str, done))
    rest = [n for n in touched if n not in done]
    if rest:
        note += "\nAlso changed but not rendered: slide(s) %s (use pptread if needed)." % ", ".join(map(str, rest))
    return note


def run_tool(runner: Runner, name: str, args: dict):
    """Returns (result_text, png_path_or_None)."""
    if name == "pptread":
        slide = args.get("slide")
        r = runner.call({"read": "" if slide is None else str(slide)})
        if not r.get("text"):
            return ("read failed: " + str(r.get("error")), None)
        renders = [x for x in (render_slide(runner, n) for n in r.get("slides") or []) if x]
        return (r["text"], renders or None)
    if name == "zimage":
        var = args.get("name") or "img"
        prompt = args.get("prompt") or ""
        size = args.get("size") or "1024x1024"
        # The runner builds the `name = Image(...)` line itself, as the add-in does; Image() is not
        # accepted in a pptcom script.
        r = runner.call({"zimage": {"name": var, "prompt": prompt, "size": size}})
        if r.get("success"):
            w, h = (int(x) for x in size.split("x"))
            import math
            g = math.gcd(w, h)
            text = ("Generated image bound to %s (%s px, aspect %d:%d). Use it in pptcom: shapes.AddPicture(%s, msoFalse, msoTrue, left, top, width, height) with width/height in the same aspect ratio."
                    % (var, size, w // g, h // g, var))
            t = runner.call({"image": var}) if ATTACH_GENERATED_IMAGES else {}
            if t.get("png"):
                return (text + "\nThe image is attached in the next message.", [("image " + var, t["png"], None)])
            return (text, None)
        return ("Image generation failed: " + str(r.get("error") or r.get("report")), None)
    if name == "pptcom":
        script = args.get("script") or ""
        r = runner.call({"script": script})
        MISSING_SET.extend(r.get("missingSet") or [])
        renders = auto_render(runner, r.get("touched") or [])
        note = render_note(r.get("touched") or [], renders)
        if r.get("success"):
            n = r.get("lines", 0)
            ids = r.get("definedIds") or []
            vals = r.get("values") or ""
            changed = r.get("changed") or []
            return ("OK: %d lines executed. Defined: %s.%s%s" % (n, ", ".join(ids) if ids else "(none)",
                    (" Changed: " + ", ".join(changed) + ".") if changed else "", (" Values: " + vals) if vals else "") + note, renders)
        if r.get("stage") == "parse":
            wd = r.get("wouldDefine") or []
            # only a script rejected as a whole (empty, too many lines); a statement that does not
            # parse comes back in the per-line report below, as in the add-in
            return ("Script rejected - NOTHING in this script was executed.\n%s%s"
                    "\nFix it and resend the WHOLE script."
                    % (r.get("error"),
                       ("\nThese variables were NOT created and do not exist: " + ", ".join(wd) + ".") if wd else ""), None)
        # The runner builds the same report as the add-in (ComScriptResult.Describe): failed lines
        # with their errors, lines skipped because they depend on them; everything else applied.
        vals = r.get("values") or ""
        changed = r.get("changed") or []
        return ((r.get("report") or r.get("error") or "(no report)") + (("\nChanged: " + ", ".join(changed) + ".") if changed else "") +
                (("\nValues set: " + vals) if vals else "") + note, renders)
    return ("unknown tool: " + name, None)


def slide_count(runner: Runner) -> str:
    r = runner.call({"script": "cnt = ActivePresentation.Slides.Count"})
    return "?" if not r.get("success") else "(see deck)"


# Same as the add-in (OpenAiCompatibleClient.StreamIdleLimit): a reply's stream that sends nothing for this long
# is a failure on the server side. urlopen's timeout bounds every socket wait, so it is the long one up to the
# headers and this one after them.
STREAM_IDLE_SEC = 60


class StreamFailure(Exception):
    """A failure on the server side (the add-in's LlmServerException): HTTP 5xx, a connection that failed or
    dropped, an error event in the stream, a stream that stopped sending or was cut off. The run ends there,
    with the deck and the summary saved."""


class RequestRefused(Exception):
    """HTTP 4xx: the server refused the request (the add-in ends the request, with no 再試行). The run ends
    there, with the deck and the summary saved."""


def tls_context() -> ssl.SSLContext:
    """The server's certificate and host name are checked, against the certificates Windows trusts (as the
    add-in's HttpClient does). VERIFY_X509_STRICT, on by default since Python 3.13, is left off: Windows does not
    apply it either, and it refuses an organization's own CA whose Basic Constraints are not marked critical."""
    ctx = ssl.create_default_context()
    ctx.verify_flags &= ~ssl.VERIFY_X509_STRICT
    return ctx


def wait_idle(resp) -> None:
    """After the headers, a socket wait longer than STREAM_IDLE_SEC is a stall (urlopen's timeout is the long
    one up to the headers)."""
    sock = getattr(getattr(getattr(resp, "fp", None), "raw", None), "_sock", None)
    if sock is not None:
        sock.settimeout(STREAM_IDLE_SEC)


def error_body(err: urllib.error.HTTPError) -> str:
    """An error's body, shown with its status, under the same limit as the stream: a gateway can send the
    headers of a 503 and never finish the body (the add-in's ReadErrorBodyAsync)."""
    wait_idle(err.fp)
    try:
        return err.read(4000).decode("utf-8", errors="replace").strip()
    except (OSError, http.client.HTTPException):
        return "(the body did not arrive within %d s)" % STREAM_IDLE_SEC


def stream(payload: dict) -> dict:
    req = urllib.request.Request(
        BASE_URL + "/v1/chat/completions",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=dict({"Content-Type": "application/json"}, **({"Authorization": "Bearer " + API_KEY} if API_KEY else {})),
    )
    started = datetime.datetime.now()
    content, reasoning, tool_args = [], [], {}
    usage, finish = None, None
    first_token = first_output = None  # seconds; see ChatReply.FirstToken / FirstOutput
    done = False
    try:
        resp = urllib.request.urlopen(req, context=tls_context(), timeout=1200)
    except urllib.error.HTTPError as ex:
        detail = "LLM call failed: %d %s" % (ex.code, error_body(ex))
        ex.close()
        if ex.code >= 500:
            raise StreamFailure(detail) from ex
        raise RequestRefused(detail) from ex
    except (OSError, http.client.HTTPException) as ex:  # URLError is an OSError
        raise StreamFailure("LLM connection failed: %s" % ex) from ex
    with resp:
        wait_idle(resp)
        lines = iter(resp)
        while True:
            try:
                raw = next(lines, None)
            except TimeoutError as ex:
                raise StreamFailure("LLM stream stalled: nothing arrived for %d s" % STREAM_IDLE_SEC) from ex
            except (OSError, http.client.HTTPException) as ex:
                raise StreamFailure("LLM connection dropped: %s" % ex) from ex
            if raw is None:
                break
            line = raw.decode("utf-8", errors="replace").strip()
            if not line.startswith("data:"):
                continue
            data = line[5:].strip()
            if data == "[DONE]":
                done = True
                break
            try:
                chunk = json.loads(data)
            except json.JSONDecodeError:
                continue  # as the add-in: a line that is not JSON is skipped
            if chunk.get("error"):
                # LiteLLM reports a backend failure after the 200 header as an error event
                raise StreamFailure("LLM stream error: " + json.dumps(chunk["error"], ensure_ascii=False))
            if chunk.get("usage"):
                usage = chunk["usage"]
            for choice in chunk.get("choices") or []:
                d = choice.get("delta") or {}
                has_out = bool(d.get("content")) or bool(d.get("tool_calls"))
                if (has_out or d.get("reasoning_content")) and first_token is None:
                    first_token = (datetime.datetime.now() - started).total_seconds()
                if has_out and first_output is None:
                    first_output = (datetime.datetime.now() - started).total_seconds()
                if d.get("content"):
                    content.append(d["content"])
                if d.get("reasoning_content"):
                    reasoning.append(d["reasoning_content"])
                for tc in d.get("tool_calls") or []:
                    idx = tc.get("index", 0)
                    fn = tc.get("function") or {}
                    slot = tool_args.setdefault(idx, {"id": None, "name": "", "arguments": ""})
                    if tc.get("id"):
                        slot["id"] = tc["id"]
                    if fn.get("name"):
                        slot["name"] += fn["name"]
                    if fn.get("arguments"):
                        slot["arguments"] += fn["arguments"]
                if choice.get("finish_reason"):
                    finish = choice["finish_reason"]
    if not done and finish is None:
        raise StreamFailure("LLM stream ended before the reply was complete (no finish_reason, no [DONE])")
    calls = []
    for k in sorted(tool_args):
        t = tool_args[k]
        calls.append({"id": t["id"] or ("call_%d" % k), "type": "function",
                      "function": {"name": t["name"], "arguments": t["arguments"]}})
    return {
        "elapsed_sec": round((datetime.datetime.now() - started).total_seconds(), 1),
        "first_token_sec": None if first_token is None else round(first_token, 1),
        "first_output_sec": None if first_output is None else round(first_output, 1),
        "usage": usage, "finish_reason": finish,
        "content": "".join(content), "reasoning": "".join(reasoning), "tool_calls": calls,
    }


def main() -> int:
    if not BASE_URL:
        sys.exit("endpoint is not set in HKCU\\" + REGISTRY_PATH + " (set it once in the add-in's task pane)")
    run_name = sys.argv[1]
    request = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_REQUEST
    out_dir = os.path.join(GENERATED_DIR, run_name)
    os.makedirs(out_dir, exist_ok=True)

    # Same as the add-in: local date/time on top of the user message (not in SYSTEM, which stays
    # identical across requests for the server's prefix cache).
    now = datetime.datetime.now().astimezone()
    off = now.strftime("%z")
    time_line = "Current date and time: %s (UTC%s:%s)" % (now.strftime("%Y-%m-%d (%a) %H:%M"), off[:3], off[3:])
    messages = [{"role": "system", "content": SYSTEM}, {"role": "user", "content": time_line + "\n\n" + request}]
    runner = Runner()
    if os.environ.get("OPEN_PPTX"):
        print("opened:", json.dumps(runner.call({"open": os.path.abspath(os.environ["OPEN_PPTX"])}), ensure_ascii=False), flush=True)
    # Deck-dependent cheat-sheet numbers (%SLIDE_W% ...), measured like the add-in does (DeckProbe).
    probe = runner.call({"probe": True})
    print("deck defaults:", probe.get("summary"), flush=True)
    system = SYSTEM
    for key, value in (probe.get("placeholders") or {}).items():
        system = system.replace(key, value)
    messages[0]["content"] = system
    turns = []
    t0 = datetime.datetime.now()
    try:
        outcome = run_turns(runner, messages, turns, out_dir)
    except Exception as ex:  # the deck and the summary are still saved, and say what ended the run
        outcome = "error: %s: %s" % (type(ex).__name__, ex)
        print("   !!", outcome, flush=True)
        traceback.print_exc()
    total = (datetime.datetime.now() - t0).total_seconds()
    try:
        saved = runner.call({"save": os.path.join(out_dir, "deck.pptx")})
        print("saved:", json.dumps(saved, ensure_ascii=False), flush=True)
    except Exception as ex:
        print("   !! deck not saved: %s: %s" % (type(ex).__name__, ex), flush=True)
    finally:
        runner.close()

    summary = {"model": MODEL, "request": request, "outcome": outcome, "total_sec": round(total, 1),
               "turns": len(turns), "reasoning_chars_total": sum(t["reasoning_chars"] for t in turns),
               "llm_sec_total": round(sum(t["elapsed_sec"] for t in turns), 1),
               "missing_set": len(MISSING_SET), "missing_set_examples": MISSING_SET[:5], "turn_details": turns}
    with open(os.path.join(out_dir, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=1)
    with open(os.path.join(out_dir, "transcript.json"), "w", encoding="utf-8") as f:
        json.dump(messages, f, ensure_ascii=False, indent=1)
    print(json.dumps({k: v for k, v in summary.items() if k != "turn_details"}, ensure_ascii=False))
    # A run the model ended (completed, max_turns, max_tokens) is a result; one a failure ended is not.
    return 1 if outcome.startswith(("server_error", "error")) else 0


def run_turns(runner: "Runner", messages: list, turns: list, out_dir: str) -> str:
    """The turns of the run, until the model answers without a tool call, the turn limit, or a failure that
    ends the run. Returns the outcome."""
    outcome = "max_turns"  # unless the model stops calling tools before the limit (same as the add-in)
    for turn in range(1, MAX_TURNS + 1):
        payload = {"model": MODEL, "messages": strip_private(messages), "tools": TOOLS, "tool_choice": "auto",
                   "max_tokens": 32000, "stream": True, "stream_options": {"include_usage": True},
                   # same as the add-in: keep images at native resolution (server default upscales ~4x)
                   "mm_processor_kwargs": {"min_pixels": 3136}}
        try:
            r = stream(payload)
        except StreamFailure as ex:
            # The add-in waits for 再試行 here; the rig has nobody to press it, so the run ends.
            outcome = "server_error: " + str(ex)
            print("   !!", outcome, flush=True)
            break
        except RequestRefused as ex:
            outcome = "error: " + str(ex)
            print("   !!", outcome, flush=True)
            break
        rt = (r["usage"] or {}).get("completion_tokens_details", {}).get("reasoning_tokens")
        info = {"turn": turn, "elapsed_sec": r["elapsed_sec"],
                "first_token_sec": r["first_token_sec"], "first_output_sec": r["first_output_sec"],
                "reasoning_chars": len(r["reasoning"]),
                "prompt_tokens": (r["usage"] or {}).get("prompt_tokens"),
                "reasoning_tokens": rt, "finish": r["finish_reason"],
                "tools": [(c["function"]["name"], len(c["function"]["arguments"])) for c in r["tool_calls"]],
                "content": r["content"][:200]}
        print(json.dumps(info, ensure_ascii=False), flush=True)
        with open(os.path.join(out_dir, "turn%02d.reasoning.txt" % turn), "w", encoding="utf-8") as f:
            f.write(r["reasoning"])

        assistant = {"role": "assistant", "content": r["content"] or None}
        if r["tool_calls"]:
            assistant["tool_calls"] = r["tool_calls"]
        if PASS_REASONING and r["reasoning"]:
            # both names, as the add-in: current vLLM reads "reasoning" and ignores "reasoning_content"
            assistant["reasoning"] = r["reasoning"]
            assistant["reasoning_content"] = r["reasoning"]
        messages.append(assistant)
        if (info["prompt_tokens"] or 0) > SCRIPT_ELISION_PROMPT_TOKENS:
            n = elide_old_scripts(messages, KEEP_RECENT_SCRIPTS)
            if n:
                print("   (prompt %d tok: %d older pptcom scripts replaced by a comment)" % (info["prompt_tokens"], n), flush=True)

        results = []
        images = []
        for c in r["tool_calls"]:
            try:
                args = json.loads(c["function"]["arguments"])
            except Exception as ex:
                result, png = "Invalid JSON in tool arguments: " + str(ex), None
            else:
                result, png = run_tool(runner, c["function"]["name"], args)
            if png:
                images.extend(png)
            results.append({"name": c["function"]["name"], "args": c["function"]["arguments"], "result": result})
            messages.append({"role": "tool", "tool_call_id": c["id"], "content": result})
            print("   ->", c["function"]["name"], ":", result.splitlines()[0][:160], flush=True)
        if images:
            # Chat Completions defines a tool message's content as text only. Some servers (vLLM with a
            # vision model's chat template) take image parts there too, but another server or template
            # may drop them without an error, so the rendering rides on a user message right after them,
            # as an OpenAI-style image_url part. Only the newest rendering per slide is kept (same rule
            # as the add-in): older image parts of a re-rendered slide are replaced by a one-line note.
            import base64
            newest = {label: (png, inventory) for label, png, inventory in images}
            for m in messages:
                if m.get("role") != "user" or not isinstance(m.get("content"), list):
                    continue
                for i, part in enumerate(m["content"]):
                    if part.get("type") == "image_url" and part.get("_label") in newest:
                        m["content"][i] = {"type": "text", "text": SUPERSEDED_NOTE.format(part["_label"])}
                # the shape list goes with its image
                m["content"] = [part for part in m["content"] if part.get("_inventory") not in newest]
            parts = [{"type": "text", "text": "Image(s) for the tool results above:"}]
            for label, (png, inventory) in newest.items():
                b64 = base64.b64encode(open(png, "rb").read()).decode("ascii")
                parts.append({"type": "text", "text": label[0].upper() + label[1:] + ":"})
                parts.append({"type": "image_url", "image_url": {"url": "data:image/png;base64," + b64}, "_label": label})
                if inventory:
                    parts.append({"type": "text", "text": inventory, "_inventory": label})
            messages.append({"role": "user", "content": parts})
        info["results"] = results
        turns.append(info)
        with open(os.path.join(out_dir, "transcript.json"), "w", encoding="utf-8") as f:
            # images become a count; the text parts (labels, shape lists) are kept
            slim = [m if not isinstance(m.get("content"), list) else
                    {**m, "content": "\n".join(p["text"] for p in m["content"] if p.get("type") == "text") +
                                     "\n[%d image(s)]" % sum(1 for p in m["content"] if p.get("type") == "image_url")}
                    for m in messages]
            json.dump(slim, f, ensure_ascii=False, indent=1)
        if not r["tool_calls"]:
            # max_tokens: the reply reached the output limit before it was done (same as the add-in)
            outcome = "max_tokens" if r["finish_reason"] == "length" else "completed"
            break
    return outcome

if __name__ == "__main__":
    sys.exit(main())
