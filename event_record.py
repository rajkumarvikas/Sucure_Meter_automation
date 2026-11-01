#!/usr/bin/env python3
"""
task_recorder.py - Record a web task once, so it can be turned into an automation.

What it captures while you work in the Chrome window it opens:
  * Mouse: clicks, double-clicks, right-clicks (with position, element, selector)
  * Keyboard: typed text per field (passwords are hidden), Enter/Tab/Esc/shortcuts
  * Dropdowns, checkboxes, radio buttons, file uploads, scrolling
  * Page navigations and new tabs
  * Screenshots of the Chrome page at every click / navigation (click is highlighted)
  * Your voice (microphone) -> transcribed locally with Whisper, merged into the timeline

Stop recording: close the Chrome window, or press Ctrl+C in the terminal.

Output (folder recordings/<date_time>/ and a .zip):
  timeline.md      human + Claude readable step-by-step timeline  <- send this
  events.json      raw events
  transcript.json  voice segments with timestamps
  screenshots/     PNGs referenced from the timeline
  audio.wav        your voice (kept locally, NOT put in the zip)

Install:
  pip install playwright sounddevice numpy faster-whisper
  playwright install chromium        (only needed if Google Chrome is not installed)

Run:
  python task_recorder.py https://your-site.com
  python task_recorder.py https://your-site.com --translate      # Hindi/Hinglish speech -> English text
  python task_recorder.py https://your-site.com --profile chrome_profile   # keep logins between runs
  python task_recorder.py https://your-site.com --no-voice

Limits: only the web page inside this Playwright-opened Chrome is captured (not the
address bar, OS file dialogs, or other apps).
"""
import argparse
import json
import queue
import sys
import time
import wave
import zipfile
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

try:
    import numpy as np
    import sounddevice as sd
except Exception:  # voice is optional
    np = sd = None


# --------------------------------------------------------------------------
# JavaScript injected into every page / iframe
# --------------------------------------------------------------------------
INIT_JS = r"""
(() => {
  if (window.__recInstalled) return;
  window.__recInstalled = true;

  const send = (e) => {
    try {
      window.__rec(Object.assign({
        url: location.href,
        frame: window === window.top ? 'top' : 'iframe',
        ts: Date.now()
      }, e));
    } catch (_) {}
  };

  const cssEsc = (s) => (window.CSS && CSS.escape) ? CSS.escape(s)
                        : String(s).replace(/[^a-zA-Z0-9_-]/g, '\\$&');

  function selectorFor(el) {
    if (!(el instanceof Element)) return null;
    if (el.id) return '#' + cssEsc(el.id);
    for (const a of ['data-testid', 'data-test', 'data-qa', 'name', 'aria-label', 'placeholder', 'title']) {
      const v = el.getAttribute(a);
      if (v) {
        const s = el.tagName.toLowerCase() + '[' + a + '="' + v.replace(/"/g, '\\"') + '"]';
        try { if (document.querySelectorAll(s).length === 1) return s; } catch (_) {}
      }
    }
    const parts = [];
    let cur = el;
    while (cur && cur.nodeType === 1 && cur !== document.documentElement) {
      let part = cur.tagName.toLowerCase();
      if (cur.id) { parts.unshift('#' + cssEsc(cur.id)); break; }
      const parent = cur.parentElement;
      if (parent) {
        const sibs = [...parent.children].filter(c => c.tagName === cur.tagName);
        if (sibs.length > 1) part += ':nth-of-type(' + (sibs.indexOf(cur) + 1) + ')';
      }
      parts.unshift(part);
      cur = parent;
      if (parts.length > 6) break;
    }
    return parts.join(' > ');
  }

  function describe(el) {
    if (!el || !el.tagName) return {};
    const r = el.getBoundingClientRect();
    const tag = el.tagName.toLowerCase();
    const isField = ['input', 'textarea', 'select'].includes(tag);
    const label = (el.labels && el.labels[0] && el.labels[0].innerText) || '';
    return {
      tag: tag,
      type: el.type || null,
      id: el.id || null,
      name: el.getAttribute('name'),
      role: el.getAttribute('role'),
      text: isField ? '' : (el.innerText || '').trim().slice(0, 80),
      placeholder: el.getAttribute('placeholder'),
      label: label.trim().slice(0, 60),
      selector: selectorFor(el),
      box: { x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height) }
    };
  }

  const pick = (el) => (el.closest && el.closest('a,button,input,select,textarea,label,summary,[role=button],[onclick]')) || el;

  // ---- typing (debounced per field; passwords hidden) ----
  let pend = null, timer = null;
  function flush() {
    if (!pend) return;
    const el = pend; pend = null; clearTimeout(timer);
    const t = (el.type || '').toLowerCase();
    const val = t === 'password' ? '<hidden>' : (el.isContentEditable ? el.innerText : el.value);
    send({ type: 'fill', value: String(val == null ? '' : val).slice(0, 500), secret: t === 'password', target: describe(el) });
  }
  document.addEventListener('input', (e) => {
    const el = e.target;
    const t = (el.type || '').toLowerCase();
    if (el.tagName === 'SELECT' || ['checkbox', 'radio', 'file', 'range', 'color'].includes(t)) return;
    if (pend && pend !== el) flush();
    pend = el; clearTimeout(timer); timer = setTimeout(flush, 800);
  }, true);
  document.addEventListener('focusout', flush, true);
  window.addEventListener('pagehide', flush);

  // ---- mouse ----
  for (const type of ['click', 'dblclick', 'contextmenu']) {
    document.addEventListener(type, (e) => {
      flush();
      send({ type: type, x: e.clientX, y: e.clientY, button: e.button, target: describe(pick(e.target)) });
    }, true);
  }

  // ---- dropdown / checkbox / file ----
  document.addEventListener('change', (e) => {
    const el = e.target;
    const t = (el.type || '').toLowerCase();
    if (el.tagName === 'SELECT') {
      send({ type: 'select', value: [...el.selectedOptions].map(o => o.text.trim()).join(', '), target: describe(el) });
    } else if (t === 'checkbox' || t === 'radio') {
      send({ type: 'check', checked: el.checked, target: describe(el) });
    } else if (t === 'file') {
      send({ type: 'upload', files: [...el.files].map(f => f.name), target: describe(el) });
    }
  }, true);

  // ---- special keys & shortcuts (normal characters are captured via 'fill') ----
  const SPECIAL = new Set(['Enter', 'Tab', 'Escape', 'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight',
                           'PageUp', 'PageDown', 'Home', 'End']);
  document.addEventListener('keydown', (e) => {
    const mods = [];
    if (e.ctrlKey) mods.push('Ctrl');
    if (e.altKey) mods.push('Alt');
    if (e.metaKey) mods.push('Meta');
    if (e.key === 'Enter' || e.key === 'Tab') flush();
    const isMod = ['Control', 'Alt', 'Meta', 'Shift'].includes(e.key);
    if ((SPECIAL.has(e.key) || (mods.length && !isMod))) {
      if (e.shiftKey) mods.push('Shift');
      send({ type: 'key', key: e.key, mods: mods, target: describe(document.activeElement) });
    }
  }, true);

  // ---- scroll (debounced) ----
  let st = null;
  window.addEventListener('scroll', () => {
    clearTimeout(st);
    st = setTimeout(() => send({ type: 'scroll', x: Math.round(window.scrollX), y: Math.round(window.scrollY) }), 500);
  });
})();
"""

HILITE_JS = """
(ev) => {
  const wrap = document.createElement('div');
  wrap.id = '__rec_hl';
  wrap.style.cssText = 'position:fixed;inset:0;pointer-events:none;z-index:2147483647';
  const b = (ev.target && ev.target.box) || null;
  if (b) {
    const r = document.createElement('div');
    r.style.cssText = `position:fixed;left:${b.x}px;top:${b.y}px;width:${b.w}px;height:${b.h}px;border:3px solid red;box-sizing:border-box`;
    wrap.appendChild(r);
  }
  if (ev.x != null) {
    const d = document.createElement('div');
    d.style.cssText = `position:fixed;left:${ev.x - 9}px;top:${ev.y - 9}px;width:18px;height:18px;border-radius:50%;background:rgba(255,0,0,.5);border:2px solid red`;
    wrap.appendChild(d);
  }
  document.documentElement.appendChild(wrap);
}
"""
CLEAR_JS = "() => { const n = document.getElementById('__rec_hl'); if (n) n.remove(); }"


# --------------------------------------------------------------------------
# Voice
# --------------------------------------------------------------------------
class VoiceRecorder:
    RATE = 16000

    def __init__(self):
        self.chunks = []
        self.stream = None
        self.t_start = None

    def start(self):
        def cb(indata, frames, t, status):
            self.chunks.append(indata.copy())
        self.stream = sd.InputStream(samplerate=self.RATE, channels=1, dtype="float32", callback=cb)
        self.stream.start()
        self.t_start = time.time()

    def stop(self):
        try:
            if self.stream:
                self.stream.stop()
                self.stream.close()
        except Exception:
            pass
        if not self.chunks:
            return None
        return np.concatenate(self.chunks).flatten()


def save_wav(path, audio, rate):
    pcm = (np.clip(audio, -1, 1) * 32767).astype(np.int16)
    with wave.open(str(path), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(rate)
        w.writeframes(pcm.tobytes())


def transcribe(audio, model_name, lang, translate):
    from faster_whisper import WhisperModel
    model = WhisperModel(model_name, device="cpu", compute_type="int8")
    segs, info = model.transcribe(
        audio, language=lang, task="translate" if translate else "transcribe", vad_filter=True
    )
    out = [{"start": s.start, "end": s.end, "text": s.text.strip()} for s in segs if s.text.strip()]
    return out, info.language


# --------------------------------------------------------------------------
# Timeline formatting
# --------------------------------------------------------------------------
def clock(t):
    m = int(t // 60)
    return f"[{m:02d}:{t - m * 60:04.1f}]"


def target_desc(tg):
    if not tg:
        return ""
    name = tg.get("text") or tg.get("label") or tg.get("placeholder") or tg.get("name") or tg.get("id") or ""
    s = f"<{tg.get('tag', '?')}>"
    if name:
        s += f' "{name}"'
    return s


def fmt(ev):
    k, tg = ev["type"], ev.get("target") or {}
    sel = f"  | selector: `{tg['selector']}`" if tg.get("selector") else ""
    shot = f"  | screenshot: {ev['screenshot']}" if ev.get("screenshot") else ""
    if k == "voice":
        return f"{clock(ev['t'])} VOICE: \"{ev['text']}\""
    if k == "navigate":
        return f"{clock(ev['t'])} NAVIGATE -> {ev['url']}{shot}"
    if k in ("click", "dblclick", "contextmenu"):
        label = {"click": "CLICK", "dblclick": "DOUBLE-CLICK", "contextmenu": "RIGHT-CLICK"}[k]
        return f"{clock(ev['t'])} {label} {target_desc(tg)} at ({ev['x']},{ev['y']}){sel}{shot}"
    if k == "fill":
        val = "<hidden password>" if ev.get("secret") else f"\"{ev['value']}\""
        return f"{clock(ev['t'])} TYPE {val} into {target_desc(tg)}{sel}"
    if k == "select":
        return f"{clock(ev['t'])} SELECT \"{ev['value']}\" in {target_desc(tg)}{sel}"
    if k == "check":
        return f"{clock(ev['t'])} {'CHECK' if ev['checked'] else 'UNCHECK'} {target_desc(tg)}{sel}"
    if k == "upload":
        return f"{clock(ev['t'])} UPLOAD file(s) {ev['files']} via {target_desc(tg)}{sel}"
    if k == "key":
        combo = "+".join(ev.get("mods", []) + [ev["key"]])
        return f"{clock(ev['t'])} KEY {combo} (focus: {target_desc(tg)})"
    if k == "scroll":
        return f"{clock(ev['t'])} SCROLL to y={ev['y']}"
    if k == "tab_opened":
        return f"{clock(ev['t'])} NEW TAB opened"
    if k == "tab_closed":
        return f"{clock(ev['t'])} TAB closed"
    return f"{clock(ev['t'])} {k.upper()}"


# --------------------------------------------------------------------------
# Browser
# --------------------------------------------------------------------------
def open_context(p, profile):
    common = dict(headless=False, args=["--start-maximized"])
    last = None
    for channel in ("chrome", None):
        try:
            kw = dict(common)
            if channel:
                kw["channel"] = channel
            if profile:
                return p.chromium.launch_persistent_context(profile, no_viewport=True, **kw)
            browser = p.chromium.launch(**kw)
            return browser.new_context(no_viewport=True)
        except Exception as e:
            last = e
    raise last


def main():
    ap = argparse.ArgumentParser(description="Record a browser task (mouse, keyboard, screenshots, voice).")
    ap.add_argument("url", nargs="?", help="URL to open")
    ap.add_argument("--out", default="recordings", help="output folder (default: recordings)")
    ap.add_argument("--profile", help="folder to keep Chrome logins/cookies between runs (use a NEW folder)")
    ap.add_argument("--no-voice", action="store_true", help="do not record the microphone")
    ap.add_argument("--model", default="base", help="Whisper model: tiny/base/small/medium (default: base)")
    ap.add_argument("--lang", default=None, help="speech language code, e.g. en, hi (default: auto)")
    ap.add_argument("--translate", action="store_true", help="translate speech to English text")
    args = ap.parse_args()

    url = args.url or input("URL to open: ").strip()
    if not url.startswith(("http://", "https://", "file://")):
        url = "https://" + url

    out = Path(args.out) / datetime.now().strftime("%Y%m%d_%H%M%S")
    shots = out / "screenshots"
    shots.mkdir(parents=True, exist_ok=True)

    events, q = [], queue.Queue()
    state = {"closed": False, "shots": 0}
    t0 = time.time()

    voice = None
    if not args.no_voice:
        if sd is None:
            print("! sounddevice/numpy not installed - continuing WITHOUT voice (pip install sounddevice numpy).")
        else:
            try:
                voice = VoiceRecorder()
            except Exception as e:
                print(f"! Microphone problem ({e}) - continuing without voice.")
                voice = None

    def take_shot(page, ev, path):
        try:
            if ev["type"] == "navigate":
                page.wait_for_load_state("domcontentloaded", timeout=4000)
                page.wait_for_timeout(400)
            elif ev.get("frame") == "top":
                page.evaluate(HILITE_JS, ev)
            page.screenshot(path=str(path))
            return True
        except Exception:
            return False
        finally:
            try:
                page.evaluate(CLEAR_JS)
            except Exception:
                pass

    def handle(page, ev):
        ev = dict(ev)
        ev["t"] = round(ev.get("ts", time.time() * 1000) / 1000 - t0, 2)
        ev.pop("ts", None)
        if ev["type"] in ("click", "dblclick", "contextmenu", "navigate"):
            state["shots"] += 1
            name = f"{state['shots']:03d}_{ev['type']}.png"
            if take_shot(page, ev, shots / name):
                ev["screenshot"] = f"screenshots/{name}"
        events.append(ev)

    with sync_playwright() as p:
        context = open_context(p, args.profile)
        context.on("close", lambda *_: state.update(closed=True))
        context.expose_binding("__rec", lambda source, ev: q.put((source["page"], ev)))
        context.add_init_script(INIT_JS)

        def hook(page):
            events.append({"t": round(time.time() - t0, 2), "type": "tab_opened"})

            def on_nav(frame, pg=page):
                if frame == pg.main_frame and frame.url not in ("", "about:blank"):
                    q.put((pg, {"type": "navigate", "url": frame.url, "ts": time.time() * 1000}))

            page.on("framenavigated", on_nav)
            page.on("close", lambda *_: events.append({"t": round(time.time() - t0, 2), "type": "tab_closed"}))

        context.on("page", hook)

        if voice:
            try:
                voice.start()
                print("Microphone recording started - speak your instructions while you work.")
            except Exception as e:
                print(f"! Could not start microphone ({e}) - continuing without voice.")
                voice = None

        page = context.pages[0] if context.pages else context.new_page()
        page.goto(url)
        print("\nRECORDING...  Do your task in the Chrome window.")
        print("Stop: close the Chrome window or press Ctrl+C here.\n")

        try:
            while not state["closed"]:
                try:
                    while True:
                        pg, ev = q.get_nowait()
                        handle(pg, ev)
                except queue.Empty:
                    pass
                pages = context.pages
                if not pages:
                    break
                pages[-1].wait_for_timeout(100)
        except KeyboardInterrupt:
            print("\nStopped by Ctrl+C.")
        except Exception:
            pass  # browser/page was closed by the user
        try:
            context.close()
        except Exception:
            pass

    duration = round(time.time() - t0, 1)
    print(f"Recording finished ({duration}s, {len(events)} events). Processing...")

    # ---- voice -> text ----
    transcript, speech_lang = [], None
    if voice:
        audio = voice.stop()
        if audio is not None and len(audio) > voice.RATE:  # > 1 second
            save_wav(out / "audio.wav", audio, voice.RATE)
            offset = voice.t_start - t0
            try:
                print("Transcribing your voice (first run downloads the Whisper model)...")
                segs, speech_lang = transcribe(audio, args.model, args.lang, args.translate)
                for s in segs:
                    transcript.append({"t": round(s["start"] + offset, 2),
                                       "end": round(s["end"] + offset, 2), "text": s["text"]})
            except Exception as e:
                print(f"! Transcription failed: {e}\n  The audio is saved in {out / 'audio.wav'}")
        else:
            print("No speech captured.")

    for s in transcript:
        events.append({"t": s["t"], "type": "voice", "text": s["text"]})
    events.sort(key=lambda e: e["t"])

    # ---- write files ----
    (out / "events.json").write_text(json.dumps(events, indent=2, ensure_ascii=False), encoding="utf-8")
    (out / "transcript.json").write_text(json.dumps(transcript, indent=2, ensure_ascii=False), encoding="utf-8")

    lines = [
        "# Task recording",
        f"- Start URL: {url}",
        f"- Duration: {duration}s",
        f"- Speech language: {speech_lang or 'n/a'}",
        "- Passwords are hidden. VOICE lines are what the user said at that moment.",
        "- Goal: turn this into a reliable Playwright (Python) automation.",
        "",
        "## Timeline",
        "",
    ]
    lines += [f"{i}. {fmt(e)}" for i, e in enumerate(events, 1)]
    (out / "timeline.md").write_text("\n".join(lines), encoding="utf-8")

    zip_path = out.with_suffix(".zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in out.rglob("*"):
            if f.is_file() and f.suffix != ".wav":
                z.write(f, f.relative_to(out.parent))

    print(f"\nDone!\n  Folder : {out}\n  Zip    : {zip_path}   <- send this (or timeline.md + screenshots)")


if __name__ == "__main__":
    sys.exit(main())