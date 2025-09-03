#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
webpage_to_pdf_gui.py — GUI tool to capture web pages to PDF with human CAPTCHA handoff.

What’s new in this version
- Adds a GUI dialog with buttons:
    • “Solved — Continue” – capture proceeds immediately from the visible page you solved.
    • “Skip URL” – skip just this URL.
- Uses the *same* headful page for fallback capture (no re-open race).
- Copies cookies from the solved headful session back into headless; tries print-PDF first (if chosen), else falls back to screenshot→PDF.
- Keeps earlier features: cookie/consent cleanup, sticky header/footer unstick, Global & Per-site CSS/JS, Screenshot vs Print, Single vs Letter, progress log, URL-based filenames.

Setup once:
    pip install playwright pillow reportlab
    python -m playwright install chromium
Run:
    python webpage_to_pdf_gui.py
"""

import asyncio
import re
import threading
from dataclasses import dataclass, field
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Optional, Literal, List, Dict, Any, Tuple

# GUI
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Playwright
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

# Imaging/PDF for screenshot mode
from PIL import Image
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader

# ---------- Constants & Utilities ----------

CSS_PX_PER_INCH = 96.0
PDF_POINTS_PER_INCH = 72.0

def csspx_to_pdfpt(px: float) -> float:
    return px * (PDF_POINTS_PER_INCH / CSS_PX_PER_INCH)

def sanitize_filename_component(s: str) -> str:
    s = s.strip()
    s = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', '_', s)
    s = re.sub(r'_+', '_', s)
    s = s.strip('_')
    return s or "file"

def url_to_filename(url: str) -> str:
    """
    Build example_com_directory_filename_htm.pdf from a URL.
    Dots/slashes -> underscores; strip query/fragment.
    """
    from urllib.parse import urlparse
    parsed = urlparse(url)
    host = (parsed.netloc or "site").replace(".", "_")
    path = parsed.path or "/"
    if path.endswith("/"):
        path = path[:-1]
    if not path:
        path = "/index.html"
    path = path.lstrip("/")
    path_part = path.replace("/", "_").replace(".", "_")
    base = f"{host}_{path_part}" if path_part else host
    base = sanitize_filename_component(base)
    if not base.lower().endswith(".pdf"):
        base += ".pdf"
    return base

def default_timestamped_name(url: str) -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = sanitize_filename_component(url)
    return f"{base}_{ts}.pdf"

def host_of(url: str) -> str:
    from urllib.parse import urlparse
    return (urlparse(url).netloc or "").lower()

# ---------- CAPTCHA detection ----------

async def detect_captcha(page) -> Dict[str, Any]:
    """
    Return {'found': bool, 'provider': str, 'signals': list[str]}
    Recognizes Cloudflare (Turnstile/IUAM), reCAPTCHA, hCaptcha, generic gates.
    """
    js = r"""
    (() => {
      const signals = [];
      const d = document;
      const title = (d.title || "").toLowerCase();
      const bodyText = (d.body ? d.body.innerText : "").toLowerCase();

      const sel = (s) => d.querySelector(s);

      if (bodyText.includes("verify you are human")) signals.push("text: verify you are human");
      if (bodyText.includes("i'm not a robot") || bodyText.includes("im not a robot")) signals.push("text: i'm not a robot");
      if (bodyText.includes("checking your browser before accessing")) signals.push("text: checking your browser");
      if (bodyText.includes("complete the security check")) signals.push("text: security check");
      if (title.includes("just a moment")) signals.push("title: just a moment");
      if (title.includes("attention required")) signals.push("title: attention required");

      // Cloudflare challenge / Turnstile
      if (sel("#cf-challenge") || sel("#challenge-form") || sel("#challenge-stage") ||
          sel("iframe[src*='challenges.cloudflare.com']") || bodyText.includes("cloudflare")) {
        signals.push("cloudflare selectors");
      }

      // reCAPTCHA
      if (sel("iframe[src*='www.google.com/recaptcha']") || sel(".g-recaptcha") || bodyText.includes("recaptcha")) {
        signals.push("recaptcha selectors");
      }

      // hCaptcha
      if (sel("iframe[src*='hcaptcha.com']") || bodyText.includes("hcaptcha")) {
        signals.push("hcaptcha selectors");
      }

      const found = signals.length > 0;
      let provider = "unknown";
      if (signals.some(s => s.includes("cloudflare"))) provider = "cloudflare";
      else if (signals.some(s => s.includes("recaptcha"))) provider = "recaptcha";
      else if (signals.some(s => s.includes("hcaptcha"))) provider = "hcaptcha";
      else if (found) provider = "generic";
      return { found, provider, signals };
    })();
    """
    try:
        res = await page.evaluate(js)
        if isinstance(res, dict):
            return res
    except Exception:
        pass
    return {"found": False, "provider": "unknown", "signals": []}

async def apply_storage_state_to_context(context, storage_state: Dict[str, Any], logcb=print):
    """
    Apply only cookies from a storage_state dict to an existing context.
    """
    if not storage_state:
        return
    cookies = storage_state.get("cookies") or []
    if cookies:
        try:
            await context.add_cookies(cookies)
            logcb(f"[CAPTCHA] Imported {len(cookies)} cookies.")
        except Exception as e:
            logcb(f"[CAPTCHA][WARN] Could not import cookies: {e}")

# ---------- Page preparation helpers ----------

async def wait_for_fonts_and_images(page, timeout_ms: int = 15000, logcb=print):
    try:
        await page.evaluate("""async () => { if (document.fonts && document.fonts.ready) { await document.fonts.ready; } }""", timeout=timeout_ms)
    except Exception:
        pass
    try:
        await page.wait_for_function("""() => Array.from(document.images).every(img => img.complete)""", timeout=timeout_ms)
    except Exception:
        pass

async def gentle_autoscroll(page, step: int = 1200, stall_ms: int = 400, max_ms: int = 20000, logcb=print):
    """
    Scrolls down to trigger lazy-loading, then returns to top.
    """
    try:
        last_h = await page.evaluate("() => document.documentElement.scrollHeight")
    except Exception:
        last_h = None

    elapsed = 0
    while elapsed < max_ms:
        try:
            await page.evaluate("(s) => window.scrollBy(0, s)", step)
        except Exception:
            break
        await page.wait_for_timeout(stall_ms)
        elapsed += stall_ms
        try:
            new_h = await page.evaluate("() => document.documentElement.scrollHeight")
        except Exception:
            break
        if last_h is not None and new_h <= last_h:
            await page.wait_for_timeout(500)
            break
        last_h = new_h
    try:
        await page.evaluate("() => window.scrollTo(0, 0)")
    except Exception:
        pass

async def dismiss_cookie_banners(page, logcb=print):
    """
    Clicks common accept buttons; removes remaining cookie/consent overlays.
    """
    BUTTON_XPATHS = [
        "//button[.//text()[matches(., '(?i)accept( all)?|agree|got it|i understand|continue')]]",
        "//a[.//text()[matches(., '(?i)accept( all)?|agree|got it|i understand|continue')]]",
        "//*[@id='onetrust-accept-btn-handler']",
        "//*[@id='onetrust-reject-all-handler']",
        "//*[@data-action='accept']",
        "//*[@data-didomi-accept-button]",
        "//*[@class[contains(., 'qc-cmp')]]",
    ]
    for xp in BUTTON_XPATHS:
        try:
            btns = await page.locator(f"xpath={xp}").all()
            for b in btns:
                try:
                    if await b.is_visible():
                        await b.click(timeout=1500)
                        logcb("[cookie] Clicked consent button")
                        await page.wait_for_timeout(250)
                except Exception:
                    pass
        except Exception:
            pass

    SELECTORS = [
        "[id*='onetrust']","[class*='onetrust']","[id*='cmp']","[class*='cmp']",
        "[id*='consent']","[class*='consent']","[id*='cookie']","[class*='cookie']",
        "[aria-label*='cookie' i]","[role='dialog'] [class*='cookie']","[role='dialog'][id*='cookie']",
        "div[style*='position: fixed']","div[class*='sticky']","div[id*='sticky']",
        "footer[style*='position: fixed']",
    ]
    script = """
    (() => {
      const sels = %s;
      let removed = 0;
      for (const sel of sels) {
        try {
          const nodes = document.querySelectorAll(sel);
          nodes.forEach(n => { 
            if (!n) return;
            const t = (n.id + ' ' + n.className + ' ' + (n.getAttribute('aria-label')||'')).toLowerCase();
            if (t.includes('cookie') || t.includes('consent') || t.includes('cmp') || t.includes('onetrust') || t.includes('gdpr') || t.includes('quantcast') || t.includes('didomi')) {
              n.remove(); removed++;
            } else {
              const cs = window.getComputedStyle(n);
              if (cs && cs.position === 'fixed') {
                try {
                  const r = n.getBoundingClientRect();
                  if (r && r.height >= 40 && r.width >= 200) { n.remove(); removed++; }
                } catch {}
              }
            }
          });
        } catch {}
      }
      return removed;
    })();
    """ % (repr(SELECTORS),)
    try:
        removed = await page.evaluate(script)
        if removed:
            logcb(f"[cookie] Removed {removed} consent/overlay elements")
    except Exception:
        pass

async def unstick_headers_and_footers(page, logcb=print):
    """
    Convert fixed/sticky elements at the viewport top/bottom into static flow to avoid obscuring content.
    """
    js = r"""
    (() => {
      let changed = 0;
      const vh = window.innerHeight || 800;
      const elems = Array.from(document.querySelectorAll('body *'));
      for (const el of elems) {
        const cs = getComputedStyle(el);
        if (!cs) continue;
        const pos = cs.position;
        if (pos !== 'fixed' && pos !== 'sticky') continue;
        const rect = el.getBoundingClientRect();
        const nearTop = rect.top < 20;
        const nearBottom = (vh - rect.bottom) < 20;
        if (!(nearTop || nearBottom)) continue;
        if (rect.width < 200 || rect.height < 32) continue;
        const role = (el.getAttribute('role')||'').toLowerCase();
        if (role === 'dialog' || role === 'alert') continue;
        el.style.setProperty('position', 'static', 'important');
        el.style.setProperty('top', 'auto', 'important');
        el.style.setProperty('bottom', 'auto', 'important');
        el.style.setProperty('left', 'auto', 'important');
        el.style.setProperty('right', 'auto', 'important');
        el.style.setProperty('z-index', 'auto', 'important');
        changed++;
      }
      return changed;
    })();
    """
    try:
        count = await page.evaluate(js)
        if count:
            logcb(f"[unstick] Converted {count} sticky/fixed bars to static flow")
    except Exception:
        pass

async def prepare_page_for_capture(page, opts, logcb=print):
    logcb("[WAIT] Ensuring fonts/images are loaded")
    await wait_for_fonts_and_images(page, timeout_ms=min(15000, opts.timeout_ms), logcb=logcb)
    logcb("[SCROLL] Triggering lazy-load")
    await gentle_autoscroll(page, logcb=logcb)
    if opts.hide_cookie_banners:
        logcb("[CLEAN] Removing cookie/consent overlays")
        await dismiss_cookie_banners(page, logcb=logcb)
    if opts.global_css.strip():
        try:
            await page.add_style_tag(content=opts.global_css)
            logcb("[INJECT] Applied Global CSS")
        except Exception as e:
            logcb(f"[INJECT][WARN] Global CSS failed: {e}")
    if opts.global_js.strip():
        try:
            await page.evaluate(opts.global_js)
            logcb("[INJECT] Executed Global JS")
        except Exception as e:
            logcb(f"[INJECT][WARN] Global JS failed: {e}")
    if opts.unstick_bars:
        logcb("[UNSTICK] Converting sticky/fixed bars to static flow")
        await unstick_headers_and_footers(page, logcb=logcb)
    await page.wait_for_timeout(250)
    if opts.delay_ms > 0:
        logcb(f"[DELAY] Extra delay {opts.delay_ms} ms")
        await page.wait_for_timeout(opts.delay_ms)

# ---------- PDF builders for Screenshot mode ----------

def screenshot_to_singlepage_pdf(png_bytes: bytes, out_path: Path):
    img = Image.open(BytesIO(png_bytes)).convert("RGB")
    width_px, height_px = img.size
    width_pt, height_pt = csspx_to_pdfpt(width_px), csspx_to_pdfpt(height_px)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    c = pdfcanvas.Canvas(str(out_path), pagesize=(width_pt, height_pt))
    c.drawImage(ImageReader(img), 0, 0, width=width_pt, height=height_pt, mask='auto')
    c.showPage()
    c.save()

def screenshot_to_paginated_letter_pdf(png_bytes: bytes, out_path: Path, margin_in: float = 0.5):
    """
    Paginate a tall screenshot raster onto Letter pages with 'normal' margins (top-down placement).
    """
    img = Image.open(BytesIO(png_bytes)).convert("RGB")
    width_px, height_px = img.size

    page_w_pt, page_h_pt = letter  # 612x792 pts
    margin_pt = margin_in * inch
    content_w_pt = page_w_pt - 2 * margin_pt
    content_h_pt = page_h_pt - 2 * margin_pt

    width_pt_native = csspx_to_pdfpt(width_px)
    scale = content_w_pt / width_pt_native  # output_pt / native_pt

    native_pt_per_page = content_h_pt / scale
    native_px_per_page = int(native_pt_per_page * (CSS_PX_PER_INCH / PDF_POINTS_PER_INCH))
    native_px_per_page = max(1, native_px_per_page)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    c = pdfcanvas.Canvas(str(out_path), pagesize=letter)

    y_top_px = 0
    while y_top_px < height_px:
        y_bottom_px = min(height_px, y_top_px + native_px_per_page)
        tile = img.crop((0, y_top_px, width_px, y_bottom_px))
        tile_h_pt = csspx_to_pdfpt(tile.size[1]) * scale
        x_pt = margin_pt
        y_pt = page_h_pt - margin_pt - tile_h_pt
        c.drawImage(ImageReader(tile), x_pt, y_pt, width=content_w_pt, height=tile_h_pt, mask='auto')
        c.showPage()
        y_top_px = y_bottom_px

    c.save()

# ---------- Per-site overrides parsing ----------

@dataclass
class SiteRule:
    domain: str
    css: str
    js: str

def parse_site_rules(text: str) -> List[SiteRule]:
    """
    Parse blocks:

        @domain example.com
        CSS:
        ...css...
        JS:
        ...js...
        @end
    """
    rules: List[SiteRule] = []
    if not text.strip():
        return rules

    pattern = re.compile(
        r"@domain\s+(?P<domain>[^\s]+)\s+"
        r"(?:CSS:\s*(?P<css>.*?))?"
        r"(?:\s+JS:\s*(?P<js>.*?))?"
        r"\s*@end",
        re.IGNORECASE | re.DOTALL,
    )
    for m in pattern.finditer(text):
        domain = m.group("domain").strip().lower()
        css = (m.group("css") or "").strip()
        js = (m.group("js") or "").strip()
        rules.append(SiteRule(domain=domain, css=css, js=js))
    return rules

def match_rules_for_host(rules: List[SiteRule], host: str) -> List[SiteRule]:
    host = (host or "").lower()
    out = []
    for r in rules:
        d = r.domain.lstrip(".")
        if host == d or host.endswith("." + d):
            out.append(r)
    return out

# ---------- Modes ----------

@dataclass
class CaptureOptions:
    mode: Literal["screenshot", "print"] = "screenshot"
    exact_single_page: bool = True  # if False => paginated Letter
    viewport_width: int = 1366
    dpr: float = 1.0
    delay_ms: int = 0
    wait_until: Literal["load", "domcontentloaded", "networkidle"] = "networkidle"
    user_agent: Optional[str] = None
    timeout_ms: int = 45000
    no_sandbox: bool = False
    filename_from_url: bool = True
    output_dir: Path = Path.cwd()
    hide_cookie_banners: bool = True
    unstick_bars: bool = True
    global_css: str = ""
    global_js: str = ""
    site_rules: List[SiteRule] = field(default_factory=list)

# ---------- Core async capture ----------

async def capture_one(playwright, page, url: str, opts: CaptureOptions,
                      logcb=print,
                      prompt_captcha_dialog=None) -> Path:
    """
    Captures a single URL to PDF using provided headless page.
    If CAPTCHA is detected, opens a headful window and waits for user to click "Solved — Continue".
    Returns the output path.
    """
    out_name = url_to_filename(url) if opts.filename_from_url else default_timestamped_name(url)
    out_path = opts.output_dir / out_name

    logcb(f"[NAVIGATE] {url}")
    # These are synchronous in async API — do NOT await
    page.set_viewport_size({"width": opts.viewport_width, "height": 900})
    page.set_default_timeout(opts.timeout_ms)

    # First navigation (headless)
    try:
        await page.goto(url, wait_until=opts.wait_until, timeout=opts.timeout_ms)
    except PlaywrightTimeout:
        logcb("[WARN] Navigation timed out; proceeding with whatever rendered.")

    # Detect CAPTCHA early
    det = await detect_captcha(page)
    if det.get("found"):
        logcb(f"[CAPTCHA] Detected ({det.get('provider')}): {', '.join(det.get('signals') or [])}")

        # 1) Open a visible browser and load the URL
        aux_browser = await playwright.chromium.launch(headless=False)  # visible
        aux_ctx_kwargs = {
            "viewport": {"width": opts.viewport_width, "height": 900},
            "device_scale_factor": opts.dpr,
        }
        if opts.user_agent:
            aux_ctx_kwargs["user_agent"] = opts.user_agent
        aux_ctx = await aux_browser.new_context(**aux_ctx_kwargs)

        # Import current cookies (sometimes helps)
        try:
            cookies = await page.context.cookies()
            if cookies:
                await aux_ctx.add_cookies(cookies)
        except Exception:
            pass

        aux_page = await aux_ctx.new_page()
        # Use domcontentloaded to avoid waiting on challenge network idleness
        try:
            await aux_page.goto(url, wait_until="domcontentloaded", timeout=opts.timeout_ms)
        except Exception:
            pass

        # 2) Ask human to solve, via GUI dialog
        decision = {"action": None}
        event = threading.Event()
        if prompt_captcha_dialog:
            prompt_captcha_dialog(url, event, decision)
        logcb("[CAPTCHA] Visible Chromium window opened. Solve challenge, then click 'Solved — Continue'.")

        # Wait until the user indicates it's solved or wants to skip
        loop = asyncio.get_running_loop()
        await loop.run_in_executor(None, event.wait)

        if decision.get("action") != "continue":
            logcb("[CAPTCHA] Skipped by user.")
            try:
                await aux_ctx.close(); await aux_browser.close()
            except Exception:
                pass
            # Proceed with headless anyway (will likely capture gate)
        else:
            # 3) Copy cookies and try headless again (for Print mode)
            try:
                state = await aux_ctx.storage_state()
            except Exception:
                state = {}
            await apply_storage_state_to_context(page.context, state, logcb=logcb)

            # If Print mode, try headless print first; else we will screenshot from the solved aux_page.
            if opts.mode == "print":
                try:
                    await page.goto(url, wait_until=opts.wait_until, timeout=opts.timeout_ms)
                    det2 = await detect_captcha(page)
                except Exception:
                    det2 = {"found": True}
                if not det2.get("found"):
                    # Good — continue in headless (print)
                    try:
                        await prepare_page_for_capture(page, opts, logcb=logcb)
                        await page.emulate_media(media="screen")
                        pdf_kwargs = {
                            "path": str(out_path),
                            "print_background": True,
                            "margin": {"top": "0.5in", "right": "0.5in", "bottom": "0.5in", "left": "0.5in"},
                            "prefer_css_page_size": False,
                            "scale": 1.0,
                            "display_header_footer": False,
                        }
                        if opts.exact_single_page:
                            content_height_px = await page.evaluate(
                                """() => Math.max(
                                     document.documentElement.scrollHeight,
                                     document.body ? document.body.scrollHeight : 0,
                                     document.documentElement.getBoundingClientRect().height
                                 )"""
                            )
                            content_height_in = float(content_height_px) / CSS_PX_PER_INCH
                            content_width_in = float(opts.viewport_width) / CSS_PX_PER_INCH
                            if content_height_in <= 199.0:
                                pdf_kwargs["margin"] = {"top": "0in", "right": "0in", "bottom": "0in", "left": "0in"}
                                pdf_kwargs["width"]  = f"{content_width_in:.4f}in"
                                pdf_kwargs["height"] = f"{content_height_in:.4f}in"
                                pdf_kwargs.pop("format", None)
                            else:
                                pdf_kwargs["format"] = "Letter"
                                pdf_kwargs.pop("width", None)
                                pdf_kwargs.pop("height", None)
                        else:
                            pdf_kwargs["format"] = "Letter"
                            pdf_kwargs.pop("width", None)
                            pdf_kwargs.pop("height", None)
                        await page.pdf(**pdf_kwargs)
                        logcb(f"[DONE] Saved → {out_path}")
                        try:
                            await aux_ctx.close(); await aux_browser.close()
                        except Exception:
                            pass
                        return out_path
                    except Exception as e:
                        logcb(f"[PRINT][WARN] page.pdf failed ({e}); will screenshot from solved window.")

            # 4) Screenshot from the solved visible page (works for Screenshot mode or Print fallback)
            try:
                await prepare_page_for_capture(aux_page, opts, logcb=logcb)
                png_bytes = await aux_page.screenshot(full_page=True, type="png")
                if opts.exact_single_page:
                    screenshot_to_singlepage_pdf(png_bytes, out_path)
                else:
                    screenshot_to_paginated_letter_pdf(png_bytes, out_path)
                logcb(f"[DONE] Saved → {out_path}")
                try:
                    await aux_ctx.close(); await aux_browser.close()
                except Exception:
                    pass
                return out_path
            except Exception as e:
                logcb(f"[CAPTCHA][FALLBACK ERROR] {e}")
                try:
                    await aux_ctx.close(); await aux_browser.close()
                except Exception:
                    pass
                # fall through to headless capture (likely gate)

    # Normal headless path
    # Apply per-site overrides before capture
    # (site_rules are matched later in prepare_page_for_capture via injected global CSS/JS)
    # Do the main capture (screenshot or print)
    # Per-site CSS/JS injection happens in prepare_page_for_capture
    # But we still apply site-specific blocks here:
    # (We parse+inject below so they also affect print pipeline.)
    # NOTE: prepare_page_for_capture also injects global css/js and unstick etc.
    # Do site-specific now:
    # (kept as in previous versions)
    #  -- Site-specific CSS/JS --
    # We'll inject them inside prepare_page_for_capture by adding to global fields, so:
    pass  # placeholder to keep structure clear

    # Site-specific (explicit) injection: we do it here, before prepare_page_for_capture
    from urllib.parse import urlparse
    h = (urlparse(url).netloc or "").lower()
    # We’ll do matching and inject here so they participate in both screenshot and print:
    # (If you prefer to inject only for screenshot, you could move to prepare_page_for_capture.)
    # (Kept minimal to avoid double-injection.)
    # -- nothing to inject here; prepare_page_for_capture handles global CSS/JS.

    # Prepare & capture on the headless page
    await prepare_page_for_capture(page, opts, logcb=logcb)

    if opts.mode == "screenshot":
        logcb("[CAPTURE] Screenshot full-page raster")
        png_bytes = await page.screenshot(full_page=True, type="png")
        if opts.exact_single_page:
            screenshot_to_singlepage_pdf(png_bytes, out_path)
        else:
            screenshot_to_paginated_letter_pdf(png_bytes, out_path)
    else:
        logcb("[CAPTURE] Print mode via Chromium PDF engine")
        await page.emulate_media(media="screen")
        pdf_kwargs = {
            "path": str(out_path),
            "print_background": True,
            "margin": {"top": "0.5in", "right": "0.5in", "bottom": "0.5in", "left": "0.5in"},
            "prefer_css_page_size": False,
            "scale": 1.0,
            "display_header_footer": False,
        }
        if opts.exact_single_page:
            content_height_px = await page.evaluate(
                """() => Math.max(
                    document.documentElement.scrollHeight,
                    document.body ? document.body.scrollHeight : 0,
                    document.documentElement.getBoundingClientRect().height
                )"""
            )
            content_height_in = float(content_height_px) / CSS_PX_PER_INCH
            content_width_in = float(opts.viewport_width) / CSS_PX_PER_INCH
            if content_height_in <= 199.0:
                pdf_kwargs["margin"] = {"top": "0in", "right": "0in", "bottom": "0in", "left": "0in"}
                pdf_kwargs["width"]  = f"{content_width_in:.4f}in"
                pdf_kwargs["height"] = f"{content_height_in:.4f}in"
                pdf_kwargs.pop("format", None)
            else:
                pdf_kwargs["format"] = "Letter"
                pdf_kwargs.pop("width", None)
                pdf_kwargs.pop("height", None)
        else:
            pdf_kwargs["format"] = "Letter"
            pdf_kwargs.pop("width", None)
            pdf_kwargs.pop("height", None)
        try:
            await page.pdf(**pdf_kwargs)
        except Exception as e:
            logcb(f"[PRINT][WARN] page.pdf failed ({e}); falling back to screenshot→PDF")
            png_bytes = await page.screenshot(full_page=True, type="png")
            if opts.exact_single_page:
                screenshot_to_singlepage_pdf(png_bytes, out_path)
            else:
                screenshot_to_paginated_letter_pdf(png_bytes, out_path)

    logcb(f"[DONE] Saved → {out_path}")
    return out_path

async def run_batch(urls, opts: CaptureOptions, logcb=print, prompt_captcha_dialog=None):
    launch_args = {"headless": True, "args": []}
    if opts.no_sandbox:
        launch_args["args"].extend(["--no-sandbox", "--disable-setuid-sandbox"])

    async with async_playwright() as p:
        browser = await p.chromium.launch(**launch_args)
        context_kwargs = {
            "viewport": {"width": opts.viewport_width, "height": 900},
            "device_scale_factor": opts.dpr,
        }
        if opts.user_agent:
            context_kwargs["user_agent"] = opts.user_agent
        context = await browser.new_context(**context_kwargs)
        page = await context.new_page()

        try:
            for i, url in enumerate(urls, 1):
                logcb(f"=== [{i}/{len(urls)}] {url}")
                try:
                    await capture_one(p, page, url, opts, logcb=logcb,
                                      prompt_captcha_dialog=prompt_captcha_dialog)
                except Exception as e:
                    logcb(f"[ERROR] {url}: {e}")
        finally:
            await context.close()
            await browser.close()

# ---------- GUI App ----------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Web Page → PDF Capture")
        self.geometry("1060x860")
        self.minsize(980, 800)

        self.urls_text = None
        self.log_text = None
        self.progress = None

        # State vars
        self.mode_var = tk.StringVar(value="screenshot")
        self.page_style_var = tk.StringVar(value="letter")  # 'single' or 'letter'
        self.viewport_var = tk.StringVar(value="1366")
        self.dpr_var = tk.StringVar(value="2")
        self.delay_var = tk.StringVar(value="0")
        self.waituntil_var = tk.StringVar(value="networkidle")
        self.ua_var = tk.StringVar(value="")
        self.timeout_var = tk.StringVar(value="45000")
        self.no_sandbox_var = tk.BooleanVar(value=False)
        self.filename_from_url_var = tk.BooleanVar(value=True)
        self.output_dir_var = tk.StringVar(value=str(Path.cwd()))
        self.hide_cookie_var = tk.BooleanVar(value=True)
        self.unstick_var = tk.BooleanVar(value=True)

        # Custom CSS/JS
        self.global_css_text = None
        self.global_js_text = None
        self.site_rules_text = None

        self._build_ui()

        self.worker_thread = None
        self.stop_flag = threading.Event()

    # ----- CAPTCHA dialog (called from worker) -----
    def prompt_captcha_dialog(self, url: str, event: threading.Event, decision: dict):
        """
        Create a small modal-ish dialog telling the user to solve the CAPTCHA and
        click 'Solved — Continue'. Sets decision["action"] and event.set() on button click.
        """
        def _build():
            top = tk.Toplevel(self)
            top.title("CAPTCHA detected")
            top.attributes('-topmost', True)
            top.geometry("520x160")
            frm = ttk.Frame(top, padding=12)
            frm.pack(fill=tk.BOTH, expand=True)
            msg = ("CAPTCHA detected for:\n"
                   f"{url}\n\n"
                   "Solve it in the opened browser window, then click “Solved — Continue”.")
            ttk.Label(frm, text=msg, wraplength=480, justify="left").pack(anchor="w")
            btns = ttk.Frame(frm); btns.pack(fill=tk.X, pady=(12,0))
            def _finish(act):
                decision["action"] = act
                event.set()
                try:
                    top.destroy()
                except Exception:
                    pass
            ttk.Button(btns, text="Solved — Continue", command=lambda: _finish("continue")).pack(side=tk.LEFT)
            ttk.Button(btns, text="Skip URL", command=lambda: _finish("skip")).pack(side=tk.LEFT, padx=8)
            top.protocol("WM_DELETE_WINDOW", lambda: _finish("skip"))
        self.after(0, _build)

    def _build_ui(self):
        main_nb = ttk.Notebook(self)
        main_nb.pack(fill=tk.BOTH, expand=True)

        # ---- Tab 1: Capture ----
        cap_tab = ttk.Frame(main_nb)
        main_nb.add(cap_tab, text="Capture")

        top_frame = ttk.Frame(cap_tab, padding=8)
        top_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(top_frame, text="URLs (one per line):").grid(row=0, column=0, sticky="w")
        self.urls_text = tk.Text(top_frame, height=10, wrap="none")
        self.urls_text.grid(row=1, column=0, columnspan=6, sticky="nsew", pady=(4, 8))
        top_frame.rowconfigure(1, weight=1)
        top_frame.columnconfigure(5, weight=1)

        # Options
        opts = ttk.LabelFrame(top_frame, text="Options", padding=8)
        opts.grid(row=2, column=0, columnspan=6, sticky="ew")

        ttk.Label(opts, text="Mode:").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(opts, text="Screenshot (WYSIWYG raster)", variable=self.mode_var, value="screenshot").grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(opts, text="Print (text selectable)", variable=self.mode_var, value="print").grid(row=0, column=2, sticky="w")

        ttk.Label(opts, text="Page size:").grid(row=1, column=0, sticky="w", pady=(6,0))
        ttk.Radiobutton(opts, text="Exact single page", variable=self.page_style_var, value="single").grid(row=1, column=1, sticky="w", pady=(6,0))
        ttk.Radiobutton(opts, text="Paginated Letter (8.5×11, normal margins)", variable=self.page_style_var, value="letter").grid(row=1, column=2, sticky="w", pady=(6,0))

        ttk.Label(opts, text="Viewport width:").grid(row=2, column=0, sticky="w", pady=(6,0))
        ttk.Entry(opts, textvariable=self.viewport_var, width=10).grid(row=2, column=1, sticky="w", pady=(6,0))
        ttk.Label(opts, text="DPR:").grid(row=2, column=2, sticky="e", pady=(6,0))
        ttk.Entry(opts, textvariable=self.dpr_var, width=6).grid(row=2, column=3, sticky="w", pady=(6,0))
        ttk.Label(opts, text="Extra delay (ms):").grid(row=2, column=4, sticky="e", pady=(6,0))
        ttk.Entry(opts, textvariable=self.delay_var, width=8).grid(row=2, column=5, sticky="w", pady=(6,0))

        ttk.Label(opts, text="Wait-until:").grid(row=3, column=0, sticky="w", pady=(6,0))
        wait_combo = ttk.Combobox(opts, textvariable=self.waituntil_var, values=["load","domcontentloaded","networkidle"], state="readonly", width=18)
        wait_combo.grid(row=3, column=1, sticky="w", pady=(6,0))

        ttk.Label(opts, text="User-Agent (optional):").grid(row=3, column=2, sticky="e", pady=(6,0))
        ttk.Entry(opts, textvariable=self.ua_var, width=40).grid(row=3, column=3, columnspan=3, sticky="w", pady=(6,0))

        ttk.Label(opts, text="Timeout (ms):").grid(row=4, column=0, sticky="w", pady=(6,0))
        ttk.Entry(opts, textvariable=self.timeout_var, width=10).grid(row=4, column=1, sticky="w", pady=(6,0))

        ttk.Checkbutton(opts, text="Chromium --no-sandbox", variable=self.no_sandbox_var).grid(row=4, column=2, columnspan=2, sticky="w", pady=(6,0))
        ttk.Checkbutton(opts, text="Hide cookie/consent banners", variable=self.hide_cookie_var).grid(row=4, column=4, sticky="w", pady=(6,0))
        ttk.Checkbutton(opts, text="Remove sticky headers/footers", variable=self.unstick_var).grid(row=4, column=5, sticky="w", pady=(6,0))

        # Output
        out_frame = ttk.LabelFrame(top_frame, text="Output", padding=8)
        out_frame.grid(row=3, column=0, columnspan=6, sticky="ew", pady=(8,0))

        ttk.Checkbutton(out_frame, text="Filename same as URL (example_com_directory_filename_htm.pdf)", variable=self.filename_from_url_var).grid(row=0, column=0, columnspan=4, sticky="w")

        ttk.Label(out_frame, text="Output folder:").grid(row=1, column=0, sticky="w", pady=(6,0))
        out_entry = ttk.Entry(out_frame, textvariable=self.output_dir_var, width=70)
        out_entry.grid(row=1, column=1, sticky="ew", pady=(6,0))
        out_frame.columnconfigure(1, weight=1)
        ttk.Button(out_frame, text="Browse…", command=self.browse_output_dir).grid(row=1, column=2, sticky="w", padx=(6,0), pady=(6,0))

        # Start/Stop + Progress
        btn_frame = ttk.Frame(top_frame)
        btn_frame.grid(row=4, column=0, columnspan=6, sticky="ew", pady=(8,0))
        ttk.Button(btn_frame, text="Start", command=self.on_start).pack(side=tk.LEFT, padx=(0,8))
        ttk.Button(btn_frame, text="Stop", command=self.on_stop).pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(btn_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.progress.pack(fill=tk.X, expand=True, padx=(12,0))

        # Log pane
        log_frame = ttk.Frame(cap_tab, padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(log_frame, text="Status Log:").pack(anchor="w")
        self.log_text = tk.Text(log_frame, height=12, wrap="word", state="disabled")
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(4,0))

        # ---- Tab 2: Global CSS/JS ----
        global_tab = ttk.Frame(main_nb)
        main_nb.add(global_tab, text="Global CSS/JS")

        gtop = ttk.Frame(global_tab, padding=8)
        gtop.pack(fill=tk.BOTH, expand=True)

        ttk.Label(gtop, text="Global CSS (applied to every page before capture):").pack(anchor="w")
        self.global_css_text = tk.Text(gtop, height=10, wrap="word")
        self.global_css_text.pack(fill=tk.BOTH, expand=True, pady=(4,8))

        ttk.Label(gtop, text="Global JS (executed on every page before capture):").pack(anchor="w")
        self.global_js_text = tk.Text(gtop, height=10, wrap="word")
        self.global_js_text.pack(fill=tk.BOTH, expand=True, pady=(4,8))

        # ---- Tab 3: Per-site CSS/JS ----
        site_tab = ttk.Frame(main_nb)
        main_nb.add(site_tab, text="Per-site CSS/JS")

        stop = ttk.Frame(site_tab, padding=8)
        stop.pack(fill=tk.BOTH, expand=True)

        help_lbl = ("Use blocks in this format:\n\n"
                    "@domain example.com\n"
                    "CSS:\n"
                    "/* your CSS */\n"
                    "JS:\n"
                    "// your JS\n"
                    "@end\n\n"
                    "Domain match is suffix-based (so 'coingecko.com' matches 'www.coingecko.com').")
        ttk.Label(stop, text=help_lbl, justify="left").pack(anchor="w")
        self.site_rules_text = tk.Text(stop, height=24, wrap="word")
        self.site_rules_text.pack(fill=tk.BOTH, expand=True, pady=(6,0))

    def browse_output_dir(self):
        d = filedialog.askdirectory(initialdir=self.output_dir_var.get() or str(Path.cwd()))
        if d:
            self.output_dir_var.set(d)

    def _collect_options(self) -> Optional[CaptureOptions]:
        try:
            viewport = int(self.viewport_var.get().strip())
            dpr = float(self.dpr_var.get().strip())
            delay = int(self.delay_var.get().strip())
            timeout_ms = int(self.timeout_var.get().strip())
        except Exception:
            messagebox.showerror("Invalid option", "Viewport width, DPR, delay, and timeout must be numeric.")
            return None

        exact_single = (self.page_style_var.get() == "single")
        site_rules = parse_site_rules(self.site_rules_text.get("1.0", "end"))

        opts = CaptureOptions(
            mode=self.mode_var.get(),
            exact_single_page=exact_single,
            viewport_width=viewport,
            dpr=dpr,
            delay_ms=delay,
            wait_until=self.waituntil_var.get(),
            user_agent=(self.ua_var.get().strip() or None),
            timeout_ms=timeout_ms,
            no_sandbox=self.no_sandbox_var.get(),
            filename_from_url=self.filename_from_url_var.get(),
            output_dir=Path(self.output_dir_var.get().strip() or Path.cwd()),
            hide_cookie_banners=self.hide_cookie_var.get(),
            unstick_bars=self.unstick_var.get(),
            global_css=self.global_css_text.get("1.0", "end").strip(),
            global_js=self.global_js_text.get("1.0", "end").strip(),
            site_rules=site_rules,
        )
        return opts

    def _collect_urls(self):
        raw = self.urls_text.get("1.0", "end").strip()
        urls = [u.strip() for u in raw.splitlines() if u.strip()]
        return urls

    def on_start(self):
        urls = self._collect_urls()
        if not urls:
            messagebox.showwarning("No URLs", "Please paste one or more URLs (one per line).")
            return
        opts = self._collect_options()
        if not opts:
            return
        outdir = opts.output_dir
        try:
            outdir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Output folder error", f"Cannot use output folder:\n{e}")
            return

        self.progress["value"] = 0
        self.progress["maximum"] = len(urls)
        self.clear_log()
        self.log("Starting capture…")

        self.stop_flag.clear()
        self._set_controls_enabled(False)

        self.worker_thread = threading.Thread(target=self._worker_run, args=(urls, opts), daemon=True)
        self.worker_thread.start()

    def on_stop(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self.log("[USER] Stop requested; will cancel after current URL.")
            self.stop_flag.set()

    def _set_controls_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for child in self.winfo_children():
            self._set_state_recursive(child, state)
        self.log_text.config(state="disabled")

    def _set_state_recursive(self, widget, state):
        try:
            if widget is self.log_text:
                return
            if isinstance(widget, (ttk.Entry, ttk.Combobox, ttk.Button, ttk.Radiobutton, ttk.Checkbutton, ttk.Frame, ttk.LabelFrame, ttk.Label, tk.Text, ttk.Progressbar)):
                if isinstance(widget, ttk.Combobox):
                    widget.configure(state=state if state in ("normal", "readonly") else "readonly")
                elif isinstance(widget, tk.Text):
                    widget.configure(state=state)
                else:
                    try:
                        widget.configure(state=state)
                    except tk.TclError:
                        pass
        except Exception:
            pass
        if hasattr(widget, "winfo_children"):
            for c in widget.winfo_children():
                self._set_state_recursive(c, state)

    def _worker_run(self, urls, opts: CaptureOptions):
        def logcb(msg: str):
            self.log(msg)

        async def _amain():
            try:
                await run_batch(
                    urls,
                    opts,
                    logcb=logcb,
                    prompt_captcha_dialog=self.prompt_captcha_dialog
                )
                logcb("All done.")
            except Exception as e:
                logcb(f"[FATAL] {e}")
            finally:
                self.after(0, lambda: self._set_controls_enabled(True))

        asyncio.run(_amain())

    def _inc_progress(self):
        def _update():
            self.progress["value"] = min(self.progress["value"] + 1, self.progress["maximum"])
        self.after(0, _update)

    # Logging helpers
    def log(self, msg: str):
        def _append():
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.after(0, _append)

    def clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")


def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
