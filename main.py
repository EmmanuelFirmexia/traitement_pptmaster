"""
main.py — traitement_pptmaster
FastAPI service wrapping the PPT Master skill pipeline.

Pipeline (headless, non-interactive):
  Phase A – Strategist LLM call  → design_spec.md + spec_lock.md
  Phase B – Executor  LLM call   → SVG slides (svg_output/)
  Phase C – Scripts              → finalize_svg.py → svg_to_pptx.py → PPTX
"""

import asyncio
import base64
import functools
import json
import logging
import os
import re
import subprocess
import sys
import time
import uuid
from pathlib import Path
from typing import Literal

import anthropic
import httpx
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

load_dotenv()

# ─────────────────────────────────────────────
# PATHS
# ─────────────────────────────────────────────
SKILL_DIR    = Path(__file__).parent / "skills" / "ppt-master"
SCRIPTS_DIR  = SKILL_DIR / "scripts"
PROJECTS_DIR = Path(__file__).parent / "projects"
PROJECTS_DIR.mkdir(exist_ok=True)
EXAMPLES_DIR = Path(__file__).parent / "examples"

SUPABASE_URL         = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")

sys.path.insert(0, str(SCRIPTS_DIR))

# ─────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# MODELS
# ─────────────────────────────────────────────

class Palette(BaseModel):
    primary:   str
    secondary: str
    accent:    str


class GenerateRequest(BaseModel):
    content:          str
    prompt_injection: str = ""
    style:            str = "professional"
    palette:          Palette
    slides_count:     int = 6
    tenant_id:        str
    title:            str = ""
    layout:           str = "free"
    provider:         Literal["claude", "mistral"] = "claude"

# ─────────────────────────────────────────────
# LAYOUT CATALOGUE
# ─────────────────────────────────────────────

LAYOUT_MAP: dict[str, str | None] = {
    "free":           None,
    "glassmorphism":  "ppt169_glassmorphism_demo",
    "swiss_grid":     "ppt169_swiss_grid_systems",
    "editorial":      "ppt169_pritzker_2026",
    "data":           "ppt169_global_ai_capital_2026",
    "brutalist":      "ppt169_brutalist_ai_newspaper_2026",
    "blueprint":      "ppt169_kubernetes_blueprint_2026",
    "dark_tech":      "ppt169_building_effective_agents",
    "consulting":     "ppt169_kimsoong_loyalty_programme",
    "showcase":       "ppt169_image_text_showcase",
    "magazine":       "ppt169_home_design_trends_2026",
    "attention":      "ppt169_attention_is_all_you_need",
    "cangzhuo":       "ppt169_cangzhuo",
    "fashion":        "ppt169_fashion_weekly_digest",
    "general_dark":   "ppt169_general_dark_tech_claude_code_auto_mode",
    "high_rise":      "ppt169_high_rise_renewal",
    "lin_huiyin":     "ppt169_lin_huiyin_architect",
    "lin_huiyin_rev": "ppt169_lin_huiyin_architect_revised",
    "liziqi":         "ppt169_liziqi_plant_dye_colors",
    "lora":           "ppt169_lora_hu_2021",
    "sugar_rush":     "ppt169_sugar_rush_memphis",
    "zine":           "ppt169_indie_bookstore_zine_guide",
}

_LAYOUTS: list[dict] = [
    {"id": "free",           "label": "Design libre IA",       "sublabel": "L'IA compose librement selon votre charte",  "color": "#6366F1"},
    {"id": "glassmorphism",  "label": "Glassmorphism SaaS",    "sublabel": "Moderne, translucide, product UI",           "color": "#0EA5E9"},
    {"id": "swiss_grid",     "label": "Swiss Grid",            "sublabel": "Typographique, structuré, épuré",            "color": "#EF4444"},
    {"id": "editorial",      "label": "Editorial Magazine",    "sublabel": "Photographique, aéré, premium",              "color": "#1E293B"},
    {"id": "data",           "label": "Data Journalism",       "sublabel": "Sombre, graphiques, Bloomberg-style",        "color": "#0F172A"},
    {"id": "brutalist",      "label": "Brutalist",             "sublabel": "Impact fort, typographie dense",             "color": "#DC2626"},
    {"id": "blueprint",      "label": "Blueprint Tech",        "sublabel": "Schémas, isométrique, IT",                   "color": "#0891B2"},
    {"id": "dark_tech",      "label": "Dark Tech",             "sublabel": "Sombre, tech, consulting digital",           "color": "#1E293B"},
    {"id": "consulting",     "label": "Corporate Consulting",  "sublabel": "Sobre, conseil, corporate propre",           "color": "#334155"},
    {"id": "showcase",       "label": "Editorial Showcase",    "sublabel": "Riche en images, mise en page moderne",      "color": "#7C3AED"},
    {"id": "magazine",       "label": "Magazine Tendances",    "sublabel": "Lifestyle premium, visuels forts",           "color": "#BE185D"},
    {"id": "attention",      "label": "Academic Blueprint",    "sublabel": "Research, schémas, académique",              "color": "#F59E0B"},
    {"id": "cangzhuo",       "label": "Chinese Ink Aesthetic", "sublabel": "Encre, minimalisme, culture",                "color": "#78716C"},
    {"id": "fashion",        "label": "Fashion Editorial",     "sublabel": "Mode, luxe, magazine",                      "color": "#EC4899"},
    {"id": "general_dark",   "label": "Dark Tech Général",     "sublabel": "Dark theme généraliste, tech",              "color": "#1E293B"},
    {"id": "high_rise",      "label": "Architecture Urbaine",  "sublabel": "Editorial, urban renewal",                   "color": "#64748B"},
    {"id": "lin_huiyin",     "label": "Portrait Culturel",     "sublabel": "Biographie, culture, photo",                 "color": "#92400E"},
    {"id": "lin_huiyin_rev", "label": "Portrait Culturel v2",  "sublabel": "Version révisée",                            "color": "#92400E"},
    {"id": "liziqi",         "label": "Nature & Couleurs",     "sublabel": "Couleurs naturelles, artisanat",             "color": "#16A34A"},
    {"id": "lora",           "label": "Technical Paper",       "sublabel": "Académique, technique, dense",               "color": "#6366F1"},
    {"id": "sugar_rush",     "label": "Memphis Pop",           "sublabel": "Coloré, playful, énergie",                   "color": "#F97316"},
    {"id": "zine",           "label": "Risograph Zine",        "sublabel": "Duotone, artisanal, culture",                "color": "#84CC16"},
]

_LAYOUTS_BY_ID: dict[str, dict] = {l["id"]: l for l in _LAYOUTS}

# ─────────────────────────────────────────────
# JOB STORE
# ─────────────────────────────────────────────

JOBS_FILE = "/tmp/pptmaster_jobs.json"


def _load_jobs() -> dict:
    if os.path.exists(JOBS_FILE):
        try:
            with open(JOBS_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_job(job_id: str, data: dict) -> None:
    jobs = _load_jobs()
    jobs[job_id] = data
    with open(JOBS_FILE, "w") as f:
        json.dump(jobs, f)


def _set_job(job_id: str, **kwargs: object) -> None:
    jobs = _load_jobs()
    if job_id in jobs:
        jobs[job_id].update(kwargs)
        with open(JOBS_FILE, "w") as f:
            json.dump(jobs, f)

# ─────────────────────────────────────────────
# APP
# ─────────────────────────────────────────────
app = FastAPI(
    title="PPT Master API",
    description="Génération PPTX via le pipeline SVG PPT Master",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────
# LAYOUTS ENDPOINT
# ─────────────────────────────────────────────

@app.get("/layouts")
async def list_layouts():
    return _LAYOUTS

# ─────────────────────────────────────────────
# SKILL FILE LOADER
# ─────────────────────────────────────────────

def _skill(rel: str, max_chars: int = 0) -> str:
    p = SKILL_DIR / rel
    if not p.exists():
        return ""
    txt = p.read_text(encoding="utf-8")
    return txt[:max_chars] if max_chars else txt

# ─────────────────────────────────────────────
# PHASE A — STRATEGIST PROMPT
# ─────────────────────────────────────────────

_STRATEGIST_SYSTEM = """
You are PPT Master acting as the STRATEGIST role.
You run in HEADLESS API MODE — skip all BLOCKING stops, browser UIs, and interactive confirmations.
Auto-approve all Eight Confirmations using the parameters provided.

Your task: produce ONLY design_spec.md and spec_lock.md for the requested presentation.
{design_example_section}
REFERENCE — spec_lock.md skeleton (follow this EXACTLY):
{spec_lock_ref}

REFERENCE — shared technical standards (SVG/PPTX rules):
{shared_standards}

OUTPUT FORMAT — respond with EXACTLY this structure, no markdown around the blocks:
DESIGN_SPEC_START
[full content of design_spec.md as plain text]
DESIGN_SPEC_END
SPEC_LOCK_START
[full content of spec_lock.md as plain text]
SPEC_LOCK_END
""".strip()

_STRATEGIST_USER = """
Generate spec files for the following presentation:

CONTENT:
{content}

CONFIRMED PARAMETERS:
- Canvas: PPT 16:9  →  viewBox 0 0 1280 720
- Slides: {slides_count}
- Style: {style}
- Palette:
    primary:   {primary}
    secondary: {secondary}
    accent:    {accent}
{extra}

Rules:
- Lock colors: use primary={primary} as `primary`, secondary={secondary} as `secondary_accent`, accent={accent} as `accent`.
  Set `bg: #FFFFFF` unless the style clearly calls for a dark background.
- Auto-select mode (pyramid/narrative/instructional/showcase/briefing) from content type.
- Auto-select visual_style matching the style parameter.
- page_layouts: assign one layout per page (free design — no template SVGs needed).
- No images section needed (placeholder rectangles will be used).
- font_family: "Calibri", Arial, sans-serif
- body: 22, title: 40, subtitle: 28, annotation: 14
- Use the DESIGN_SPEC_START / DESIGN_SPEC_END and SPEC_LOCK_START / SPEC_LOCK_END markers exactly.
""".strip()

# ─────────────────────────────────────────────
# PHASE B — EXECUTOR PROMPT
# ─────────────────────────────────────────────

_EXECUTOR_SYSTEM = """
You are PPT Master acting as the EXECUTOR role.
You run in HEADLESS API MODE — generate all SVG slides automatically without any confirmation.

REFERENCE — executor guidelines:
{executor_base}

REFERENCE — shared technical standards:
{shared_standards}

CRITICAL SVG RULES:
1. viewBox MUST be "0 0 1280 720" for every slide.
2. All colors MUST come from spec_lock.md — no invented values.
3. No external URLs (no href="http://...") — use inline shapes / gradients instead of images.
4. Fonts: "Calibri", Arial, sans-serif — PPT-safe stacks only.
5. Each SVG must be a complete, self-contained <svg> element.
6. Slide files: 01_cover.svg, 02_*.svg, …, {slides_count_padded}_closing.svg

OUTPUT FORMAT — return exactly this JSON array and nothing else:
[
  {{"path": "svg_output/01_cover.svg",   "content": "<svg viewBox=\\"0 0 1280 720\\" ...>...</svg>"}},
  {{"path": "svg_output/02_agenda.svg",  "content": "<svg ...>...</svg>"}},
  ...
]
""".strip()

_EXECUTOR_USER = """
Generate exactly {slides_count} SVG slides based on these specs:

DESIGN SPEC:
{design_spec_md}

SPEC LOCK:
{spec_lock_md}

SOURCE CONTENT:
{content}

VISUAL STYLE: {layout_label}
Apply this aesthetic consistently — colors, shapes, gradients, and layout patterns must reflect this style.

{extra}

Instructions:
- Slide 01: cover (title + subtitle + decorative element using primary color {primary})
- Slides 02 to {slides_count_minus_1}: content slides (1 key idea per slide, strong visual hierarchy)
- Slide {slides_count_padded}: closing (thank you / call to action)
- Use SVG <rect>, <text>, <line>, <path>, <g> — no <image> tags (no external assets)
- Apply colors strictly from spec_lock.md
- Output ONLY the JSON array — no prose, no markdown.
""".strip()

# ─────────────────────────────────────────────
# LLM CALLERS
# ─────────────────────────────────────────────

def _claude_call(system: str, user: str, max_tokens: int = 8192) -> str:
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    full_response = ""
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
    ) as stream:
        for text in stream.text_stream:
            full_response += text
    return full_response


async def _mistral_call(system: str, user: str, max_tokens: int = 8192) -> str:
    key = os.environ.get("MISTRAL_API_KEY", "")
    if not key:
        raise HTTPException(500, "MISTRAL_API_KEY not set")
    async with httpx.AsyncClient(timeout=300) as client:
        r = await client.post(
            "https://api.mistral.ai/v1/chat/completions",
            headers={"Authorization": f"Bearer {key}"},
            json={
                "model": "mistral-large-latest",
                "max_tokens": max_tokens,
                "messages": [
                    {"role": "system", "content": system},
                    {"role": "user",   "content": user},
                ],
            },
        )
        r.raise_for_status()
        return r.json()["choices"][0]["message"]["content"]


async def _llm(provider: str, system: str, user: str, max_tokens: int = 8192) -> str:
    if provider == "claude":
        # Run blocking Claude call in thread pool so polling requests stay responsive
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(
            None, functools.partial(_claude_call, system, user, max_tokens)
        )
    return await _mistral_call(system, user, max_tokens)

# ─────────────────────────────────────────────
# RESPONSE PARSERS
# ─────────────────────────────────────────────

def _extract_json_safe(raw: str) -> str:
    """Strip markdown fences and extract the outermost JSON object or array."""
    cleaned = re.sub(r"^```json\s*", "", raw.strip(), flags=re.MULTILINE)
    cleaned = re.sub(r"\s*```\s*$", "", cleaned, flags=re.MULTILINE)
    for start_char, end_char in [('{', '}'), ('[', ']')]:
        start = cleaned.find(start_char)
        end   = cleaned.rfind(end_char)
        if start != -1 and end != -1 and end > start:
            candidate = cleaned[start:end + 1]
            try:
                json.loads(candidate)
                return candidate
            except json.JSONDecodeError:
                continue
    raise ValueError("No valid JSON object found in response")


def _parse_strategist(raw: str) -> tuple[str, str]:
    """Parse Strategist response using DESIGN_SPEC / SPEC_LOCK text markers.

    Tolerant: missing END markers fall back to the next START or end-of-string.
    """
    ds_start = raw.find("DESIGN_SPEC_START")
    sl_start = raw.find("SPEC_LOCK_START")

    if ds_start == -1 or sl_start == -1:
        raise ValueError(
            f"Strategist markers not found. Raw (first 400 chars): {raw[:400]}"
        )

    ds_end = raw.find("DESIGN_SPEC_END")
    if ds_end == -1:
        ds_end = sl_start  # fallback: everything before SPEC_LOCK_START

    sl_end = raw.find("SPEC_LOCK_END")
    if sl_end == -1:
        sl_end = len(raw)  # fallback: to end of string

    design_spec = raw[ds_start + len("DESIGN_SPEC_START"):ds_end].strip()
    spec_lock   = raw[sl_start + len("SPEC_LOCK_START"):sl_end].strip()

    if not design_spec or not spec_lock:
        raise ValueError(
            f"Empty design_spec or spec_lock after parsing. Raw (first 400 chars): {raw[:400]}"
        )

    return design_spec, spec_lock


def _parse_executor(raw: str) -> dict[str, str]:
    """Parse Executor JSON array of {path, content} objects."""
    data = json.loads(_extract_json_safe(raw))
    if not isinstance(data, list):
        raise ValueError("Expected a JSON array of {path, content} objects")
    return {item["path"]: item["content"] for item in data}

# ─────────────────────────────────────────────
# PROJECT HELPERS
# ─────────────────────────────────────────────

def _write(project_dir: Path, rel: str, content: str) -> None:
    path = project_dir / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    logger.info("wrote %s", path.relative_to(project_dir))


def _run_script(script: str, project_dir: Path) -> None:
    cmd = [sys.executable, str(SCRIPTS_DIR / script), str(project_dir)]
    logger.info("run: %s", " ".join(cmd))
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=str(SCRIPTS_DIR))
    if r.stdout:
        logger.info("[%s stdout] %s", script, r.stdout[:500])
    if r.returncode != 0:
        logger.warning("[%s stderr] %s", script, r.stderr[:500])


async def _run_async(script: str, project_dir: Path) -> None:
    loop = asyncio.get_running_loop()
    await loop.run_in_executor(None, functools.partial(_run_script, script, project_dir))

# ─────────────────────────────────────────────
# SUPABASE HELPERS
# ─────────────────────────────────────────────

async def upload_to_supabase(pptx_bytes: bytes, tenant_id: str, title: str) -> str:
    safe_title = re.sub(r"[^\w\-]", "_", title)[:50]
    filename   = f"{tenant_id}/{int(time.time())}_{safe_title}.pptx"
    url        = f"{SUPABASE_URL}/storage/v1/object/studio-documents/{filename}"
    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(
            url,
            content=pptx_bytes,
            headers={
                "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
                "Content-Type":  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            },
        )
        resp.raise_for_status()
    return f"{SUPABASE_URL}/storage/v1/object/public/studio-documents/{filename}"


async def _sb_upsert_document(
    tenant_id: str,
    titre: str,
    pptx_url: str,
    svg_slides: list | None,
) -> str:
    payload = {
        "tenant_id":   tenant_id,
        "titre":       titre,
        "format":      "paysage",
        "design":      "ppt-master",
        "provider":    "claude",
        "pptx_url":    pptx_url,
        "html_output": json.dumps(svg_slides),
    }
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.post(
            f"{SUPABASE_URL}/rest/v1/studio_documents",
            headers={
                "apikey":        SUPABASE_SERVICE_KEY,
                "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
                "Content-Type":  "application/json",
                "Prefer":        "return=representation",
            },
            json=payload,
        )
        r.raise_for_status()
        return r.json()[0]["id"]

# ─────────────────────────────────────────────
# PIPELINE (background task)
# ─────────────────────────────────────────────

async def _pipeline(job_id: str, req: GenerateRequest) -> None:
    """
    Full async pipeline — runs as a background task.
    Progress updates are written to _jobs[job_id] for polling.
    """
    try:
        proj_name   = f"ppt_{req.tenant_id[:8]}_{job_id}"
        project_dir = PROJECTS_DIR / proj_name

        for sub in ["svg_output", "notes", "exports", "images", "svg_final"]:
            (project_dir / sub).mkdir(parents=True, exist_ok=True)

        logger.info("[%s] START provider=%s slides=%d layout=%s",
                    job_id, req.provider, req.slides_count, req.layout)

        n     = req.slides_count
        n_pad = str(n).zfill(2)
        extra = f"\nADDITIONAL INSTRUCTIONS:\n{req.prompt_injection}" if req.prompt_injection else ""

        # ── Layout reference ──────────────────────────────────
        layout_key    = req.layout or "free"
        layout_folder = LAYOUT_MAP.get(layout_key)
        layout_info   = _LAYOUTS_BY_ID.get(layout_key, _LAYOUTS[0])
        layout_label  = f"{layout_info['label']} — {layout_info['sublabel']}"

        design_example_section = ""
        if layout_folder:
            example_spec_path = EXAMPLES_DIR / layout_folder / "design_spec.md"
            if example_spec_path.exists():
                example_text = example_spec_path.read_text(encoding="utf-8")
                logger.info("[%s] design_spec_reference length: %d chars", job_id, len(example_text))
                _write(project_dir, "design_spec_reference.md", example_text)
                design_example_section = (
                    "\nVISUAL STYLE REFERENCE — reproduce this exact design language, "
                    "adapted to the new content below:\n"
                    f"{example_text}\n"
                )
            else:
                logger.warning("[%s] design_spec.md not found in %s", job_id, layout_folder)
                logger.info("[%s] Layout reference: %s", job_id, layout_folder)

        # ── Phase A : Strategist ──────────────────────────────
        _set_job(job_id, step=0, progress=5)

        sys_a = _STRATEGIST_SYSTEM.format(
            spec_lock_ref=_skill("templates/spec_lock_reference.md",   4000),
            shared_standards=_skill("references/shared-standards.md", 3000),
            design_example_section=design_example_section,
        )
        usr_a = _STRATEGIST_USER.format(
            content=req.content,
            slides_count=n,
            style=req.style,
            primary=req.palette.primary,
            secondary=req.palette.secondary,
            accent=req.palette.accent,
            extra=extra,
        )

        logger.info("[%s] Phase A — Strategist", job_id)
        raw_a = await _llm(req.provider, sys_a, usr_a, max_tokens=4096)
        design_spec_md, spec_lock_md = _parse_strategist(raw_a)
        _write(project_dir, "design_spec.md", design_spec_md)
        _write(project_dir, "spec_lock.md",   spec_lock_md)
        _set_job(job_id, step=1, progress=30)

        # ── Phase B : Executor ────────────────────────────────
        sys_b = _EXECUTOR_SYSTEM.format(
            executor_base=_skill("references/executor-base.md", 3000),
            shared_standards=_skill("references/shared-standards.md", 2000),
            slides_count_padded=n_pad,
        )
        usr_b = _EXECUTOR_USER.format(
            slides_count=n,
            slides_count_minus_1=n - 1,
            slides_count_padded=n_pad,
            design_spec_md=design_spec_md,
            spec_lock_md=spec_lock_md,
            content=req.content,
            primary=req.palette.primary,
            layout_label=layout_label,
            extra=extra,
        )

        logger.info("[%s] Phase B — Executor", job_id)
        raw_b = await _llm(req.provider, sys_b, usr_b, max_tokens=32768)
        svg_files = _parse_executor(raw_b)

        if not svg_files:
            raise ValueError("Executor produced no SVG files")

        for rel_path, content in svg_files.items():
            _write(project_dir, rel_path, content)
        logger.info("[%s] %d SVG files written", job_id, len(svg_files))
        _set_job(job_id, step=2, progress=80)

        # ── Phase C : Scripts ─────────────────────────────────
        notes_total = project_dir / "notes" / "total.md"
        if not notes_total.exists():
            notes_total.write_text("", encoding="utf-8")

        await _run_async("total_md_split.py", project_dir)
        await _run_async("finalize_svg.py",   project_dir)
        await _run_async("svg_to_pptx.py",    project_dir)
        _set_job(job_id, step=3, progress=90)

        # ── Collect SVG slides (svg_output first, then svg_final) ────────────
        svg_slides: list[str] = []
        for _svg_dir_name in ("svg_output", "svg_final"):
            _svg_dir   = project_dir / _svg_dir_name
            _svg_files = sorted(_svg_dir.glob("*.svg")) if _svg_dir.exists() else []
            if _svg_files:
                svg_slides = [p.read_text(encoding="utf-8") for p in _svg_files]
                logger.info("[%s] %d SVG slides from %s", job_id, len(svg_slides), _svg_dir_name)
                break
        if not svg_slides:
            logger.warning("[%s] No SVG slides found in svg_output or svg_final", job_id)

        # ── Find PPTX ─────────────────────────────────────────
        exports = sorted(
            (project_dir / "exports").glob("*.pptx"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not exports:
            raise ValueError("svg_to_pptx.py produced no PPTX")

        pptx         = exports[0]
        pptx_bytes  = pptx.read_bytes()
        pptx_b64    = base64.b64encode(pptx_bytes).decode("utf-8")

        title       = req.title or req.content[:60].split('\n')[0].strip() or f"Présentation {job_id}"
        pptx_url    = ""
        document_id = ""

        if SUPABASE_URL and SUPABASE_SERVICE_KEY:
            try:
                pptx_url    = await upload_to_supabase(pptx_bytes, req.tenant_id, title)
                document_id = await _sb_upsert_document(
                    tenant_id=req.tenant_id,
                    titre=title,
                    pptx_url=pptx_url,
                    svg_slides=svg_slides,
                )
                logger.info("[%s] Supabase OK → doc=%s", job_id, document_id)
            except Exception as e:
                logger.warning("[%s] Supabase storage/DB error: %s", job_id, e)

        _set_job(job_id,
                 status="done",
                 step=3,
                 progress=100,
                 result={
                     "status":      "done",
                     "pptx_base64": pptx_b64,
                     "pptx_url":    pptx_url,
                     "document_id": document_id,
                     "svg_slides":  svg_slides,
                 })
        logger.info("[%s] DONE", job_id)

    except Exception as e:
        logger.error("[%s] Pipeline error: %s", job_id, e, exc_info=True)
        _set_job(job_id, status="error", error=str(e))

# ─────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────

@app.post("/generate-pptx")
async def start_generate(req: GenerateRequest):
    """Start async pipeline — returns job_id immediately for polling."""
    job_id = uuid.uuid4().hex[:8]
    _save_job(job_id, {
        "status":   "running",
        "step":     0,
        "progress": 2,
        "result":   None,
        "error":    None,
    })
    asyncio.create_task(_pipeline(job_id, req))
    logger.info("[%s] Job created, background task started", job_id)
    return {"job_id": job_id}


@app.get("/generate-pptx/{job_id}")
async def poll_generate(job_id: str):
    """Poll job status. Returns step (0-3), progress (0-100), status, result."""
    job = _load_jobs().get(job_id)
    if not job:
        raise HTTPException(404, f"Job {job_id!r} not found")
    return job

# ─────────────────────────────────────────────
# HEALTH
# ─────────────────────────────────────────────

@app.get("/")
async def root():
    return {
        "service":   "traitement_pptmaster",
        "status":    "ok",
        "version":   "1.0.0",
        "endpoints": ["GET /layouts", "POST /generate-pptx", "GET /generate-pptx/{job_id}"],
    }


@app.get("/health")
async def health():
    return {"status": "ok"}
