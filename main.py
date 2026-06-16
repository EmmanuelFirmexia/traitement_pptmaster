"""
main.py — traitement_pptmaster
FastAPI service wrapping the PPT Master skill pipeline.

Pipeline (headless, non-interactive):
  Phase A – Strategist LLM call  → design_spec.md + spec_lock.md
  Phase B – Executor  LLM call   → SVG slides (svg_output/)
  Phase C – Scripts              → finalize_svg.py → svg_to_pptx.py → PPTX
"""

import json
import logging
import os
import re
import subprocess
import sys
import uuid
from pathlib import Path
from typing import Literal

import anthropic
import httpx
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

load_dotenv()

# ─────────────────────────────────────────────
# PATHS
# ─────────────────────────────────────────────
SKILL_DIR   = Path(__file__).parent / "skills" / "ppt-master"
SCRIPTS_DIR = SKILL_DIR / "scripts"
PROJECTS_DIR = Path(__file__).parent / "projects"
PROJECTS_DIR.mkdir(exist_ok=True)

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
    slides_count:     int = 10
    tenant_id:        str
    provider:         Literal["claude", "mistral"] = "claude"

# ─────────────────────────────────────────────
# APP
# ─────────────────────────────────────────────
app = FastAPI(
    title="PPT Master API",
    description="Génération PPTX via le pipeline SVG PPT Master",
    version="1.0.0",
)

ALLOWED_ORIGINS = [
    "https://app-leadpme.lovable.app",
    "https://leadpme.firmexia.com",
    "https://*.lovable.app",
    "https://*.lovableproject.com",
    "http://localhost:3000",
    "http://localhost:5173",
]
extra = os.environ.get("EXTRA_CORS_ORIGIN")
if extra:
    ALLOWED_ORIGINS.append(extra)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_origin_regex=r"https://.*\.(lovable\.app|lovableproject\.com|firmexia\.com|railway\.app)$",
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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

REFERENCE — spec_lock.md skeleton (follow this EXACTLY):
{spec_lock_ref}

REFERENCE — shared technical standards (SVG/PPTX rules):
{shared_standards}

OUTPUT FORMAT — return exactly this JSON object and nothing else:
{{
  "design_spec_md": "<full content of design_spec.md>",
  "spec_lock_md":   "<full content of spec_lock.md>"
}}
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
- Output ONLY the JSON object — no prose, no markdown fences.
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
    msg = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    return msg.content[0].text


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
        return _claude_call(system, user, max_tokens)
    return await _mistral_call(system, user, max_tokens)

# ─────────────────────────────────────────────
# RESPONSE PARSERS
# ─────────────────────────────────────────────

def _extract_json(text: str) -> str:
    """Strip prose / fences around a JSON block."""
    text = text.strip()
    # Remove ```json ... ``` fences
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"\s*```\s*$", "", text, flags=re.MULTILINE)
    # Find the outermost { } or [ ]
    for start_char, end_char in [('{', '}'), ('[', ']')]:
        start = text.find(start_char)
        end   = text.rfind(end_char)
        if start != -1 and end != -1 and end > start:
            candidate = text[start:end + 1]
            try:
                json.loads(candidate)
                return candidate
            except json.JSONDecodeError:
                continue
    return text


def _parse_strategist(raw: str) -> tuple[str, str]:
    data = json.loads(_extract_json(raw))
    return data["design_spec_md"], data["spec_lock_md"]


def _parse_executor(raw: str) -> dict[str, str]:
    data = json.loads(_extract_json(raw))
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


def _run(script: str, project_dir: Path) -> None:
    cmd = [sys.executable, str(SCRIPTS_DIR / script), str(project_dir)]
    logger.info("run: %s", " ".join(cmd))
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=str(SCRIPTS_DIR))
    if r.stdout:
        logger.info("[%s stdout] %s", script, r.stdout[:500])
    if r.returncode != 0:
        logger.warning("[%s stderr] %s", script, r.stderr[:500])

# ─────────────────────────────────────────────
# ROUTE — POST /generate-pptx
# ─────────────────────────────────────────────

@app.post("/generate-pptx")
async def generate_pptx(req: GenerateRequest):
    """
    Full PPT Master pipeline:
      A → Strategist LLM  → design_spec.md + spec_lock.md
      B → Executor  LLM  → SVG slides
      C → Scripts        → finalize_svg.py + svg_to_pptx.py → PPTX
    """
    job_id      = uuid.uuid4().hex[:8]
    proj_name   = f"ppt_{req.tenant_id[:8]}_{job_id}"
    project_dir = PROJECTS_DIR / proj_name

    for sub in ["svg_output", "notes", "exports", "images", "svg_final"]:
        (project_dir / sub).mkdir(parents=True, exist_ok=True)

    logger.info("[%s] START provider=%s slides=%d", job_id, req.provider, req.slides_count)

    n      = req.slides_count
    n_pad  = str(n).zfill(2)
    extra  = f"\nADDITIONAL INSTRUCTIONS:\n{req.prompt_injection}" if req.prompt_injection else ""

    # ── PHASE A : STRATEGIST ──────────────────────────────────
    sys_a = _STRATEGIST_SYSTEM.format(
        spec_lock_ref=_skill("templates/spec_lock_reference.md",   4000),
        shared_standards=_skill("references/shared-standards.md", 3000),
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
    try:
        raw_a = await _llm(req.provider, sys_a, usr_a, max_tokens=4096)
        design_spec_md, spec_lock_md = _parse_strategist(raw_a)
    except Exception as e:
        logger.error("[%s] Strategist error: %s\nRaw: %s", job_id, e, raw_a[:400] if 'raw_a' in dir() else "")
        raise HTTPException(502, f"Strategist LLM error: {e}")

    _write(project_dir, "design_spec.md", design_spec_md)
    _write(project_dir, "spec_lock.md",   spec_lock_md)

    # ── PHASE B : EXECUTOR ────────────────────────────────────
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
        extra=extra,
    )

    logger.info("[%s] Phase B — Executor", job_id)
    try:
        raw_b = await _llm(req.provider, sys_b, usr_b, max_tokens=32768)
        svg_files = _parse_executor(raw_b)
    except Exception as e:
        logger.error("[%s] Executor error: %s", job_id, e)
        raise HTTPException(502, f"Executor LLM error: {e}")

    if not svg_files:
        raise HTTPException(502, "Executor produced no SVG files")

    for rel_path, content in svg_files.items():
        _write(project_dir, rel_path, content)
    logger.info("[%s] %d SVG files written", job_id, len(svg_files))

    # ── PHASE C : SCRIPTS ─────────────────────────────────────
    # Speaker notes stub
    notes_total = project_dir / "notes" / "total.md"
    if not notes_total.exists():
        notes_total.write_text("", encoding="utf-8")

    _run("total_md_split.py", project_dir)
    _run("finalize_svg.py",   project_dir)
    _run("svg_to_pptx.py",   project_dir)

    # ── FIND PPTX ─────────────────────────────────────────────
    exports = sorted(
        (project_dir / "exports").glob("*.pptx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not exports:
        raise HTTPException(500, "svg_to_pptx.py produced no PPTX")

    pptx = exports[0]
    logger.info("[%s] DONE → %s", job_id, pptx.name)

    return FileResponse(
        path=str(pptx),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=f"presentation_{job_id}.pptx",
    )

# ─────────────────────────────────────────────
# HEALTH
# ─────────────────────────────────────────────

@app.get("/")
async def root():
    return {
        "service": "traitement_pptmaster",
        "status":  "ok",
        "version": "1.0.0",
        "endpoints": ["POST /generate-pptx"],
    }


@app.get("/health")
async def health():
    return {"status": "ok"}
