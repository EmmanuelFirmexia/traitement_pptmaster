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
import shutil
import subprocess
import sys
import time
import uuid
from pathlib import Path
from typing import Literal, Optional

import anthropic
import httpx
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from strict_parser import parse_strict_content

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
    content:            str
    prompt_injection:   str = ""
    style:              str = "professional"
    palette:            Palette
    tenant_id:          str
    title:              str = ""
    layout:             str = "free"
    provider:           Literal["claude", "mistral"] = "mistral"
    content_mode:       str = "marketing"       # "marketing" | "strict"
    document_type:      str = "presentation"    # "presentation" | "report" | "diagnostic"
    # ── NOUVEAU : options utilisateur ──────────────────────────────────────
    titre:              str = ""               # alias frontend pour title
    palette_key:        str = "theme"          # "theme" | "tenant" | "neutral"
    include_cta:        bool = False            # Ajouter une slide CTA finale
    target_slide_count: Optional[int] = None   # None = adaptatif, int = contrainte exacte
    document_id:        Optional[str] = None   # présent = mode édition (UPDATE)
    mistral_api_key:    Optional[str] = None   # clé Mistral/Scaleway propre au tenant (sinon env)

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
    version="1.1.0",
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
# PHASE A — STRATEGIST PROMPTS
# ─────────────────────────────────────────────

# ── Mode instructions ─────────────────────────────────────────────────────────

CONTENT_MODE_INSTRUCTIONS: dict[str, str] = {
    "strict": """
## MODE DE GÉNÉRATION : STRICT — INVIOLABLE

Tu es un transcripteur fidèle, pas un marketeur.
Les règles suivantes sont absolues et priment sur toute autre instruction.

### CE QUI EST INTERDIT (violation = échec critique) :
- Inventer du texte, des slogans, des titres, des transitions non présents dans la source
- Ajouter des chiffres, statistiques, pourcentages non présents dans la source
- Ajouter des URLs, numéros de téléphone, adresses email non présents dans la source
- Ajouter des témoignages, noms de clients, logos, certifications non présents dans la source
- Formuler des promesses ou garanties non présentes dans la source
- Appliquer une structure narrative forcée (hook/problem/solution/CTA) si le contenu ne la supporte pas
- Créer des slides pour remplir un quota — chaque slide doit avoir du contenu source réel

### CE QUI EST AUTORISÉ :
- Découper le contenu source en slides logiques (1 idée principale par slide)
- Créer une slide de couverture si un titre est identifiable dans la source
- Réorganiser l'ordre des idées pour la lisibilité (sans altérer le sens)
- Fusionner deux idées courtes sur une même slide si < 25 mots chacune et même thème

### RÈGLE DU NOMBRE DE SLIDES :
- Nombre de slides = nombre d'idées distinctes dans la source
- Si TARGET_SLIDE_COUNT est fourni : respecter exactement, fusionner ou découper en conséquence
- Si TARGET_SLIDE_COUNT est null : adaptatif, aucun plafond artificiel
- 1 slide par bloc de contenu substantiel (80-120 mots indicatif)
- Ne jamais créer une slide vide ou avec du contenu inventé pour atteindre un nombre

### RÈGLE DU CTA :
- Slide CTA générée UNIQUEMENT si : include_cta=true ET un CTA explicite existe dans la source
- Si include_cta=true mais aucun CTA dans la source : ne pas créer de slide CTA, le signaler dans spec_lock.md
- Si include_cta=false : pas de slide CTA, quelle que soit la source
- Interdit : inventer une URL, un numéro, un email même si include_cta=true
""".strip(),

    "marketing": """
## MODE DE GÉNÉRATION : MARKETING — CRÉATIF MAÎTRISÉ

Tu es un consultant marketing B2B qui restructure le contenu pour maximiser l'impact commercial.
Tu t'appuies sur les 8 dimensions de l'ADN marketing pour structurer la présentation.

### LES 8 DIMENSIONS ADN MARKETING :
1. Problème / Douleur — Quel problème concret ? Coût de l'inaction ?
2. Cible & Périmètre — Pour qui exactement ? Qui est exclu ?
3. Différenciateur unique — Pourquoi nous plutôt qu'un autre ?
4. Modalités d'engagement & Format — Comment ça marche concrètement ?
5. Facilité d'adoption — Pourquoi c'est simple de démarrer ?
6. Résultats attendus — Ce qu'on obtient, chiffré si possible
7. Bénéfice client — La transformation vécue (avant / après)
8. Preuve sociale & Réassurance — Qui nous fait confiance ? Quelles garanties ?

### STRUCTURE NARRATIVE RECOMMANDÉE :
Hook → Problem → Solution → Proof → Adoption → [CTA]
Correspondance :
- Hook      → Bénéfice client + Cible
- Problem   → Problème / Douleur + Coût de l'inaction
- Solution  → Différenciateur + Modalités d'engagement
- Proof     → Preuve sociale + Réassurance (AVANT les bénéfices — les PME françaises ont besoin d'être rassurées avant d'écouter les promesses)
- Adoption  → Facilité d'adoption + Résultats attendus
- CTA       → Option utilisateur uniquement (voir règle CTA)

### CE QUI EST AUTORISÉ :
- Réécrire pour plus d'impact commercial (ton B2B, direct, vouvoiement)
- Réorganiser le contenu selon la structure narrative
- Déduire le problème depuis le bénéfice (ex: si "gain de temps" → problème implicite "perte de temps")
- Reformuler qualitativement (ex: "un accompagnement simple" si la source dit "déploiement rapide")
- Fusionner des dimensions faibles avec des dimensions adjacentes
- CTA générique sans coordonnées si include_cta=true et aucun contact dans la source

### CE QUI EST INTERDIT — MÊME EN MODE MARKETING :
- Inventer des chiffres, statistiques, pourcentages (ex: "réduction de 37%" sans source)
- Inventer des URLs, numéros de téléphone, adresses email
- Inventer des témoignages, noms de clients, certifications
- Inventer des résultats chiffrés non présents dans la source
- Inventer des données sectorielles génériques (ex: "les PME perdent 15h/semaine")

### RÈGLE DU NOMBRE DE SLIDES :
- Si TARGET_SLIDE_COUNT est fourni par l'utilisateur : respecter exactement, sans plafond
- Si TARGET_SLIDE_COUNT est null : nombre adaptatif selon richesse du contenu :
    0-1 dimension riche  → 3 slides
    2-3 dimensions riches → 4-5 slides
    4-5 dimensions riches → 6 slides
    6+ dimensions riches  → 7-8 slides (conseil indicatif, pas de plafond dur)
  Note : au-delà de 8 slides, envisager de découper en plusieurs documents
- Une dimension est "riche" si ≥ 2 phrases substantielles ou données chiffrées dans la source

### RÈGLE DU CTA :
- Si include_cta=true ET source contient un CTA/contact → utiliser le texte exact de la source
- Si include_cta=true ET aucun contact dans la source → CTA générique : "Contactez-nous pour en savoir plus" (SANS URL ni numéro inventés)
- Si include_cta=false → pas de slide CTA

### GESTION DES DIMENSIONS ABSENTES :
Ordre de priorité si TARGET_SLIDE_COUNT force une réduction :
Solution > Bénéfice client > Différenciateur > Résultats > Facilité > Preuve > CTA
- Dimension absente → omettre la slide, ne jamais inventer
- Dimension pauvre → fusionner avec la dimension adjacente
- Jamais de slide vide
""".strip(),
}

# ── Document type instructions ────────────────────────────────────────────────

DOCUMENT_TYPE_INSTRUCTIONS: dict[str, str] = {
    "presentation": """
## TYPE DE DOCUMENT : PRÉSENTATION COMMERCIALE
- Structure narrative : Hook → Problem → Solution → Proof → Adoption → [CTA]
- Nombre de slides : adaptatif au contenu (voir règles du mode)
- Densité : visuel, minimal text par slide (max 40 mots par slide)
- Une idée clé par slide
- Ton : B2B, professionnel, direct, vouvoiement
- La slide CTA est une OPTION utilisateur — ne jamais l'imposer
""".strip(),

    "report": """
## TYPE DE DOCUMENT : RAPPORT ANALYTIQUE
- Structure : contexte → méthodologie → données → analyse → recommandations
- Nombre de slides : adaptatif, typiquement 8-15
- Densité : plus dense, tableaux et données encouragés
- Ton : factuel, précis
- Dernière slide : plan d'action priorisé (sauf si mode strict sans CTA source)
""".strip(),

    "diagnostic": """
## TYPE DE DOCUMENT : DIAGNOSTIC / AUDIT
- Structure : périmètre → état actuel → écarts → causes → recommandations → feuille de route
- Nombre de slides : adaptatif, typiquement 6-12
- Densité : findings quantifiés, niveaux de risque
- Ton : structuré, recommandations numérotées
- Dernière slide : quick wins vs actions long terme (sauf si mode strict sans CTA source)
""".strip(),
}

# ── Strategist system + user prompts ─────────────────────────────────────────

_STRATEGIST_SYSTEM_BASE = """
You are PPT Master acting as the STRATEGIST role.
You run in HEADLESS API MODE — skip all BLOCKING stops, browser UIs, and interactive confirmations.
Auto-approve all Eight Confirmations using the parameters provided.

Your task: produce ONLY design_spec.md and spec_lock.md for the requested presentation.

REFERENCE — spec_lock.md skeleton (follow this EXACTLY):
{spec_lock_ref}

REFERENCE — shared technical standards (SVG/PPTX rules):
{shared_standards}

## SPEC_LOCK.MD FORMAT — MANDATORY STRUCTURE
The spec_lock.md MUST begin with a CONFIGURATION section that the Executor reads
as its sole source of truth for slide count and CTA policy:

```
## CONFIGURATION
generation_mode: [STRICT|MARKETING]
slide_count: [exact integer — number of slides you will generate]
include_cta: [true|false]
cta_source: [explicit_source|generic_placeholder|none]
target_slide_count_requested: [integer|null]
```

Then one section per slide:
```
## SLIDE 01: [TYPE]
- Layout Type: [COVER|HOOK|PROBLEM|SOLUTION|PROOF|ADOPTION|BENEFITS|CTA|CONTENT|CLOSING]
- Title: [exact text]
- Content:
  * [bullet 1 — max 15 words]
  * [bullet 2 — max 15 words]
- Visual Hint: [brief SVG layout suggestion — e.g. "3 columns metrics", "left text right icon"]
- ADN Dimension: [dimension name or "none"]
```

IMPORTANT: The Executor reads slide_count from ## CONFIGURATION and generates
EXACTLY that number of slides. It does NOT add or remove slides on its own.

OUTPUT FORMAT — respond with EXACTLY this structure, no markdown around the blocks:
DESIGN_SPEC_START
[full content of design_spec.md as plain text]
DESIGN_SPEC_END
SPEC_LOCK_START
[full content of spec_lock.md as plain text — starting with ## CONFIGURATION]
SPEC_LOCK_END
""".strip()

_STRATEGIST_USER = """
Generate spec files for the following presentation:

CONTENT:
{content}
{content_lock_section}
CONFIRMED PARAMETERS:
- Canvas: PPT 16:9  →  viewBox 0 0 1280 720
- Style: {style}
- Palette:
    primary:   {primary}
    secondary: {secondary}
    accent:    {accent}
- CTA slide option: {include_cta}
- Target slide count: {target_slide_count}
{extra}

Rules:
- Lock colors: use primary={primary} as `primary`, secondary={secondary} as `secondary_accent`, accent={accent} as `accent`.
  Set `bg: #FFFFFF` unless the style clearly calls for a dark background.
- font_family: "Calibri", Arial, sans-serif
- body: 22, title: 40, subtitle: 28, annotation: 14
- Use the DESIGN_SPEC_START / DESIGN_SPEC_END and SPEC_LOCK_START / SPEC_LOCK_END markers exactly.
- The spec_lock.md MUST start with ## CONFIGURATION containing slide_count as an exact integer.
""".strip()

# ─────────────────────────────────────────────
# PHASE B — EXECUTOR PROMPTS
# ─────────────────────────────────────────────

_EXECUTOR_SYSTEM = """
You are PPT Master acting as the EXECUTOR role.
You run in HEADLESS API MODE — generate all SVG slides automatically without any confirmation.

REFERENCE — executor guidelines:
{executor_base}

REFERENCE — shared technical standards:
{shared_standards}

## YOUR SOLE SOURCE OF TRUTH: spec_lock.md
Read the ## CONFIGURATION section of spec_lock.md FIRST.
- slide_count: generate EXACTLY this number of SVG files — no more, no less
- include_cta: if false, DO NOT generate a CTA slide even if you think one is needed
- If spec_lock.md contains no CTA slide section, DO NOT generate one

## ABSOLUTE RULES — EXECUTOR HAS NO NARRATIVE AUTONOMY:
1. You are an execution engine, not a content creator
2. NEVER add slides not listed in spec_lock.md
3. NEVER remove slides listed in spec_lock.md
4. NEVER invent text, titles, slogans, URLs, phone numbers, emails, metrics
5. NEVER add a closing/CTA slide that is not explicitly defined in spec_lock.md
6. If a spec_lock.md slide section has empty content fields → generate a minimal visual slide with no text rather than inventing content
7. Any deviation from spec_lock.md is a critical execution error

## CRITICAL SVG RULES:
1. viewBox MUST be "0 0 1280 720" for every slide
2. All colors MUST come from spec_lock.md — no invented values
3. No external URLs (no href="http://...") — inline shapes / gradients only
4. Fonts: "Calibri", Arial, sans-serif — PPT-safe stacks only
5. Each SVG must be a complete, self-contained <svg> element
6. Slide files named: 01_cover.svg, 02_*.svg, …, NN_closing.svg (only if CTA is in spec)

OUTPUT FORMAT — return exactly this JSON array and nothing else:
[
  {{"path": "svg_output/01_cover.svg",   "content": "<svg viewBox=\\"0 0 1280 720\\" ...>...</svg>"}},
  {{"path": "svg_output/02_problem.svg", "content": "<svg ...>...</svg>"}},
  ...
]
""".strip()

_EXECUTOR_USER = """
Generate SVG slides based on these specs.

READ FIRST — CONFIGURATION from spec_lock.md:
{spec_lock_config_section}

DESIGN SPEC:
{design_spec_md}

FULL SPEC LOCK:
{spec_lock_md}

SOURCE CONTENT (reference only — do not add content not in spec_lock.md):
{content}
{content_lock_executor_note}
VISUAL STYLE: {layout_label}
Apply this aesthetic consistently — colors, shapes, gradients, and layout patterns must reflect this style.

{extra}

Instructions:
- Generate EXACTLY {slide_count} slides as specified in ## CONFIGURATION above
- Follow the slide plan in spec_lock.md exactly — slide types, titles, content, order
- Use SVG <rect>, <text>, <line>, <path>, <g> — no <image> tags
- Apply colors strictly from spec_lock.md
- For text layout: use <tspan> with explicit dy attributes for multi-line text (prevents overflow)
- Output ONLY the JSON array — no prose, no markdown
""".strip()

# ─────────────────────────────────────────────
# LLM CALLERS
# ─────────────────────────────────────────────

def _claude_call(system: str, user: str, max_tokens: int = 8192) -> tuple[str, int, int]:
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
        final             = stream.get_final_message()
        prompt_tokens     = final.usage.input_tokens
        completion_tokens = final.usage.output_tokens
    return full_response, prompt_tokens, completion_tokens


async def _mistral_call(system: str, user: str, max_tokens: int = 8192,
                        mistral_api_key: Optional[str] = None) -> tuple[str, int, int]:
    key = mistral_api_key or os.environ.get("SCALEWAY_API_KEY_MEDIUM", "")
    if not key:
        raise HTTPException(500, "SCALEWAY_API_KEY_MEDIUM not set")
    url = os.environ.get("SCALEWAY_API_URL", "https://api.scaleway.ai/v1/chat/completions")
    model_name = os.environ.get("MISTRAL_MODEL_MEDIUM", "mistral/mistral-medium-3.5-128b:fp8")
    async with httpx.AsyncClient(timeout=300) as client:
        logger.info("[mistral_call] payload: model=%s url=%s", model_name, url)
        r = await client.post(
            url,
            headers={"Authorization": f"Bearer {key}"},
            json={
                "model": model_name,
                "max_tokens": max_tokens,
                "messages": [
                    {"role": "system", "content": system},
                    {"role": "user",   "content": user},
                ],
            },
        )
        if r.status_code >= 400:
            logger.error("[mistral_call] error body: %s", r.text)
        r.raise_for_status()
        data = r.json()
        usage = data.get("usage", {})
        return (
            data["choices"][0]["message"]["content"],
            usage.get("prompt_tokens", 0),
            usage.get("completion_tokens", 0),
        )


# Provider réellement appelé → (label provider, nom du modèle) pour le logging usage
def _provider_meta(provider: str) -> tuple[str, str]:
    if provider == "mistral":
        return "mistral", "mistral-medium-3.5-128b"
    return "anthropic", "claude-sonnet-4-6"


async def _llm(provider: str, system: str, user: str, max_tokens: int = 8192,
               mistral_api_key: Optional[str] = None) -> tuple[str, int, int]:
    if provider == "claude":
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(
            None, functools.partial(_claude_call, system, user, max_tokens)
        )
    return await _mistral_call(system, user, max_tokens, mistral_api_key=mistral_api_key)


async def _log_ai_usage(tenant_id: str, action: str, model: str,
                        prompt_tokens: int, completion_tokens: int,
                        provider: str = "anthropic") -> None:
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        return
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            await client.post(
                f"{SUPABASE_URL}/rest/v1/tenant_ai_usage",
                headers={
                    "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
                    "apikey":        SUPABASE_SERVICE_KEY,
                    "Content-Type":  "application/json",
                },
                json={
                    "tenant_id":         tenant_id,
                    "action_type":       action,
                    "model":             model,
                    "prompt_tokens":     prompt_tokens,
                    "completion_tokens": completion_tokens,
                    "total_tokens":      prompt_tokens + completion_tokens,
                    "provider":          provider,
                },
            )
    except Exception as e:
        logger.warning("AI usage log failed: %s", e)

# ─────────────────────────────────────────────
# RESPONSE PARSERS
# ─────────────────────────────────────────────

def _extract_json_safe(raw: str) -> str:
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


def _strip_fences(text: str) -> str:
    text = re.sub(r"^```[a-zA-Z]*\r?\n?", "", text.lstrip())
    text = re.sub(r"\r?\n?```\s*$",        "", text.rstrip())
    return text.strip()


def _parse_strategist(raw: str) -> tuple[str, str]:
    ds_start = raw.find("DESIGN_SPEC_START")
    sl_start = raw.find("SPEC_LOCK_START")

    if ds_start == -1 or sl_start == -1:
        raise ValueError(
            f"Strategist markers not found. Raw (first 400 chars): {raw[:400]}"
        )

    if sl_start <= ds_start:
        raise ValueError(
            f"Marker order invalid: SPEC_LOCK_START (pos {sl_start}) "
            f"must appear after DESIGN_SPEC_START (pos {ds_start})"
        )

    ds_end = raw.find("DESIGN_SPEC_END")
    if ds_end == -1 or ds_end <= ds_start:
        ds_end = sl_start

    sl_end = raw.find("SPEC_LOCK_END")
    if sl_end == -1 or sl_end <= sl_start:
        sl_end = len(raw)

    design_spec = _strip_fences(raw[ds_start + len("DESIGN_SPEC_START"):ds_end])
    spec_lock   = _strip_fences(raw[sl_start + len("SPEC_LOCK_START"):sl_end])

    if not design_spec or not spec_lock:
        raise ValueError(
            f"Empty design_spec or spec_lock after parsing. Raw (first 400 chars): {raw[:400]}"
        )

    return design_spec, spec_lock


def _extract_spec_lock_config(spec_lock_md: str) -> tuple[int, bool]:
    """
    Extract slide_count and include_cta from the ## CONFIGURATION section
    of spec_lock.md. Falls back to counting ## SLIDE sections if not found.
    Returns (slide_count, include_cta).
    """
    slide_count = None
    include_cta = False

    # Try parsing ## CONFIGURATION section
    config_match = re.search(r"##\s*CONFIGURATION(.*?)(?=##|\Z)", spec_lock_md, re.DOTALL | re.IGNORECASE)
    if config_match:
        config_text = config_match.group(1)
        sc_match = re.search(r"slide_count\s*:\s*(\d+)", config_text)
        if sc_match:
            slide_count = int(sc_match.group(1))
        cta_match = re.search(r"include_cta\s*:\s*(true|false)", config_text, re.IGNORECASE)
        if cta_match:
            include_cta = cta_match.group(1).lower() == "true"

    # Fallback: count ## SLIDE sections
    if slide_count is None:
        slide_sections = re.findall(r"##\s*SLIDE\s+\d+", spec_lock_md, re.IGNORECASE)
        slide_count = len(slide_sections) if slide_sections else 6

    return slide_count, include_cta


def _parse_executor(raw: str) -> dict[str, str]:
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
    layout: str = "free",
    source_content: str = "",
    palette_key: str = "theme",
    include_cta: bool = False,
    target_slide_count: Optional[int] = None,
    gen_mode: str = "marketing",
    document_id: Optional[str] = None,
) -> str:
    payload = {
        "tenant_id":          tenant_id,
        "titre":              titre,
        "format":             "paysage",
        "design":             "ppt-master",
        "provider":           "claude",
        "pptx_url":           pptx_url,
        "html_output":        json.dumps(svg_slides),
        "layout":             layout,
        "source_content":     source_content,
        "palette_key":        palette_key,
        "include_cta":        include_cta,
        "target_slide_count": target_slide_count,
        "gen_mode":           gen_mode,
    }
    async with httpx.AsyncClient(timeout=30) as client:
        if document_id:
            update_payload = {k: v for k, v in payload.items() if k != "tenant_id"}
            r = await client.patch(
                f"{SUPABASE_URL}/rest/v1/studio_documents?id=eq.{document_id}",
                headers={
                    "apikey":        SUPABASE_SERVICE_KEY,
                    "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
                    "Content-Type":  "application/json",
                    "Prefer":        "return=representation",
                },
                json=update_payload,
            )
            r.raise_for_status()
            rows = r.json()
            return rows[0]["id"] if rows else document_id
        else:
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
    try:
        proj_name   = f"ppt_{req.tenant_id[:8]}_{job_id}"
        project_dir = PROJECTS_DIR / proj_name

        for sub in ["svg_output", "notes", "exports", "images", "svg_final"]:
            (project_dir / sub).mkdir(parents=True, exist_ok=True)

        # ── Phase 0 : Parser déterministe (mode strict uniquement) ──
        content_lock = None
        if req.content_mode == "strict":
            content_lock = parse_strict_content(req.content)
            logger.info("[%s] content_lock: %d slides parsées (mode strict)",
                        job_id, len(content_lock))
            _write(project_dir, "content_lock.json", json.dumps(content_lock, ensure_ascii=False, indent=2))

        logger.info(
            "[%s] START provider=%s layout=%s mode=%s include_cta=%s target_slides=%s",
            job_id, req.provider, req.layout, req.content_mode,
            req.include_cta, req.target_slide_count,
        )

        extra = f"\nADDITIONAL INSTRUCTIONS:\n{req.prompt_injection}" if req.prompt_injection else ""

        # ── Layout reference ──────────────────────────────────
        layout_key    = req.layout or "free"
        layout_folder = LAYOUT_MAP.get(layout_key)
        layout_info   = _LAYOUTS_BY_ID.get(layout_key, _LAYOUTS[0])
        layout_label  = f"{layout_info['label']} — {layout_info['sublabel']}"

        design_example_section = ""
        spec_lock_path: Path | None = None

        if layout_folder:
            example_dir       = EXAMPLES_DIR / layout_folder
            example_spec_path = example_dir / "design_spec.md"
            spec_lock_path    = example_dir / "spec_lock.md"

            if example_spec_path.exists():
                example_text = example_spec_path.read_text(encoding="utf-8")
                logger.info("[%s] design_spec_reference length: %d chars", job_id, len(example_text))
                _write(project_dir, "design_spec_reference.md", example_text)

                spec_lock_reference = ""
                if spec_lock_path.exists():
                    spec_lock_reference = spec_lock_path.read_text(encoding="utf-8")
                    logger.info("[%s] spec_lock_reference length: %d chars", job_id, len(spec_lock_reference))

                design_example_section = f"""
VISUAL STYLE REFERENCE — MANDATORY BINDING CONTRACT:
The following defines the EXACT visual language you MUST use.
Copy colors, fonts, layout patterns verbatim. Do NOT invent new colors or styles.

DESIGN SPEC (narrative reference):
{example_text}

SPEC LOCK (binding execution contract — copy exactly):
{spec_lock_reference}

Your output spec_lock.md MUST replicate these exact values:
- Same color palette (hex codes verbatim)
- Same typography (font families verbatim)
- Same layout patterns
- Same visual effects
Only the CONTENT (texts, data) changes. The STYLE is locked.
"""
            else:
                logger.warning("[%s] design_spec.md not found in %s", job_id, layout_folder)
                spec_lock_path = None

        # ── Phase A : Strategist ──────────────────────────────
        _set_job(job_id, step=0, progress=5)

        if req.content_mode == "strict" and content_lock:
            # ── Mode strict : Phase A entièrement ignorée (aucun appel LLM Strategist) ──
            # spec_lock_md est construit directement depuis content_lock.
            spec_lock_md     = json.dumps(content_lock, ensure_ascii=False, indent=2)
            design_spec_md   = ""
            spec_slide_count = len(content_lock)
            spec_include_cta = False
            _write(project_dir, "design_spec.md", design_spec_md)
            _write(project_dir, "spec_lock.md",   spec_lock_md)
            logger.info("[%s] Mode strict : Phase A ignorée, spec_lock = content_lock", job_id)
            logger.info("[%s] content_lock.json :\n%s", job_id, spec_lock_md)
        else:
            content_instructions = CONTENT_MODE_INSTRUCTIONS.get(
                req.content_mode, CONTENT_MODE_INSTRUCTIONS["marketing"]
            )
            type_instructions = DOCUMENT_TYPE_INSTRUCTIONS.get(
                req.document_type, DOCUMENT_TYPE_INSTRUCTIONS["presentation"]
            )

            sys_base = _STRATEGIST_SYSTEM_BASE.format(
                spec_lock_ref=_skill("templates/spec_lock_reference.md", 4000),
                shared_standards=_skill("references/shared-standards.md", 3000),
            )
            sys_a = "\n\n".join(filter(None, [
                sys_base,
                type_instructions,
                content_instructions,
                design_example_section or "",
            ]))

            # target_slide_count label for prompt
            if req.target_slide_count:
                target_label = f"{req.target_slide_count} slides (exact — user constraint)"
            else:
                target_label = "null (adaptive — determine from content richness)"

            usr_a = _STRATEGIST_USER.format(
                content=req.content,
                content_lock_section="",
                style=req.style,
                primary=req.palette.primary,
                secondary=req.palette.secondary,
                accent=req.palette.accent,
                include_cta=str(req.include_cta).lower(),
                target_slide_count=target_label,
                extra=extra,
            )

            logger.info(
                "[%s] Phase A — Strategist content_mode=%s document_type=%s include_cta=%s target_slides=%s",
                job_id, req.content_mode, req.document_type, req.include_cta, req.target_slide_count,
            )
            raw_a, pt_a, ct_a = await _llm(req.provider, sys_a, usr_a, max_tokens=4096,
                                           mistral_api_key=req.mistral_api_key)
            logger.info("[%s] Phase A tokens — prompt=%d completion=%d", job_id, pt_a, ct_a)
            _prov_label, _prov_model = _provider_meta(req.provider)
            await _log_ai_usage(req.tenant_id, "pptmaster_strategist", _prov_model, pt_a, ct_a, _prov_label)

            design_spec_md, spec_lock_md = _parse_strategist(raw_a)
            _write(project_dir, "design_spec.md", design_spec_md)
            _write(project_dir, "spec_lock.md",   spec_lock_md)

            # Extract slide_count and include_cta from spec_lock
            spec_slide_count, spec_include_cta = _extract_spec_lock_config(spec_lock_md)
            logger.info(
                "[%s] spec_lock parsed: slide_count=%d include_cta=%s",
                job_id, spec_slide_count, spec_include_cta,
            )

        _set_job(job_id, step=1, progress=30)

        # ── Phase B : Executor ────────────────────────────────
        # Override spec_lock.md with example's locked version for visual style
        if spec_lock_path is not None and spec_lock_path.exists() and req.content_mode != "strict":
            # Merge: keep CONFIGURATION from generated spec_lock, override style from example
            example_spec_lock = spec_lock_path.read_text(encoding="utf-8")
            shutil.copy(spec_lock_path, project_dir / "spec_lock.md")
            spec_lock_md = example_spec_lock
            logger.info("[%s] spec_lock overridden with example (%d chars)", job_id, len(spec_lock_md))
        else:
            logger.info("[%s] spec_lock kept from Strategist (mode=%s)", job_id, req.content_mode)

        # Extract config section for explicit injection into Executor user prompt
        config_section_match = re.search(
            r"(##\s*CONFIGURATION.*?)(?=##\s*SLIDE|\Z)", spec_lock_md, re.DOTALL | re.IGNORECASE
        )
        spec_lock_config_section = config_section_match.group(1).strip() if config_section_match else (
            f"## CONFIGURATION\ngeneration_mode: {req.content_mode.upper()}\n"
            f"slide_count: {spec_slide_count}\ninclude_cta: {str(spec_include_cta).lower()}"
        )

        sys_b = _EXECUTOR_SYSTEM.format(
            executor_base=_skill("references/executor-base.md", 3000),
            shared_standards=_skill("references/shared-standards.md", 2000),
        )

        content_lock_executor_note = (
            "\nMODE STRICT — Le texte exact de chaque slide est dans content_lock.json — "
            "ne modifier aucun mot.\n"
            if req.content_mode == "strict" else ""
        )

        usr_b = _EXECUTOR_USER.format(
            spec_lock_config_section=spec_lock_config_section,
            slide_count=spec_slide_count,
            design_spec_md=design_spec_md,
            spec_lock_md=spec_lock_md,
            content=req.content,
            content_lock_executor_note=content_lock_executor_note,
            primary=req.palette.primary,
            layout_label=layout_label,
            extra=extra,
        )

        logger.info("[%s] Phase B — Executor content_mode=%s slide_count=%d", job_id, req.content_mode, spec_slide_count)
        raw_b, pt_b, ct_b = await _llm(req.provider, sys_b, usr_b, max_tokens=32768,
                                       mistral_api_key=req.mistral_api_key)
        logger.info("[%s] Phase B tokens — prompt=%d completion=%d", job_id, pt_b, ct_b)
        _prov_label_b, _prov_model_b = _provider_meta(req.provider)
        await _log_ai_usage(req.tenant_id, "pptmaster_executor", _prov_model_b, pt_b, ct_b, _prov_label_b)

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

        # ── Collect SVG slides ────────────────────────────────
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

        pptx        = exports[0]
        pptx_bytes  = pptx.read_bytes()
        pptx_b64    = base64.b64encode(pptx_bytes).decode("utf-8")

        title       = req.titre or req.title or req.content[:60].split('\n')[0].strip() or f"Présentation {job_id}"
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
                    layout=req.layout or "free",
                    source_content=req.content,
                    palette_key=req.palette_key,
                    include_cta=req.include_cta,
                    target_slide_count=req.target_slide_count,
                    gen_mode=req.content_mode,
                    document_id=req.document_id,
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
                     "slide_count": spec_slide_count,
                 })
        logger.info("[%s] DONE — %d slides", job_id, spec_slide_count)

    except Exception as e:
        logger.error("[%s] Pipeline error: %s", job_id, e, exc_info=True)
        _set_job(job_id, status="error", error=str(e))

# ─────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────

@app.post("/generate-pptx")
async def start_generate(req: GenerateRequest):
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
        "version":   "1.1.0",
        "endpoints": ["GET /layouts", "POST /generate-pptx", "GET /generate-pptx/{job_id}", "POST /parse-strict"],
    }


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/parse-strict")
async def parse_strict_endpoint(data: dict):
    """
    Endpoint de debug — retourne le content_lock sans générer de PPTX.
    Body: { "content": "texte source..." }
    """
    from strict_parser import parse_strict_content
    content = data.get("content", "")
    if not content:
        return {"error": "content vide", "slides": []}

    slides = parse_strict_content(content)
    return {
        "slide_count": len(slides),
        "slides": slides
    }
