"""
strict_parser.py
─────────────────────────────────────────────────────────────────────────────
Parsing STRICT (zéro IA) du texte source en slides.

Les fonctions de découpe ci-dessous sont copiées depuis le repo
`EmmanuelFirmexia/leadpme-presentations` (main.py) — elles constituent le
moteur de parsing déterministe utilisé pour transformer du Markdown en blocs
puis en slides, sans jamais inventer un seul mot.

Fonctions copiées :
    - markdown_to_blocs()
    - proposer_slides()
    - eclater_blocs_mixtes()
    - fusionner_blocs_texte()
    - couper_texte_en_phrases()
    - _est_fin_de_phrase()
    - detecter_stat()
    - est_titrise()

Helpers nécessaires (dépendances des fonctions ci-dessus) :
    - md_inline_to_html()   (appelé par markdown_to_blocs)
    - SEUIL_CHARS, ABREVIATIONS (constantes)

Point d'entrée :
    - parse_strict_content(source_text) -> list[dict]
"""

import re
import unicodedata

from bs4 import BeautifulSoup


# ─────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────

SEUIL_CHARS = 600  # caractères max de texte par slide

ABREVIATIONS = {"etc", "ex", "cf", "art", "al", "fig", "vol", "p", "pp", "n°", "dr", "mr", "mme", "st"}


# ─────────────────────────────────────────
# DÉTECTION
# ─────────────────────────────────────────

def detecter_stat(texte: str) -> bool:
    return bool(re.match(r'^[\d\s,\.]+\s*[%€$KkMm]', texte)) and len(texte) < 20


def est_titrise(html: str) -> bool:
    """Retourne True si tout le contenu visible est en gras."""
    import unicodedata
    soup = BeautifulSoup(html, "html.parser")
    texte_brut = unicodedata.normalize("NFC", soup.get_text()).strip()
    if not texte_brut:
        return False
    strongs = soup.find_all("strong")
    texte_strong = unicodedata.normalize("NFC", "".join(s.get_text() for s in strongs)).strip()
    # Normaliser aussi les espaces pour éviter les faux négatifs
    return texte_strong.replace(" ", "") == texte_brut.replace(" ", "")


# ─────────────────────────────────────────
# DÉCOUPE EN BLOCS
# ─────────────────────────────────────────

def eclater_blocs_mixtes(blocs: list) -> list:
    """Si un bloc texte commence par du contenu entièrement gras suivi de texte normal,
    le couper en deux : un bloc titrisé + un bloc texte normal.
    """
    import unicodedata
    result = []
    for bloc in blocs:
        if bloc.get("type") != "texte":
            result.append(bloc)
            continue
        html = bloc.get("html", "")
        if not html or est_titrise(html):
            result.append(bloc)
            continue
        soup = BeautifulSoup(html, "html.parser")
        children = list(soup.children)
        titre_parts = []
        reste_start = 0
        for i, child in enumerate(children):
            tag = getattr(child, 'name', None)
            text = str(child)
            if tag == 'strong':
                titre_parts.append(child.get_text())
                reste_start = i + 1
            elif text.strip() in ('', '<br/>', '<br>'):
                if titre_parts:
                    reste_start = i + 1
                    break
            else:
                break
        if not titre_parts:
            result.append(bloc)
            continue
        titre_texte = unicodedata.normalize("NFC", "".join(titre_parts)).strip()
        reste_html = "".join(str(c) for c in children[reste_start:]).strip()
        if titre_texte:
            result.append({"type": "titre", "texte": titre_texte, "html": f"<strong>{titre_texte}</strong>"})
        if reste_html:
            reste_texte = BeautifulSoup(reste_html, "html.parser").get_text().strip()
            result.append({"type": "texte", "texte": reste_texte, "html": reste_html})
    return result


def fusionner_blocs_texte(blocs: list) -> list:
    """Fusionne les blocs 'texte' consecutifs NON-titrisés en un seul bloc.
    Un bloc titrisé (= titre) reste toujours isolé pour que le layout le détecte.
    """
    blocs_fusionnes = []
    buffer_texte = []
    buffer_html = []

    def flush():
        if buffer_texte:
            blocs_fusionnes.append({
                "type": "texte",
                "texte": "\n\n".join(buffer_texte),
                "html": "<br><br>".join(buffer_html)
            })
            buffer_texte.clear()
            buffer_html.clear()

    for bloc in blocs:
        if bloc["type"] == "texte":
            html = bloc.get("html", bloc.get("texte", ""))
            if est_titrise(html):
                # Titre : flush le buffer courant, puis ajouter le titre seul
                flush()
                blocs_fusionnes.append(bloc)
            else:
                buffer_texte.append(bloc.get("texte", ""))
                buffer_html.append(html)
        else:
            flush()
            blocs_fusionnes.append(bloc)

    flush()
    return blocs_fusionnes


def _est_fin_de_phrase(texte: str, pos: int) -> bool:
    """Retourne True si texte[pos] est un vrai point de fin de phrase (pas une abréviation)."""
    if texte[pos] not in ".!?":
        return False
    if texte[pos] in "!?":
        return True  # ! et ? sont toujours des fins de phrase
    # Vérifier si c'est une abréviation : regarder le mot précédent
    avant = texte[:pos].rstrip()
    mot_avant = re.split(r'\s+', avant)[-1].lower().rstrip('.')
    if mot_avant in ABREVIATIONS:
        return False
    # Vérifier si la lettre suivante est minuscule (pas une fin de phrase)
    apres = texte[pos+1:].lstrip()
    if apres and apres[0].islower():
        return False
    return True


def couper_texte_en_phrases(texte: str, html: str, seuil: int) -> list:
    """Découpe un texte long en morceaux de ~seuil caractères,
    en coupant à la fin de phrases (. ! ?) ou à défaut aux <br>."""
    if len(texte) <= seuil:
        return [{"texte": texte, "html": html}]

    # Chercher les positions de fin de phrase
    morceaux = []
    debut = 0
    while debut < len(texte):
        fin_ideale = debut + seuil
        if fin_ideale >= len(texte):
            morceaux.append(texte[debut:])
            break
        # Chercher en arrière depuis fin_ideale la fin de phrase la plus proche
        coupe = -1
        for pos in range(min(fin_ideale, len(texte) - 1), max(debut, fin_ideale - 200), -1):
            if _est_fin_de_phrase(texte, pos):
                coupe = pos + 1
                break
        if coupe == -1:
            # Pas de fin de phrase — couper au dernier espace
            for pos in range(fin_ideale, max(debut, fin_ideale - 100), -1):
                if pos < len(texte) and texte[pos] == " ":
                    coupe = pos
                    break
        if coupe == -1:
            coupe = fin_ideale
        morceaux.append(texte[debut:coupe].strip())
        debut = coupe

    # Reconstruire les blocs html correspondants (approx. — on recoupe le html pareil)
    result = []
    for i, morceau in enumerate(morceaux):
        if not morceau:
            continue
        # Pour le html : essayer de retrouver la portion correspondante
        # On cherche le texte brut dans le html
        result.append({
            "type": "texte",
            "texte": morceau,
            "html": morceau  # simplifié — le gras/italique sera perdu sur la coupure
        })
    return result


def proposer_slides(blocs: list) -> list:
    """Découpe intelligente en slides :
    - Un titre force une nouvelle slide
    - Si le texte cumulé d'une slide dépasse SEUIL_CHARS, on cherche la meilleure coupure :
      1. Entre deux blocs (coupure propre)
      2. À l'intérieur du bloc le plus long, en fin de phrase
    - Les blocs non-texte (tableau, image, stat) ne sont jamais découpés
    """
    if not blocs:
        return blocs

    # Première passe : éclater les blocs texte trop longs individuellement
    blocs_eclates = []
    for bloc in blocs:
        if bloc.get("type") == "texte" and not est_titrise(bloc.get("html", "")):
            texte = bloc.get("texte", "")
            if len(texte) > SEUIL_CHARS:
                morceaux = couper_texte_en_phrases(texte, bloc.get("html", texte), SEUIL_CHARS)
                blocs_eclates.extend(morceaux)
            else:
                blocs_eclates.append(bloc)
        else:
            blocs_eclates.append(bloc)

    # Deuxième passe : regrouper en slides en respectant le seuil
    slides_blocs = []  # liste de listes de blocs
    slide_courante = []
    chars_courants = 0

    for bloc in blocs_eclates:
        t = bloc.get("type", "texte")
        longueur = len(bloc.get("texte", bloc.get("valeur", "")))
        est_titre = t == "titre" or (t == "texte" and est_titrise(bloc.get("html", "")))

        # Un titre force une nouvelle slide (sauf si c'est le tout premier bloc)
        if est_titre and slide_courante:
            slides_blocs.append(slide_courante)
            slide_courante = []
            chars_courants = 0

        # Si ajouter ce bloc dépasse le seuil ET qu'il y a déjà du contenu → nouvelle slide
        if chars_courants + longueur > SEUIL_CHARS and slide_courante and t in ("texte",) and not est_titre:
            slides_blocs.append(slide_courante)
            slide_courante = []
            chars_courants = 0

        slide_courante.append(bloc)
        if t == "texte":
            chars_courants += longueur

    if slide_courante:
        slides_blocs.append(slide_courante)

    # Aplatir en marquant nouvelle_slide
    result = []
    for i, groupe in enumerate(slides_blocs):
        for j, bloc in enumerate(groupe):
            if i > 0 and j == 0:
                bloc["nouvelle_slide"] = True
            elif i == 0 and j == 0:
                bloc["nouvelle_slide"] = True
            result.append(bloc)

    return result


# ─────────────────────────────────────────
# MARKDOWN → BLOCS
# ─────────────────────────────────────────

def md_inline_to_html(text: str) -> str:
    """Convertit le markdown inline (**gras**, *italique*) en HTML."""
    import re
    text = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
    text = re.sub(r'__(.+?)__', r'<strong>\1</strong>', text)
    text = re.sub(r'\*(.+?)\*', r'<em>\1</em>', text)
    text = re.sub(r'_(.+?)_', r'<em>\1</em>', text)
    return text


def markdown_to_blocs(md_text: str, image_positions: dict) -> list:
    """Convertit du Markdown en blocs leadPME.
    Règles :
    - Texte en gras (**...**) ou heading (#) → titre
    - Liste à puces (- item) → point_cle (toujours, peu importe le nombre)
    - Textes normaux consécutifs → fusionnés en un seul bloc texte
    - Tableau Markdown → bloc tableau
    - Stat (38%, 1200€...) → bloc stat
    """
    import re
    blocs = []
    lines = md_text.split('\n')
    buffer_texte = []
    buffer_html = []

    def flush_buffer():
        if buffer_texte:
            texte = ' '.join(buffer_texte).strip()
            html = '<br>'.join(buffer_html).strip()
            if texte:
                blocs.append({"type": "texte", "texte": texte, "html": html})
            buffer_texte.clear()
            buffer_html.clear()

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Heading Markdown (# ## ###)
        heading_match = re.match(r'^(#{1,3})\s+(.+)', stripped)
        if heading_match:
            flush_buffer()
            texte = heading_match.group(2).strip()
            blocs.append({"type": "titre", "texte": texte, "html": f"<strong>{texte}</strong>"})
            i += 1
            continue

        # Texte entièrement en gras (**texte**) → titre
        bold_match = re.match(r'^\*\*(.+?)\*\*$', stripped)
        if bold_match:
            flush_buffer()
            texte = bold_match.group(1).strip()
            blocs.append({"type": "titre", "texte": texte, "html": f"<strong>{texte}</strong>"})
            i += 1
            continue

        # Tableau Markdown — lignes avec |
        if stripped.startswith('|'):
            flush_buffer()
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            data_lines = [l for l in table_lines if not re.match(r'^\|[-| :]+\|$', l)]
            if data_lines:
                colonnes = [c.strip() for c in data_lines[0].split('|') if c.strip()]
                lignes = []
                for tl in data_lines[1:]:
                    cells = [c.strip() for c in tl.split('|') if c.strip()]
                    if cells:
                        lignes.append(cells)
                if colonnes:
                    blocs.append({"type": "tableau", "colonnes": colonnes, "lignes": lignes})
            continue

        # Liste à puces → bloc texte avec <ul><li>...</li></ul>
        list_match = re.match(r'^[-*]\s+(.+)', stripped) or re.match(r'^\d+\.\s+(.+)', stripped)
        if list_match:
            flush_buffer()
            # Accumuler toutes les puces consécutives
            items = []
            while i < len(lines):
                l = lines[i].strip()
                m = re.match(r'^[-*]\s+(.+)', l) or re.match(r'^\d+\.\s+(.+)', l)
                if m:
                    item_texte = m.group(1).strip()
                    item_html = md_inline_to_html(item_texte)
                    if detecter_stat(item_texte):
                        # stat isolée dans la liste → flush la liste courante puis stat
                        if items:
                            texte_liste = " ".join(it[0] for it in items)
                            html_liste = "<ul>" + "".join(f"<li>{it[1]}</li>" for it in items) + "</ul>"
                            blocs.append({"type": "texte", "texte": texte_liste, "html": html_liste})
                            items = []
                        blocs.append({"type": "stat", "valeur": item_texte, "texte": item_texte, "label": ""})
                    else:
                        items.append((item_texte, item_html))
                    i += 1
                elif not l:  # ligne vide — on continue à chercher des puces
                    i += 1
                else:
                    break
            if items:
                texte_liste = " ".join(it[0] for it in items)
                html_liste = "<ul>" + "".join(f"<li>{it[1]}</li>" for it in items) + "</ul>"
                blocs.append({"type": "texte", "texte": texte_liste, "html": html_liste})
            continue

        # Ligne vide → on ne flush pas, les textes se fusionnent
        if not stripped:
            i += 1
            continue

        # Stat standalone
        if detecter_stat(stripped):
            flush_buffer()
            blocs.append({"type": "stat", "valeur": stripped, "texte": stripped, "label": ""})
            i += 1
            continue

        # Texte normal → accumuler
        html = md_inline_to_html(stripped)
        buffer_texte.append(stripped)
        buffer_html.append(html)
        i += 1

    flush_buffer()

    for url in image_positions.values():
        blocs.append({"type": "image", "url": url, "legende": ""})

    return blocs


# ─────────────────────────────────────────
# POINT D'ENTRÉE — PARSING STRICT (zéro IA)
# ─────────────────────────────────────────

def parse_strict_content(source_text: str) -> list[dict]:
    """
    Parse le texte source en slides sans aucune IA.
    Retourne une liste de dicts :
    [
      {"slide_index": 1, "type": "cover", "title": "...", "content": ["...", "..."]},
      {"slide_index": 2, "type": "content", "title": "...", "content": ["..."]},
      ...
    ]
    Zéro mot inventé — tout vient du texte source.
    """
    if not source_text or not source_text.strip():
        return []

    # 1. Markdown → blocs (titres, textes, listes, tableaux, stats)
    blocs = markdown_to_blocs(source_text, {})

    # 2. Éclater les blocs mixtes (gras en tête + texte normal → titre + texte)
    blocs = eclater_blocs_mixtes(blocs)

    # 3. Découpe en slides (titre force nouvelle slide, seuil de caractères respecté)
    blocs = proposer_slides(blocs)

    # 4. Regrouper les blocs en slides via le marqueur `nouvelle_slide`
    groupes: list[list[dict]] = []
    courant: list[dict] = []
    for bloc in blocs:
        if bloc.get("nouvelle_slide") and courant:
            groupes.append(courant)
            courant = []
        courant.append(bloc)
    if courant:
        groupes.append(courant)

    # 5. Transformer chaque groupe en slide structurée
    slides: list[dict] = []
    for idx, groupe in enumerate(groupes, start=1):
        title = ""
        content: list[str] = []
        for bloc in groupe:
            t = bloc.get("type", "texte")
            est_titre = t == "titre" or (t == "texte" and est_titrise(bloc.get("html", "")))
            if est_titre and not title:
                # Premier titre du groupe → titre de la slide
                title = (bloc.get("texte", "") or "").strip()
            else:
                # Tout le reste → contenu (texte, stat, tableau, image…)
                valeur = (bloc.get("texte") or bloc.get("valeur") or "").strip()
                if valeur:
                    content.append(valeur)
        slides.append({
            "slide_index": idx,
            "type": "cover" if idx == 1 else "content",
            "title": title,
            "content": content,
        })

    return slides
