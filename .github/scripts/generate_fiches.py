"""
Génère les fiches animateur PPTX pour chaque atelier analysé.
Lit slide-plans.json + assets/pptx/KEY.pptx → assets/fiches/KEY-fiche.pptx
"""
import json, re
from pathlib import Path
from pptx import Presentation
from pptx.oxml.ns import qn

ROOT        = Path(__file__).parent.parent.parent
PPTX_DIR    = ROOT / "assets" / "pptx"
FICHES_DIR  = ROOT / "assets" / "fiches"
TEMPLATE    = ROOT / "assets" / "fiche-template.pptx"
PLANS_FILE  = ROOT / "assets" / "slide-plans.json"

MATERIEL_DEFAULT = [
    "1 PC ou tablette par participant",
    "Connexion Wi-Fi stable",
    "Grand écran / vidéoprojecteur",
    "Fiches mémo (1 par participant)",
]

CONSEILS_DEFAULT = [
    "Adapter le rythme selon le niveau du groupe.",
    "Laisser les participants manipuler dès que possible.",
    "Prévoir un compte de démonstration pour ne pas exposer de données personnelles.",
    "Terminer en distribuant la fiche mémo — c'est le support qu'ils garderont.",
]


def clean(text):
    """Supprime emojis et caractères parasites."""
    return re.sub(r'[^\x00-\x7FÀ-ɏ‘’–—«»]', '', text).strip()


def trunc(t, n):
    return t[:n] + "…" if len(t) > n else t


def get_slide_title(slide):
    if slide.shapes.title:
        t = slide.shapes.title.text.strip()
        if t: return t
    for shape in slide.shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if t and len(t) < 120:
                return t.split("\n")[0].strip()
    return "(sans titre)"


def get_first_bullet(slide, title):
    """Retourne la première phrase de contenu d'une slide (hors titre)."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            t = para.text.strip()
            if t and t != title and len(t) > 10:
                t = clean(t)
                return t[:90] + ("…" if len(t) > 90 else "")
    return ""


def get_bullets(slide, title, max_bullets=3):
    """Retourne jusqu'à max_bullets phrases de contenu d'une slide (hors titre)."""
    bullets = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            t = para.text.strip()
            if t and t != title and len(t) > 10:
                bullets.append(trunc(clean(t), 75))
                if len(bullets) >= max_bullets:
                    return bullets
    return bullets


def extract_fiche(pptx_path, slides_meta, indices=None):
    """indices : liste 1-indexée des slides à inclure (version preset). None = toutes."""
    prs = Presentation(pptx_path)

    # Titre atelier (slide 1 du PPTX original, toujours)
    titre = clean(get_slide_title(prs.slides[0]))

    # Filtrer les paires slide/meta selon la version
    idx_set = set(indices) if indices is not None else None
    pairs = [(s, m) for s, m in zip(prs.slides, slides_meta)
             if idx_set is None or m["n"] in idx_set]

    # Durée totale = slides sélectionnées uniquement
    duree_min = sum(m["min"] for _, m in pairs)
    h, mv = divmod(duree_min, 60)
    duree_str = f"{h}h{mv:02d}" if h and mv else (f"{h}h" if h else f"{mv} min")

    # Slides de contenu (hors titre et conclusion)
    content_slides = [
        (slide, meta) for slide, meta in pairs
        if meta["role"] in ("contenu", "demo", "chapitre") and meta["n"] > 1
    ]

    # Objectifs : titres des 4 premières slides de contenu (max 55 car.)
    objectifs = [trunc(clean(m["titre"]), 55) for _, m in content_slides[:4]]
    while len(objectifs) < 4:
        objectifs.append("")

    # Déroulé : slides non-titre non-optionnelles (max 6)
    deroulé = []
    cumtime = 0
    for slide, meta in pairs:
        if meta["role"] == "titre":
            cumtime += meta["min"]
            continue
        if meta.get("optionnel"):
            cumtime += meta["min"]
            continue
        h2, m2 = divmod(cumtime, 60)
        heure_str = f"{h2}:{m2:02d}"
        bullet = get_first_bullet(slide, meta["titre"])
        deroulé.append({
            "heure": heure_str,
            "min": meta["min"],
            "titre": trunc(clean(meta["titre"]), 48),
            "bullet": bullet,
        })
        cumtime += meta["min"]
        if len(deroulé) == 6:
            break

    # Blocs optionnels (max 2, pour la table en bas de fiche)
    optional_slides = [
        (slide, meta) for slide, meta in pairs
        if meta.get("optionnel") and meta["role"] not in ("titre", "conclusion")
    ]
    optionnels = []
    for slide, meta in optional_slides[:2]:
        titre_opt = trunc(clean(meta["titre"]), 38)
        bullets = get_bullets(slide, meta["titre"], max_bullets=3)
        optionnels.append({
            "min": meta["min"],
            "titre": titre_opt,
            "bullets": bullets,
        })

    return {
        "titre": titre,
        "duree": duree_str,
        "objectifs": objectifs,
        "materiel": MATERIEL_DEFAULT,
        "deroulé": deroulé,
        "optionnels": optionnels,
        "conseils": CONSEILS_DEFAULT,
    }


def set_para_text(para, text):
    """Remplace le texte d'un paragraphe en conservant le style du premier run."""
    p_xml = para._p
    runs = p_xml.findall(qn("a:r"))
    if not runs:
        return
    first_run_xml = runs[0]
    for r in runs[1:]:
        p_xml.remove(r)
    t_el = first_run_xml.find(qn("a:t"))
    if t_el is not None:
        t_el.text = text


def clear_para_runs(para):
    """Supprime tous les runs d'un paragraphe (évite les artefacts de style)."""
    p_xml = para._p
    for r in p_xml.findall(qn("a:r")):
        p_xml.remove(r)


def set_badge_para(para, badge_text, bullet_text):
    """Para avec badge stylé (run[0] conservé) + bullet en run plain séparé."""
    import copy
    p_xml = para._p
    runs = p_xml.findall(qn("a:r"))
    if not runs:
        return
    # Remettre le texte du badge dans run[0] (garde son style orange)
    t_el = runs[0].find(qn("a:t"))
    if t_el is not None:
        t_el.text = badge_text
    # Supprimer tous les runs suivants
    for r in runs[1:]:
        p_xml.remove(r)
    # Ajouter un run plain pour le bullet (copie sans couleur ni gras)
    new_run = copy.deepcopy(runs[0])
    rPr = new_run.find(qn("a:rPr"))
    if rPr is not None:
        for fill in rPr.findall(qn("a:solidFill")):
            rPr.remove(fill)
        rPr.attrib.pop("b", None)
        rPr.set("sz", "900")
    new_t = new_run.find(qn("a:t"))
    if new_t is not None:
        new_t.text = f"\t{bullet_text}"
    p_xml.append(new_run)


def clear_table_row(row):
    """Vide une ligne de tableau : texte + fond de cellule (évite les rectangles oranges résiduels)."""
    from lxml import etree
    cell = row.cells[0]
    for p in cell.text_frame.paragraphs:
        clear_para_runs(p)
    # Effacer aussi le fond de cellule (solidFill/gradFill au niveau <a:tcPr>)
    tc = cell._tc
    tcPr = tc.find(qn("a:tcPr"))
    if tcPr is not None:
        for fill_tag in (qn("a:solidFill"), qn("a:gradFill"), qn("a:pattFill"), qn("a:blipFill")):
            for el in tcPr.findall(fill_tag):
                tcPr.remove(el)


def fill_optional_table(shape, optionnels):
    """Remplit (ou vide) les lignes 1 et 2 de la table Blocs optionnels."""
    tbl = shape.table
    for row_idx in [1, 2]:
        row = tbl.rows[row_idx]
        cell = row.cells[0]
        paras = cell.text_frame.paragraphs
        data_idx = row_idx - 1

        if data_idx >= len(optionnels):
            clear_table_row(row)
            continue

        opt = optionnels[data_idx]
        bullets = opt.get("bullets", [])
        if len(paras) > 0:
            set_para_text(paras[0], f"{opt['min']} min")
        if len(paras) > 1:
            set_para_text(paras[1], opt["titre"])
        # para[2] : badge "OPT." (run[0] stylé) + premier bullet en run plain
        if len(paras) > 2:
            set_badge_para(paras[2], "OPT.", f". {bullets[0]}" if bullets else "")
        if len(paras) > 3:
            set_para_text(paras[3], f". {bullets[1]}" if len(bullets) > 1 else "")
        if len(paras) > 4:
            set_para_text(paras[4], f". {bullets[2]}" if len(bullets) > 2 else "")


def set_norm_autofit(text_frame):
    """Remplace spAutoFit par normAutofit pour que le texte rétrécisse au lieu de déborder."""
    from lxml import etree
    bodyPr = text_frame._txBody.find(qn("a:bodyPr"))
    if bodyPr is None:
        return
    for tag in (qn("a:spAutoFit"), qn("a:noAutofit"), qn("a:normAutofit")):
        for el in bodyPr.findall(tag):
            bodyPr.remove(el)
    bodyPr.append(etree.SubElement(bodyPr, qn("a:normAutofit")))


def hide_shape(shape):
    """Déplace la forme hors de la slide (left très négatif) pour la masquer."""
    shape.left = -10000000  # ~-11 cm hors diapo


def fill_template(fiche_data, output_path):
    prs = Presentation(TEMPLATE)
    slide = prs.slides[0]
    shapes = {s.name: s for s in slide.shapes if s.has_text_frame}

    # ── Titre + durée
    s = shapes.get("object 2")
    if s:
        tf = s.text_frame
        set_norm_autofit(tf)
        if len(tf.paragraphs) > 0:
            set_para_text(tf.paragraphs[0], fiche_data['titre'])
        if len(tf.paragraphs) > 1:
            set_para_text(tf.paragraphs[1],
                f"Fiche Action | {fiche_data['duree']} | Guide d'animation — Conseillers Numériques CD47")

    # ── Objectifs
    s = shapes.get("object 7")
    if s:
        tf = s.text_frame
        for i, obj in enumerate(fiche_data["objectifs"][:4], 1):
            if i < len(tf.paragraphs):
                set_para_text(tf.paragraphs[i], obj)

    # ── Matériel
    s = shapes.get("object 11")
    if s:
        tf = s.text_frame
        for i, item in enumerate(fiche_data["materiel"][:4], 1):
            if i < len(tf.paragraphs):
                set_para_text(tf.paragraphs[i], item)

    # ── Déroulé (6 étapes : paires heure/contenu)
    step_pairs = [
        ("object 16", "object 17"),
        ("object 22", "object 23"),
        ("object 28", "object 29"),
        ("object 34", "object 35"),
        ("object 40", "object 41"),
        ("object 46", "object 47"),
    ]
    for idx, (time_name, content_name) in enumerate(step_pairs):
        if idx >= len(fiche_data["deroulé"]):
            s_time = shapes.get(time_name)
            s_cont = shapes.get(content_name)
            if s_time:
                set_para_text(s_time.text_frame.paragraphs[0], "")
            if s_cont:
                for p in s_cont.text_frame.paragraphs:
                    set_para_text(p, "")
            continue

        step = fiche_data["deroulé"][idx]

        s_time = shapes.get(time_name)
        if s_time:
            set_para_text(s_time.text_frame.paragraphs[0], step["heure"])

        s_cont = shapes.get(content_name)
        if s_cont:
            tf = s_cont.text_frame
            paras = tf.paragraphs
            if len(paras) > 0:
                set_para_text(paras[0], f"{step['min']} min")
            if len(paras) > 1:
                set_para_text(paras[1], step["titre"])
            if len(paras) > 2:
                set_para_text(paras[2], f". {step['bullet']}" if step["bullet"] else "")
            for i in range(3, len(paras)):
                set_para_text(paras[i], "")

    # ── Blocs optionnels (table object 53 + décoration object 48)
    has_opt = bool(fiche_data["optionnels"])
    for shape in slide.shapes:
        if shape.name == "object 53" and shape.shape_type == 19:  # TABLE
            if has_opt:
                fill_optional_table(shape, fiche_data["optionnels"])
            else:
                hide_shape(shape)
        elif shape.name == "object 48":
            # GROUP décoratif du bloc OPT : toujours masqué car on ne peut pas
            # cibler ses enfants individuellement (1 décoration / ligne optionnelle)
            hide_shape(shape)

    # ── Conseils
    s = shapes.get("object 57")
    if s:
        tf = s.text_frame
        for i, conseil in enumerate(fiche_data["conseils"][:4], 1):
            if i < len(tf.paragraphs):
                set_para_text(tf.paragraphs[i], conseil)

    prs.save(output_path)


def main():
    FICHES_DIR.mkdir(exist_ok=True)

    with open(PLANS_FILE) as f:
        plans = json.load(f)

    generated = []
    for pptx_path in sorted(PPTX_DIR.glob("*.pptx")):
        key = pptx_path.stem
        plan = plans.get(key, {})
        if not plan.get("analyse") or not plan.get("slides"):
            print(f"SKIP {key} — non analysé")
            continue

        versions = plan.get("versions") or {}
        if not versions:
            # Fallback : version unique complète
            versions = {"complet": {"indices": [s["n"] for s in plan["slides"]], "label": "Complet"}}

        ok = []
        for vk, vdata in versions.items():
            try:
                fiche_data = extract_fiche(pptx_path, plan["slides"], indices=vdata["indices"])
                out = FICHES_DIR / f"{key}-fiche-{vk}.pptx"
                fill_template(fiche_data, out)
                ok.append(f"{vk}({fiche_data['duree']})")
            except Exception as e:
                print(f"ERREUR {key}/{vk}: {e}")

        if ok:
            print(f"OK {key}: {', '.join(ok)}")
            generated.append(key)

    print(f"\n{len(generated)} atelier(s) traité(s) : {', '.join(generated)}")


if __name__ == "__main__":
    main()
