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
    return re.sub(r'[^\x00-\x7FÀ-ɏ’–—«»]', '', text).strip()


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
                # Tronquer à 90 caractères
                return t[:90] + ("…" if len(t) > 90 else "")
    return ""


def extract_fiche(pptx_path, slides_meta):
    prs = Presentation(pptx_path)

    # Titre atelier (slide 1)
    titre = clean(get_slide_title(prs.slides[0]))

    # Durée totale
    duree_min = sum(s["min"] for s in slides_meta)
    h, m = divmod(duree_min, 60)
    duree_str = f"{h}h{m:02d}" if h and m else (f"{h}h" if h else f"{m} min")

    # Slides de contenu (hors titre et conclusion)
    content_slides = [
        (slide, meta) for slide, meta in zip(prs.slides, slides_meta)
        if meta["role"] in ("contenu", "demo", "chapitre") and meta["n"] > 1
    ]

    # Objectifs : titres des 4 premières slides de contenu (max 55 car.)
    def trunc(t, n): return t[:n] + "…" if len(t) > n else t
    objectifs = [trunc(clean(m["titre"]), 55) for _, m in content_slides[:4]]
    while len(objectifs) < 4:
        objectifs.append("")

    # Déroulé : toutes les slides non-titre (max 6)
    deroulé = []
    cumtime = 0
    for slide, meta in zip(prs.slides, slides_meta):
        if meta["role"] == "titre":
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

    return {
        "titre": titre,
        "duree": duree_str,
        "objectifs": objectifs,
        "materiel": MATERIEL_DEFAULT,
        "deroulé": deroulé,
        "conseils": CONSEILS_DEFAULT,
    }


def set_para_text(para, text):
    """Remplace le texte d'un paragraphe en conservant le style du premier run."""
    p_xml = para._p
    runs = p_xml.findall(qn("a:r"))
    if not runs:
        return
    # Conserver le premier run, supprimer les autres
    first_run_xml = runs[0]
    for r in runs[1:]:
        p_xml.remove(r)
    # Mettre à jour le texte dans le premier run
    t_el = first_run_xml.find(qn("a:t"))
    if t_el is not None:
        t_el.text = text


def fill_template(fiche_data, output_path):
    prs = Presentation(TEMPLATE)
    slide = prs.slides[0]
    shapes = {s.name: s for s in slide.shapes if s.has_text_frame}

    # ── Titre + durée
    s = shapes.get("object 2")
    if s:
        tf = s.text_frame
        if len(tf.paragraphs) > 0:
            set_para_text(tf.paragraphs[0], f"Fiche Action - Atelier {fiche_data['titre']}")
        if len(tf.paragraphs) > 1:
            set_para_text(tf.paragraphs[1],
                f"{fiche_data['duree']} | Guide d'animation pour les Conseillers Numériques CD47")

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
            # Vider l'étape si pas de données
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

        try:
            fiche_data = extract_fiche(pptx_path, plan["slides"])
            out = FICHES_DIR / f"{key}-fiche.pptx"
            fill_template(fiche_data, out)
            print(f"OK {key}: {fiche_data['titre']} ({fiche_data['duree']})")
            generated.append(key)
        except Exception as e:
            print(f"ERREUR {key}: {e}")

    print(f"\n{len(generated)} fiche(s) générée(s) : {', '.join(generated)}")


if __name__ == "__main__":
    main()
