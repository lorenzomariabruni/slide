from fastmcp import FastMCP
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
import copy
import os
import re
from pathlib import Path
from lxml import etree

TEMPLATE_PATH = os.environ.get("SLIDE_TEMPLATE", "template.pptx")

mcp = FastMCP(
    name="slide-server",
    version="1.0.0",
    description="MCP Server per generare presentazioni PowerPoint con animazioni a partire da un template .pptx",
)


# ─────────────────────────────────────────────────────────────
# Helper: clona layout dal template e applica testo
# ─────────────────────────────────────────────────────────────

def _clone_slide(prs: Presentation, layout_index: int = 0):
    """Aggiunge uno slide clonando il layout specificato."""
    layout = prs.slide_layouts[layout_index]
    return prs.slides.add_slide(layout)


def _set_placeholder(slide, ph_idx: int, text: str, font_size: int | None = None):
    """Imposta il testo di un placeholder per indice."""
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == ph_idx:
            ph.text = text
            if font_size:
                for para in ph.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(font_size)
            return


def _add_speaker_notes(slide, notes_text: str):
    """Scrive il testo nelle note relatore."""
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text


def _add_fade_animation(slide):
    """
    Aggiunge un'animazione Fade-In (apparizione) a tutti gli oggetti
    del corpo dello slide tramite Open XML.
    """
    spTree = slide.shapes._spTree
    # timing element
    nsmap = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    }
    timing_xml = (
        '<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<p:tnLst>'
        '<p:par>'
        '<p:cTn id="1" dur="indefinite" restart="whenNotActive" nodeType="tmRoot">'
        '<p:childTnLst>'
        '<p:par>'
        '<p:cTn id="2" fill="hold">'
        '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
        '<p:childTnLst>'
        '<p:par>'
        '<p:cTn id="3" presetID="10" presetClass="entr" presetSubtype="0"'
        ' fill="hold" grpId="0" nodeType="clickEffect">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        '<p:childTnLst>'
        '<p:set>'
        '<p:cBhvr>'
        '<p:cTn id="4" dur="1" fill="hold"/>'
        '<p:tgtEl><p:spTgt spid=""/></p:tgtEl>'
        '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
        '</p:cBhvr>'
        '<p:to><p:strVal val="visible"/></p:to>'
        '</p:set>'
        '</p:childTnLst>'
        '</p:cTn>'
        '</p:par>'
        '</p:childTnLst>'
        '</p:cTn>'
        '</p:par>'
        '</p:childTnLst>'
        '</p:cTn>'
        '</p:par>'
        '</p:tnLst>'
        '<p:bldLst/>'
        '</p:timing>'
    )
    timing_el = etree.fromstring(timing_xml)
    # inserisce timing nello spTree del parent (slide xml)
    slide._element.append(timing_el)


def _build_presentation(
    topic: str,
    slides_content: list[dict],
) -> Presentation:
    """
    Costruisce una Presentation partendo dal template.
    slides_content: lista di dict con chiavi 'title', 'bullets', 'notes'
    """
    prs = Presentation(TEMPLATE_PATH)

    # rimuovi eventuali slide di esempio già nel template
    xml_slides = prs.slides._sldIdLst
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].get('r:id') or prs.slides._sldIdLst[0].get(
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        )
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

    # ─── Slide 1: Titolo ───────────────────────────────────────
    slide_title = _clone_slide(prs, 0)
    _set_placeholder(slide_title, 0, topic, font_size=40)
    _add_speaker_notes(
        slide_title,
        f"Benvenuti. Oggi parleremo di: {topic}. "
        "Introducetevi brevemente e ricordate al pubblico l'obiettivo della presentazione.",
    )
    _add_fade_animation(slide_title)

    # ─── Slide 2: Agenda ──────────────────────────────────────
    slide_agenda = _clone_slide(prs, 1)
    _set_placeholder(slide_agenda, 0, "Agenda")
    agenda_items = [
        f"{i + 1}. {s['title']}" for i, s in enumerate(slides_content)
    ]
    agenda_text = "\n".join(agenda_items)
    _set_placeholder(slide_agenda, 1, agenda_text, font_size=18)
    _add_speaker_notes(
        slide_agenda,
        "Ecco gli argomenti che tratteremo oggi. " + ", ".join([s['title'] for s in slides_content]) + ".",
    )
    _add_fade_animation(slide_agenda)

    # ─── Slide N: Contenuto ───────────────────────────────────
    for i, sc in enumerate(slides_content):
        layout_idx = 1 if len(prs.slide_layouts) == 1 else min(1, len(prs.slide_layouts) - 1)
        slide = _clone_slide(prs, layout_idx)
        _set_placeholder(slide, 0, sc["title"])

        bullets = sc.get("bullets", [])
        body_text = "\n".join(f"• {b}" for b in bullets)
        _set_placeholder(slide, 1, body_text, font_size=18)
        _add_speaker_notes(slide, sc.get("notes", ""))
        _add_fade_animation(slide)

    # ─── Ultima Slide: Ringraziamenti ─────────────────────────
    slide_thanks = _clone_slide(prs, 0)
    _set_placeholder(slide_thanks, 0, "Grazie!", font_size=48)
    _set_placeholder(
        slide_thanks, 1,
        "Domande? Siamo felici di rispondere.\n\nContatti: [email] | [LinkedIn]",
        font_size=20,
    )
    _add_speaker_notes(
        slide_thanks,
        "Grazie mille per la vostra attenzione. Siamo ora disponibili per rispondere a qualsiasi domanda.",
    )
    _add_fade_animation(slide_thanks)

    return prs


# ─────────────────────────────────────────────────────────────
# Tool: genera presentazione
# ─────────────────────────────────────────────────────────────

@mcp.tool(
    description="""
Genera una presentazione PowerPoint (.pptx) a partire da un argomento e da N slide di contenuto.
La presentazione avrà sempre N+3 slide:
  - Slide 1: Titolo con il nome dell'argomento
  - Slide 2: Agenda numerata con tutti i titoli
  - Slide 3..N+2: Slide di contenuto richieste
  - Slide N+3: Ringraziamenti finali

Ogni slide include animazioni Fade-In e note relatore pronte per essere lette.
Il layout grafico viene preso automaticamente da template.pptx.

Convenzioni rispettate:
  - Titolo: massimo 7 parole
  - Bullet point per slide: massimo 5
  - Ogni bullet: massimo 15 parole
  - Niente paragrafi lunghi: frasi brevi e telegrafiche

Few-shot examples
-----------------
Esempio 1 – singola slide:
  topic = "Intelligenza Artificiale nel 2025"
  slides = [
    {
      "title": "Modelli Linguistici",
      "bullets": ["GPT-4 e successori", "RAG e knowledge graph", "Costi in calo del 60%"],
      "notes": "Iniziate citando i principali vendor e mostrate il grafico di adozione."
    }
  ]

Esempio 2 – tre slide:
  topic = "Sostenibilità Aziendale"
  slides = [
    {
      "title": "Perché la Sostenibilità",
      "bullets": ["Riduzione emissioni CO2", "Risparmio energetico", "Reputazione del brand"],
      "notes": "Citate il rapporto IPCC 2024 e i dati Eurostat."
    },
    {
      "title": "Strumenti e Standard",
      "bullets": ["ESG reporting", "ISO 14001", "Carbon footprint calculator"],
      "notes": "Mostrate un esempio di dashboard ESG reale."
    },
    {
      "title": "Piano d'Azione",
      "bullets": ["Quick wins nei prossimi 3 mesi", "Obiettivi annuali", "KPI di monitoraggio"],
      "notes": "Lasciate spazio per domande al termine di questa slide."
    }
  ]
"""
)
def generate_presentation(
    topic: str,
    slides: list[dict],
    output_path: str = "output.pptx",
) -> str:
    """
    Genera una presentazione PowerPoint.

    Args:
        topic: Il titolo/argomento principale della presentazione.
        slides: Lista di dict, ognuno con:
                  - title (str): titolo della slide (max 7 parole)
                  - bullets (list[str]): elenco puntato, max 5 voci da max 15 parole
                  - notes (str): discorso completo da leggere come note relatore
        output_path: Percorso del file .pptx da generare (default: output.pptx)

    Returns:
        Messaggio con il percorso del file generato.
    """
    # Validazione convenzioni
    errors = []
    for i, s in enumerate(slides):
        title_words = len(s.get("title", "").split())
        if title_words > 7:
            errors.append(f"Slide {i+1}: il titolo ha {title_words} parole (max 7).")
        bullets = s.get("bullets", [])
        if len(bullets) > 5:
            errors.append(f"Slide {i+1}: troppi bullet ({len(bullets)}, max 5).")
        for j, b in enumerate(bullets):
            bw = len(b.split())
            if bw > 15:
                errors.append(f"Slide {i+1}, bullet {j+1}: {bw} parole (max 15).")
    if errors:
        return "Errori di validazione:\n" + "\n".join(errors)

    prs = _build_presentation(topic, slides)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    total = len(slides) + 3
    return (
        f"Presentazione generata con successo: {out.resolve()}\n"
        f"Slide totali: {total} ({len(slides)} di contenuto + titolo + agenda + ringraziamenti)"
    )


# ─────────────────────────────────────────────────────────────
# Tool: lista template disponibili
# ─────────────────────────────────────────────────────────────

@mcp.tool(
    description="""
Restituisce informazioni sul template PowerPoint attualmente configurato.
Utile per verificare quanti layout sono disponibili prima di generare una presentazione.

Few-shot examples
-----------------
Esempio 1:
  Input: nessun parametro richiesto
  Output: "Template: template.pptx | Layout disponibili: 11 | Slide example già presenti: 0"

Esempio 2 (template non trovato):
  Output: "Errore: template.pptx non trovato. Assicurati che il file esista nella directory corrente."
"""
)
def template_info() -> str:
    """
    Ritorna informazioni sul template .pptx configurato.

    Returns:
        Stringa descrittiva con path, numero di layout e slide presenti.
    """
    path = Path(TEMPLATE_PATH)
    if not path.exists():
        return f"Errore: {TEMPLATE_PATH} non trovato. Assicurati che il file esista nella directory corrente."
    prs = Presentation(str(path))
    return (
        f"Template: {path.resolve()}\n"
        f"Layout disponibili: {len(prs.slide_layouts)}\n"
        f"Slide già presenti nel template: {len(prs.slides)}\n"
        f"Dimensioni slide: {prs.slide_width.inches:.1f}\" x {prs.slide_height.inches:.1f}\""
    )


# ─────────────────────────────────────────────────────────────
# Tool: valida contenuto slide
# ─────────────────────────────────────────────────────────────

@mcp.tool(
    description="""
Valida il contenuto di una lista di slide rispetto alle convenzioni standard delle presentazioni:
  - Titolo slide: max 7 parole
  - Bullet per slide: max 5
  - Parole per bullet: max 15
  - Note relatore: obbligatorie (almeno 10 parole)

Ritorna OK se tutto è valido, oppure un elenco dettagliato degli errori.

Few-shot examples
-----------------
Esempio 1 – tutto valido:
  slides = [
    {
      "title": "Cloud Computing",
      "bullets": ["Scalabilità on-demand", "Riduzione costi IT"],
      "notes": "Spiegate come il cloud ha trasformato le infrastrutture negli ultimi 10 anni."
    }
  ]
  Output: "✅ Tutte le slide rispettano le convenzioni."

Esempio 2 – errori rilevati:
  slides = [
    {
      "title": "Questa è una slide con un titolo molto lungo che supera il limite",
      "bullets": ["b1","b2","b3","b4","b5","b6"],
      "notes": "ok"
    }
  ]
  Output:
    "❌ Errori trovati:\n"
    "- Slide 1: titolo ha 12 parole (max 7)\n"
    "- Slide 1: 6 bullet points (max 5)\n"
    "- Slide 1: note troppo brevi (1 parola, min 10)"
"""
)
def validate_slides(slides: list[dict]) -> str:
    """
    Valida una lista di slide rispetto alle convenzioni.

    Args:
        slides: Lista di dict con title, bullets, notes.

    Returns:
        Stringa con esito validazione.
    """
    errors = []
    for i, s in enumerate(slides, 1):
        title = s.get("title", "")
        bullets = s.get("bullets", [])
        notes = s.get("notes", "")

        tw = len(title.split())
        if tw > 7:
            errors.append(f"Slide {i}: titolo ha {tw} parole (max 7)")

        if len(bullets) > 5:
            errors.append(f"Slide {i}: {len(bullets)} bullet points (max 5)")

        for j, b in enumerate(bullets, 1):
            bw = len(b.split())
            if bw > 15:
                errors.append(f"Slide {i}, bullet {j}: {bw} parole (max 15)")

        nw = len(notes.split())
        if nw < 10:
            errors.append(f"Slide {i}: note troppo brevi ({nw} parole, min 10)")

    if errors:
        return "❌ Errori trovati:\n" + "\n".join(f"- {e}" for e in errors)
    return "✅ Tutte le slide rispettano le convenzioni."


if __name__ == "__main__":
    mcp.run()
