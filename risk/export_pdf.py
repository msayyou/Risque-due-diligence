from io import BytesIO
from typing import Dict, Any

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER


def generate_pdf_bytes(asset_name: str, inputs: Dict[str, Any], C: Dict[str, Any]) -> BytesIO:
    """Generate a compact PDF report and return it as BytesIO.

    The PDF is intentionally concise but includes the global score and a small
    table with key KPIs and recommendations.
    """
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('title', parent=styles['Heading1'], alignment=TA_CENTER, textColor=colors.HexColor('#1A2E44'))
    normal = styles['Normal']
    bold = ParagraphStyle('bold', parent=styles['Heading4'], alignment=TA_CENTER)

    story = []
    story.append(Paragraph(f'RAPPORT — {asset_name}', title_style))
    story.append(Spacer(1, 8))

    # Score block
    sc = C.get('global_score', '—')
    story.append(Paragraph(f'<b>Score global:</b> {sc} / 100', normal))
    story.append(Spacer(1, 8))

    # KPIs table
    kpis = [
        ['Indicateur', 'Valeur'],
        ['RevPAR', f"{C.get('revpar','—')}€"],
        ['GOPPAR', f"{C.get('goppar','—')}€"],
        ['Marge GOP', f"{C.get('gop_ratio','—')}%"],
        ['LTV', f"{inputs.get('ltv','—')}%"],
        ['DSCR', f"{inputs.get('dscr','—')}x"]
    ]
    t = Table(kpis, colWidths=[8*cm, 8*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A2E44')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.25, colors.HexColor('#D3D1C7')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F1EFE8')])
    ]))
    story.append(t)
    story.append(Spacer(1, 12))

    # Recommendations (simple)
    recs = []
    if C.get('fin', 100) < 50:
        recs.append('Risque financier élevé — revoir la structure de financement (LTV/DSCR)')
    if C.get('ops', 100) < 50:
        recs.append('Performance opérationnelle faible — auditer revenue management et coûts')
    if C.get('legal', 100) < 50:
        recs.append('Non-conformités réglementaires — corriger DPE/RGPD/sécurité')
    if not recs:
        recs.append('Aucun signal critique détecté. Maintenir le suivi mensuel des KPIs.')

    story.append(Paragraph('<b>Recommandations</b>', normal))
    for r in recs:
        story.append(Paragraph(f'• {r}', normal))
        story.append(Spacer(1, 4))

    doc.build(story)
    buf.seek(0)
    return buf
