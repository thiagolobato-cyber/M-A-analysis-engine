#!/usr/bin/env python3
"""
Fase 6 — Output Generator.

Lê os resultados de um deal já processado (agent_runs + synthesis_runs) e
gera:
  output/{deal_id}_analise.xlsx  — 8 abas (7 do Prompt Master + Financial Analysis)
  output/{deal_id}_executivo.pptx — 5 slides, estilo CFO/IC/McKinsey

Uso:
    python scripts/generate_outputs.py --deal-id <uuid>

Sobre a entrega: até o conector do Microsoft Drive existir (Fase 5, hoje
bloqueada esperando acesso de TI), esses dois arquivos ficam disponíveis
como "Artifact" da própria execução do GitHub Actions — dá pra baixar
direto da página do run, sem precisar de mais nenhum acesso. Quando o
Drive estiver pronto, isso vira um upload automático além do artifact
(ou no lugar dele).
"""
import argparse
import json
import os
import sys
import urllib.request
import urllib.error
from datetime import datetime, timezone

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------------------------------------------------------------------
# Paleta — reaproveita as cores de marca da BHub (mesmo tom da arquitetura)
# ---------------------------------------------------------------------------
NAVY = "0F1727"
CORAL = "F25461"
SUCCESS = "2E7D4F"
WARN = "B6592A"
DANGER = "B03A3A"
TEXT_DARK = "141414"
TEXT_MUTED = "6A6A68"
BORDER_GRAY = "D8D8D6"
WHITE = "FFFFFF"

RECOMMENDATION_COLOR = {"GO": SUCCESS, "NO-GO": DANGER, "CONDITIONAL_GO": WARN}


# ---------------------------------------------------------------------------
# Supabase (mesma abordagem do run_agent.py)
# ---------------------------------------------------------------------------
def supabase_request(method: str, path: str) -> dict:
    url = os.environ["SUPABASE_URL"].rstrip("/") + "/rest/v1/" + path
    key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    headers = {"apikey": key, "Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    req = urllib.request.Request(url, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        print(f"Erro Supabase [{method} {path}]: {e.code} {e.read().decode()}", file=sys.stderr)
        raise


def fetch_deal_bundle(deal_id: str) -> dict:
    deals = supabase_request("GET", f"deals?id=eq.{deal_id}")
    if not deals:
        raise SystemExit(f"Deal {deal_id} não encontrado.")

    runs = supabase_request(
        "GET",
        f"agent_runs?deal_id=eq.{deal_id}&select=*,agent_versions(version,agent_id,agents(name))",
    )
    by_agent = {}
    for r in runs:
        name = r["agent_versions"]["agents"]["name"]
        by_agent[name] = r

    synthesis = supabase_request(
        "GET",
        f"synthesis_runs?deal_id=eq.{deal_id}&select=*,agent_versions(version)&order=created_at.desc&limit=1",
    )

    return {
        "deal": deals[0],
        "agents": by_agent,
        "synthesis": synthesis[0] if synthesis else None,
    }


# ---------------------------------------------------------------------------
# EXCEL — 8 abas
# ---------------------------------------------------------------------------
def build_excel(bundle: dict, path: str):
    deal = bundle["deal"]
    agents = bundle["agents"]
    synth = bundle["synthesis"]

    wb = Workbook()
    header_font = Font(bold=True, color=WHITE, size=11)
    header_fill = PatternFill("solid", fgColor=NAVY)
    title_font = Font(bold=True, size=14, color=TEXT_DARK)
    thin_border = Border(*[Side(style="thin", color=BORDER_GRAY)] * 4)
    wrap = Alignment(wrap_text=True, vertical="top")

    def style_header_row(ws, row, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(vertical="center")

    def autosize(ws, widths):
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

    # 1. Executive Summary --------------------------------------------------
    ws = wb.active
    ws.title = "Executive Summary"
    ws["A1"] = deal["name"]
    ws["A1"].font = Font(bold=True, size=18, color=NAVY)
    ws.merge_cells("A1:B1")

    complexity = agents.get("complexity", {}).get("output", {}) or {}
    rec = synth.get("recommendation") if synth else None
    rows = [
        ("Deal", deal["name"]),
        ("Status", deal.get("status", "")),
        ("Complexidade", complexity.get("classificacao", "não avaliado")),
        ("Recomendação", rec or "não avaliado"),
        ("Confiança da recomendação", synth.get("confidence") if synth else None),
    ]
    r = 3
    for label, value in rows:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        ws.cell(row=r, column=2, value=value)
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Sumário executivo").font = title_font
    r += 1
    if synth:
        cell = ws.cell(row=r, column=1, value=synth.get("output", {}).get("sumario_executivo", ""))
        cell.alignment = wrap
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.row_dimensions[r].height = 90
    r += 2

    def write_list_section(ws, r, title, items):
        ws.cell(row=r, column=1, value=title).font = title_font
        r += 1
        for item in items or []:
            ws.cell(row=r, column=1, value=f"• {item}")
            r += 1
        return r + 1

    if synth:
        out = synth.get("output", {})
        r = write_list_section(ws, r, "Principais riscos", out.get("principais_riscos"))
        r = write_list_section(ws, r, "Próximos passos / condições", out.get("condicoes"))
    autosize(ws, [28, 60, 20, 20])

    # 2. Complexity -----------------------------------------------------
    ws = wb.create_sheet("Complexity")
    if "score_total" in complexity:
        ws["A1"] = "Score total (Bloco B)"; ws["B1"] = complexity.get("score_total")
        ws["A2"] = "Classificação"; ws["B2"] = complexity.get("classificacao")
        header_row = 4
    else:
        header_row = 1
    headers = ["Criterion", "Finding", "Impact", "Evidence"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=i, value=h)
    style_header_row(ws, header_row, len(headers))
    row = header_row + 1
    for crit in complexity.get("criterios_acionados", []):
        ws.cell(row=row, column=1, value=crit.get("criterio"))
        ws.cell(row=row, column=2, value=crit.get("detalhe"))
        ws.cell(row=row, column=3, value=crit.get("nivel_apontado"))
        ws.cell(row=row, column=4, value=json.dumps(complexity.get("evidence", []), ensure_ascii=False))
        row += 1
    autosize(ws, [28, 60, 14, 40])

    # 3. Integration Risks ------------------------------------------------
    ws = wb.create_sheet("Integration Risks")
    ir = agents.get("integration_risks", {}).get("output", {}) or {}
    ws["A1"] = "Nível de risco"; ws["B1"] = ir.get("nivel_risco", "n/d")
    row = 3
    ws.cell(row=row, column=1, value="Riscos de integração").font = title_font
    row += 1
    for item in ir.get("riscos_integracao", []):
        ws.cell(row=row, column=1, value=f"• {item}"); row += 1
    autosize(ws, [90])

    # 4. Operational Risks --------------------------------------------------
    ws = wb.create_sheet("Operational Risks")
    op = agents.get("operational_risks", {}).get("output", {}) or {}
    ws["A1"] = "Concentração top 10 (%)"; ws["B1"] = op.get("concentracao_receita_top10_pct")
    ws["A2"] = "Nível de concentração"; ws["B2"] = op.get("nivel_concentracao")
    row = 4
    ws.cell(row=row, column=1, value="Riscos operacionais").font = title_font
    row += 1
    for item in op.get("riscos_operacionais", []):
        ws.cell(row=row, column=1, value=f"• {item}"); row += 1
    autosize(ws, [90])

    # 5. Strategic Opinion ----------------------------------------------
    ws = wb.create_sheet("Strategic Opinion")
    op_agent = agents.get("opinion", {}).get("output", {}) or {}
    ws["A1"] = "Recomendação"; ws["B1"] = op_agent.get("recomendacao")
    ws["A2"] = "Parecer do time"; ws["A2"].font = title_font
    cell = ws.cell(row=3, column=1, value=op_agent.get("parecer_do_time", ""))
    cell.alignment = wrap
    ws.merge_cells("A3:D3")
    ws.row_dimensions[3].height = 100
    autosize(ws, [28, 40, 20, 20])

    # 6. CFO Synthesis --------------------------------------------------
    ws = wb.create_sheet("CFO Synthesis")
    if synth:
        out = synth.get("output", {})
        ws["A1"] = "Recomendação final"; ws["B1"] = synth.get("recommendation")
        ws["A2"] = "Confiança"; ws["B2"] = synth.get("confidence")
        ws["A3"] = "Sumário executivo"; ws["A3"].font = title_font
        cell = ws.cell(row=4, column=1, value=out.get("sumario_executivo", ""))
        cell.alignment = wrap
        ws.merge_cells("A4:D4")
        ws.row_dimensions[4].height = 100
        row = 6
        ws.cell(row=row, column=1, value="Reconciliação entre agentes").font = title_font
        row += 1
        for rec_item in out.get("reconciliacao", []):
            ws.cell(row=row, column=1, value=f"Divergência: {rec_item.get('divergencia')}")
            row += 1
            ws.cell(row=row, column=1, value=f"Resolução: {rec_item.get('resolucao')}")
            row += 2
    autosize(ws, [28, 60, 20, 20])

    # 7. Financial Analysis — reescrito para o schema forense v2.0 (19/08):
    #    DRE de verdade, anomalias mês a mês, simulação de cenário, DD ------
    ws = wb.create_sheet("Financial Analysis")
    fin = agents.get("financial_analysis", {}).get("output", {}) or {}
    row = 1

    bridge = fin.get("ebitda_bridge") or {}
    ws.cell(row=row, column=1, value="1. EBITDA Bridge").font = title_font
    row += 1
    bridge_rows = [
        ("Lucro Líquido", bridge.get("lucro_liquido")), ("Resultado Financeiro", bridge.get("resultado_financeiro")),
        ("Tributos sobre Lucro", bridge.get("tributos_sobre_lucro")), ("D&A", bridge.get("d_a")),
        ("EBITDA Reportado", bridge.get("ebitda_reportado")), ("Margem Reportada %", bridge.get("margem_reportada_pct")),
        ("EBITDA Ajustado", bridge.get("ebitda_ajustado")), ("Margem Ajustada %", bridge.get("margem_ajustada_pct")),
    ]
    for label, value in bridge_rows:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=2, value=value)
        row += 1
    ajustes = bridge.get("ajustes") or []
    if ajustes:
        row += 1
        ws.cell(row=row, column=1, value="Ajustes de não-recorrência").font = Font(bold=True)
        row += 1
        for aj in ajustes:
            ws.cell(row=row, column=1, value=aj.get("descricao"))
            ws.cell(row=row, column=2, value=aj.get("valor"))
            ws.cell(row=row, column=3, value=aj.get("justificativa"))
            row += 1

    row += 2
    ws.cell(row=row, column=1, value="2. Anomalias Detectadas").font = title_font
    row += 1
    headers = ["Conta", "Período", "Severidade", "Narrativa"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=row, column=i, value=h)
    style_header_row(ws, row, len(headers))
    row += 1
    for a in fin.get("anomalias", []):
        ws.cell(row=row, column=1, value=a.get("conta")); ws.cell(row=row, column=2, value=a.get("periodo"))
        ws.cell(row=row, column=3, value=a.get("severidade")); ws.cell(row=row, column=4, value=a.get("narrativa"))
        row += 1
    if not fin.get("anomalias"):
        ws.cell(row=row, column=1, value="Nenhuma anomalia detectada (ou sem série temporal suficiente para avaliar)")
        row += 1

    row += 2
    ws.cell(row=row, column=1, value="3. Red Flags").font = title_font
    row += 1
    headers = ["Severidade", "Título", "Detalhe", "Valor"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=row, column=i, value=h)
    style_header_row(ws, row, len(headers))
    row += 1
    for flag in fin.get("red_flags", []):
        ws.cell(row=row, column=1, value=flag.get("severidade")); ws.cell(row=row, column=2, value=flag.get("titulo"))
        ws.cell(row=row, column=3, value=flag.get("detalhe")); ws.cell(row=row, column=4, value=flag.get("valor"))
        row += 1

    row += 2
    ws.cell(row=row, column=1, value="4. Perguntas para Due Diligence").font = title_font
    row += 1
    headers = ["Prioridade", "Pergunta", "Contexto"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=row, column=i, value=h)
    style_header_row(ws, row, len(headers))
    row += 1
    for q in fin.get("perguntas_dd", []):
        ws.cell(row=row, column=1, value=q.get("prioridade")); ws.cell(row=row, column=2, value=q.get("pergunta"))
        ws.cell(row=row, column=3, value=q.get("contexto"))
        row += 1

    row += 2
    kpis = fin.get("kpis_valuation") or {}
    if kpis:
        ws.cell(row=row, column=1, value="5. KPIs para Valuation").font = title_font
        row += 1
        for k, v in kpis.items():
            ws.cell(row=row, column=1, value=k); ws.cell(row=row, column=2, value=v)
            row += 1

    limitacoes = fin.get("limitacoes_dados") or []
    if limitacoes:
        row += 2
        ws.cell(row=row, column=1, value="Limitações dos dados disponíveis").font = title_font
        row += 1
        for lim in limitacoes:
            cell = ws.cell(row=row, column=1, value=f"• {lim}")
            cell.alignment = wrap
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            row += 1

    autosize(ws, [26, 40, 16, 16, 12, 45])

    # 9. Viabilidade Financeira (Bloco A — 100% código, ver regras_negocio.py)
    ws = wb.create_sheet("Viabilidade Financeira")
    vf = agents.get("viabilidade_financeira", {}).get("output", {}) or {}
    ws["A1"] = "Classificação"; ws["A1"].font = title_font
    ws["B1"] = vf.get("classificacao", "não avaliado")
    row = 3
    ws.cell(row=row, column=1, value="Critério").font = title_font
    ws.cell(row=row, column=2, value="Faixa").font = title_font
    row += 1
    for criterio, faixa in (vf.get("criterios") or {}).items():
        ws.cell(row=row, column=1, value=criterio)
        ws.cell(row=row, column=2, value=faixa or "não avaliado")
        row += 1
    row += 1
    ws.cell(row=row, column=1, value="Perfil restrito presente?")
    ws.cell(row=row, column=2, value=vf.get("perfil_restrito_presente"))
    row += 1
    if vf.get("criterios_nao_avaliados"):
        row += 1
        ws.cell(row=row, column=1, value="Critérios não avaliados (dado ausente)").font = Font(bold=True)
        row += 1
        for c in vf.get("criterios_nao_avaliados", []):
            ws.cell(row=row, column=1, value=f"• {c}"); row += 1
    autosize(ws, [32, 24])

    # 8. AI Audit Log --------------------------------------------------
    ws = wb.create_sheet("AI Audit Log")
    headers = ["Agent", "Version", "Timestamp", "Status"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header_row(ws, 1, len(headers))
    row = 2
    for name, run in agents.items():
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=run["agent_versions"]["version"])
        ws.cell(row=row, column=3, value=run.get("created_at", ""))
        ws.cell(row=row, column=4, value=run.get("status", ""))
        row += 1
    if synth:
        ws.cell(row=row, column=1, value="cfo_synthesis")
        ws.cell(row=row, column=2, value=synth["agent_versions"]["version"])
        ws.cell(row=row, column=3, value=synth.get("created_at", ""))
        ws.cell(row=row, column=4, value=synth.get("status", ""))
    autosize(ws, [22, 14, 24, 14])

    wb.save(path)
    print(f"Excel salvo: {path}")


# ---------------------------------------------------------------------------
# PPT — 5 slides, estilo CFO / Investment Committee
# ---------------------------------------------------------------------------
SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.6)


def _rgb(hexstr):
    return RGBColor.from_string(hexstr)


def _add_textbox(slide, left, top, width, height, text, size=14, bold=False,
                  color=TEXT_DARK, align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font
    run.font.color.rgb = _rgb(color)
    return box


def _add_bullets(slide, left, top, width, height, items, size=14, color=TEXT_DARK, space_after=8):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = f"—  {item}"
        run.font.size = Pt(size)
        run.font.color.rgb = _rgb(color)
        run.font.name = "Calibri"
    return box


def _badge(slide, left, top, width, height, text, fill_hex, text_color=WHITE, size=14):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill_hex)
    shape.line.fill.background()
    # shape.shadow.inherit = False não é suficiente pro LibreOffice — remove
    # o elemento de efeito diretamente na árvore XML da forma.
    sp = shape._element.spPr
    for tag in ("a:effectLst", "a:effectDag"):
        el = sp.find("{http://schemas.openxmlformats.org/drawingml/2006/main}" + tag.split(":")[1])
        if el is not None:
            sp.remove(el)
    from pptx.oxml.ns import qn
    effect_lst = sp.makeelement(qn("a:effectLst"), {})
    sp.append(effect_lst)
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = _rgb(text_color)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return shape


def _blank_slide(prs, bg_hex=WHITE):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(bg_hex)
    return slide


def _page_title(slide, text):
    _add_textbox(slide, MARGIN, Inches(0.5), Inches(11.5), Inches(0.8),
                 text, size=30, bold=True, color=NAVY)


def build_pptx(bundle: dict, path: str):
    deal = bundle["deal"]
    agents = bundle["agents"]
    synth = bundle["synthesis"]
    complexity = agents.get("complexity", {}).get("output", {}) or {}
    opinion = agents.get("opinion", {}).get("output", {}) or {}
    integration = agents.get("integration_risks", {}).get("output", {}) or {}
    operational = agents.get("operational_risks", {}).get("output", {}) or {}
    financial = agents.get("financial_analysis", {}).get("output", {}) or {}
    synth_out = (synth or {}).get("output", {})
    recommendation = (synth or {}).get("recommendation", "N/D")

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # ---- Slide 1: Deal Overview (fundo navy, alto contraste) -------------
    s = _blank_slide(prs, bg_hex=NAVY)
    _add_textbox(s, MARGIN, Inches(2.3), Inches(12.1), Inches(1.2),
                 deal["name"], size=40, bold=True, color=WHITE)
    _add_textbox(s, MARGIN, Inches(3.5), Inches(12.1), Inches(0.6),
                 "M&A Analysis Engine — Sumário Executivo", size=16, color="C9CDD6")

    complexity_label = complexity.get("classificacao", "não avaliado")
    stats = [
        ("Complexidade", complexity_label),
        ("Recomendação", recommendation),
        ("Confiança", f"{round((synth or {}).get('confidence', 0) * 100)}%" if synth else "n/d"),
    ]
    x = MARGIN
    for label, value in stats:
        _add_textbox(s, x, Inches(4.6), Inches(3.8), Inches(0.4), label.upper(), size=12, color="8B93A8")
        _add_textbox(s, x, Inches(5.0), Inches(3.8), Inches(0.7), str(value), size=22, bold=True, color=WHITE)
        x += Inches(4.0)

    # ---- Slide 2: Complexity Assessment ------------------------------------
    s = _blank_slide(prs)
    _page_title(s, "Complexity Assessment")
    _badge(s, MARGIN, Inches(1.4), Inches(3.2), Inches(0.55), complexity_label.upper(),
           fill_hex=DANGER if "Alta" in complexity_label else (WARN if "Média" in complexity_label else SUCCESS))

    criteria = complexity.get("criterios_acionados", [])
    if criteria:
        top = Inches(2.3)
        for crit in criteria[:4]:
            _add_textbox(s, MARGIN, top, Inches(3.0), Inches(0.5), crit.get("criterio", ""), size=15, bold=True, color=NAVY)
            _add_textbox(s, Inches(4.0), top, Inches(8.5), Inches(0.7), crit.get("detalhe", ""), size=14, color=TEXT_MUTED)
            top += Inches(1.0)
    else:
        _add_textbox(s, MARGIN, Inches(2.3), Inches(11.5), Inches(1), "Nenhum critério de complexidade foi acionado.", size=14, color=TEXT_MUTED)

    # ---- Slide 3: Key Risks (integração + operacional) ---------------------
    s = _blank_slide(prs)
    _page_title(s, "Key Risks")
    _add_textbox(s, MARGIN, Inches(1.6), Inches(5.8), Inches(0.4), "INTEGRAÇÃO", size=14, bold=True, color=CORAL)
    _add_bullets(s, MARGIN, Inches(2.15), Inches(5.8), Inches(4.5),
                 (integration.get("riscos_integracao") or ["Sem riscos de integração identificados."])[:5], size=15, space_after=14)

    _add_textbox(s, Inches(6.9), Inches(1.6), Inches(5.8), Inches(0.4), "OPERACIONAL", size=14, bold=True, color=CORAL)
    op_pct = operational.get("concentracao_receita_top10_pct")
    op_items = operational.get("riscos_operacionais") or ["Sem riscos operacionais identificados."]
    body_top = Inches(2.15)
    if op_pct is not None:
        level = operational.get("nivel_concentracao", "")
        _badge(s, Inches(6.9), Inches(1.05), Inches(3.4), Inches(0.42),
               f"Top 10: {op_pct}% ({level})",
               fill_hex=DANGER if level == "Crítico" else (WARN if level == "Médio" else SUCCESS), size=13)
    _add_bullets(s, Inches(6.9), body_top, Inches(5.8), Inches(4.5), op_items[:5], size=15, space_after=14)

    # ---- Slide 4: CFO / Strategic Assessment -------------------------------
    s = _blank_slide(prs)
    _page_title(s, "CFO / Strategic Assessment")

    fin_result = financial.get("resultado", "n/d")
    _badge(s, MARGIN, Inches(1.4), Inches(3.0), Inches(0.5), f"Financeiro: {fin_result}",
           fill_hex=DANGER if fin_result == "Crítico" else (WARN if fin_result == "Atenção" else SUCCESS), size=13)

    fin_lines = []
    bridge = financial.get("ebitda_bridge") or {}
    if bridge.get("ebitda_reportado") is not None:
        fin_lines.append(f"EBITDA: R$ {bridge['ebitda_reportado']:,.0f} ({bridge.get('margem_reportada_pct', '?')}%)".replace(",", "."))
    if bridge.get("ajustes"):
        fin_lines.append(f"EBITDA Ajustado: R$ {bridge.get('ebitda_ajustado', 0):,.0f} ({len(bridge['ajustes'])} ajuste(s) de não-recorrência)".replace(",", "."))
    for a in (financial.get("anomalias") or [])[:2]:
        fin_lines.append(f"Anomalia: {a.get('conta')} em {a.get('periodo')} ({a.get('narrativa', '')[:80]})")
    for flag in (financial.get("red_flags") or [])[:2]:
        fin_lines.append(f"Red flag ({flag.get('severidade')}): {flag.get('titulo')}")
    if fin_lines:
        _add_bullets(s, MARGIN, Inches(2.2), Inches(5.8), Inches(3.5), fin_lines, size=14, space_after=12)

    parecer = opinion.get("parecer_do_time", "")
    _add_textbox(s, Inches(6.9), Inches(1.4), Inches(5.8), Inches(0.4), "PARECER DO TIME", size=14, bold=True, color=CORAL)
    _add_textbox(s, Inches(6.9), Inches(1.9), Inches(5.8), Inches(3.8), parecer, size=15, color=TEXT_DARK)

    # ---- Slide 5: Recommendation & Next Steps ------------------------------
    s = _blank_slide(prs)
    _page_title(s, "Recommendation & Next Steps")
    _badge(s, MARGIN, Inches(1.5), Inches(4.2), Inches(0.7), recommendation,
           fill_hex=RECOMMENDATION_COLOR.get(recommendation, TEXT_MUTED), size=20)

    if synth_out.get("sumario_executivo"):
        _add_textbox(s, MARGIN, Inches(2.6), Inches(11.8), Inches(1.6),
                     synth_out["sumario_executivo"], size=16, color=TEXT_DARK)

    next_steps = synth_out.get("condicoes") or ["Nenhuma condição adicional."]
    _add_textbox(s, MARGIN, Inches(4.5), Inches(11.8), Inches(0.4), "PRÓXIMOS PASSOS", size=14, bold=True, color=CORAL)
    _add_bullets(s, MARGIN, Inches(5.0), Inches(11.8), Inches(2), next_steps, size=16, space_after=12)

    prs.save(path)
    print(f"PPT salvo: {path}")


def upload_to_storage(local_path: str, storage_path: str, content_type: str):
    """Sobe um arquivo pro bucket privado deal-outputs, via service_role
    (o frontend nunca escreve aqui, só lê depois com signed URL)."""
    url = os.environ["SUPABASE_URL"].rstrip("/") + "/storage/v1/object/deal-outputs/" + storage_path
    key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    with open(local_path, "rb") as f:
        data = f.read()
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": content_type,
        "x-upsert": "true",  # permite reprocessar sem dar erro de arquivo já existente
    }
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    with urllib.request.urlopen(req) as resp:
        print(f"Upload OK: {storage_path} ({resp.status})")


def record_output_link(deal_id: str, synthesis_run_id: str | None, excel_path: str, pptx_path: str):
    """Sobe os 2 arquivos pro Storage e grava as referências em public.outputs
    — é isso que o frontend usa para gerar o link de download real (signed
    URL), sem precisar passar pelo GitHub Actions."""
    if not os.environ.get("GITHUB_REPOSITORY"):
        print("Fora do GitHub Actions (teste local) — não sobe nem grava outputs.")
        return

    excel_storage_ref = f"{deal_id}/analise.xlsx"
    pptx_storage_ref = f"{deal_id}/executivo.pptx"
    upload_to_storage(
        excel_path, excel_storage_ref,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    upload_to_storage(
        pptx_path, pptx_storage_ref,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

    body = {
        "deal_id": deal_id,
        "synthesis_run_id": synthesis_run_id,
        "excel_ref": excel_storage_ref,
        "ppt_ref": pptx_storage_ref,
    }
    url = os.environ["SUPABASE_URL"].rstrip("/") + "/rest/v1/outputs"
    key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    headers = {"apikey": key, "Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    req = urllib.request.Request(url, data=json.dumps(body).encode(), headers=headers, method="POST")
    with urllib.request.urlopen(req) as resp:
        print(f"outputs gravado: {resp.status}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--deal-id", required=True)
    args = parser.parse_args()

    os.makedirs("output", exist_ok=True)
    bundle = fetch_deal_bundle(args.deal_id)

    excel_path = f"output/{args.deal_id}_analise.xlsx"
    pptx_path = f"output/{args.deal_id}_executivo.pptx"
    build_excel(bundle, excel_path)
    build_pptx(bundle, pptx_path)
    record_output_link(args.deal_id, (bundle.get("synthesis") or {}).get("id"), excel_path, pptx_path)
