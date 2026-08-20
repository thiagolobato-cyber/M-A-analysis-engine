"""
Parser de DRE/Balancete — a parte que o Thiago pediu pra virar "agente
especialista de baixo custo". A ideia não é um agente de IA mais esperto:
é reconhecer em CÓDIGO os dois formatos que se repetem entre planilhas
de escritórios diferentes (confirmado contra o arquivo real da Nacional
Controladoria em 20/08), e só chamar IA pra interpretar números que o
código já extraiu — nunca pra vasculhar a planilha inteira.

Os dois padrões:

1. BALANCETE CONSOLIDADO NUMA ABA SÓ (ex.: "BALANCETE 2025"): mesma
   informação que as abas MES 01..MES 12 já tratadas em run_agent.py,
   só que numa tabela só, com uma coluna por mês. Quando essa aba existe,
   ela é preferível às 12 abas mensais — mesmo dado, 1 leitura em vez de
   12, e elimina a duplicação que existia hoje (as duas eram enviadas).

2. DRE JÁ AGREGADA (aba "DRE"): ~40 linhas nomeadas (Receita, CMV,
   Despesas, Resultado) por mês — muito mais compacto e confiável que
   deixar o modelo tentar reconstruir isso a partir de 500 contas cruas
   do plano de contas.

Se NENHUM dos dois padrões for reconhecido, o código não trava — ele
deixa pro caminho antigo (excel_to_text) processar a aba como texto
genérico. Isso é o "fallback pra formato novo" que sustenta a promessa
de baixo custo sem quebrar em planilhas que a gente ainda não viu.

IMPORTANTE (bug real encontrado em 20/08 testando contra o arquivo da
Nacional Controladoria): `ws.max_row` vem **None** em modo `read_only`
para essa planilha — o export não declara a dimensão da aba no XML.
Qualquer código que dependa de `ws.max_row` pra limitar um loop quebra
silenciosamente (vira `range(1, 1)`, ou seja, zero linhas, sem erro
nenhum). Por isso este módulo usa `iter_rows(values_only=True)` em toda
parte — o mesmo padrão que `merge_monthly_balancete` já usava — que
funciona por streaming e não depende de `max_row` estar correto.
"""
from __future__ import annotations

import re

MESES_PT = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]


def _norm(v) -> str:
    return str(v).strip().lower() if v is not None else ""


def _find_month_header_row(ws, max_scan_rows: int = 10, min_months: int = 6):
    """Procura, nas primeiras linhas da aba, uma linha com pelo menos
    `min_months` nomes de mês em português — isso identifica o cabeçalho
    de qualquer tabela mensal, seja DRE ou Balancete consolidado, sem
    depender do nome da aba (que pode variar: "BALANCETE 2025",
    "Balancete Anual", "Consolidado" etc.) nem de `ws.max_row` (que pode
    vir None — ver nota do módulo)."""
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan_rows, values_only=True), start=1):
        row_vals = [_norm(v) for v in row]
        meses_encontrados = {v: i for i, v in enumerate(row_vals) if v in MESES_PT}
        if len(meses_encontrados) >= min_months:
            return r, row_vals, meses_encontrados
    return None


def detect_consolidated_balancete(wb):
    """Retorna um dict com a localização do Balancete consolidado (aba,
    linha de cabeçalho, colunas de código/descrição, mapa mês->coluna) se
    achar uma aba com todos os meses juntos. None se não achar (planilha
    só tem os moldes MES 01..12, ou formato totalmente diferente)."""
    for name in wb.sheetnames:
        ws = wb[name]
        found = _find_month_header_row(ws)
        if not found:
            continue
        header_row, row_vals, meses_col = found
        # A coluna de descrição normalmente se chama "descrição" — preferida.
        # "conta" só serve de fallback quando não existe "descrição" clara,
        # porque em planilhas como a testada em 20/08 "CONTA" é na verdade
        # a coluna do CÓDIGO da conta, não do nome (achado testando contra
        # o arquivo real: sem essa ordem de prioridade, "conta" era pego
        # primeiro e as contas saíam com o código no lugar do nome).
        col_descricao = next((i for i, v in enumerate(row_vals) if v in ("descrição", "descricao")), None)
        if col_descricao is None:
            col_descricao = next((i for i, v in enumerate(row_vals) if v == "conta"), None)
        col_codigo = next((i for i, v in enumerate(row_vals) if v in ("código", "codigo", "nº", "no", "cód", "conta")), None)
        if col_descricao is None:
            continue
        return {
            "aba": name, "linha_cabecalho": header_row,
            "col_codigo": col_codigo, "col_descricao": col_descricao,
            "meses_para_coluna": meses_col,
        }
    return None


def parse_consolidated_balancete(wb, deteccao: dict, max_data_rows: int = 5000) -> list[dict]:
    """Extrai {conta, codigo, valores} de um Balancete consolidado — mesma
    forma de saída que merge_monthly_balancete(), pra plugar direto onde
    select_relevant_accounts() já espera. Usa iter_rows por streaming —
    não depende de ws.max_row (ver nota do módulo)."""
    ws = wb[deteccao["aba"]]
    col_desc = deteccao["col_descricao"]
    col_cod = deteccao["col_codigo"]
    meses_col = deteccao["meses_para_coluna"]
    header_row = deteccao["linha_cabecalho"]

    contas = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + max_data_rows, values_only=True):
        if col_desc >= len(row):
            continue
        nome = row[col_desc]
        if nome is None or not str(nome).strip():
            continue
        codigo = row[col_cod] if col_cod is not None and col_cod < len(row) else None
        valores = {}
        for mes_idx, (mes_nome, col_idx) in enumerate(meses_col.items(), start=1):
            if col_idx < len(row):
                v = row[col_idx]
                if isinstance(v, (int, float)):
                    valores[f"mes_{mes_idx:02d}"] = v
        if valores:
            contas.append({"conta": str(nome).strip(), "codigo": codigo, "valores": valores})
    return contas


# Termos-chave que aparecem quase sempre numa DRE brasileira, na ordem em
# que normalmente aparecem — usado só pra dar um "apelido" (categoria)
# estável a cada linha, já que o texto exato varia (ex.: "Receitas
# Serviços" vs "Receita de Serviços" vs "Faturamento Serviços").
DRE_CATEGORIAS = [
    (re.compile(r"(?i)receita.*(serv|venda|faturamento)"), "receita_bruta"),
    (re.compile(r"(?i)impostos?\s*s[/.]?\s*(venda|serviço|faturamento)"), "impostos_sobre_receita"),
    (re.compile(r"(?i)receita.*l[ií]quida"), "receita_liquida"),
    (re.compile(r"(?i)^cmv$|custo.*(servi|mercadoria|venda)"), "cmv"),
    (re.compile(r"(?i)total.*despesas?\s*operacionais?"), "despesas_operacionais_total"),
    (re.compile(r"(?i)resultado.*(ap[óo]s|antes).*(ir|csll)"), "resultado_liquido"),
    (re.compile(r"(?i)margem\s*l[íi]quida"), "margem_liquida"),
]


def categorizar_linha_dre(rotulo: str) -> str | None:
    for regex, categoria in DRE_CATEGORIAS:
        if regex.search(rotulo):
            return categoria
    return None


def detect_dre_sheet(wb):
    """Mesma lógica de detecção de cabeçalho mensal, mas focada em achar
    uma aba de DRE (poucas dezenas de linhas nomeadas, não centenas de
    contas). Diferencia de Balancete pela presença de rótulos como
    'receita'/'despesa'/'resultado' logo abaixo do cabeçalho."""
    for name in wb.sheetnames:
        ws = wb[name]
        found = _find_month_header_row(ws)
        if not found:
            continue
        header_row, row_vals, meses_col = found
        col_rotulo = next((i for i, v in enumerate(row_vals) if v and v not in MESES_PT), None)
        if col_rotulo is None:
            continue
        # Confere se há linhas com termos típicos de DRE logo abaixo do
        # cabeçalho — evita casar com um Balancete de layout parecido.
        termos_dre_achados = 0
        for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + 60, values_only=True):
            if col_rotulo >= len(row):
                continue
            if categorizar_linha_dre(_norm(row[col_rotulo])):
                termos_dre_achados += 1
        if termos_dre_achados >= 3:
            return {"aba": name, "linha_cabecalho": header_row, "col_rotulo": col_rotulo, "meses_para_coluna": meses_col}
    return None


def parse_dre_sheet(wb, deteccao: dict, max_data_rows: int = 300) -> dict:
    """Extrai a DRE como {categoria_ou_rotulo_original: {mes_01: valor, ...}}.
    Linhas que batem numa categoria conhecida (DRE_CATEGORIAS) usam o nome
    estável da categoria; as demais mantêm o rótulo original (pra não
    perder informação específica do negócio, só não ganham um apelido)."""
    ws = wb[deteccao["aba"]]
    col_rotulo = deteccao["col_rotulo"]
    meses_col = deteccao["meses_para_coluna"]
    header_row = deteccao["linha_cabecalho"]

    linhas = {}
    for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + max_data_rows, values_only=True):
        if col_rotulo >= len(row):
            continue
        rotulo = row[col_rotulo]
        if rotulo is None or not str(rotulo).strip():
            continue
        rotulo_str = str(rotulo).strip()
        valores = {}
        for mes_idx, (mes_nome, col_idx) in enumerate(meses_col.items(), start=1):
            if col_idx < len(row):
                v = row[col_idx]
                if isinstance(v, (int, float)):
                    valores[f"mes_{mes_idx:02d}"] = v
        if not valores:
            continue
        categoria = categorizar_linha_dre(rotulo_str)
        chave = categoria or rotulo_str
        linhas[chave] = valores
    return linhas
