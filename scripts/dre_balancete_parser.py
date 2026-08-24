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
        # Prioridade importa: "código"/"cód"/"conta" são o código hierárquico
        # de verdade; "nº"/"no" é numeração sequencial de linha (1, 2, 3...),
        # não hierarquia — usar isso como código quebra qualquer lógica que
        # dependa de prefixo (leaf detection, distinção ativo/passivo vs
        # receita/despesa). Só cai pra "nº" se nada melhor existir.
        col_codigo = next((i for i, v in enumerate(row_vals) if v in ("código", "codigo", "cód")), None)
        if col_codigo is None:
            col_codigo = next((i for i, v in enumerate(row_vals) if v == "conta" and i != col_descricao), None)
        if col_codigo is None:
            col_codigo = next((i for i, v in enumerate(row_vals) if v in ("nº", "no")), None)
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
    (re.compile(r"(?i)deprecia[çc][ãa]o|amortiza[çc][ãa]o"), "d_a"),
    (re.compile(r"(?i)total.*resultado\s*financeiro|resultado\s*financeiro\s*l[íi]quido"), "resultado_financeiro"),
    (re.compile(r"(?i)resultado.*(ap[óo]s|antes).*(ir|csll)"), "resultado_liquido"),
    (re.compile(r"(?i)irpj|csll|tributos?\s*sobre\s*(o\s*)?lucro"), "tributos_sobre_lucro"),
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


def calcular_ebitda_de_dre(dre_estruturada: dict) -> dict:
    """EBITDA a partir da DRE já parseada (fonte PRIMÁRIA e confiável — sem
    risco de duplicar hierarquia, porque já veio como categoria de negócio,
    não conta contábil crua). Usa só as categorias padronizadas que
    `categorizar_linha_dre` já reconhece (ver dre_balancete_parser.py).

    Convenção de sinal: `resultado_financeiro` é POSITIVO quando é receita
    financeira líquida (convenção usada nas DREs testadas) — por isso EBITDA
    SUBTRAI esse valor (removendo o efeito não-operacional do lucro líquido),
    nunca soma. Testado contra dado real: bate exatamente com
    (receita_líquida − despesas_operacionais_total + D&A).

    Retorna None nos campos que a DRE não trouxe — nunca estima um valor que
    não pode ser calculado com o dado disponível."""
    def total_anual(chave):
        linha = dre_estruturada.get(chave)
        if not linha:
            return None
        valores = [v for v in linha.values() if isinstance(v, (int, float))]
        return sum(valores) if valores else None

    receita_liquida = total_anual("receita_liquida")
    despesas_op = total_anual("despesas_operacionais_total")
    d_a = total_anual("d_a") or 0.0
    resultado_financeiro = total_anual("resultado_financeiro") or 0.0
    tributos_sobre_lucro = total_anual("tributos_sobre_lucro") or 0.0
    lucro_liquido = total_anual("resultado_liquido")

    ebitda_bottom_up = None
    if receita_liquida is not None and despesas_op is not None:
        ebitda_bottom_up = round(receita_liquida - despesas_op + d_a, 2)

    ebitda_top_down = None
    if lucro_liquido is not None:
        ebitda_top_down = round(lucro_liquido - resultado_financeiro + tributos_sobre_lucro + d_a, 2)

    return {
        "ebitda_bottom_up_receita_menos_despesas": ebitda_bottom_up,
        "ebitda_top_down_lucro_liquido": ebitda_top_down,
        "diferenca_bottom_vs_top": (
            round(ebitda_bottom_up - ebitda_top_down, 2)
            if ebitda_bottom_up is not None and ebitda_top_down is not None else None
        ),
        "margem_ebitda_pct": round(100 * ebitda_bottom_up / receita_liquida, 1) if ebitda_bottom_up and receita_liquida else None,
    }


# Chaves agregadas/derivadas — são o RESULTADO de outras linhas, não uma
# causa-raiz. Apontar "resultado_liquido teve uma variação" não ajuda quem
# lê a análise; a linha de despesa/receita específica por trás, sim.
CHAVES_AGREGADAS_DRE = {
    "receita_bruta", "impostos_sobre_receita", "receita_liquida", "cmv",
    "despesas_operacionais_total", "resultado_financeiro",
    "tributos_sobre_lucro", "resultado_liquido", "margem_liquida",
}


def dre_linhas_para_contas(dre_estruturada: dict, incluir_agregados: bool = False) -> list[dict]:
    """Converte a DRE já parseada pro mesmo formato {conta, valores} que
    select_relevant_accounts/detectar_anomalias_run_rate já esperam — permite
    rodar a MESMA detecção de anomalia em cima das linhas já nomeadas da DRE
    (menos ruído que a conta crua do Balancete, porque já vem categorizada
    pela própria planilha da empresa).

    Por padrão, exclui as linhas agregadas (receita_líquida, resultado_líquido
    etc.) da lista — elas são o resultado de outras linhas, não uma causa-raiz
    de anomalia. Passe `incluir_agregados=True` só se quiser os totais também
    disponíveis pro Claude como contexto (não recomendado pra detecção)."""
    itens = dre_estruturada.items()
    if not incluir_agregados:
        itens = [(nome, valores) for nome, valores in itens if nome not in CHAVES_AGREGADAS_DRE]
    return [{"conta": nome, "codigo": None, "valores": valores} for nome, valores in itens]



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
