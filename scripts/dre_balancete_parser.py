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
from datetime import date

MESES_PT = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]

MESES_PT_ABREV = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}

# "01/2024", "01-2024", "1/24", "jan/24", "jan/2024" — formatos de data
# em TEXTO (não objeto datetime real) comuns em export de sistema de
# gestão. Grupo 1 = mês (número ou abreviação PT), grupo 2 = ano
# (2 ou 4 dígitos).
_RE_DATA_TEXTO = re.compile(
    r"^(?P<mes>0?[1-9]|1[0-2]|jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)"
    r"[/\-]"
    r"(?P<ano>\d{2}|\d{4})$"
)


def _norm(v) -> str:
    return str(v).strip().lower() if v is not None else ""


def _parse_marcador_mes(v) -> tuple[int, int] | None:
    """Tenta interpretar uma célula como marcador de mês, em QUALQUER um
    dos formatos observados na prática (achado real em 25/08, testando
    contra arquivo do Thiago + bateria de 1200 casos sintéticos):
    - Objeto de data real do Excel (datetime/date).
    - Nome de mês por extenso em português.
    - Nome de mês abreviado em português ("jan", "fev"...).
    - Texto no formato "MM/AAAA", "M/AA", "jan/24", "jan/2024" etc.
    Retorna (ano, mes) se reconhecer, None caso contrário — NUNCA levanta
    exceção (célula pode ser número solto, texto qualquer, etc.)."""
    if isinstance(v, date):
        return (v.year, v.month)
    texto = _norm(v)
    if not texto:
        return None
    if texto in MESES_PT:
        return (0, MESES_PT.index(texto) + 1)  # ano desconhecido — só ordem importa
    if texto in MESES_PT_ABREV:
        return (0, MESES_PT_ABREV[texto])  # abreviação sozinha, sem ano ("jan", "fev"...)
    m = _RE_DATA_TEXTO.match(texto)
    if m:
        mes_str, ano_str = m.group("mes"), m.group("ano")
        mes = MESES_PT_ABREV.get(mes_str, int(mes_str) if mes_str.isdigit() else None)
        if mes is None or not (1 <= mes <= 12):
            return None
        ano = int(ano_str)
        if ano < 100:
            ano += 2000
        return (ano, mes)
    return None


def _find_month_header_row(ws, max_scan_rows: int = 10, min_months: int = 6):
    """Procura, nas primeiras linhas da aba, uma linha com pelo menos
    `min_months` marcadores de mês — isso identifica o cabeçalho de
    qualquer tabela mensal, seja DRE ou Balancete consolidado, sem
    depender do nome da aba (que pode variar: "BALANCETE 2025",
    "Balancete Anual", "Consolidado" etc.) nem de `ws.max_row` (que pode
    vir None — ver nota do módulo).

    Reconhece DOIS formatos de marcador de mês na mesma linha de
    cabeçalho (achado real em 25/08, testando contra a DRE que o Thiago
    enviou e contra uma bateria de 1200 casos sintéticos): nome por
    extenso, nome abreviado, data real do Excel, ou texto "MM/AAAA" —
    ver `_parse_marcador_mes`, que cobre todos de uma vez.

    IMPORTANTE: esta função exige granularidade MENSAL de propósito —
    é usada tanto por DRE quanto por Balancete consolidado, e o
    Balancete alimenta cálculo que precisa de mês a mês (RBT12,
    waterfall). Pra DRE que só tem cabeçalho ANUAL, ver
    `_find_period_header_row` — não misturamos os dois aqui pra nunca
    fingir granularidade mensal onde só existe anual."""
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan_rows, values_only=True), start=1):
        row_vals = [_norm(v) for v in row]
        marcadores = {}
        indices_mes = set()  # TODOS os índices reconhecidos como marcador de
        # mês, mesmo quando duas colunas caem no mesmo ano-mês (achado real
        # em 25/08: planilha do Thiago tem uma coluna "Anual" repetindo a
        # data do último mês — sem isso, a coluna duplicada "vazava" como
        # candidata a coluna de rótulo, porque só sobrevive 1 no dict).
        for i, v in enumerate(row):
            parsed = _parse_marcador_mes(v)
            if parsed is not None:
                indices_mes.add(i)
                ano, mes = parsed
                chave = f"{ano:04d}-{mes:02d}" if ano else f"m{mes:02d}-{i:03d}"
                if chave not in marcadores:  # primeira ocorrência vence — a
                    marcadores[chave] = i    # segunda é duplicata/total, não mês novo
        if len(marcadores) >= min_months:
            row_vals_limpo = ["" if i in indices_mes else rv for i, rv in enumerate(row_vals)]
            return r, row_vals_limpo, marcadores
    return None


_RE_ANO_FISCAL = re.compile(r"^fy\s?'?(\d{2}|\d{4})$")


def _parse_marcador_ano(v) -> int | None:
    """Tenta interpretar uma célula como marcador de ANO (não mês) —
    achado real em 25/08 testando arquivos de deals reais: várias DREs
    reais são anuais (colunas "2021".."2025"), não mensais. Aceita
    número/float que pareça ano (faixa 1900-2100, sem casas decimais —
    isso evita confundir com valor monetário: um valor de receita
    raramente cai exatamente nessa faixa E sem decimal ao mesmo tempo,
    mas mesmo que caia, só conta como cabeçalho de ano se `min_anos`
    colunas bateram na mesma linha, o que uma linha de dados financeiros
    normais não faria por acaso), objeto de data (usa só o ano), ou
    "FY23"/"FY2023" (ano fiscal)."""
    if isinstance(v, date):
        return v.year
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if float(v) == int(v) and 1900 <= int(v) <= 2100:
            return int(v)
        return None
    texto = _norm(v)
    if not texto:
        return None
    if texto.isdigit() and 1900 <= int(texto) <= 2100:
        return int(texto)
    m = _RE_ANO_FISCAL.match(texto)
    if m:
        ano = int(m.group(1))
        return ano + 2000 if ano < 100 else ano
    return None


_RE_COLUNA_CALCULADA = re.compile(r"(?i)%|\bvar\b|varia[çc][ãa]o")


def _find_free_period_header_row(ws, max_scan_rows: int = 10, min_periodos: int = 2, max_scan_dados: int = 30):
    """Terceiro fallback de cabeçalho (depois de mensal e anual) — achado
    real em 25/08, arquivo real de deal (BPO Innova): a DRE compara dois
    PERÍODOS NOMEADOS LIVREMENTE ("2025 (Ano Completo)", "2026 (Q1
    Jan-Mar)"), não mês nem ano numérico puro — nenhum dos dois
    fallbacks anteriores reconhece isso. Aceita qualquer linha com pelo
    menos `min_periodos` células de texto que NÃO pareçam ser uma coluna
    CALCULADA (variação %, "Var" no nome) — essa exclusão importa: a
    mesma planilha real tinha uma 3ª coluna "% Var (2026 anualizado)"
    que, se tratada como período de dado bruto, misturaria percentual
    com valor monetário na soma (número sem sentido nenhum).

    IMPORTANTE (bug real corrigido em 25/08): também exige que a coluna
    tenha valores NUMÉRICOS reais nas linhas de dados logo abaixo — sem
    essa checagem, a própria coluna de RÓTULO ("CONTA / SUBCONTA", texto
    em toda linha) passava como se fosse mais um "período", porque
    também tem texto não-vazio no cabeçalho."""
    linhas_cache = list(ws.iter_rows(min_row=1, max_row=max_scan_rows + max_scan_dados, values_only=True))
    for idx, row in enumerate(linhas_cache[:max_scan_rows]):
        r = idx + 1
        row_vals = [_norm(v) for v in row]
        candidatos_texto = {}
        for i, v in enumerate(row_vals):
            if not v or _RE_COLUNA_CALCULADA.search(v):
                continue
            candidatos_texto[i] = f"periodo_{i:03d}_{v[:24]}"
        if len(candidatos_texto) < min_periodos:
            continue
        # Confirma que cada coluna candidata tem valor numérico de
        # verdade abaixo — só assim é período de dado, não rótulo.
        linhas_dados = linhas_cache[idx + 1: idx + 1 + max_scan_dados]
        candidatos = {}
        for i, chave in candidatos_texto.items():
            tem_numero = any(
                isinstance(dr[i], (int, float)) for dr in linhas_dados if i < len(dr)
            )
            if tem_numero:
                candidatos[chave] = i
        if len(candidatos) >= min_periodos:
            row_vals_limpo = ["" if i in candidatos.values() else rv for i, rv in enumerate(row_vals)]
            return r, row_vals_limpo, candidatos
    return None


def _find_period_header_row(ws, max_scan_rows: int = 10, min_months: int = 6, min_anos: int = 3):
    """Acha cabeçalho de período pra DRE — tenta MENSAL primeiro (mais
    granular, preferível sempre que existir); se não achar, cai pra
    ANUAL (achado real em 25/08: 3 dos 11 arquivos reais testados eram
    DRE anual, não mensal — sem esse fallback, ficavam 100% não
    reconhecidos); se ainda assim não achar, cai pra PERÍODO LIVRE (ver
    `_find_free_period_header_row` — achado real com um deal de verdade,
    BPO Innova, cujo cabeçalho não é mês nem ano). Retorna (linha,
    row_vals_limpo, periodos_col, granularidade) — granularidade é
    sempre marcada explicitamente no retorno, pra quem consumir NUNCA
    tratar coluna anual/livre como se fosse um mês (isso produziria
    número financeiro errado silenciosamente, ex.: RBT12 ou waterfall
    mensal aplicado em cima de total anual)."""
    encontrado = _find_month_header_row(ws, max_scan_rows, min_months)
    if encontrado:
        r, row_vals, periodos = encontrado
        return r, row_vals, periodos, "mensal"

    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan_rows, values_only=True), start=1):
        row_vals = [_norm(v) for v in row]
        marcadores = {}
        indices_ano = set()
        for i, v in enumerate(row):
            ano = _parse_marcador_ano(v)
            if ano is not None:
                indices_ano.add(i)
                chave = f"{ano:04d}"
                if chave not in marcadores:
                    marcadores[chave] = i
        if len(marcadores) >= min_anos:
            row_vals_limpo = ["" if i in indices_ano else rv for i, rv in enumerate(row_vals)]
            return r, row_vals_limpo, marcadores, "anual"

    encontrado_livre = _find_free_period_header_row(ws)
    if encontrado_livre:
        r, row_vals, periodos = encontrado_livre
        return r, row_vals, periodos, "periodo_livre"
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


def _inferir_coluna_rotulo(ws, header_row: int, meses_col: dict, max_scan_rows: int = 60) -> int | None:
    """Fallback pra quando o rótulo da conta NÃO está na própria linha de
    cabeçalho (caso real achado em 25/08: cabeçalho com datas reais tem
    só as datas na linha, o nome de cada conta — "RECEITAS", "TOTAL GERAL
    DE DESPESAS" etc. — fica em linhas ABAIXO, numa coluna à esquerda das
    colunas de valor). Escaneia as linhas de dados e escolhe, dentre as
    colunas ANTES da primeira coluna de mês, a que tem mais texto —
    essa é a coluna de rótulo na prática observada."""
    primeira_col_mes = min(meses_col.values())
    if primeira_col_mes == 0:
        return None
    contagem_texto = [0] * primeira_col_mes
    for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + max_scan_rows, values_only=True):
        for i in range(min(primeira_col_mes, len(row))):
            v = row[i]
            if isinstance(v, str) and v.strip():
                contagem_texto[i] += 1
    if not any(contagem_texto):
        return None
    return contagem_texto.index(max(contagem_texto))


def detect_dre_sheet(wb, min_linhas: int = 3, max_linhas: int = 400):
    """Mesma lógica de detecção de cabeçalho mensal, mas focada em achar
    uma aba de DRE (poucas dezenas de linhas nomeadas, não centenas de
    contas cruas de um Balancete).

    IMPORTANTE (corrigido em 25/08, achado rodando bateria de 1200 casos
    sintéticos): a confirmação de "isso é uma DRE" NÃO pode depender de
    bater com `DRE_CATEGORIAS` (nomenclatura fixa) — isso é o mesmo erro,
    na camada de detecção, que fazia a extração cair em modo genérico só
    porque a empresa usa nomenclatura fora do padrão ("RECEITAS" em vez
    de "Receita Líquida", por exemplo — caso real do Thiago em 25/08).
    Cada DRE é única na nomenclatura, mas a ESTRUTURA se repete: cabeçalho
    de mês OU ano (ver `_find_period_header_row`) + coluna de rótulo com
    um número razoável de linhas de conta (dezenas, não centenas — isso
    sim distingue de Balancete, sem depender de nenhuma palavra
    específica). A categorização semântica de cada linha (que precisa de
    DRE_CATEGORIAS ou de julgamento) fica inteiramente em
    `parse_dre_sheet` / na camada que vier depois — a detecção não deve
    mais barrar nada por causa disso."""
    candidatas = []
    for name in wb.sheetnames:
        ws = wb[name]
        found = _find_period_header_row(ws)
        if not found:
            continue
        header_row, row_vals, meses_col, granularidade = found
        # Prioridade: a inferência por linhas de DADOS é mais confiável —
        # a própria linha de cabeçalho pode ter texto residual que não é
        # a coluna de rótulo (achado real em 25/08: colunas "Anual" de
        # totalização no meio do cabeçalho da DRE do Thiago enganavam o
        # critério ingênuo de "primeira célula de texto na linha").
        col_rotulo = _inferir_coluna_rotulo(ws, header_row, meses_col)
        if col_rotulo is None:
            col_rotulo = next((i for i, v in enumerate(row_vals) if v and v not in MESES_PT), None)
        if col_rotulo is None:
            continue
        # Confirmação estrutural: conta linhas com rótulo de texto E pelo
        # menos 1 valor numérico numa coluna de mês — sem olhar o TEXTO do
        # rótulo. Uma DRE tem dezenas de linhas assim; um Balancete cru
        # tem centenas (é isso que `max_linhas` filtra).
        linhas_validas = 0
        for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + max_linhas + 50, values_only=True):
            if col_rotulo >= len(row):
                continue
            rotulo = row[col_rotulo]
            if not (rotulo and str(rotulo).strip()):
                continue
            tem_valor = any(
                isinstance(row[c], (int, float)) for c in meses_col.values() if c < len(row)
            )
            if tem_valor:
                linhas_validas += 1
        if min_linhas <= linhas_validas <= max_linhas:
            candidatas.append({
                "aba": name, "linha_cabecalho": header_row, "col_rotulo": col_rotulo,
                "meses_para_coluna": meses_col, "granularidade": granularidade,
                "linhas_validas": linhas_validas,
            })

    if not candidatas:
        return None

    # Quando mais de uma aba bate o critério estrutural (achado real em
    # 25/08: arquivos com Balanço Patrimonial + DRE, ou com aba "DRE"
    # sendo só um resumo/pivô e o detalhe de verdade estar noutra aba
    # chamada algo genérico como "Base de Dados") — escolhe rodando a
    # EXTRAÇÃO DE VERDADE em cada candidata e comparando o resultado
    # real (% de raízes classificadas), não um proxy. Tentei nome da aba
    # e contagem de linhas antes — os dois erraram em casos reais (a aba
    # com mais linhas às vezes é uma versão pior/duplicada, não a
    # detalhada de verdade), então não dá pra decidir sem rodar a
    # extração mesmo — o custo é local (sem IA), poucas abas candidatas
    # na prática, compensa.
    if len(candidatas) == 1:
        melhor = candidatas[0]
        del melhor["linhas_validas"]
        return melhor

    def _qualidade(c):
        hier = extrair_hierarquia_dre(wb, c)
        total = len(hier["raizes_classificadas"]) + len(hier["raizes_ambiguas"])
        pct = len(hier["raizes_classificadas"]) / total if total else 0
        return (pct, total)

    candidatas.sort(key=lambda c: _qualidade(c), reverse=True)
    melhor = candidatas[0]
    del melhor["linhas_validas"]
    return melhor


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


# ===========================================================================
# EXTRAÇÃO POR HIERARQUIA DE INDENTAÇÃO (achado real em 25/08)
# ===========================================================================
#
# `categorizar_linha_dre` (acima) depende de bater com nomenclatura
# conhecida ("Receita Líquida", "Total Despesas Operacionais"...) — e
# como o próprio Thiago apontou, cada DRE é única na nomenclatura, não
# dá pra cobrir por regex. Testando contra o arquivo real de um deal
# (CSF Hotelaria), a nomenclatura era toda diferente ("RECEITAS", "TOTAL
# GERAL DE DESPESAS", "(-) DESPESAS COM PESSOAL"...) — zero bateu com
# DRE_CATEGORIAS.
#
# O que NÃO varia de empresa pra empresa é a ESTRUTURA: quase toda DRE
# exportada de um sistema de gestão usa INDENTAÇÃO (espaços/tabs antes
# do texto) pra marcar hierarquia — linha-mãe (nível 0, sem indentação)
# vs. linha-detalhe (indentada, "filha" da linha-mãe mais próxima acima
# com nível menor). Isso é convenção de planilha, não de idioma/negócio.
#
# Testando contra o arquivo real: das ~90 linhas de detalhe, só ~11 são
# linha-mãe (nível 0) — são as únicas que realmente importam pra montar
# um EBITDA bridge; o resto é detalhe que soma pra cima. E dessas 11,
# ~10 já têm sinal MECÂNICO óbvio (prefixo "(-)" = despesa, "receita"/
# "faturamento" no nome = receita, "resultado"/"lucro" no nome = linha
# final) — sobra só 1 rótulo genuinamente ambíguo. Essa é a única parte
# que precisa de julgamento (IA mínima, ou pergunta direta), e o volume
# é ínfimo perto do dump bruto da planilha inteira (~250 caracteres
# contra ~35.000 do texto cru — 99%+ de redução).


def _nivel_indentacao(texto_bruto: str) -> int:
    """Conta espaços/tabs no início do texto — é a única fonte de
    hierarquia usada aqui, não depende de nenhuma palavra específica.
    Tab conta como 1 nível; múltiplos espaços contam 1 nível a cada 1
    espaço (testado contra o arquivo real: o padrão lá era 1 espaço por
    nível, não 2 ou 4 — mas o código não assume isso, só usa a CONTAGEM
    relativa entre linhas, então funciona também com 2/4 espaços por
    nível, desde que seja consistente dentro do mesmo arquivo)."""
    sem_indent = texto_bruto.lstrip(" \t")
    return len(texto_bruto) - len(sem_indent)


_RE_SINAL_AJUSTE = re.compile(r"^\(?\+/-\)?\s*")  # "(+/-)" — ajuste que pode ir nos dois sentidos, checar ANTES do sinal simples
_RE_SINAL = re.compile(r"^\(?([+\-=])\)?\s*")
_PALAVRAS_RECEITA = re.compile(r"(?i)receita|faturamento|entrada")
_PALAVRAS_DESPESA = re.compile(r"(?i)despesa|custo|imposto|tribut[oaá]")
_PALAVRAS_RESULTADO = re.compile(r"(?i)resultado|lucro|preju[íi]zo|ebitda|\bebit\b|margem")


def _classificar_por_sinal_mecanico(rotulo: str) -> str | None:
    """Classifica um rótulo de linha-mãe SEM olhar nomenclatura de
    negócio — só convenções estruturais/tipográficas que se repetem
    entre planilhas de empresas diferentes:
    - Prefixo "(-)"/"-" = despesa. "(+)"/"+" = receita. "(=)"/"="
      = linha calculada (subtotal ou resultado — Receita Líquida,
      EBITDA, Lucro Líquido todas usam "=", não dá pra saber qual sem
      olhar o nome; fica marcado como "resultado_calculado" de propósito,
      sem fingir mais certeza do que se tem). Achado real em 25/08
      (arquivos reais 06 e 08 usam esses 3 sinais explicitamente).
    - "(+/-)" = ajuste que pode ir nos dois sentidos (ex.: Resultado
      Financeiro, Resultado Não Operacional) — categoria PRÓPRIA
      ("ajuste"), nunca somada como receita nem despesa. Bug real
      corrigido em 25/08 (DRE real BPO Innova): sem checar "(+/-)" ANTES
      do sinal simples, o regex de "+" batia primeiro e classificava
      como receita — um "Resultado Financeiro" negativo então SUBTRAÍA
      da receita total, inflando artificialmente a margem calculada.
    - Presença de "receita"/"faturamento"/"entrada" = receita.
    - Presença de "resultado"/"lucro"/"prejuízo" = linha de resultado final.
    Retorna None se nenhum sinal bater — aí sim precisa de julgamento
    (IA mínima ou confirmação manual), não antes."""
    rotulo_strip = rotulo.strip()
    if _RE_SINAL_AJUSTE.match(rotulo_strip):
        return "ajuste"
    m = _RE_SINAL.match(rotulo_strip)
    if m:
        sinal = m.group(1)
        if sinal == "-":
            return "despesa"
        if sinal == "+":
            return "receita"
        if sinal == "=":
            return "resultado_calculado"
    if _PALAVRAS_RECEITA.search(rotulo):
        return "receita"
    if re.search(r"(?i)deprecia[çc][ãa]o|amortiza[çc][ãa]o", rotulo):
        return "d_a"  # subcategoria de despesa — separado pra poder somar de volta no EBITDA aproximado
    if _PALAVRAS_DESPESA.search(rotulo):
        return "despesa"
    if _PALAVRAS_RESULTADO.search(rotulo):
        return "resultado_final"
    return None


def _detectar_colunas_hierarquia(ws, header_row: int, meses_col: dict, col_rotulo_base: int, max_scan_rows: int = 300) -> list[int]:
    """Acha TODAS as colunas de texto candidatas a nível de hierarquia —
    tanto ANTES quanto DEPOIS das colunas de valor (achado real em
    25/08: no arquivo "Projeto Bravo" a coluna de categoria, "(-)
    TRIBUTOS", vem DEPOIS das colunas de valor FY23/24/25, não antes —
    sem olhar os dois lados, essa coluna nunca era vista). A ORDEM final
    (qual é "raiz" e qual é "detalhe") não é decidida aqui — é decidida
    por cardinalidade em `_extrair_hierarquia_por_colunas`, porque a
    posição física (esquerda/direita) não é confiável (ver lá).

    IMPORTANTE (bug real corrigido em 25/08, achado testando em produção
    com `read_only=True`): o limiar de "2% das linhas" não pode ser
    calculado sobre o número BRUTO de linhas que o iterator produziu —
    em modo normal (`read_only=False`), `ws.iter_rows(max_row=N)` sempre
    devolve N linhas, completando com linhas vazias até esse total; em
    `read_only=True`, ele para no fim real dos dados (pode ser bem menos
    que N). Isso fazia o mesmo arquivo produzir limiares diferentes
    (300 vs 225 linhas → limiar 6 vs 4) dependendo de qual modo carregou
    o workbook — e como só `run_agent.py` usa `read_only=True` de
    verdade, os testes locais (sem esse parâmetro) nunca pegavam isso. A
    correção: o denominador é o número de linhas com ALGUM dado real
    (qualquer célula não vazia, em qualquer coluna), que é estável nos
    dois modos — não o total de linhas que o iterator devolveu."""
    primeira_col_valor = min(meses_col.values())
    ultima_col_valor = max(meses_col.values())
    contagem: dict[int, int] = {}
    linhas_com_dado = 0
    for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + max_scan_rows, values_only=True):
        teve_dado_na_linha = False
        for i, v in enumerate(row):
            if v is not None and (not isinstance(v, str) or v.strip()):
                teve_dado_na_linha = True
            if primeira_col_valor <= i <= ultima_col_valor:
                continue  # coluna de valor, não candidata a hierarquia
            if isinstance(v, str) and v.strip():
                contagem[i] = contagem.get(i, 0) + 1
        if teve_dado_na_linha:
            linhas_com_dado += 1
    if linhas_com_dado == 0 or not contagem:
        return [col_rotulo_base]
    # Piso de 3 ocorrências, não 1 — uma célula de texto isolada (ex.:
    # "Realizado" aparecendo só numa linha, ruído/anotação solta na
    # planilha) nunca é padrão real de coluna de hierarquia. Achado real
    # em 25/08: no arquivo original do Thiago, uma célula "Realizado"
    # com 1 única ocorrência passava o limiar antigo (`max(1, ...)`
    # sempre aceitava >=1), fazendo o código escolher "colunas_separadas"
    # por engano nesse arquivo, que na verdade é hierarquia por
    # indentação numa coluna só.
    limiar = max(3, int(linhas_com_dado * 0.02))
    colunas = sorted(i for i, c in contagem.items() if c >= limiar)
    return colunas if colunas else [col_rotulo_base]


def _rotulos_legiveis_periodo(meses_col: dict) -> list[str]:
    """Extrai nomes de período LEGÍVEIS a partir das chaves internas de
    `meses_col`, na mesma ordem usada pra numerar mes_01/mes_02/... —
    sem isso, quem consome só via "mes_01"/"mes_02" genérico, mesmo
    quando o período real é "2025 (Ano Completo)"/"2026 (Q1 Jan-Mar)"
    (granularidade `periodo_livre`) ou um ano solto (`anual`). Pra
    granularidade `mensal` (chaves tipo "2023-01" ou "mMM-III"), o nome
    genérico "mes_01" já é claro o bastante — aqui só enriquece quando
    a chave carrega informação que "mes_NN" sozinho não tem."""
    rotulos = []
    for chave in meses_col:
        if chave.startswith("periodo_") and "_" in chave[8:]:
            # A chave interna vem normalizada (minúscula) — .title() bate
            # de volta com o texto original ("2025 (ano completo)" ->
            # "2025 (Ano Completo)"), sem precisar guardar o texto bruto
            # em outro lugar.
            rotulos.append(chave.split("_", 2)[-1].title())
        elif re.fullmatch(r"\d{4}", chave):
            rotulos.append(chave)  # granularidade anual — "2021", "2022"...
        else:
            rotulos.append(None)  # mensal — "mes_NN" sequencial já é claro
    return rotulos


def extrair_hierarquia_dre(wb, deteccao: dict, max_data_rows: int = 300) -> dict:
    """Reconstrói a hierarquia da DRE SEM nenhuma suposição de
    nomenclatura, em QUALQUER um dos dois formatos observados na prática
    (achado real em 25/08, testando 11 DREs de deals reais):

    1. INDENTAÇÃO numa coluna só (espaços/tabs antes do texto marcam
       nível — caso original, arquivo CSF Hotelaria).
    2. COLUNAS SEPARADAS pra cada nível de hierarquia (ex.: "Categoria" |
       "Descrição" — mais comum na amostra real: 4 de 11 arquivos). A
       coluna mais à esquerda com texto é a raiz; quando uma célula vem
       vazia, herda o último valor não-vazio visto naquela coluna
       (forward-fill — necessário pra arquivos onde a categoria só
       aparece na 1ª linha do bloco, ex.: EC 2026 revisado).

    Detecta automaticamente qual dos dois formatos o arquivo usa — não
    precisa configurar nada por arquivo.

    Já resolve por sinal mecânico tudo que dá — retorna só as poucas
    linhas-raiz que sobrarem genuinamente ambíguas, prontas pra uma
    chamada de IA mínima (ou confirmação manual), no lugar de mandar a
    planilha inteira pra IA reconstruir.

    Retorna:
    {
      "modo": "indentacao" | "colunas_separadas",
      "raizes_classificadas": {rotulo: {"tipo": "despesa"|"receita"|"resultado_final"|"resultado_calculado",
                                          "valores": {mes_01: ..., ...},
                                          "detalhes": [rotulos filhos]}},
      "raizes_ambiguas": [{"rotulo": str, "valores": {...}, "detalhes": [...]}],
      "total_linhas_lidas": int,
    }"""
    ws = wb[deteccao["aba"]]
    col_rotulo = deteccao["col_rotulo"]
    meses_col = deteccao["meses_para_coluna"]
    header_row = deteccao["linha_cabecalho"]

    colunas_hier = _detectar_colunas_hierarquia(ws, header_row, meses_col, col_rotulo, max_data_rows)

    resultado_indentacao = _extrair_hierarquia_por_indentacao(ws, header_row, meses_col, col_rotulo, max_data_rows)
    resultado_indentacao["modo"] = "indentacao"
    resultado_indentacao["periodos_rotulos"] = _rotulos_legiveis_periodo(meses_col)

    if len(colunas_hier) <= 1:
        return resultado_indentacao

    # IMPORTANTE (bug real corrigido em 25/08): tentar achar um limiar
    # perfeito pra decidir "essa coluna é hierarquia real ou ruído?" se
    # provou frágil — uma anotação esparsa do próprio usuário na
    # planilha (4 células soltas tipo "pessoal"/"software", sem relação
    # com estrutura nenhuma) passava qualquer limiar razoável, fazendo o
    # código escolher "colunas_separadas" por engano num arquivo que na
    # verdade é hierarquia por indentação simples. A correção: SEMPRE
    # calcular os dois modos possíveis e comparar o RESULTADO real (%
    # de raízes classificadas, igual já fazemos pra escolher entre abas
    # candidatas em `detect_dre_sheet`) — não tentar adivinhar de
    # antemão qual é o formato certo.
    resultado_colunas = _extrair_hierarquia_por_colunas(ws, header_row, meses_col, colunas_hier, max_data_rows)
    resultado_colunas["modo"] = "colunas_separadas"
    resultado_colunas.setdefault("hierarquia_confiavel", True)  # colunas explícitas = sempre confiável, nunca inferida

    def _qualidade(r):
        total = len(r["raizes_classificadas"]) + len(r["raizes_ambiguas"])
        pct = len(r["raizes_classificadas"]) / total if total else 0
        return (pct, total)

    vencedor = max(resultado_indentacao, resultado_colunas, key=_qualidade)
    vencedor["periodos_rotulos"] = _rotulos_legiveis_periodo(meses_col)
    return vencedor


def _extrair_valores_da_linha(row, meses_col: dict) -> dict:
    valores = {}
    for mes_idx, (_, col_idx) in enumerate(meses_col.items(), start=1):
        if col_idx < len(row):
            v = row[col_idx]
            if isinstance(v, (int, float)):
                valores[f"mes_{mes_idx:02d}"] = v
    return valores


def _extrair_hierarquia_por_indentacao(ws, header_row: int, meses_col: dict, col_rotulo: int, max_data_rows: int) -> dict:
    """Modo original (25/08, arquivo CSF Hotelaria) — hierarquia por
    contagem de espaço/tab numa única coluna de rótulo."""
    linhas_brutas = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + max_data_rows, values_only=True):
        if col_rotulo >= len(row):
            continue
        rotulo = row[col_rotulo]
        if rotulo is None or not str(rotulo).strip():
            continue
        texto_bruto = str(rotulo)
        nivel = _nivel_indentacao(texto_bruto)
        valores = _extrair_valores_da_linha(row, meses_col)
        if not valores:
            continue
        linhas_brutas.append({"rotulo": texto_bruto.strip(), "nivel": nivel, "valores": valores})

    if not linhas_brutas:
        return {"raizes_classificadas": {}, "raizes_ambiguas": [], "total_linhas_lidas": 0, "hierarquia_confiavel": True}

    # IMPORTANTE (bug real corrigido em 25/08, DRE real "BPO Innova"):
    # "nível mínimo = raiz" sozinho não basta quando a planilha mistura
    # dois tipos de linha no MESMO nível mínimo — subtotais/resultados
    # ("(=) SUBTOTAL X", nível 0) E linhas de despesa detalhada só um
    # pouco mais indentadas ("(-) Despesa com Pessoal...", nível 1).
    # Sem essa correção, a primeira linha "(=)..." de nível 0 (ex.: "(=)
    # RECEITA LÍQUIDA") "engolia" como detalhe TODAS as despesas de
    # nível 1 seguintes — que nunca eram somadas, derrubando o EBITDA
    # calculado pra uma fração do valor real. A correção: qualquer linha
    # com SINAL MECÂNICO PRÓPRIO ("(+)"/"(-)"/"(=)" no início do texto)
    # sempre vira raiz, não importa o nível de indentação — o sinal já é
    # uma marcação estrutural explícita de "isso é uma linha-chave",
    # mais confiável que indentação sozinha quando os dois sinais
    # coexistem. Continua usando só indentação (comportamento original)
    # quando a planilha não tem sinal nenhum (ex.: CSF Hotelaria).
    nivel_raiz = min(l["nivel"] for l in linhas_brutas)
    raizes = []
    atual = None
    for l in linhas_brutas:
        eh_raiz_por_nivel = l["nivel"] <= nivel_raiz
        eh_raiz_por_sinal = bool(_RE_SINAL.match(l["rotulo"]))
        if eh_raiz_por_nivel or eh_raiz_por_sinal:
            atual = {"rotulo": l["rotulo"], "valores": l["valores"], "detalhes": []}
            raizes.append(atual)
        elif atual is not None:
            atual["detalhes"].append(l["rotulo"])
        else:
            atual = {"rotulo": l["rotulo"], "valores": l["valores"], "detalhes": []}
            raizes.append(atual)

    # Sinal de baixa confiança (bug real achado em 25/08, arquivo "DRE
    # Gerencial"): se TODAS as linhas ficaram no mesmo nível (nenhuma
    # indentação real na planilha), pode existir hierarquia IMPLÍCITA
    # que o código não consegue ver (ex.: "SG&A" sendo o TOTAL de
    # "Pessoal"+"Sócios"+"S&M"+"G&A" logo abaixo, sem nenhum espaço/tab
    # nem coluna separada marcando isso — só visível por negrito/ordem
    # visual, que célula bruta não carrega). Somar tudo nesse caso
    # arrisca dupla contagem (total + detalhes juntos). Não tentamos
    # adivinhar — só marcamos a confiança como baixa e deixamos quem
    # consome decidir (ex.: preferir o caminho antigo/`parse_dre_sheet`
    # se disponível, ou pedir confirmação manual antes de usar o EBITDA
    # aproximado como número final).
    sem_hierarquia_real = len(set(l["nivel"] for l in linhas_brutas)) <= 1 and len(linhas_brutas) > 1

    return {**_classificar_raizes(raizes, len(linhas_brutas)), "hierarquia_confiavel": not sem_hierarquia_real}


def _extrair_hierarquia_por_colunas(ws, header_row: int, meses_col: dict, colunas_hier: list[int], max_data_rows: int) -> dict:
    """Modo colunas separadas (achado real em 25/08, maioria dos
    arquivos reais testados) — cada coluna em `colunas_hier` é um nível
    de hierarquia. A ORDEM (qual é raiz, qual é detalhe) é decidida por
    CARDINALIDADE, não posição física na planilha: a coluna com MENOS
    valores distintos é mais genérica (raiz — ex.: "(-) TRIBUTOS" se
    repete em poucas dezenas de linhas), a com MAIS valores distintos é
    mais específica (detalhe — ex.: "COFINS", "ISS", "PIS"... uma linha
    cada). Isso generaliza mesmo quando a coluna de categoria vem
    DEPOIS das colunas de valor na planilha (achado real: arquivo
    "Projeto Bravo"), onde a posição física sozinha erraria a ordem.
    Forward-fill nas células vazias (a categoria-mãe às vezes só aparece
    na 1ª linha do bloco, não em toda linha)."""
    # Primeiro passo: ler todas as linhas cruas (sem decidir ordem ainda).
    linhas_cruas = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + max_data_rows, values_only=True):
        valores = _extrair_valores_da_linha(row, meses_col)
        if not valores:
            continue
        textos = {c: (str(row[c]).strip() if c < len(row) and row[c] is not None and str(row[c]).strip() else None)
                  for c in colunas_hier}
        linhas_cruas.append({"textos": textos, "valores": valores})

    if not linhas_cruas:
        return {"raizes_classificadas": {}, "raizes_ambiguas": [], "total_linhas_lidas": 0, "hierarquia_confiavel": True}

    # Forward-fill por coluna, na ordem em que as linhas aparecem.
    ultimo_valor = {c: None for c in colunas_hier}
    for l in linhas_cruas:
        for c in colunas_hier:
            if l["textos"][c] is not None:
                ultimo_valor[c] = l["textos"][c]
            l["textos"][c] = ultimo_valor[c]
        # reset do forward-fill não é necessário entre blocos — cada
        # coluna carrega seu último valor visto até a próxima mudança,
        # que é exatamente o comportamento de planilha com célula
        # mesclada visualmente (mesma lógica que Excel usa pra exibir).
        ultimo_valor = {c: l["textos"][c] for c in colunas_hier}

    # Cardinalidade por coluna (após forward-fill) — decide ordem raiz->detalhe.
    cardinalidade = {c: len({l["textos"][c] for l in linhas_cruas if l["textos"][c] is not None}) for c in colunas_hier}
    colunas_por_generalidade = sorted(colunas_hier, key=lambda c: cardinalidade[c])
    col_raiz = colunas_por_generalidade[0]
    col_detalhe = colunas_por_generalidade[-1]

    raizes_dict: dict[str, dict] = {}
    for l in linhas_cruas:
        raiz = l["textos"][col_raiz]
        detalhe = l["textos"][col_detalhe]
        if raiz is None:
            continue
        if raiz not in raizes_dict:
            raizes_dict[raiz] = {"rotulo": raiz, "valores": {}, "detalhes": []}
        for k, v in l["valores"].items():
            raizes_dict[raiz]["valores"][k] = raizes_dict[raiz]["valores"].get(k, 0) + v
        if detalhe and detalhe != raiz:
            raizes_dict[raiz]["detalhes"].append(detalhe)

    return _classificar_raizes(list(raizes_dict.values()), len(linhas_cruas))


def _classificar_raizes(raizes: list[dict], total_linhas_lidas: int) -> dict:
    """Último passo comum aos dois modos: aplica o classificador
    mecânico e separa o que sobrar sem sinal — só isso precisa de
    julgamento (IA mínima ou confirmação manual)."""
    raizes_classificadas = {}
    raizes_ambiguas = []
    for r in raizes:
        tipo = _classificar_por_sinal_mecanico(r["rotulo"])
        if tipo:
            raizes_classificadas[r["rotulo"]] = {
                "tipo": tipo, "valores": r["valores"], "detalhes": r["detalhes"],
            }
        else:
            raizes_ambiguas.append(r)

    return {
        "raizes_classificadas": raizes_classificadas,
        "raizes_ambiguas": raizes_ambiguas,
        "total_linhas_lidas": total_linhas_lidas,
    }


def calcular_resultado_de_hierarquia(hierarquia: dict) -> dict:
    """Resultado operacional e EBITDA APROXIMADO a partir do resultado de
    `extrair_hierarquia_dre` — usado como FALLBACK quando o caminho antigo
    (`detect_dre_sheet` + `parse_dre_sheet` + `calcular_ebitda_de_dre`, que
    usa `DRE_CATEGORIAS` fino: receita_liquida, cmv, d_a etc. separados)
    não reconhece a nomenclatura da DRE. Esta versão é deliberadamente mais
    grosseira — soma tudo que foi classificado como "receita" menos tudo
    que foi classificado como "despesa" (D&A incluído nessa soma, porque é
    uma despesa operacional de verdade).

    IMPORTANTE (bug real corrigido em 25/08, achado testando 11 arquivos
    reais): não existe convenção universal de sinal — algumas DREs
    armazenam despesa como número NEGATIVO na própria célula (ex.: Grupo
    Roma, "-5.098.345,10"), outras como POSITIVO com o sinal só no rótulo
    (ex.: CSF Hotelaria, "(-) DESPESAS COM PESSOAL" = 9.619.143,08
    positivo). Assumir uma convenção fixa deu margem de 26.000%+ num
    arquivo real. A correção: detecta a convenção pela soma agregada das
    linhas de despesa (predominantemente negativa ou positiva) e ajusta a
    fórmula — nunca força um sinal, só observa o que já está lá.

    NÃO diferencia impostos sobre venda de impostos sobre lucro (ambos
    entram como "despesa") — por isso é uma APROXIMAÇÃO, nunca deve
    substituir silenciosamente o cálculo fino quando ele está disponível.
    Linhas em `raizes_ambiguas` (sem classificação) NÃO entram nessa soma
    — ficam de fora até serem resolvidas (IA mínima ou confirmação
    manual), nunca silenciosamente ignoradas sem registro (ver
    `linhas_nao_somadas` no retorno)."""
    receita_total: dict[str, float] = {}
    despesa_bruta: dict[str, float] = {}  # sinal ORIGINAL da planilha, sem normalizar ainda
    d_a_bruto: dict[str, float] = {}

    for rotulo, dados in hierarquia["raizes_classificadas"].items():
        alvo = {"receita": receita_total, "despesa": despesa_bruta, "d_a": d_a_bruto}.get(dados["tipo"])
        if alvo is None:  # resultado_final / resultado_calculado — não soma, é linha derivada
            continue
        for mes, v in dados["valores"].items():
            alvo[mes] = alvo.get(mes, 0) + v
        if dados["tipo"] == "d_a":  # D&A também é despesa operacional pro resultado bruto
            for mes, v in dados["valores"].items():
                despesa_bruta[mes] = despesa_bruta.get(mes, 0) + v

    soma_receita = sum(receita_total.values())
    soma_despesa_bruta = sum(despesa_bruta.values())
    soma_d_a_bruto = sum(d_a_bruto.values())
    receita_fonte = "capturada_direto" if receita_total else "ausente"

    # Checagem de sanidade (bug real achado em 25/08, arquivo "DRE
    # Gerencial"): quando a única linha de receita de verdade vem marcada
    # com sinal "(=)" (ex.: "(=) RECEITA BRUTA" — é um subtotal calculado,
    # mas também é a fonte primária de receita numa DRE resumida sem
    # linha de detalhe separada), ela cai em "resultado_calculado" e não
    # é somada — a receita capturada fica artificialmente pequena (só
    # "outras receitas"/"receitas financeiras"), dando margem sem
    # sentido (-26.000%+ no caso real). Se a receita capturada for menor
    # que 20% da despesa (sinal forte de que a receita "de verdade" não
    # foi capturada), busca entre as linhas "(=)" uma que contenha
    # "receita" (prioridade: líquida > bruta) e usa o valor dela — SEMPRE
    # sinalizando que foi inferência, nunca escondendo.
    if soma_despesa_bruta and abs(soma_receita) < 0.2 * abs(soma_despesa_bruta):
        candidatas_receita = [
            r for r in hierarquia["raizes_classificadas"].items()
            if r[1]["tipo"] == "resultado_calculado" and _PALAVRAS_RECEITA.search(r[0])
        ]
        candidatas_receita.sort(key=lambda r: 0 if "líquida" in r[0].lower() or "liquida" in r[0].lower() else 1)
        if candidatas_receita:
            rotulo_receita, dados_receita = candidatas_receita[0]
            soma_receita = sum(dados_receita["valores"].values())
            receita_fonte = f"inferida_de_resultado_calculado:{rotulo_receita}"

    # Detecta a convenção de sinal pela soma agregada — se as despesas já
    # vêm negativas na planilha, SOMAR (o sinal já embute a subtração);
    # se vêm positivas, SUBTRAIR (fórmula clássica). Convenção mista
    # linha-a-linha (ex.: "Recuperação de Despesas" positiva dentro de um
    # grupo majoritariamente negativo) já fica correta ao usar a soma
    # agregada com seu sinal original, sem re-inverter nada.
    if soma_despesa_bruta < 0:
        soma_resultado_operacional = soma_receita + soma_despesa_bruta
        soma_d_a_para_ebitda = -soma_d_a_bruto if soma_d_a_bruto < 0 else soma_d_a_bruto
    else:
        soma_resultado_operacional = soma_receita - soma_despesa_bruta
        soma_d_a_para_ebitda = soma_d_a_bruto

    ebitda_aproximado = soma_resultado_operacional + soma_d_a_para_ebitda
    tem_dado = bool(receita_total or despesa_bruta) or receita_fonte.startswith("inferida")

    return {
        "receita_total": round(soma_receita, 2) if tem_dado else None,
        "receita_total_fonte": receita_fonte,
        "despesa_total": round(abs(soma_despesa_bruta), 2) if despesa_bruta else None,
        "d_a_total": round(abs(soma_d_a_bruto), 2) if d_a_bruto else None,
        "resultado_operacional_total": round(soma_resultado_operacional, 2) if tem_dado else None,
        "ebitda_aproximado": round(ebitda_aproximado, 2) if tem_dado else None,
        "margem_operacional_pct": round(100 * soma_resultado_operacional / soma_receita, 1) if soma_receita else None,
        "convencao_sinal_despesa_detectada": "negativo" if soma_despesa_bruta < 0 else "positivo",
        "linhas_nao_somadas": [r["rotulo"] for r in hierarquia["raizes_ambiguas"]],
        # Achado real em 25/08 (DRE "BPO Innova"): "margem_operacional_pct"
        # acima mistura TODOS os níveis de despesa numa métrica só (do
        # custo direto até tributos) — não corresponde a "Margem Bruta"
        # nem "Margem EBITDA" especificamente, só a uma aproximação
        # ampla. Quando a planilha já tem seus PRÓPRIOS subtotais/
        # resultados calculados ("=" — Receita Líquida, Resultado
        # Operacional, Lucro Líquido etc.), eles são MAIS CONFIÁVEIS que
        # qualquer soma nossa (vêm prontos da fonte, sem risco de
        # duplicar ou misturar nível de agregação errado). Em vez de
        # tentar adivinhar qual desses é "Margem Bruta" vs "EBITDA" vs
        # "Líquida" (isso varia de empresa pra empresa e exige juízo de
        # negócio), entrega todos eles estruturados — o agente
        # (financial_analysis) já demonstrou competência em interpretar
        # isso corretamente num teste real (CSF Hotelaria, 25/08).
        "linhas_resultado_da_fonte": [
            {"rotulo": rotulo, "valores": dados["valores"]}
            for rotulo, dados in hierarquia["raizes_classificadas"].items()
            if dados["tipo"] == "resultado_calculado"
        ],
        "fonte": "aproximacao_hierarquia_v1",  # nunca confundir com "rule_engine" fino — é aproximação
        "hierarquia_confiavel": hierarquia.get("hierarquia_confiavel", True),
    }


_RE_MB_RECEITA_BRUTA = re.compile(r"(?i)^\W*receitas?\W*$|receita.*(bruta|serv|venda|faturamento)")
_RE_MB_RECEITA_LIQUIDA = re.compile(r"(?i)receita.*l[íi]quida")
_RE_MB_DESPESA_PESSOAL = re.compile(r"(?i)despesa.*pessoal|folha\s*de\s*pagamento|custo.*folha")
_RE_MB_CUSTO_SISTEMAS = re.compile(r"(?i)custo.*sistemas?(?!.*financeiro)|servi[çc]os?\s*de\s*sistema")
_RE_MB_DEDUCAO_RECEITA = re.compile(r"(?i)dedu[çc][ãa]o.*receita|impostos?\s*s[/.]?\s*(venda|serviço|faturamento)|pis.*cofins")


def _achar_linha_por_padrao(fonte: dict, padrao: re.Pattern, formato: str) -> dict | None:
    """Busca uma linha por PALAVRA-CHAVE no rótulo — funciona tanto no
    formato de `dre_estruturada` (caminho fino: {rótulo_ou_categoria:
    {mes: valor}}) quanto em `raizes_classificadas` (hierarquia:
    {rótulo: {"tipo":..., "valores": {mes: valor}}}).

    IMPORTANTE (bug real corrigido em 26/08, achado testando a DRE do
    BPO Innova): quando MAIS DE UMA linha bate o mesmo padrão (ex.:
    "Custo Sistemas" batendo tanto "Serviços de Sistema" quanto "Custos
    Sistemas" QUANTO o "SUBTOTAL CUSTO COM SISTEMAS" que já soma os
    dois), pegar a primeira cegamente subestima o valor — o subtotal é
    sempre a versão mais completa/agregada quando existe. Por isso
    prioriza qualquer candidato com "subtotal"/"total" no rótulo; só cai
    pro primeiro candidato "normal" se nenhum subtotal bater."""
    candidatos = [(rotulo, dado) for rotulo, dado in fonte.items() if padrao.search(rotulo)]
    if not candidatos:
        return None
    subtotais = [c for c in candidatos if re.search(r"(?i)subtotal|total\s*geral", c[0])]
    rotulo, dado = (subtotais or candidatos)[0]
    return dado["valores"] if formato == "hierarquia" else dado


def extrair_margem_bruta_de_dre(dre_estruturada: dict | None, hierarquia: dict | None, regime_tributario: str | None) -> dict | None:
    """Calcula Margem Bruta e Custo Folha % DIRETO DA DRE — decisão
    explícita do Thiago em 26/08: "quem manda é a DRE, não o
    formulário" pra margem. Antes, o Bloco A sempre usava
    faturamento_mensal/folha_informada/custo_sistemas do FORMULÁRIO —
    testando contra o BPO Innova real, isso deu 46,89% (formulário) vs.
    55%/58% que o time humano já tinha validado usando a DRE.

    Busca as 4 linhas necessárias (Receita Bruta, Despesa com Pessoal,
    Custo Sistemas, Dedução de Receita) por PALAVRA-CHAVE no rótulo
    original — tenta primeiro no caminho fino (`dre_estruturada`, mais
    confiável quando reconhece), senão na hierarquia (`raizes_classificadas`
    do fallback por indentação/colunas). Calcula margem POR PERÍODO (não
    soma anos diferentes juntos — uma DRE com 2025 anual + 2026
    trimestral nunca deveria ter os dois períodos somados numa margem
    só, distorce a métrica).

    Retorna None se não achar TODAS as 4 linhas necessárias em nenhuma
    das duas fontes — nesse caso, quem chama deve cair pro formulário
    como fallback, não travar."""
    fontes = []
    if dre_estruturada:
        fontes.append((dre_estruturada, "fino"))
    if hierarquia and hierarquia.get("raizes_classificadas"):
        fontes.append((hierarquia["raizes_classificadas"], "hierarquia"))

    for fonte, formato in fontes:
        receita = _achar_linha_por_padrao(fonte, _RE_MB_RECEITA_BRUTA, formato)
        if receita is None:
            receita = _achar_linha_por_padrao(fonte, _RE_MB_RECEITA_LIQUIDA, formato)
        despesa_pessoal = _achar_linha_por_padrao(fonte, _RE_MB_DESPESA_PESSOAL, formato)
        custo_sistemas = _achar_linha_por_padrao(fonte, _RE_MB_CUSTO_SISTEMAS, formato)
        deducao = _achar_linha_por_padrao(fonte, _RE_MB_DEDUCAO_RECEITA, formato)

        if receita is None or despesa_pessoal is None or custo_sistemas is None:
            continue  # não achou o mínimo necessário nesta fonte — tenta a próxima

        margens_por_periodo = {}
        custo_folha_por_periodo = {}
        for periodo, valor_receita in receita.items():
            if not valor_receita:
                continue
            vd = abs(despesa_pessoal.get(periodo, 0))
            vc = abs(custo_sistemas.get(periodo, 0))
            vi = abs(deducao.get(periodo, 0)) if deducao else 0
            margens_por_periodo[periodo] = round(100 * (valor_receita - vd - vc - vi) / valor_receita, 2)
            custo_folha_por_periodo[periodo] = round(100 * vd / valor_receita, 2)

        if margens_por_periodo:
            return {
                "margem_bruta_pct_por_periodo": margens_por_periodo,
                "custo_folha_pct_por_periodo": custo_folha_por_periodo,
                "fonte": formato,
            }
    return None


def montar_mini_dre(dre_estruturada: dict | None, hierarquia: dict | None, periodos_rotulos: list | None = None) -> dict | None:
    """Monta uma DRE simplificada (Receita Bruta -> Deduções -> Receita
    Líquida -> Despesas -> Resultado), por período — pra exibição direta
    em PPT/Excel, sem depender do agente de IA reconstruir isso no seu
    texto (achado real em 26/08: pedido do Thiago pra ter uma "mini DRE"
    visível no slide financeiro, no estilo do teaser de M&A).

    Determinística, 100% código — reaproveita a mesma busca por
    palavra-chave já usada em `extrair_margem_bruta_de_dre` pra achar
    Receita Bruta/Líquida, Despesa com Pessoal, Custo Sistemas e Dedução
    de Receita; o restante das despesas (tudo que não caiu nessas 4
    linhas) fica agrupado como "Outras Despesas Operacionais", calculado
    por diferença a partir do resultado que a própria planilha já
    reportou (`linhas_resultado_da_fonte`) — nunca inventa um número
    novo, só reorganiza o que já foi extraído.

    Retorna None se não achar o mínimo (receita + pelo menos 1 linha de
    resultado calculado pela própria fonte) — quem chama cai pro EBITDA
    Bridge clássico (bottom-up) nesse caso."""
    fontes = []
    if dre_estruturada:
        fontes.append((dre_estruturada, "fino"))
    raizes_classificadas = (hierarquia or {}).get("raizes_classificadas")
    if raizes_classificadas:
        fontes.append((raizes_classificadas, "hierarquia"))

    for fonte, formato in fontes:
        receita = _achar_linha_por_padrao(fonte, _RE_MB_RECEITA_BRUTA, formato)
        if receita is None:
            receita = _achar_linha_por_padrao(fonte, _RE_MB_RECEITA_LIQUIDA, formato)
        if receita is None:
            continue
        deducao = _achar_linha_por_padrao(fonte, _RE_MB_DEDUCAO_RECEITA, formato)
        despesa_pessoal = _achar_linha_por_padrao(fonte, _RE_MB_DESPESA_PESSOAL, formato)
        custo_sistemas = _achar_linha_por_padrao(fonte, _RE_MB_CUSTO_SISTEMAS, formato)

        # Resultado final: prioriza a ÚLTIMA linha "(=)" da hierarquia
        # (mais próxima de Lucro Líquido); se só tiver dre_estruturada
        # (caminho fino), usa resultado_liquido/margem_liquida quando
        # reconhecidos, senão calcula por diferença.
        resultado = None
        if formato == "hierarquia" and raizes_classificadas:
            resultados_calc = [
                dados["valores"] for rotulo, dados in raizes_classificadas.items()
                if dados["tipo"] == "resultado_calculado"
            ]
            if resultados_calc:
                resultado = resultados_calc[-1]
        if resultado is None:
            resultado = fonte.get("resultado_liquido")
        despesas_op_total = fonte.get("despesas_operacionais_total") if formato == "fino" else None
        if resultado is None and despesas_op_total:
            # Fallback: caminho fino reconhece "despesas_operacionais_total"
            # (subtotal completo) mas não teve uma linha "resultado_liquido"
            # reconhecível — calcula resultado = receita_líquida - despesas
            # totais (mesma lógica do EBITDA bottom-up já usado em
            # `calcular_ebitda_de_dre`, só reaproveitada aqui).
            resultado = {
                periodo: receita.get(periodo, 0) - abs(deducao.get(periodo, 0) if deducao else 0) - abs(v)
                for periodo, v in despesas_op_total.items()
            }

        linhas = []
        periodos = list(receita.keys())
        for i, periodo in enumerate(periodos):
            rotulo_periodo = periodos_rotulos[i] if periodos_rotulos and i < len(periodos_rotulos) and periodos_rotulos[i] else periodo
            v_receita = receita.get(periodo, 0)
            v_deducao = abs(deducao.get(periodo, 0)) if deducao else 0
            v_pessoal = abs(despesa_pessoal.get(periodo, 0)) if despesa_pessoal else 0
            v_sistemas = abs(custo_sistemas.get(periodo, 0)) if custo_sistemas else 0
            v_resultado = resultado.get(periodo) if resultado else None
            receita_liquida = v_receita - v_deducao
            if v_resultado is not None:
                v_outras = receita_liquida - v_pessoal - v_sistemas - v_resultado
            else:
                v_outras = None
            linhas.append({
                "periodo": rotulo_periodo,
                "receita_bruta": round(v_receita, 2),
                "deducoes": -round(v_deducao, 2),
                "receita_liquida": round(receita_liquida, 2),
                "despesa_pessoal": -round(v_pessoal, 2) if despesa_pessoal else None,
                "custo_sistemas": -round(v_sistemas, 2) if custo_sistemas else None,
                "outras_despesas": -round(v_outras, 2) if v_outras is not None else None,
                "resultado": round(v_resultado, 2) if v_resultado is not None else None,
                "margem_pct": round(100 * v_resultado / v_receita, 1) if v_resultado is not None and v_receita else None,
            })

        if linhas and any(l["resultado"] is not None for l in linhas):
            return {"linhas": linhas, "fonte": formato}
    return None
