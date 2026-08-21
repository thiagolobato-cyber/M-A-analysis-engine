"""
financial_engine.py — motor determinístico de análise financeira, baseado na
metodologia consolidada em Instrucoes_Projeto_Claude.md (aprendida pelo
Thiago em múltiplas análises reais: BPO Innova, Cifali, Pretorian, Ruhling,
Contabilivre, BWA 360).

Objetivo: fazer em código o que é regra e cálculo (waterfall, EBITDA,
anomalia por limiar), deixando pro Claude só a interpretação final e curta —
não a descoberta. Isso ataca o lado caro (a IA "pensando" sobre 300+ contas),
não só o tamanho da prosa de saída.

IMPORTANTE (confirmado pelo Thiago em 21/08): DRE/Balancete muda de perfil
pra empresa — nome de conta, nível de hierarquia e nomenclatura variam. Este
módulo NUNCA assume nome de conta fixo — usa correspondência por palavra-
chave (regex), com uma categoria explícita "nao_mapeado" pra tudo que não
bate com confiança, em vez de arriscar categorizar errado e entregar um
número que parece certo mas não é. Mesma filosofia do
`extraction_prompt_addition.txt`: null/não-mapeado é sempre preferível a
chute.
"""
from __future__ import annotations

import re
import statistics


# ---------------------------------------------------------------------------
# 1. WATERFALL — mapeamento de conta para categoria gerencial
# ---------------------------------------------------------------------------
# Cada entrada: (regex, categoria). Ordem importa — a primeira que bater vence,
# então padrões mais específicos vêm antes dos genéricos (ex.: "pró-labore"
# antes de "pessoal", senão pró-labore cairia em OPERACIONAL por engano).
WATERFALL_CATEGORIAS = [
    (re.compile(r"(?i)comiss[ãa]o"), "OPERACIONAL"),           # sempre sai do operacional, nunca some com custo direto genérico
    (re.compile(r"(?i)pr[óo].?labore"), "DESPESAS_COM_SOCIOS"),
    (re.compile(r"(?i)distribui[çc][ãa]o.*lucro"), "DESPESAS_COM_SOCIOS"),
    (re.compile(r"(?i)pis|cofins|^iss$|iss\s|simples\s*nacional"), "TRIBUTOS"),
    (re.compile(r"(?i)irpj|csll"), "TRIBUTOS_SOBRE_LUCRO"),
    (re.compile(r"(?i)deprecia[çc][ãa]o|amortiza[çc][ãa]o"), "D_A"),
    (re.compile(r"(?i)multa"), "DESPESAS_INDEDUTIVEIS"),
    (re.compile(r"(?i)sistema|software|licen[çc]a"), "CUSTO_COM_SISTEMAS"),
    (re.compile(r"(?i)custos?\s*(s[\/.]?|sobre)?\s*vendas|^cmv$|custo.*(servi[çc]o|mercadoria).*prestad"), "OPERACIONAL"),  # CMV/custo direto — nunca confundir com despesa comercial
    (re.compile(r"(?i)marketing|s&?m|despesas?\s*comerciais?"), "DESPESAS_COMERCIAIS"),
    (re.compile(r"(?i)resultado\s*financeiro|juros|receita\s*financeira"), "DESPESAS_ADMINISTRATIVAS"),
    (re.compile(r"(?i)pessoal\s*operacional|custo\s*direto|csp"), "OPERACIONAL"),
    (re.compile(r"(?i)g&?a|administrativ|outros\s*custos"), "DESPESAS_ADMINISTRATIVAS"),
    (re.compile(r"(?i)receitas?\b|presta[çc][ãa]o\s*de\s*servi[çc]o|faturamento|venda\s*de\s*servi[çc]o"), "RECEBIMENTOS"),
]


def _codigo_normalizado(codigo) -> str | None:
    """Normaliza o código da conta pra string comparável (remove '.0' de
    float, mantém dígitos/pontos). Retorna None se não der pra usar."""
    if codigo is None:
        return None
    s = str(codigo)
    if s.endswith(".0"):
        s = s[:-2]
    return s if s else None


def filtrar_contas_folha(contas: list[dict]) -> list[dict]:
    """Remove contas-pai (sintéticas) que também têm filhas na mesma lista —
    somar pai + filhas conta o mesmo valor 2-3x, exatamente o erro que a
    metodologia do Thiago avisa pra nunca cometer. Detecta hierarquia por
    prefixo de código (funciona tanto pra código com ponto — "3.4.1" — quanto
    pra código inteiro concatenado — "341010038" — sem assumir qual dos dois
    formatos o arquivo usa).

    Contas sem código utilizável (None) são mantidas como estão — mais
    seguro somar uma conta sem hierarquia clara do que descartá-la e
    subestimar o total."""
    codigos = [_codigo_normalizado(c.get("codigo")) for c in contas]
    codigos_validos = {c for c in codigos if c}

    def tem_filha(codigo: str) -> bool:
        for outro in codigos_validos:
            if outro == codigo:
                continue
            if outro.startswith(codigo + ".") or (outro.startswith(codigo) and len(outro) > len(codigo)):
                return True
        return False

    folhas = []
    for conta, codigo in zip(contas, codigos):
        if codigo is None or not tem_filha(codigo):
            folhas.append(conta)
    return folhas


def categorizar_conta_waterfall(nome_conta: str) -> str | None:
    """Retorna a categoria do waterfall pra uma conta, ou None se não bater
    com confiança em nenhum padrão conhecido (fica pro chamador decidir o
    que fazer com o resíduo — nunca advinha aqui)."""
    for regex, categoria in WATERFALL_CATEGORIAS:
        if regex.search(nome_conta):
            return categoria
    return None


def mapear_waterfall(contas: list[dict]) -> dict:
    """Soma cada conta na sua categoria do waterfall. `contas` é a lista já
    consolidada (mesmo formato de select_relevant_accounts / series_contabeis:
    [{"conta": str, "codigo": ..., "valores": {"mes_01": float, ...}}]).

    Retorna {categoria: soma_total, "nao_mapeado": [{"conta":..., "total":...}]}
    — o resíduo não mapeado vem à parte, explícito, nunca escondido dentro de
    uma categoria genérica por engano.

    Soma só contas-folha (ver filtrar_contas_folha) — nunca conta-pai junto
    com suas filhas, que duplicaria/triplicaria o valor."""
    contas_folha = filtrar_contas_folha(contas)
    totais = {cat: 0.0 for _, cat in WATERFALL_CATEGORIAS}
    nao_mapeado = []
    for conta in contas_folha:
        nome = conta.get("conta", "")
        valores = conta.get("valores", {})
        total_conta = sum(v for v in valores.values() if isinstance(v, (int, float)))
        categoria = categorizar_conta_waterfall(nome)
        if categoria:
            totais[categoria] += total_conta
        elif abs(total_conta) > 0.01:
            nao_mapeado.append({"conta": nome, "total": round(total_conta, 2)})
    totais["nao_mapeado"] = sorted(nao_mapeado, key=lambda x: abs(x["total"]), reverse=True)
    return totais


# ---------------------------------------------------------------------------
# 2. EBITDA EM 3 CAMADAS
# ---------------------------------------------------------------------------
def calcular_ebitda_3_camadas(waterfall: dict, lucro_liquido: float | None,
                               tributos_sobre_lucro: float | None,
                               resultado_financeiro: float | None) -> dict:
    """(a) bottom-up: soma das categorias operacionais do waterfall.
    (b) top-down: parte do lucro líquido já fechado, sempre exclui resultado
    financeiro (fórmula de mercado padrão).
    Se (a) e (b) baterem, a diferença deve ser exatamente o resultado
    financeiro — reconcilia lado a lado, nunca esconde a diferença."""
    # Receita é credora — aparece NEGATIVA no arquivo por convenção contábil.
    # Inverte o sinal só aqui (nunca nas despesas, que já são positivas/devedoras).
    receitas = abs(waterfall.get("RECEBIMENTOS", 0.0))
    custos_operacionais = sum(
        waterfall.get(cat, 0.0) for cat in
        ("TRIBUTOS", "OPERACIONAL", "CUSTO_COM_SISTEMAS", "DESPESAS_ADMINISTRATIVAS",
         "DESPESAS_COMERCIAIS", "DESPESAS_COM_SOCIOS", "DESPESAS_INDEDUTIVEIS")
    )
    ebitda_bottom_up = receitas - abs(custos_operacionais)

    ebitda_top_down = None
    diferenca = None
    if lucro_liquido is not None and tributos_sobre_lucro is not None and resultado_financeiro is not None:
        d_a = abs(waterfall.get("D_A", 0.0))
        ebitda_top_down = lucro_liquido + tributos_sobre_lucro + resultado_financeiro + d_a
        if ebitda_bottom_up:
            diferenca = round(ebitda_bottom_up - ebitda_top_down, 2)

    return {
        "ebitda_gerencial_bottom_up": round(ebitda_bottom_up, 2),
        "ebitda_padrao_mercado_top_down": round(ebitda_top_down, 2) if ebitda_top_down is not None else None,
        "diferenca_bottom_vs_top": diferenca,
        "margem_ebitda_gerencial_pct": round(100 * ebitda_bottom_up / receitas, 1) if receitas else None,
    }


# ---------------------------------------------------------------------------
# 3. DETECÇÃO DE ANOMALIA POR RUN-RATE
# ---------------------------------------------------------------------------
RATIO_ALTA = 2.5
RATIO_BAIXA = 0.4
PISO_MATERIALIDADE = 10_000.0


def _e_conta_de_resultado(codigo) -> bool | None:
    """Convenção contábil brasileira: código começando em 1/2 é conta de
    balanço (ativo/passivo — saldo num instante, não flui mês a mês do
    zero); começando em 3/4 é conta de resultado (receita/despesa —
    acumula ao longo do ano, aí sim run-rate faz sentido). Retorna None se
    o código não permitir essa leitura (aí o chamador decide se inclui)."""
    codigo_norm = _codigo_normalizado(codigo)
    if not codigo_norm or not codigo_norm[0].isdigit():
        return None
    return codigo_norm[0] in ("3", "4")


def detectar_anomalias_run_rate(contas: list[dict], top_n: int = 10) -> list[dict]:
    """Pra cada conta de RESULTADO (receita/despesa — código 3/4) com série
    mensal, calcula o run-rate esperado (mediana dos meses "normais", ou
    seja, todos exceto o candidato a outlier) e marca como anomalia se
    algum mês individual romper a razão real/esperado de >2,5x ou <0,4x,
    acima do piso de materialidade de R$10 mil.

    Contas de BALANÇO (ativo/passivo, código 1/2) ficam de fora — são saldo
    num instante, não fluxo; comparar run-rate nelas gera ruído, não sinal
    (confirmado testando contra dado real: sem esse filtro, "APLICAÇÕES
    FINANCEIRAS" aparecia como "anomalia" só por movimentação normal de
    caixa). Contas sem código legível entram por padrão — mais seguro
    avaliar de mais do que perder uma anomalia real por falta de código.

    Retorna só as `top_n` mais materiais — o Claude nunca vê a lista
    completa, só o resíduo que já passou pelo filtro."""
    candidatos = []
    for conta in contas:
        if _e_conta_de_resultado(conta.get("codigo")) is False:
            continue
        nome = conta.get("conta", "")
        valores = conta.get("valores", {})
        meses_validos = {m: v for m, v in valores.items() if isinstance(v, (int, float))}
        if len(meses_validos) < 3:
            continue

        itens = sorted(meses_validos.items())
        for i, (mes, valor) in enumerate(itens):
            outros = [v for m, v in itens if m != mes]
            run_rate_esperado = statistics.median(outros) if outros else 0.0
            if abs(run_rate_esperado) < 1.0:
                continue  # evita divisão por ~zero — conta que só existe nesse mês entra em "conta nova", não aqui
            ratio = valor / run_rate_esperado
            desvio_absoluto = abs(valor - run_rate_esperado)
            if desvio_absoluto < PISO_MATERIALIDADE:
                continue
            if ratio > RATIO_ALTA or ratio < RATIO_BAIXA:
                candidatos.append({
                    "conta": nome,
                    "periodo": mes,
                    "valor_esperado_run_rate": round(run_rate_esperado, 2),
                    "valor_real": round(valor, 2),
                    "ratio": round(ratio, 2),
                    "desvio_absoluto": round(desvio_absoluto, 2),
                })

    candidatos.sort(key=lambda c: c["desvio_absoluto"], reverse=True)
    return candidatos[:top_n]


def detectar_contas_novas_ou_zeradas(contas: list[dict]) -> dict:
    """Contas que só aparecem numa parte do período (novas) ou que somem no
    meio (zeradas) — sinalizadas à parte, sem tentar adivinhar se é
    sazonalidade normal ou risco real (isso fica pro Claude decidir, com o
    dado já isolado)."""
    novas, zeradas = [], []
    for conta in contas:
        valores = conta.get("valores", {})
        meses_ordenados = sorted(valores.keys())
        vistos = [m for m in meses_ordenados if isinstance(valores.get(m), (int, float)) and abs(valores[m]) > 0.01]
        if not vistos:
            continue
        primeiro_mes_com_dado = meses_ordenados[0]
        if vistos[0] != primeiro_mes_com_dado:
            novas.append({"conta": conta.get("conta", ""), "primeiro_mes_com_valor": vistos[0]})
        if vistos[-1] != meses_ordenados[-1]:
            zeradas.append({"conta": conta.get("conta", ""), "ultimo_mes_com_valor": vistos[-1]})
    return {"contas_novas": novas, "contas_zeradas": zeradas}
