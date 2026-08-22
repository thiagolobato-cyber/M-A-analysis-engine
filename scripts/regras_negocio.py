"""
regras_negocio.py — Bloco A (Viabilidade Financeira) e Bloco B (Complexidade
Operacional), conforme "Regras - Análise de Negócios" definido pelo time do
Thiago em 21/08. Substitui inteiramente a versão anterior de
complexity_rules.py (critérios e pesos mudaram) e cria um agente novo —
viabilidade financeira — que também nunca precisou de IA: é fórmula fiscal
(Simples Nacional Anexo III, Lucro Presumido/Real) e limiar numérico.

Por que isso é seguro de virar 100% código: todo critério do documento é
comparação numérica ou palavra-chave (ex.: "Domínio e SCI" = campo contém
as duas palavras). Não existe julgamento de texto livre aqui — é exatamente
o tipo de regra que custa caro em IA e é instantâneo e determinístico em
código.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field


# ===========================================================================
# BLOCO A — VIABILIDADE FINANCEIRA
# ===========================================================================

# Tabela do Simples Nacional, Anexo III (Receita Bruta em 12 meses).
# (faixa_ate, aliquota_nominal, parcela_a_deduzir)
SIMPLES_ANEXO_III = [
    (180_000.00, 0.060, 0.00),
    (360_000.00, 0.112, 9_360.00),
    (720_000.00, 0.135, 17_640.00),
    (1_800_000.00, 0.160, 35_640.00),
    (3_600_000.00, 0.210, 125_640.00),
    (4_800_000.00, 0.330, 648_000.00),
]


def calcular_aliquota_efetiva_simples(rbt12: float) -> float | None:
    """Alíquota efetiva do Simples Nacional (Anexo III) a partir da Receita
    Bruta dos últimos 12 meses (RBT12). Retorna None se RBT12 ultrapassar o
    teto do Anexo III (sublimite) — nesse caso o cálculo não se aplica e o
    chamador deve sinalizar como dado a confirmar manualmente, nunca estimar."""
    if rbt12 is None or rbt12 <= 0:
        return None
    for faixa_ate, aliquota_nominal, parcela_a_deduzir in SIMPLES_ANEXO_III:
        if rbt12 <= faixa_ate:
            return round((rbt12 * aliquota_nominal - parcela_a_deduzir) / rbt12, 6)
    return None  # acima da última faixa — fora do Anexo III, não estimar


def calcular_imposto_mensal(faturamento_mensal: float, regime_tributario: str,
                             rbt12: float | None = None) -> dict:
    """Imposto mensal por regime — cada fórmula é exatamente a do documento,
    nada aproximado. Retorna {"imposto_mensal": ..., "aliquota_usada": ...,
    "regime_reconhecido": bool} — regime não reconhecido nunca vira zero
    silencioso, fica sinalizado."""
    regime_norm = (regime_tributario or "").strip().lower()

    if "presumido" in regime_norm:
        aliquota = 0.0865
        return {"imposto_mensal": round(faturamento_mensal * aliquota, 2),
                "aliquota_usada": aliquota, "regime_reconhecido": True}

    if "real" in regime_norm:
        aliquota = 0.1425
        return {"imposto_mensal": round(faturamento_mensal * aliquota, 2),
                "aliquota_usada": aliquota, "regime_reconhecido": True}

    if "simples" in regime_norm:
        aliquota = calcular_aliquota_efetiva_simples(rbt12)
        if aliquota is None:
            return {"imposto_mensal": None, "aliquota_usada": None, "regime_reconhecido": False,
                    "motivo": "RBT12 ausente ou fora do Anexo III — confirmar manualmente."}
        return {"imposto_mensal": round(faturamento_mensal * aliquota, 2),
                "aliquota_usada": aliquota, "regime_reconhecido": True}

    return {"imposto_mensal": None, "aliquota_usada": None, "regime_reconhecido": False,
            "motivo": f"Regime tributário '{regime_tributario}' não reconhecido (esperado: Simples Nacional, Lucro Presumido ou Lucro Real)."}


def calcular_custo_folha_total(folha_informada: float, regime_tributario: str) -> dict:
    """Custo Folha Total = Folha Informada + Encargo, onde o multiplicador de
    encargo depende do regime: 1,08x (Simples Nacional) ou 1,275x (Presumido
    ou Real) — direto do documento."""
    regime_norm = (regime_tributario or "").strip().lower()
    if "simples" in regime_norm:
        multiplicador = 1.08
    elif "presumido" in regime_norm or "real" in regime_norm:
        multiplicador = 1.275
    else:
        return {"custo_folha_total": None, "multiplicador": None,
                "motivo": f"Regime tributário '{regime_tributario}' não reconhecido."}
    return {"custo_folha_total": round(folha_informada * multiplicador, 2), "multiplicador": multiplicador}


def calcular_margem_bruta(faturamento_mensal: float, folha_informada: float,
                           custo_sistemas: float, regime_tributario: str,
                           rbt12: float | None = None) -> dict:
    """Margem Bruta (%) = (Faturamento − Custo Folha Total − Custo Sistemas
    − Imposto) ÷ Faturamento. Cada componente vem das funções acima —
    nenhum número aparece do nada aqui."""
    folha = calcular_custo_folha_total(folha_informada, regime_tributario)
    imposto = calcular_imposto_mensal(faturamento_mensal, regime_tributario, rbt12)

    if folha["custo_folha_total"] is None or imposto["imposto_mensal"] is None:
        return {
            "margem_bruta_pct": None,
            "motivo": folha.get("motivo") or imposto.get("motivo"),
            "custo_folha_total": folha["custo_folha_total"],
            "imposto_mensal": imposto["imposto_mensal"],
        }

    lucro_bruto = faturamento_mensal - folha["custo_folha_total"] - custo_sistemas - imposto["imposto_mensal"]
    margem_pct = round(100 * lucro_bruto / faturamento_mensal, 2) if faturamento_mensal else None
    return {
        "margem_bruta_pct": margem_pct,
        "custo_folha_total": folha["custo_folha_total"],
        "imposto_mensal": imposto["imposto_mensal"],
        "aliquota_usada": imposto["aliquota_usada"],
    }


# --- Faixas Verde/Amarelo/Vermelho de cada critério do Bloco A ---
def _faixa_percentual(valor: float | None, verde_op, amarelo_faixa, vermelho_op) -> str | None:
    """Helper genérico: aplica os operadores de corte de cada critério.
    Retorna None se valor for None (dado ausente, nunca chuta faixa)."""
    if valor is None:
        return None
    if verde_op(valor):
        return "Verde"
    if vermelho_op(valor):
        return "Vermelho"
    return "Amarelo"


def avaliar_viabilidade_financeira(
    margem_bruta_pct: float | None,
    custo_folha_pct_faturamento: float | None,
    inadimplencia_media_pct: float | None,
    churn_medio_pct: float | None,
    perfil_restrito_presente: bool | None,
) -> dict:
    """Bloco A completo — classifica cada critério e agrega pela regra do
    documento: qualquer Vermelho → Não Recomendado; perfil restrito = Sim →
    Fit com Ressalvas (revisar), mesmo sem Vermelho; qualquer Amarelo sem
    Vermelho → Fit com Ressalvas; tudo Verde → Fit Estratégico."""
    faixas = {
        "margem_bruta": _faixa_percentual(margem_bruta_pct, lambda v: v > 40, None, lambda v: v < 25),
        "custo_folha_faturamento": _faixa_percentual(custo_folha_pct_faturamento, lambda v: v < 45, None, lambda v: v > 60),
        "inadimplencia": _faixa_percentual(inadimplencia_media_pct, lambda v: v < 3, None, lambda v: v > 7),
        "churn": _faixa_percentual(churn_medio_pct, lambda v: v < 1.5, None, lambda v: v > 3),
    }

    nao_avaliados = [k for k, v in faixas.items() if v is None]
    tem_vermelho = any(v == "Vermelho" for v in faixas.values())
    tem_amarelo = any(v == "Amarelo" for v in faixas.values())

    if tem_vermelho:
        classificacao = "Não Recomendado"
    elif perfil_restrito_presente is True:
        classificacao = "Fit com Ressalvas (revisar manualmente — perfil restrito)"
    elif tem_amarelo:
        classificacao = "Fit com Ressalvas"
    elif nao_avaliados:
        classificacao = "Dados insuficientes"
    else:
        classificacao = "Fit Estratégico"

    return {
        "classificacao": classificacao,
        "criterios": faixas,
        "perfil_restrito_presente": perfil_restrito_presente,
        "criterios_nao_avaliados": nao_avaliados,
    }


# ===========================================================================
# BLOCO B — COMPLEXIDADE OPERACIONAL (score ponderado)
# ===========================================================================

SISTEMAS_NOMEADOS_CONHECIDOS = ("contmatic", "folhamatic", "questor", "fortes", "alterdata")

MUNICIPIOS_GRANDE_SP = {
    "são paulo", "sao paulo", "guarulhos", "osasco", "santo andré", "santo andre",
    "são bernardo do campo", "sao bernardo do campo", "são caetano do sul", "sao caetano do sul",
    "diadema", "barueri", "carapicuíba", "carapicuiba", "cotia", "embu das artes",
    "itapevi", "jandira", "taboão da serra", "taboao da serra", "mauá", "maua",
    "ribeirão pires", "ribeirao pires", "suzano", "mogi das cruzes", "itaquaquecetuba",
}


def classificar_erp_infraestrutura(sistema_texto: str | None, hospedagem: str | None) -> str:
    """'Domínio e SCI' é reconhecido quando o campo contém simultaneamente as
    palavras 'Domínio' e 'SCI' (regra literal do documento). Campo vazio ou
    sistema não reconhecido = 'Outros'. Retorna Baixa/Média/Alta direto da
    matriz ERP × Infraestrutura."""
    texto = (sistema_texto or "").strip().lower()

    if ("domínio" in texto or "dominio" in texto) and "sci" in texto:
        return "Baixa"  # Domínio e SCI Único — Baixa em qualquer infraestrutura, por definição do documento
    if any(nome in texto for nome in SISTEMAS_NOMEADOS_CONHECIDOS):
        return "Média"  # Outro sistema nomeado — Média em qualquer infraestrutura
    return "Alta"  # "Outros" — não informado ou não reconhecido


def _nivel_para_pontos(nivel: str) -> int:
    return {"Baixa": 0, "Média": 1, "Media": 1, "Alta": 2}.get(nivel, 0)


@dataclass
class ComplexidadeOperacionalResultado:
    score_total: int
    classificacao: str
    detalhe_por_criterio: list = field(default_factory=list)
    criterios_nao_avaliados: list = field(default_factory=list)

    def to_agent_run_output(self) -> dict:
        return {
            "classificacao": self.classificacao,
            "confidence": 1.0 if not self.criterios_nao_avaliados else 0.6,
            "score_total": self.score_total,
            "criterios_acionados": [
                {"criterio": d["criterio"], "detalhe": d["nivel"], "nivel_apontado": d["nivel"]}
                for d in self.detalhe_por_criterio if d["nivel"] != "Baixa"
            ],
            "evidence": [{"criterio": d["criterio"], "nivel": d["nivel"], "pontos": d["pontos"]} for d in self.detalhe_por_criterio],
            "source": "rule_engine_complexidade_operacional_v2",
            "criterios_nao_avaliados": self.criterios_nao_avaliados,
        }


def avaliar_complexidade_operacional(
    custo_sistemas_pct_faturamento: float | None,
    localizacao_fora_grande_sp: bool | None,
    numero_clientes: int | None,
    numero_colaboradores: int | None,
    concentracao_top10_pct: float | None,
    segmentos_sensiveis_presentes: bool | None,
    sistema_financeiro_e_omie: bool | None,
    sistema_utilizado_texto: str | None,
    sistema_hospedagem: str | None,
    outsourcing_pessoas_pct: float | None,
    outsourcing_sistemas_pct: float | None,
) -> ComplexidadeOperacionalResultado:
    """Score ponderado dos 9 critérios do Bloco B — soma de (pontos × peso).
    Pontuação máxima teórica: 35 (critério 5 só tem 2 níveis, não 3)."""
    detalhes = []
    nao_avaliados = []

    def registrar(criterio: str, peso: int, nivel: str | None, campo_nome: str):
        if nivel is None:
            nao_avaliados.append(campo_nome)
            return 0
        pontos = _nivel_para_pontos(nivel)
        detalhes.append({"criterio": criterio, "peso": peso, "nivel": nivel, "pontos": pontos * peso})
        return pontos * peso

    score = 0

    # 1. Custo Sistemas / Faturamento — peso 3
    nivel_bruto = _faixa_percentual(custo_sistemas_pct_faturamento, lambda v: v <= 2.5, None, lambda v: v >= 5.0)
    score += registrar("Custo Sistemas / Faturamento", 3,
                        {"Verde": "Baixa", "Amarelo": "Média", "Vermelho": "Alta"}.get(nivel_bruto), "custo_sistemas_pct_faturamento")

    # 2. Localização — peso 2 (só 2 níveis: Baixa ou Alta, sem Média)
    if localizacao_fora_grande_sp is None:
        nao_avaliados.append("localizacao_fora_grande_sp")
    else:
        nivel = "Alta" if localizacao_fora_grande_sp else "Baixa"
        score += registrar("Localização Geográfica", 2, nivel, "localizacao_fora_grande_sp")

    # 3. CNPJ/HC (carteira ÷ headcount) — peso 2. >=20 Baixa, 15-19 Média, <15 Alta.
    if numero_clientes is not None and numero_colaboradores:
        ratio = numero_clientes / numero_colaboradores
        nivel = "Baixa" if ratio >= 20 else ("Média" if ratio >= 15 else "Alta")
        detalhes.append({"criterio": f"CNPJ/HC = {ratio:.1f}", "peso": 2, "nivel": nivel, "pontos": _nivel_para_pontos(nivel) * 2})
        score += _nivel_para_pontos(nivel) * 2
    else:
        nao_avaliados.append("numero_clientes/numero_colaboradores")

    # 4. Concentração Pareto (top 10) — peso 3
    nivel_bruto = _faixa_percentual(concentracao_top10_pct, lambda v: v <= 25, None, lambda v: v > 35)
    score += registrar("Concentração Pareto (top 10 clientes)", 3,
                        {"Verde": "Baixa", "Amarelo": "Média", "Vermelho": "Alta"}.get(nivel_bruto), "concentracao_top10_pct")

    # 5. Segmentos específicos — peso 1, só 2 níveis (Baixa/Média, nunca Alta)
    if segmentos_sensiveis_presentes is None:
        nao_avaliados.append("segmentos_sensiveis_presentes")
    else:
        nivel = "Média" if segmentos_sensiveis_presentes else "Baixa"
        score += registrar("Segmentos Específicos", 1, nivel, "segmentos_sensiveis_presentes")

    # 6. Sistema Financeiro do Cliente — peso 2, só 2 níveis (Baixa/Alta, sem Média)
    if sistema_financeiro_e_omie is None:
        nao_avaliados.append("sistema_financeiro_e_omie")
    else:
        nivel = "Baixa" if sistema_financeiro_e_omie else "Alta"
        score += registrar("Sistema Financeiro do Cliente", 2, nivel, "sistema_financeiro_e_omie")

    # 7. ERP × Infraestrutura — peso 3 (matriz dedicada)
    nivel = classificar_erp_infraestrutura(sistema_utilizado_texto, sistema_hospedagem)
    score += registrar("ERP × Infraestrutura", 3, nivel, "sistema_utilizado_texto")

    # 8. Outsourcing — Pessoas Alocadas — peso 1
    nivel_bruto = _faixa_percentual(outsourcing_pessoas_pct, lambda v: v <= 10, None, lambda v: v > 15)
    score += registrar("Outsourcing — Pessoas Alocadas", 1,
                        {"Verde": "Baixa", "Amarelo": "Média", "Vermelho": "Alta"}.get(nivel_bruto), "outsourcing_pessoas_pct")

    # 9. Outsourcing — Sistemas do Cliente — peso 1
    nivel_bruto = _faixa_percentual(outsourcing_sistemas_pct, lambda v: v <= 10, None, lambda v: v > 15)
    score += registrar("Outsourcing — Sistemas do Cliente", 1,
                        {"Verde": "Baixa", "Amarelo": "Média", "Vermelho": "Alta"}.get(nivel_bruto), "outsourcing_sistemas_pct")

    if score <= 11:
        classificacao = "Baixa Complexidade"
    elif score <= 23:
        classificacao = "Média Complexidade"
    else:
        classificacao = "Alta Complexidade"

    return ComplexidadeOperacionalResultado(
        score_total=score, classificacao=classificacao,
        detalhe_por_criterio=detalhes, criterios_nao_avaliados=nao_avaliados,
    )
