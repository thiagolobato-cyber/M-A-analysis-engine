"""
Motor de regras para classificação de Complexidade — substitui o agente de
IA por código determinístico. Os critérios abaixo são uma transcrição
literal do prompt original de "Complexidade dos Parceiros" que o Thiago
trouxe; nada aqui foi inventado ou reinterpretado.

REVISADO em 20/08 depois de ler o system_prompt real do Agent 00
(Extraction): o `structured` que ele produz é PLANO (`localizacao`,
`sistema_utilizado`, `numero_clientes`...), não aninhado como eu tinha
suposto numa primeira versão. Além disso, 5 campos que este motor precisa
não existiam no prompt original — foram propostos como adição em
`extraction_prompt_addition.txt` (ainda não aplicada em produção; este
módulo já assume que a adição foi feita).

Por que isso pode virar código (e por que o texto cru do Excel não pode):
os critérios em si (Domínio? não é Domínio? >25%? >35%?) são limiares
fixos, sem julgamento — mas os DADOS chegam como texto livre em
português ("dominio e sistema do cliente", "Patos de Minas - MG"), e isso
varia de parceiro pra parceiro. Normalizar texto livre em campos tipados
(bool/enum) é trabalho do Agent 00, que já lê o texto com IA mesmo — não
é uma chamada nova. A partir do momento em que os 5 campos normalizados
existem em `structured`, classificar não precisa mais de IA nenhuma.

Nota sobre "CNPJ/HC": confirmado com o Thiago em 20/08 — é o número de
clientes da carteira dividido pelo headcount operacional. No schema real
do Agent 00 isso é `numero_clientes` ÷ `numero_colaboradores`.
"""
from __future__ import annotations

from dataclasses import dataclass, field


# Segmentos que, segundo o critério original, tornam a carteira mais
# exigente operacionalmente (indústrias, associações, hospitais,
# instituições de ensino) — mesmos valores que o Agent 00 deve devolver
# em `segmentos_sensiveis_top10_clientes` (ver extraction_prompt_addition.txt).
SEGMENTOS_MEDIA_COMPLEXIDADE = {"industria", "associacao", "hospital", "instituicao_de_ensino"}


@dataclass
class ComplexidadeResultado:
    nivel: str  # "Baixa complexidade" | "Média complexidade" | "Alta complexidade" | "Dados insuficientes"
    motivos: list[str] = field(default_factory=list)
    criterios_nao_avaliados: list[str] = field(default_factory=list)

    def to_agent_run_output(self) -> dict:
        """Formato compatível com o que o agente de IA devolvia — pra não
        quebrar o generate_outputs.py, que lê agents['complexity']['output']."""
        return {
            "result": self.nivel,
            "confidence": 1.0 if not self.criterios_nao_avaliados else 0.6,
            "findings": self.motivos,
            "drivers": self.motivos,
            "risks": [],
            "evidence": [{"criterio": m} for m in self.motivos],
            "source": "rule_engine_complexity_v1",  # não "agent_complexity" — deixa claro que não passou por LLM
            "criterios_nao_avaliados": self.criterios_nao_avaliados,
        }


def classificar_complexidade(structured: dict) -> ComplexidadeResultado:
    """Aplica os critérios de complexidade a um `deal_data.structured`
    plano, no formato que o Agent 00 (Extraction) produz — ver
    extraction_prompt_addition.txt para os 5 campos que precisam existir
    além do que o prompt original já extrai.
    """
    motivos_alta: list[str] = []
    motivos_media: list[str] = []
    nao_avaliados: list[str] = []

    # --- Critério 1: localização ---
    fora_grande_sp = structured.get("localizacao_fora_grande_sp")
    if fora_grande_sp is True:
        motivos_alta.append(
            f"Localização fora da capital de SP / região metropolitana adjacente "
            f"(cidade informada: {structured.get('localizacao', 'não informada')})."
        )
    elif fora_grande_sp is None:
        nao_avaliados.append("localizacao_fora_grande_sp")

    # --- Critério 2: sistema de ERP contábil não é o Domínio ---
    sistema_e_dominio = structured.get("sistema_e_dominio")
    if sistema_e_dominio is False:
        motivos_alta.append(
            f"Sistema de ERP contábil não é o Domínio (informado: {structured.get('sistema_utilizado', 'não informado')})."
        )
    elif sistema_e_dominio is None:
        nao_avaliados.append("sistema_e_dominio")

    # --- Critério 3: hospedagem em servidor físico (revisado 20/08: Média, não Alta) ---
    hospedagem = structured.get("sistema_hospedagem")
    if hospedagem == "servidor_fisico":
        motivos_media.append("Sistema hospedado em servidor físico (não em nuvem/web).")
    elif hospedagem is None:
        nao_avaliados.append("sistema_hospedagem")

    # --- Critério 4: sistema financeiro diferente do Omie (revisado 20/08: Média, não Alta) ---
    sistema_financeiro_omie = structured.get("sistema_financeiro_e_omie")
    if sistema_financeiro_omie is False:
        motivos_media.append(
            f"Sistema financeiro diferente do Omie (informado: {structured.get('sistema_financeiro', 'não informado')})."
        )
    elif sistema_financeiro_omie is None:
        nao_avaliados.append("sistema_financeiro_e_omie")

    # --- Critério 5: razão clientes ÷ headcount operacional ---
    numero_clientes = structured.get("numero_clientes")
    numero_colaboradores = structured.get("numero_colaboradores")
    if numero_clientes is not None and numero_colaboradores:
        ratio = numero_clientes / numero_colaboradores
        if ratio < 15:
            motivos_alta.append(
                f"Razão clientes/colaboradores = {ratio:.1f} ({numero_clientes} ÷ {numero_colaboradores}), abaixo de 15 (alta)."
            )
        elif ratio < 20:
            motivos_media.append(
                f"Razão clientes/colaboradores = {ratio:.1f} ({numero_clientes} ÷ {numero_colaboradores}), abaixo de 20 (média)."
            )
    else:
        nao_avaliados.append("numero_clientes ou numero_colaboradores")

    # --- Critério 6: concentração de receita (Pareto top 10) ---
    concentracao = structured.get("porcentagem_faturamento_top_10_clientes")
    if concentracao is not None:
        if concentracao > 35:
            motivos_alta.append(f"Concentração de receita nos top 10 clientes = {concentracao}% (>35%, alta).")
        elif concentracao > 25:
            motivos_media.append(f"Concentração de receita nos top 10 clientes = {concentracao}% (>25%, média).")
    else:
        nao_avaliados.append("porcentagem_faturamento_top_10_clientes")

    # --- Critério 7: segmento sensível entre os top 10 clientes ---
    segmentos = set(structured.get("segmentos_sensiveis_top10_clientes") or [])
    if segmentos & SEGMENTOS_MEDIA_COMPLEXIDADE:
        motivos_media.append(
            f"Atividade dos top 10 clientes inclui segmento(s) sensível(is): {sorted(segmentos & SEGMENTOS_MEDIA_COMPLEXIDADE)}."
        )
    elif "segmentos_sensiveis_top10_clientes" not in structured:
        nao_avaliados.append("segmentos_sensiveis_top10_clientes")

    if motivos_alta:
        nivel = "Alta complexidade"
    elif motivos_media:
        nivel = "Média complexidade"
    elif nao_avaliados and not (motivos_alta or motivos_media):
        # Dado insuficiente pra confirmar baixa com segurança — mesma lógica
        # de "insufficient evidence" que já existe no CFO Agent, aplicada aqui.
        nivel = "Dados insuficientes"
    else:
        nivel = "Baixa complexidade"

    return ComplexidadeResultado(
        nivel=nivel,
        motivos=motivos_alta + motivos_media,
        criterios_nao_avaliados=nao_avaliados,
    )
