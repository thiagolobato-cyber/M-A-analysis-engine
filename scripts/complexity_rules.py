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


SEGMENTOS_MEDIA_COMPLEXIDADE = {"industria", "associacao", "hospital", "instituicao_de_ensino"}


@dataclass
class ComplexidadeResultado:
    nivel: str  # "Baixa complexidade" | "Média complexidade" | "Alta complexidade" | "Dados insuficientes"
    motivos: list[str] = field(default_factory=list)
    criterios_nao_avaliados: list[str] = field(default_factory=list)

    def to_agent_run_output(self) -> dict:
        """Formato compatível com o que o agente de IA devolvia — confirmado
        contra generate_outputs.py (que lê `classificacao` e
        `criterios_acionados`, não `result`/`findings` — schema real, não o
        que eu tinha suposto numa primeira versão)."""
        return {
            "classificacao": self.nivel,
            "confidence": 1.0 if not self.criterios_nao_avaliados else 0.6,
            "criterios_acionados": [
                {"criterio": m, "detalhe": m, "nivel_apontado": self.nivel}
                for m in self.motivos
            ],
            "evidence": [{"criterio": m} for m in self.motivos],
            "source": "rule_engine_complexity_v1",
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

    fora_grande_sp = structured.get("localizacao_fora_grande_sp")
    if fora_grande_sp is True:
        motivos_alta.append(
            f"Localização fora da capital de SP / região metropolitana adjacente "
            f"(cidade informada: {structured.get('localizacao', 'não informada')})."
        )
    elif fora_grande_sp is None:
        nao_avaliados.append("localizacao_fora_grande_sp")

    sistema_e_dominio = structured.get("sistema_e_dominio")
    if sistema_e_dominio is False:
        motivos_alta.append(
            f"Sistema de ERP contábil não é o Domínio (informado: {structured.get('sistema_utilizado', 'não informado')})."
        )
    elif sistema_e_dominio is None:
        nao_avaliados.append("sistema_e_dominio")

    hospedagem = structured.get("sistema_hospedagem")
    if hospedagem == "servidor_fisico":
        motivos_media.append("Sistema hospedado em servidor físico (não em nuvem/web).")
    elif hospedagem is None:
        nao_avaliados.append("sistema_hospedagem")

    sistema_financeiro_omie = structured.get("sistema_financeiro_e_omie")
    if sistema_financeiro_omie is False:
        motivos_media.append(
            f"Sistema financeiro diferente do Omie (informado: {structured.get('sistema_financeiro', 'não informado')})."
        )
    elif sistema_financeiro_omie is None:
        nao_avaliados.append("sistema_financeiro_e_omie")

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

    concentracao = structured.get("porcentagem_faturamento_top_10_clientes")
    if concentracao is not None:
        if concentracao > 35:
            motivos_alta.append(f"Concentração de receita nos top 10 clientes = {concentracao}% (>35%, alta).")
        elif concentracao > 25:
            motivos_media.append(f"Concentração de receita nos top 10 clientes = {concentracao}% (>25%, média).")
    else:
        nao_avaliados.append("porcentagem_faturamento_top_10_clientes")

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
        nivel = "Dados insuficientes"
    else:
        nivel = "Baixa complexidade"

    return ComplexidadeResultado(
        nivel=nivel,
        motivos=motivos_alta + motivos_media,
        criterios_nao_avaliados=nao_avaliados,
    )
