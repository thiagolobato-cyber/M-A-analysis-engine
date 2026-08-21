#!/usr/bin/env python3
"""
Roda um agente (Extraction, os 4 analíticos, ou CFO Synthesis) contra um deal.

Uso:
    python scripts/run_agent.py --agent complexity --deal-id <uuid>

O que faz:
  1. Busca no Supabase o deal_data (o "Data Object" canônico) do deal.
  2. Busca a versão ativa do agente e o system_prompt correspondente.
  3. Monta o prompt final e chama `claude -p` (autenticado via
     CLAUDE_CODE_OAUTH_TOKEN, já configurado como GitHub Secret).
  4. Valida que a resposta é um JSON bem formado.
  5. Grava o resultado em agent_runs (ou synthesis_runs, se for o CFO).

Nota: este é um primeiro rascunho. Testei a sintaxe e a lógica localmente,
mas não contra o Supabase real (esta sandbox não tem acesso de rede até
supabase.co) — vamos validar isso de verdade assim que os secrets do
GitHub Actions existirem.
"""
import argparse
import base64
import hashlib
import io
import json
import os
import re
import statistics
import subprocess
import sys
import urllib.request
import urllib.error

from openpyxl import load_workbook

from dre_balancete_parser import (
    detect_consolidated_balancete,
    parse_consolidated_balancete,
    detect_dre_sheet,
    parse_dre_sheet,
    calcular_ebitda_de_dre,
    dre_linhas_para_contas,
)
from financial_engine import (
    mapear_waterfall,
    calcular_ebitda_3_camadas,
    detectar_anomalias_run_rate,
    detectar_contas_novas_ou_zeradas,
)

MAX_ROWS_PER_SHEET = 300
MAX_COLS_PER_SHEET = 40
MAX_SHEETS = 30
MONTHLY_SHEET_RE = re.compile(r'(?i)^m[eê]s\s*0?(\d{1,2})$')

# Só o agente que lê DRE/Balancete de verdade precisa da série contábil
# crua — Complexity, Opinion, Integration/Operational Risks decidem a
# partir do resumo estruturado (geografia, sistema, HC, Pareto, etc.),
# nunca abriram a tabela de centenas de contas x 12 meses mesmo antes.
AGENTS_NEED_RAW_SERIES = {"financial_analysis"}

# Quantas contas (das potencialmente 300+) realmente vão pro prompt do
# agente financeiro — as demais já foram descartadas por não serem nem
# materiais nem voláteis o suficiente pra merecer leitura humana ou de IA.
MAX_ACCOUNTS_FOR_FINANCIAL_AGENT = 25

# Alias de modelo por padrão quando agent_versions.model vier vazio —
# nunca deixar cair no default da sessão (pode ser Opus sem ninguém notar).
DEFAULT_MODEL_ALIAS = "sonnet"


def detect_monthly_sheets(wb):
    """Detecta abas no padrão 'MES 01'..'MES 12' — um mês por aba, mesmo
    layout de colunas em cada uma. Comum em exports de Balancete brasileiros."""
    found = []
    for name in wb.sheetnames:
        m = MONTHLY_SHEET_RE.match(name.strip())
        if m:
            found.append((int(m.group(1)), name))
    found.sort()
    return found if len(found) >= 3 else []


def merge_monthly_balancete(wb, monthly_sheets):
    """Cruza as abas mensais por conta (nome + código), pegando o último
    valor numérico de cada linha como saldo final do mês — validado contra
    um arquivo real: Saldo Anterior + Débito − Crédito = Saldo Atual,
    sempre na última posição."""
    accounts = {}
    for month_num, sheet_name in monthly_sheets:
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            if len(row) < 5:
                continue
            _, nome, codigo, tipo = row[0], row[1], row[2], row[3]
            if nome is None or codigo is None:
                continue
            numeric_vals = [v for v in row[4:] if isinstance(v, (int, float))]
            if not numeric_vals:
                continue
            key = (str(nome).strip(), codigo)
            if key not in accounts:
                accounts[key] = {"tipo": tipo, "meses": {}}
            accounts[key]["meses"][month_num] = numeric_vals[-1]
    return accounts


def format_merged_table(accounts, max_rows=300):
    """Formata o cruzamento como tabela, ordenada por relevância (maior
    valor absoluto em qualquer mês primeiro) — isso substitui pedir pro
    modelo cruzar 12 abas manualmente, o que não é confiável."""
    def materiality(item):
        return max((abs(v) for v in item[1]["meses"].values()), default=0)
    sorted_accounts = sorted(accounts.items(), key=materiality, reverse=True)

    lines = [
        "===== SÉRIE MENSAL CONSOLIDADA (cruzada automaticamente das abas mensais) =====",
        "Saldo final de cada mês, ordenado por relevância (maior valor absoluto primeiro).",
        "ATENÇÃO: contas de resultado (receita/despesa, geralmente código iniciando em 3 ou 4) "
        "costumam vir ACUMULADAS no ano-calendário — para o valor do mês isolado, calcule a "
        "diferença entre um mês e o anterior. Contas de balanço (ativo/passivo) já são o saldo "
        "no fim daquele mês, sem precisar de ajuste.",
        "Conta | Código | Tipo(S=sintética/A=analítica) | " + " | ".join(f"Mês{m:02d}" for m in range(1, 13)),
    ]
    for i, ((nome, codigo), data) in enumerate(sorted_accounts):
        if i >= max_rows:
            lines.append(f"[... {len(sorted_accounts) - max_rows} contas menos relevantes omitidas ...]")
            break
        valores = " | ".join(f"{data['meses'].get(m, 0):.2f}" if m in data["meses"] else "—" for m in range(1, 13))
        lines.append(f"{nome} | {codigo} | {data['tipo']} | {valores}")
    return "\n".join(lines)


def format_dre_table(dre_linhas: dict) -> str:
    """Formata a DRE já extraída em código como texto compacto pro
    contexto — como a tabela já é pequena (dezenas de linhas, não
    centenas), incluir o texto inteiro aqui é barato; o que evitamos é
    o modelo precisar RECONSTRUIR isso a partir da aba genérica."""
    lines = [
        "===== DRE (extraída automaticamente pelo código, por rótulo de linha) =====",
        "Já vem pronta — não precisa copiar de volta no seu output, use como contexto.",
        "Rótulo | " + " | ".join(f"Mês{m:02d}" for m in range(1, 13)),
    ]
    for rotulo, valores in dre_linhas.items():
        valores_fmt = " | ".join(
            f"{valores.get(f'mes_{m:02d}', 0):.2f}" if f"mes_{m:02d}" in valores else "—"
            for m in range(1, 13)
        )
        lines.append(f"{rotulo} | {valores_fmt}")
    return "\n".join(lines)


def excel_to_text(file_bytes: bytes, filename: str) -> tuple[str, list, dict]:
    """Converte um Excel em texto legível pro Claude — não interpreta,
    só descreve fielmente o que tem em cada aba, linha por linha.
    Testado localmente contra os 8 arquivos reais do BHub antes de ir
    pra produção (incluindo um com 24 abas e outro com erro #REF!).

    Retorna (texto, contas_balancete, dre_linhas) — os dois últimos já
    estruturados em código, prontos pra injetar em raw_extracted sem
    depender do LLM copiar de volta uma tabela grande.

    REVISADO em 20/08 contra um arquivo real (Nacional Controladoria) que
    tinha AS DUAS COISAS ao mesmo tempo: 12 abas mensais (MES 01..12) E
    uma aba "BALANCETE 2025" com a mesma informação já consolidada numa
    tabela só. Antes desta revisão, as duas eram enviadas — a mesma
    tabela de ~300 contas duplicada no prompt. Agora: se existe uma aba
    consolidada (qualquer nome — a detecção é pelo cabeçalho, não pelo
    nome da aba), ela é a fonte única e as abas mensais somem do despejo
    também. Só cai de volta pras 12 abas mensais se NENHUMA aba
    consolidada for encontrada."""
    try:
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    except Exception as e:
        return f"[ERRO ao abrir '{filename}' como Excel: {e}]", [], {}

    parts = [f"===== ARQUIVO: {filename} ====="]
    skip_sheets = set()
    contas: list = []

    consolidado = detect_consolidated_balancete(wb)
    if consolidado:
        contas = parse_consolidated_balancete(wb, consolidado)
        skip_sheets.add(consolidado["aba"])
        # As 12 abas mensais (se existirem) ficam redundantes com a
        # consolidada — não processa as duas, e tira ambas do despejo cru.
        monthly_sheets = detect_monthly_sheets(wb)
        skip_sheets |= {name for _, name in monthly_sheets}
        parts.append(
            f"===== SÉRIE MENSAL CONSOLIDADA (extraída automaticamente da aba '{consolidado['aba']}') =====\n"
            f"{len(contas)} contas encontradas — não precisa copiar no seu output, "
            "o código já anexa em raw_extracted.series_contabeis.\n"
            "ATENÇÃO: contas de resultado (receita/despesa) costumam vir ACUMULADAS no "
            "ano-calendário — para o valor do mês isolado, calcule a diferença entre um "
            "mês e o anterior. Contas de balanço (ativo/passivo) já são o saldo no fim "
            "daquele mês, sem precisar de ajuste."
        )
    else:
        monthly_sheets = detect_monthly_sheets(wb)
        skip_sheets |= {name for _, name in monthly_sheets}
        if monthly_sheets:
            accounts_dict = merge_monthly_balancete(wb, monthly_sheets)
            contas = [
                {"conta": nome, "codigo": codigo, "tipo": data["tipo"],
                 "valores": {f"mes_{m:02d}": v for m, v in data["meses"].items()}}
                for (nome, codigo), data in accounts_dict.items()
            ]
            parts.append(format_merged_table(accounts_dict))

    dre_linhas: dict = {}
    dre_deteccao = detect_dre_sheet(wb)
    if dre_deteccao:
        dre_linhas = parse_dre_sheet(wb, dre_deteccao)
        skip_sheets.add(dre_deteccao["aba"])
        if dre_linhas:
            parts.append(format_dre_table(dre_linhas))

    sheet_names = [s for s in wb.sheetnames if s not in skip_sheets][:MAX_SHEETS]
    if len(wb.sheetnames) - len(skip_sheets) > MAX_SHEETS:
        parts.append(f"(arquivo tem mais abas — mostrando as primeiras {MAX_SHEETS} além da série mensal/DRE)")

    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        parts.append(f"\n--- Aba: {sheet_name} ---")
        row_count = 0
        for row in ws.iter_rows(values_only=True):
            if row_count >= MAX_ROWS_PER_SHEET:
                parts.append(f"[... aba truncada em {MAX_ROWS_PER_SHEET} linhas ...]")
                break
            row_trimmed = row[:MAX_COLS_PER_SHEET]
            if all(v is None for v in row_trimmed):
                row_count += 1
                continue
            values = [("" if v is None else str(v)) for v in row_trimmed]
            parts.append(f"L{row_count+1}: " + " | ".join(values))
            row_count += 1

    return "\n".join(parts), contas, dre_linhas


def select_relevant_accounts(series: list, top_n: int = MAX_ACCOUNTS_FOR_FINANCIAL_AGENT) -> list:
    """Reduz a série contábil consolidada (potencialmente 300+ contas) às
    que interessam pra análise financeira: maior volatilidade mês a mês
    (candidata a anomalia) combinada com materialidade absoluta. Isso é
    estatística determinística — não pedimos pro modelo vasculhar a
    tabela inteira pra achar o que o código já sabe calcular em O(n)."""
    def score(item: dict) -> float:
        valores = [v for v in item.get("valores", {}).values() if isinstance(v, (int, float))]
        if not valores:
            return 0.0
        materialidade = max(abs(v) for v in valores)
        if len(valores) < 3:
            return materialidade
        deltas = [b - a for a, b in zip(valores, valores[1:])]
        volatilidade = statistics.pstdev(deltas)
        # combina os dois sinais — uma conta pequena mas muito instável
        # (ex.: pró-labore que dobra num mês) também precisa aparecer,
        # não só as maiores em valor absoluto.
        return volatilidade + 0.1 * materialidade

    return sorted(series, key=score, reverse=True)[:top_n]


def supabase_request(method: str, path: str, body: dict | None = None) -> dict:
    """Chamada simples à API REST do Supabase (PostgREST), usando a service_role key."""
    url = os.environ["SUPABASE_URL"].rstrip("/") + "/rest/v1/" + path
    key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    }
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        print(f"Erro Supabase [{method} {path}]: {e.code} {e.read().decode()}", file=sys.stderr)
        raise


def get_deal_data(deal_id: str) -> dict:
    rows = supabase_request("GET", f"deal_data?deal_id=eq.{deal_id}&order=created_at.desc&limit=1")
    if not rows:
        raise SystemExit(f"Nenhum deal_data encontrado para o deal {deal_id} — rode o Agent 00 primeiro.")
    return rows[0]


def download_from_storage(storage_ref: str) -> bytes:
    url = os.environ["SUPABASE_URL"].rstrip("/") + "/storage/v1/object/deal-files/" + storage_ref
    key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    req = urllib.request.Request(url, headers={"apikey": key, "Authorization": f"Bearer {key}"})
    with urllib.request.urlopen(req) as resp:
        return resp.read()


def run_extraction(deal_id: str):
    """Caminho próprio do Agent 00: baixa os arquivos do deal, converte pra
    texto, chama o Claude, e CRIA um deal_data novo (os outros agentes só
    leem um deal_data que já existe — este é quem produz).

    A série mensal consolidada (quando existe) é calculada pelo código e
    injetada direto no resultado — não pedimos pro Claude copiar de volta
    uma tabela de centenas de contas x 12 meses que o código já tem. Isso
    era o que estava estourando o timeout de 10 minutos (visto em 19/08)."""
    files = supabase_request("GET", f"files?deal_id=eq.{deal_id}")
    if not files:
        raise SystemExit(f"Nenhum arquivo encontrado para o deal {deal_id} — nada para extrair.")

    agent_version = get_active_agent_version("extraction")

    combined_text = []
    raw_bytes_for_checksum = b""
    code_computed_series = []  # preenchido diretamente pelo código, não pelo LLM
    code_computed_dre = {}     # idem, pela mesma razão
    for f in files:
        try:
            content = download_from_storage(f["storage_ref"])
        except urllib.error.HTTPError as e:
            combined_text.append(f"===== ARQUIVO: {f['original_filename']} =====\n[ERRO ao baixar do Storage: {e.code}]")
            continue
        raw_bytes_for_checksum += content
        filename = f["original_filename"] or f["storage_ref"]

        texto, contas, dre = excel_to_text(content, filename)
        combined_text.append(texto)
        for conta in contas:
            conta.setdefault("arquivo", filename)
            code_computed_series.append(conta)
        if dre:
            code_computed_dre[filename] = dre

    full_dump = "\n\n".join(combined_text)
    checksum = hashlib.sha256(raw_bytes_for_checksum).hexdigest()

    existing = supabase_request("GET", f"deal_data?deal_id=eq.{deal_id}&checksum=eq.{checksum}")
    if existing:
        print("[extraction] mesmo conjunto de arquivos já extraído antes — pulando (idempotência).")
        return

    output = call_claude(
        agent_version["system_prompt"], {"arquivos_em_texto": full_dump},
        model=agent_version.get("model"),
    )

    # Injeta a série e a DRE calculadas pelo código — sobrepõe o que o LLM
    # eventualmente tenha tentado copiar (nunca confiamos numa cópia manual
    # de centenas de linhas, e a DRE já vem pronta e categorizada).
    if code_computed_series:
        output.setdefault("raw_extracted", {})["series_contabeis"] = code_computed_series
    if code_computed_dre:
        output.setdefault("raw_extracted", {})["dre_estruturada"] = code_computed_dre

    deal_data_row = supabase_request("POST", "deal_data", {
        "deal_id": deal_id,
        "schema_version": "v1",
        "structured": output.get("structured", {}),
        "raw_extracted": output.get("raw_extracted", {}),
        "checksum": checksum,
    })
    deal_data_id = deal_data_row[0]["id"] if isinstance(deal_data_row, list) else deal_data_row["id"]

    # Registra também em agent_runs, pro AI Audit Log não ter um buraco
    # justamente no primeiro agente da cadeia.
    supabase_request("POST", "agent_runs", {
        "deal_id": deal_id,
        "agent_version_id": agent_version["id"],
        "deal_data_id": deal_data_id,
        "input_hash": checksum,
        "status": "completed",
        "output": output,
        "confidence": output.get("confidence"),
    })
    print(f"[extraction] deal_data criado ({deal_data_id}), confidence={output.get('confidence')}")


def get_active_agent_version(agent_name: str) -> dict:
    agents = supabase_request("GET", f"agents?name=eq.{agent_name}")
    if not agents:
        raise SystemExit(f"Agente '{agent_name}' não encontrado na tabela agents.")
    agent_id = agents[0]["id"]

    versions = supabase_request(
        "GET", f"agent_versions?agent_id=eq.{agent_id}&active=eq.true&order=created_at.desc&limit=1"
    )
    if not versions:
        raise SystemExit(f"Nenhuma versão ativa para o agente '{agent_name}'.")
    version = versions[0]

    prompts = supabase_request("GET", f"agent_prompts?agent_version_id=eq.{version['id']}")
    if not prompts:
        raise SystemExit(f"Sem system_prompt gravado para a versão {version['id']} de '{agent_name}'.")

    version["system_prompt"] = prompts[0]["system_prompt"]
    return version


def call_claude(system_prompt: str, deal_data: dict, other_outputs: dict | None = None,
                 model: str | None = None) -> dict:
    """Chama o Claude Code em modo headless e valida que a resposta é JSON.

    O prompt vai por stdin, não como argumento de linha de comando — prompts
    grandes (planilhas reais podem passar de 100 mil caracteres) estouram o
    limite do sistema operacional para argumentos (visto na prática em
    19/08: OSError "Argument list too long"). Documentação oficial confirma
    stdin como o caminho certo para isso: `echo "..." | claude -p`.

    IMPORTANTE (achado revisando o script em 20/08): a chamada aqui embaixo
    nunca teve `--model` — ou seja, toda troca de modelo feita via
    agent_versions.model no banco (inclusive tirar o CFO Synthesis de Opus)
    NUNCA teve efeito real. O agente sempre rodou no modelo padrão da sessão
    do Claude Code, não no que estava configurado. Corrigido abaixo — sem
    isso, nada do resto desta otimização de custo se sustenta."""
    user_payload = {"deal_data": deal_data}
    if other_outputs:
        user_payload["outputs_dos_outros_agentes"] = other_outputs

    full_prompt = system_prompt + "\n\n# Dados do deal\n\n" + json.dumps(user_payload, ensure_ascii=False, indent=2)

    result = subprocess.run(
        ["claude", "-p", "--model", model or DEFAULT_MODEL_ALIAS, "--output-format", "json"],
        input=full_prompt,
        capture_output=True,
        text=True,
        timeout=600,  # prompts grandes com Opus podem demorar mais que os 300s originais
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"claude -p saiu com código {result.returncode}\n"
            f"--- stderr ---\n{result.stderr!r}\n"
            f"--- stdout (primeiros 2000 chars) ---\n{result.stdout[:2000]!r}"
        )

    # --output-format json do Claude Code envolve a resposta num envelope;
    # o campo com o texto do agente é "result" (confirmado no teste de 17/08).
    raw = json.loads(result.stdout)
    agent_text = raw.get("result", raw.get("content", result.stdout))

    parsed = _extract_json(agent_text)
    if parsed is None:
        raise RuntimeError(f"Resposta do agente não é um JSON válido.\nResposta bruta: {agent_text[:500]}")
    return parsed


def _extract_json(text: str) -> dict | None:
    """Extrai JSON de uma resposta de LLM, tolerando variações comuns:
    JSON puro, cercado em ```json ... ```, cercado em ``` ... ``` simples,
    ou com texto explicativo antes/depois do bloco. Tenta na ordem do mais
    estrito pro mais tolerante."""
    text = text.strip()

    # 1. JSON puro
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 2. Bloco cercado em ```json ... ``` ou ``` ... ```, em qualquer parte do texto
    fence_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if fence_match:
        try:
            return json.loads(fence_match.group(1))
        except json.JSONDecodeError:
            pass

    # 3. Último recurso: do primeiro "{" ao último "}" no texto inteiro
    first, last = text.find("{"), text.rfind("}")
    if first != -1 and last != -1 and last > first:
        try:
            return json.loads(text[first:last + 1])
        except json.JSONDecodeError:
            pass

    return None


def compute_input_hash(deal_id: str, agent_version_id: str, checksum: str) -> str:
    raw = f"{deal_id}:{agent_version_id}:{checksum}"
    return hashlib.sha256(raw.encode()).hexdigest()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent", required=True)
    parser.add_argument("--deal-id", required=True)
    args = parser.parse_args()

    if args.agent == "extraction":
        run_extraction(args.deal_id)
        return

    deal_data = get_deal_data(args.deal_id)
    agent_version = get_active_agent_version(args.agent)

    input_hash = compute_input_hash(args.deal_id, agent_version["id"], deal_data["checksum"])

    # Idempotência: se já existe um run com esse hash exato, não reprocessa.
    table = "synthesis_runs" if args.agent == "cfo_synthesis" else "agent_runs"
    existing = supabase_request(
        "GET",
        f"{table}?deal_id=eq.{args.deal_id}&agent_version_id=eq.{agent_version['id']}"
        + ("" if table == "synthesis_runs" else f"&input_hash=eq.{input_hash}"),
    )
    if existing:
        print(f"[{args.agent}] já existe um run para este input — pulando (idempotência).")
        return

    # Só o conteúdo relevante vai para o prompt — id/checksum são bookkeeping
    # interno, não fazem parte do que o agente precisa analisar. A série
    # contábil crua só vai pro agente financeiro, e mesmo assim já reduzida
    # às contas relevantes (ver select_relevant_accounts) — os demais
    # recebem só o resumo estruturado.
    agent_input = {"structured": deal_data.get("structured", {})}
    if args.agent in AGENTS_NEED_RAW_SERIES:
        raw = dict(deal_data.get("raw_extracted", {}))

        # Motor determinístico (financial_engine.py + dre_balancete_parser.py):
        # o Claude não recebe mais a série crua pra "descobrir" EBITDA e
        # anomalia sozinho — código já calculou, ele só interpreta em poucas
        # palavras o que já foi achado. Corta contexto de entrada E o
        # "raciocínio" caro que o modelo fazia pra vasculhar 300+ contas.
        dre_estruturada = raw.get("dre_estruturada") or {}
        if dre_estruturada:
            # DRE é a fonte PRIMÁRIA — já vem categorizada pela própria
            # planilha da empresa, sem risco de duplicar hierarquia (ver
            # 21/08: o waterfall reconstruído do Balancete cru precisou de
            # 5 correções pra bater, a DRE bateu de primeira).
            primeira_dre = next(iter(dre_estruturada.values()))
            raw["ebitda_calculado"] = calcular_ebitda_de_dre(primeira_dre)
            raw["anomalias_detectadas"] = detectar_anomalias_run_rate(
                dre_linhas_para_contas(primeira_dre), top_n=10
            )
        else:
            # Sem DRE própria no arquivo — cai pro Balancete cru como plano B.
            series = raw.get("series_contabeis") or []
            if series:
                raw["waterfall_calculado"] = mapear_waterfall(series)
                raw["ebitda_calculado"] = calcular_ebitda_3_camadas(
                    raw["waterfall_calculado"], lucro_liquido=None,
                    tributos_sobre_lucro=None, resultado_financeiro=None,
                )
                raw["anomalias_detectadas"] = detectar_anomalias_run_rate(series, top_n=10)

        series = raw.get("series_contabeis") or []
        if series:
            raw["contas_novas_ou_zeradas"] = detectar_contas_novas_ou_zeradas(series)
            # A série crua completa não vai mais — o Claude já recebeu o que
            # importa dela (anomalias + novas/zeradas) calculado acima.
            raw.pop("series_contabeis", None)
            raw["series_contabeis_nota"] = (
                "Série crua não incluída — EBITDA e anomalias já calculados por código "
                "(ver ebitda_calculado / anomalias_detectadas / contas_novas_ou_zeradas)."
            )
        agent_input["raw_extracted"] = raw

    # CFO Synthesis precisa ver o que os outros 5 agentes concluíram pra
    # reconciliar divergências — o caminho genérico não levava isso (gap
    # encontrado revisando o script em 20/08). Busca os agent_runs do deal
    # e monta {nome_do_agente: output}.
    # NOTA: não testado contra o Supabase real ainda — a sintaxe de select
    # aninhado do PostgREST assume que agent_runs.agent_version_id ->
    # agent_versions.agent_id -> agents.id são as FKs reais; confirmar
    # antes do primeiro run em produção.
    other_outputs = None
    if args.agent == "cfo_synthesis":
        prior_runs = supabase_request(
            "GET",
            f"agent_runs?deal_id=eq.{args.deal_id}"
            "&select=output,agent_versions(agents(name))",
        )
        other_outputs = {
            r["agent_versions"]["agents"]["name"]: r["output"]
            for r in prior_runs
            if r.get("agent_versions", {}).get("agents", {}).get("name")
        }

    output = call_claude(
        agent_version["system_prompt"], agent_input, other_outputs,
        model=agent_version.get("model"),
    )

    row = {
        "deal_id": args.deal_id,
        "agent_version_id": agent_version["id"],
        "status": "completed",
        "output": output,
        "confidence": output.get("confidence"),
    }
    if table == "agent_runs":
        row["deal_data_id"] = deal_data["id"]
        row["input_hash"] = input_hash
    else:
        row["recommendation"] = output.get("recommendation")

    supabase_request("POST", table, row)
    print(f"[{args.agent}] concluído e gravado em {table}.")

    # O CFO Synthesis é o último passo da análise em si (Output Generator é
    # depois, separado) — é o momento certo de marcar o deal como concluído.
    if args.agent == "cfo_synthesis":
        supabase_request("PATCH", f"deals?id=eq.{args.deal_id}", {"status": "completed"})
        print("deal marcado como completed.")


if __name__ == "__main__":
    main()
