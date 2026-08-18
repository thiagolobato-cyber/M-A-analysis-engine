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
import subprocess
import sys
import urllib.request
import urllib.error

from openpyxl import load_workbook

MAX_ROWS_PER_SHEET = 300
MAX_COLS_PER_SHEET = 40
MAX_SHEETS = 30


def excel_to_text(file_bytes: bytes, filename: str) -> str:
    """Converte um Excel em texto legível pro Claude — não interpreta,
    só descreve fielmente o que tem em cada aba, linha por linha.
    Testado localmente contra os 8 arquivos reais do BHub antes de ir
    pra produção (incluindo um com 24 abas e outro com erro #REF!)."""
    try:
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    except Exception as e:
        return f"[ERRO ao abrir '{filename}' como Excel: {e}]"

    parts = [f"===== ARQUIVO: {filename} ====="]
    sheet_names = wb.sheetnames[:MAX_SHEETS]
    if len(wb.sheetnames) > MAX_SHEETS:
        parts.append(f"(arquivo tem {len(wb.sheetnames)} abas — mostrando as primeiras {MAX_SHEETS})")

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

    return "\n".join(parts)


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
    leem um deal_data que já existe — este é quem produz)."""
    files = supabase_request("GET", f"files?deal_id=eq.{deal_id}")
    if not files:
        raise SystemExit(f"Nenhum arquivo encontrado para o deal {deal_id} — nada para extrair.")

    agent_version = get_active_agent_version("extraction")

    combined_text = []
    raw_bytes_for_checksum = b""
    for f in files:
        try:
            content = download_from_storage(f["storage_ref"])
        except urllib.error.HTTPError as e:
            combined_text.append(f"===== ARQUIVO: {f['original_filename']} =====\n[ERRO ao baixar do Storage: {e.code}]")
            continue
        raw_bytes_for_checksum += content
        combined_text.append(excel_to_text(content, f["original_filename"] or f["storage_ref"]))

    full_dump = "\n\n".join(combined_text)
    checksum = hashlib.sha256(raw_bytes_for_checksum).hexdigest()

    # Idempotência: mesmo conjunto de arquivos (mesmo checksum) já foi extraído?
    existing = supabase_request("GET", f"deal_data?deal_id=eq.{deal_id}&checksum=eq.{checksum}")
    if existing:
        print("[extraction] mesmo conjunto de arquivos já extraído antes — pulando (idempotência).")
        return

    output = call_claude(agent_version["system_prompt"], {"arquivos_em_texto": full_dump})

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


def call_claude(system_prompt: str, deal_data: dict, other_outputs: dict | None = None) -> dict:
    """Chama o Claude Code em modo headless e valida que a resposta é JSON."""
    user_payload = {"deal_data": deal_data}
    if other_outputs:
        user_payload["outputs_dos_outros_agentes"] = other_outputs

    full_prompt = system_prompt + "\n\n# Dados do deal\n\n" + json.dumps(user_payload, ensure_ascii=False, indent=2)

    result = subprocess.run(
        ["claude", "-p", full_prompt, "--output-format", "json"],
        capture_output=True,
        text=True,
        timeout=300,
    )
    if result.returncode != 0:
        raise RuntimeError(f"claude -p saiu com erro: {result.stderr}")

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
    # interno, não fazem parte do que o agente precisa analisar.
    agent_input = {
        "structured": deal_data.get("structured", {}),
        "raw_extracted": deal_data.get("raw_extracted", {}),
    }
    output = call_claude(agent_version["system_prompt"], agent_input)

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
