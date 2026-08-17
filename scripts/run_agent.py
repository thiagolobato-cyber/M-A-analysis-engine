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
import hashlib
import json
import os
import subprocess
import sys
import urllib.request
import urllib.error


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
    # o campo com o texto do agente pode variar de versão — ajustar depois
    # do primeiro teste real contra o CLI.
    raw = json.loads(result.stdout)
    agent_text = raw.get("result", raw.get("content", result.stdout))

    try:
        return json.loads(agent_text)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"Resposta do agente não é um JSON válido: {e}\nResposta bruta: {agent_text[:500]}")


def compute_input_hash(deal_id: str, agent_version_id: str, checksum: str) -> str:
    raw = f"{deal_id}:{agent_version_id}:{checksum}"
    return hashlib.sha256(raw.encode()).hexdigest()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent", required=True)
    parser.add_argument("--deal-id", required=True)
    args = parser.parse_args()

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


if __name__ == "__main__":
    main()
