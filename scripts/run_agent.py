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
)

MAX_ROWS_PER_SHEET = 300
MAX_COLS_PER_SHEET = 40
MAX_SHEETS = 30
MONTHLY_SHEET_RE = re.compile(r'(?i)^m[eê]s\s*0?(\d{1,2})$')

# Só o agente que lê
