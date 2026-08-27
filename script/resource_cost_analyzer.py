"""
resource_cost_analyzer.py
==========================
Coleta o custo dos últimos N meses para uma lista de recursos informada em Excel.

REGRA PRINCIPAL
---------------
- Se o recurso informado for uma VM  -> soma custo da VM + discos anexados + NIC(s) + IP público.
- Se o recurso informado for qualquer outra coisa (disco, storage, etc.) -> soma só aquele ResourceId.
- Backup / Logs / Segurança NÃO têm vínculo direto de ResourceId com a VM no billing da Azure
  (são cobrados no Recovery Services Vault / Log Analytics Workspace / Defender, não na VM).
  Por isso são tratados à parte, por aproximação (nome da VM contido em RG/Tags), e sempre
  marcados como "Aproximado=True" na saída — nunca somados silenciosamente ao total da VM.

PRÉ-REQUISITOS
---------------
    pip install azure-identity azure-mgmt-compute azure-mgmt-network azure-mgmt-resourcegraph azure-mgmt-costmanagement azure-mgmt-subscription pandas openpyxl

Autenticação: usa a sessão do Azure CLI já logada (az login), via DefaultAzureCredential.
Não precisa de client secret nem app registration.

FALLBACK EM NUVEM
------------------
Se um nome da lista não bater com nada na base local (AzureBase_FinOps.csv), o script
consulta a Azure diretamente:
  1. Azure Resource Graph — localiza o recurso pelo nome, em todas as subscriptions que
     aparecem na sua base local (não precisa cadastrar IDs manualmente).
  2. Azure Cost Management API — busca o custo real dos últimos 6 meses direto na Azure
     para esse recurso (e, se for VM, para disco/rede associados também).
Isso cobre recursos novos, renomeados, ou que por algum motivo não entraram na extração
mensal. Linhas resolvidas assim vêm marcadas em Observações como "custo consultado direto
na Azure". Use --no-cloud-fallback para desligar esse comportamento.

ENTRADA
-------
Excel com uma coluna "Recurso" (nome do recurso, como aparece no campo "Resource" do
AzureBase_FinOps.csv). Uma linha por recurso.

SAÍDA
-----
Excel com 1 linha por recurso: Mes1 ... Mes6 (custo total) + colunas de detalhe
por categoria (VM, Disco, Rede, Backup_Aprox) + observações.

Uso:
    python resource_cost_analyzer.py --input lista_recursos.xlsx --output custos_saida.xlsx
"""

import argparse
import hashlib
import json
import logging
import os
import re
import sys
import time
import traceback
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ajuste os caminhos e nomes de coluna conforme seu ambiente
# ---------------------------------------------------------------------------

BASE_CSV_PATH = "AzureBase_FinOps.csv"     # base consolidada (semicolon, UTF-8-BOM)
CSV_SEP = ";"
CSV_ENCODING = "utf-8-sig"

COL_RESOURCE = "Resource"
COL_RESOURCE_ID = "ResourceId"
COL_RESOURCE_TYPE = "ResourceType"
COL_RESOURCE_GROUP = "ResourceGroupName"
COL_SUBSCRIPTION_NAME = "SubscriptionName"
COL_COST = "CostUSD"                        # troque para "Cost" se quiser moeda local
COL_MONTH = "Mês"                           # nome real da coluna no AzureBase_FinOps.csv
COL_TAGS = "Tags"

N_MONTHS = 6

# Tipos de serviço tratados como "aproximados" (não têm ResourceId da VM no billing)
APPROX_SERVICE_KEYWORDS = ["backup", "recovery services", "log analytics",
                            "microsoft defender", "sentinel"]

RESOURCE_ID_SUB_REGEX = re.compile(r"/subscriptions/([0-9a-fA-F-]{36})/", re.IGNORECASE)

# Este script vive em <raiz do projeto>/script/resource_cost_analyzer.py.
# CONFIG_DIR (custos_saida.xlsx.state.json) e logs/ ficam na raiz do projeto, um nível acima.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
CONFIG_DIR = os.path.join(PROJECT_ROOT, "config")

LOGGER_NAME = "resource_cost_analyzer"


def _compute_script_version() -> str:
    """
    Hash do conteúdo deste próprio arquivo. Usado no cache de estado pra detectar
    automaticamente quando o script foi atualizado — se a versão salva no cache não
    bater com a versão atual, o item é reprocessado (não fica preso num resultado
    calculado por uma lógica antiga/com bug já corrigido).
    """
    try:
        with open(os.path.abspath(__file__), "rb") as f:
            return hashlib.md5(f.read()).hexdigest()[:12]
    except Exception:
        return "unknown"


SCRIPT_VERSION = _compute_script_version()

STATUS_OK = "OK"
STATUS_ERRO = "ERRO"
STATUS_NAO_ENCONTRADO = "NAO_ENCONTRADO"

MAX_RETRIES = 3
RETRY_BASE_DELAY_SECONDS = 2


def call_with_retries(func, description: str, *args, **kwargs):
    """
    Chama func(*args, **kwargs) com até MAX_RETRIES tentativas. Cada falha é logada
    (detalhe completo vai pro log; console só mostra que está tentando de novo).
    Se todas as tentativas falharem, propaga a última exceção pro chamador decidir
    o que fazer (normalmente: marcar a linha como ERRO).
    """
    logger = logging.getLogger(LOGGER_NAME)
    last_exc = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            last_exc = e
            logger.warning("Tentativa %d/%d falhou (%s): %s", attempt, MAX_RETRIES, description, e)
            if attempt < MAX_RETRIES:
                print(f"     [tentativa {attempt}/{MAX_RETRIES} falhou: {description} — tentando de novo...]")
                time.sleep(RETRY_BASE_DELAY_SECONDS * attempt)
    logger.error("Todas as %d tentativas falharam (%s): %s\n%s",
                 MAX_RETRIES, description, last_exc, traceback.format_exc())
    raise last_exc


def setup_logging() -> str:
    """
    Cria a pasta logs/ (se não existir) na raiz do projeto — um nível acima da pasta
    onde este script está (a pasta "script"), já que o script agora fica em subpasta
    e não na raiz — e configura um logger que grava detalhes técnicos completos (erros,
    tracebacks) em arquivo de texto. A tela e a planilha de saída mostram só
    mensagens resumidas — o log é onde fica o detalhe técnico.
    """
    log_dir = os.path.join(PROJECT_ROOT, "logs")
    os.makedirs(log_dir, exist_ok=True)

    log_path = os.path.join(log_dir, f"execucao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(file_handler)

    return log_path


# ---------------------------------------------------------------------------
# CARGA DA BASE DE CUSTOS
# ---------------------------------------------------------------------------

def load_cost_base(path: str) -> pd.DataFrame:
    """
    Carrega a base de custos. A coluna de custo pode vir em formato PT-BR (vírgula
    decimal, ex.: '60,36') dependendo de como o CSV foi gerado/exportado — se lida
    direto com pd.to_numeric, isso vira NaN silenciosamente (comma não é separador
    decimal válido) e o custo real de recursos ativos aparece como zero.
    """
    df = pd.read_csv(path, sep=CSV_SEP, encoding=CSV_ENCODING)

    cost_str = df[COL_COST].astype(str)
    # só troca vírgula por ponto em valores que parecem PT-BR (têm vírgula e não têm ponto já),
    # pra não quebrar o caso em que o CSV já vier com ponto decimal (formato EN-US)
    looks_pt_br = cost_str.str.contains(",", na=False) & ~cost_str.str.contains(r"\.", na=False)
    cost_str = cost_str.where(~looks_pt_br, cost_str.str.replace(",", ".", regex=False))

    df[COL_COST] = pd.to_numeric(cost_str, errors="coerce").fillna(0)
    return df


def last_n_months(df: pd.DataFrame, n: int) -> list:
    """Retorna os últimos n valores distintos da coluna de mês, em ordem cronológica."""
    months = sorted(df[COL_MONTH].dropna().unique().tolist())
    return months[-n:]


# ---------------------------------------------------------------------------
# CACHE DE EXECUÇÃO — pula recursos que já deram certo, reprocessa só os com erro
# ---------------------------------------------------------------------------

def state_path_for(output_path: str) -> str:
    """O cache de estado vai para config\\, usando só o nome do arquivo de saída (não o caminho)."""
    os.makedirs(CONFIG_DIR, exist_ok=True)
    return os.path.join(CONFIG_DIR, os.path.basename(output_path) + ".state.json")


def load_state(state_path: str) -> dict:
    if not os.path.exists(state_path):
        return {}
    try:
        with open(state_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logging.getLogger(LOGGER_NAME).warning("Não foi possível ler o cache de estado (%s): %s — começando do zero.", state_path, e)
        return {}


def save_state(state_path: str, state: dict) -> None:
    with open(state_path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# DETECÇÃO DO RECURSO INFORMADO PELO USUÁRIO
# ---------------------------------------------------------------------------

def _candidates_from_matches(matches: pd.DataFrame) -> dict:
    if matches.empty:
        return {"found": False}
    # dedup final por ResourceId ignorando case — o mesmo recurso pode aparecer com
    # capitalização diferente em snapshots diferentes da base, e isso NÃO é um recurso
    # "duplicado" de verdade, é o mesmo recurso registrado duas vezes.
    matches = matches.copy()
    matches["_rid_lower"] = matches[COL_RESOURCE_ID].str.lower()
    matches = matches.drop_duplicates(subset=["_rid_lower"])

    candidates = []
    for _, row in matches.iterrows():
        sub_match = RESOURCE_ID_SUB_REGEX.search(row[COL_RESOURCE_ID])
        candidates.append({
            "resource_id": row[COL_RESOURCE_ID],
            "resource_type": row[COL_RESOURCE_TYPE],
            "resource_group": row[COL_RESOURCE_GROUP],
            "subscription_name": row[COL_SUBSCRIPTION_NAME],
            "subscription_id": sub_match.group(1) if sub_match else None,
        })
    return {"found": True, "candidates": candidates}


def find_resource_exact(df: pd.DataFrame, name: str) -> dict:
    """
    Match EXATO (case-insensitive) no campo Resource ou no último segmento do ResourceId.
    Isso é o que diferencia 'vm-teste01' de 'vm-teste01_disk1' — sem essa etapa, o match
    parcial abaixo casaria com qualquer coisa que contenha o nome como substring.
    """
    name_lower = name.lower()
    last_segment = df[COL_RESOURCE_ID].str.split("/").str[-1].str.lower()
    exact_mask = (df[COL_RESOURCE].str.lower() == name_lower) | (last_segment == name_lower)
    matches = df.loc[exact_mask, [COL_RESOURCE, COL_RESOURCE_ID, COL_RESOURCE_TYPE,
                                   COL_RESOURCE_GROUP, COL_SUBSCRIPTION_NAME]]
    return _candidates_from_matches(matches)


def find_resource_partial(df: pd.DataFrame, name: str) -> dict:
    """
    Match parcial (contains) — último recurso, só usado se nada bateu em nenhuma das
    etapas anteriores (exato, nem Resource Group). O ResourceId sempre contém o nome
    do Resource Group como substring, então isso teria que vir DEPOIS da checagem de
    Resource Group, senão nomes de RG nunca chegariam lá.
    """
    name_lower = name.lower()
    partial_mask = (
        df[COL_RESOURCE].str.lower().str.contains(re.escape(name_lower), na=False)
        | df[COL_RESOURCE_ID].str.lower().str.contains(re.escape(name_lower), na=False)
    )
    matches = df.loc[partial_mask, [COL_RESOURCE, COL_RESOURCE_ID, COL_RESOURCE_TYPE,
                                     COL_RESOURCE_GROUP, COL_SUBSCRIPTION_NAME]]
    return _candidates_from_matches(matches)


def is_vm(resource_type: str) -> bool:
    """
    Reconhece VM tanto pelo tipo técnico do ARM ('microsoft.compute/virtualmachines')
    quanto pelo nome amigável que aparece em algumas bases ('Virtual Machine', com
    espaço) — sem remover o espaço antes de comparar, 'virtual machine' nunca bate
    com 'virtualmachines' e a VM inteira passa a ser tratada como recurso isolado.
    """
    if not resource_type:
        return False
    t = resource_type.lower().replace(" ", "")
    return "virtualmachine" in t and "virtualmachinescaleset" not in t and "virtualmachineimage" not in t


# ---------------------------------------------------------------------------
# RESOURCE GROUP — quando o nome informado não é um recurso, mas um RG inteiro
# ---------------------------------------------------------------------------

def find_resource_group_local(df: pd.DataFrame, name: str) -> list:
    """
    Verifica se 'name' bate exatamente (case-insensitive) com um ResourceGroupName
    da base local. Retorna uma lista de candidatos {resource_group, subscription_name,
    subscription_id} — pode haver mais de um se o mesmo nome de RG existir em mais de
    uma subscription.
    """
    name_lower = name.lower()
    mask = df[COL_RESOURCE_GROUP].astype(str).str.lower() == name_lower
    matches = df.loc[mask, [COL_RESOURCE_GROUP, COL_SUBSCRIPTION_NAME, COL_RESOURCE_ID]].copy()
    matches["_rg_lower"] = matches[COL_RESOURCE_GROUP].astype(str).str.lower()
    matches = matches.drop_duplicates(subset=["_rg_lower", COL_SUBSCRIPTION_NAME])

    candidates = []
    for _, row in matches.iterrows():
        sub_match = RESOURCE_ID_SUB_REGEX.search(row[COL_RESOURCE_ID])
        candidates.append({
            "resource_group": row[COL_RESOURCE_GROUP],
            "subscription_name": row[COL_SUBSCRIPTION_NAME],
            "subscription_id": sub_match.group(1) if sub_match else None,
        })
    return candidates


def aggregate_resource_group_local(df: pd.DataFrame, resource_group: str, subscription_name: str, months: list) -> dict:
    """Soma o custo de TODOS os recursos dentro do Resource Group (base local), por mês, e conta quantos recursos há."""
    mask = (
        (df[COL_RESOURCE_GROUP].astype(str).str.lower() == resource_group.lower())
        & (df[COL_SUBSCRIPTION_NAME] == subscription_name)
    )
    subset = df.loc[mask]
    totals = {m: subset.loc[subset[COL_MONTH] == m, COL_COST].sum() for m in months}
    n_resources = subset[COL_RESOURCE_ID].nunique()
    return totals, n_resources


def cloud_search_resource_group(subscription_ids: list, name: str) -> list:
    """Verifica na Azure (Resource Graph) se 'name' é um Resource Group, em qualquer das subscriptions informadas."""
    try:
        from azure.identity import DefaultAzureCredential
        from azure.mgmt.resourcegraph import ResourceGraphClient
        from azure.mgmt.resourcegraph.models import QueryRequest
    except ImportError:
        raise RuntimeError("Biblioteca não instalada. Rode:\n  pip install azure-mgmt-resourcegraph")

    credential = DefaultAzureCredential()
    client = ResourceGraphClient(credential)
    query = f"ResourceContainers | where type =~ 'microsoft.resources/subscriptions/resourcegroups' and name =~ '{name}'"
    request = QueryRequest(subscriptions=subscription_ids, query=query)
    response = client.resources(request)
    return response.data or []


def _find_column_index(columns: list, *fragments: str):
    for i, c in enumerate(columns):
        if any(frag.lower() in c.lower() for frag in fragments):
            return i
    return None


def get_cloud_cost_for_resource_group(subscription_id: str, resource_group_name: str, months: list) -> dict:
    """Soma o custo de TODOS os recursos do Resource Group direto na Azure (Cost Management API), por mês."""
    from datetime import datetime, timedelta
    try:
        from azure.identity import DefaultAzureCredential
        from azure.mgmt.costmanagement import CostManagementClient
        from azure.mgmt.costmanagement.models import (
            QueryDefinition, QueryDataset, QueryAggregation, QueryFilter,
            QueryComparisonExpression, QueryTimePeriod,
        )
    except ImportError:
        raise RuntimeError("Biblioteca não instalada. Rode:\n  pip install azure-mgmt-costmanagement")

    credential = DefaultAzureCredential()
    client = CostManagementClient(credential)
    scope = f"/subscriptions/{subscription_id}"

    start_dt = datetime.strptime(months[0] + "-01", "%Y-%m-%d")
    y, m = map(int, months[-1].split("-"))
    end_dt = (datetime(y + 1, 1, 1) if m == 12 else datetime(y, m + 1, 1)) - timedelta(days=1)

    definition = QueryDefinition(
        type="ActualCost",
        timeframe="Custom",
        time_period=QueryTimePeriod(from_property=start_dt, to=end_dt),
        dataset=QueryDataset(
            granularity="Monthly",
            aggregation={"totalCost": QueryAggregation(name="Cost", function="Sum")},
            filter=QueryFilter(
                dimensions=QueryComparisonExpression(
                    name="ResourceGroupName", operator="In",
                    values_property=[resource_group_name.lower()]
                )
            ),
        ),
    )

    result = client.query.usage(scope, definition)
    columns = [c.name for c in result.columns]
    idx_cost = _find_column_index(columns, "cost")
    idx_date = _find_column_index(columns, "usagedate", "billingmonth", "date")

    monthly_totals = {m: 0.0 for m in months}
    for row in result.rows or []:
        cost = row[idx_cost] if idx_cost is not None else 0
        raw_date = row[idx_date] if idx_date is not None else None
        digits = re.sub(r"\D", "", str(raw_date)) if raw_date is not None else ""
        month_key = f"{digits[:4]}-{digits[4:6]}" if len(digits) >= 6 else None
        if month_key in monthly_totals:
            monthly_totals[month_key] += float(cost or 0)
    return monthly_totals


def get_cloud_cost_by_resource_type_in_group(subscription_id: str, resource_type: str,
                                              resource_group_name: str, months: list) -> dict:
    """
    Soma o custo de um TIPO específico de recurso (ex.: 'microsoft.compute/restorepointcollections')
    dentro de um Resource Group, direto na Azure. Mais preciso que somar o RG inteiro quando o RG
    tem outros tipos de recurso misturados — isola só o que interessa (ex.: só os restore points,
    ignorando qualquer outra coisa que exista no mesmo RG).
    """
    from datetime import datetime, timedelta
    try:
        from azure.identity import DefaultAzureCredential
        from azure.mgmt.costmanagement import CostManagementClient
        from azure.mgmt.costmanagement.models import (
            QueryDefinition, QueryDataset, QueryAggregation, QueryFilter,
            QueryComparisonExpression, QueryTimePeriod,
        )
    except ImportError:
        raise RuntimeError("Biblioteca não instalada. Rode:\n  pip install azure-mgmt-costmanagement")

    credential = DefaultAzureCredential()
    client = CostManagementClient(credential)
    scope = f"/subscriptions/{subscription_id}"

    start_dt = datetime.strptime(months[0] + "-01", "%Y-%m-%d")
    y, m = map(int, months[-1].split("-"))
    end_dt = (datetime(y + 1, 1, 1) if m == 12 else datetime(y, m + 1, 1)) - timedelta(days=1)

    definition = QueryDefinition(
        type="ActualCost",
        timeframe="Custom",
        time_period=QueryTimePeriod(from_property=start_dt, to=end_dt),
        dataset=QueryDataset(
            granularity="Monthly",
            aggregation={"totalCost": QueryAggregation(name="Cost", function="Sum")},
            filter=QueryFilter(
                and_property=[
                    QueryFilter(dimensions=QueryComparisonExpression(
                        name="ResourceType", operator="In", values_property=[resource_type.lower()])),
                    QueryFilter(dimensions=QueryComparisonExpression(
                        name="ResourceGroupName", operator="In", values_property=[resource_group_name.lower()])),
                ]
            ),
        ),
    )

    result = client.query.usage(scope, definition)
    columns = [c.name for c in result.columns]
    idx_cost = _find_column_index(columns, "cost")
    idx_date = _find_column_index(columns, "usagedate", "billingmonth", "date")

    monthly_totals = {m: 0.0 for m in months}
    for row in result.rows or []:
        cost = row[idx_cost] if idx_cost is not None else 0
        raw_date = row[idx_date] if idx_date is not None else None
        digits = re.sub(r"\D", "", str(raw_date)) if raw_date is not None else ""
        month_key = f"{digits[:4]}-{digits[4:6]}" if len(digits) >= 6 else None
        if month_key in monthly_totals:
            monthly_totals[month_key] += float(cost or 0)
    return monthly_totals


# ---------------------------------------------------------------------------
# DESCOBERTA DE RECURSOS RELACIONADOS À VM (via Azure API — dados reais, não CSV)
# ---------------------------------------------------------------------------

def get_vm_related_resource_ids(subscription_id: str, resource_group: str, vm_name: str) -> dict:
    """
    Retorna dict {"VM": [...], "Disco": [...], "Rede": [...]} com os ResourceIds
    reais anexados à VM, consultando a Azure API (não aproximação por nome).
    """
    try:
        from azure.identity import DefaultAzureCredential
        from azure.mgmt.compute import ComputeManagementClient
        from azure.mgmt.network import NetworkManagementClient
    except ImportError:
        raise RuntimeError(
            "Bibliotecas Azure não instaladas. Rode:\n"
            "  pip install azure-identity azure-mgmt-compute azure-mgmt-network"
        )

    credential = DefaultAzureCredential()
    compute_client = ComputeManagementClient(credential, subscription_id)
    network_client = NetworkManagementClient(credential, subscription_id)

    vm = compute_client.virtual_machines.get(resource_group, vm_name)

    related = {"VM": [vm.id], "Disco": [], "Rede": []}

    # Discos (OS + data disks)
    if vm.storage_profile.os_disk and vm.storage_profile.os_disk.managed_disk:
        related["Disco"].append(vm.storage_profile.os_disk.managed_disk.id)
    for data_disk in vm.storage_profile.data_disks or []:
        if data_disk.managed_disk:
            related["Disco"].append(data_disk.managed_disk.id)

    # NICs + IP público associado a cada NIC
    for nic_ref in vm.network_profile.network_interfaces or []:
        nic_id = nic_ref.id
        related["Rede"].append(nic_id)
        try:
            nic_rg, nic_name = nic_id.split("/resourceGroups/")[1].split("/providers/")[0], nic_id.split("/")[-1]
            nic = network_client.network_interfaces.get(nic_rg, nic_name)
            for ip_config in nic.ip_configurations or []:
                if ip_config.public_ip_address:
                    related["Rede"].append(ip_config.public_ip_address.id)
        except Exception as e:
            logging.getLogger(LOGGER_NAME).warning("Não foi possível resolver IP público da NIC %s: %s", nic_id, e)
            print(f"  [aviso] não foi possível resolver IP público de uma NIC — ver log para detalhes")

    return related


def find_approx_backup_log_security(df: pd.DataFrame, vm_name: str, resource_group: str) -> pd.DataFrame:
    """
    Aproximação: linhas cujo ServiceName sugira Backup/Log/Segurança e cujo RG ou Tags
    mencionem a VM. SEMPRE marcado como aproximado — não é vínculo exato de billing.
    """
    service_col_candidates = [c for c in df.columns if c.lower() == "servicename"]
    if not service_col_candidates:
        return pd.DataFrame()
    service_col = service_col_candidates[0]

    service_mask = df[service_col].str.lower().str.contains(
        "|".join(APPROX_SERVICE_KEYWORDS), na=False
    )
    scope_mask = (
        (df[COL_RESOURCE_GROUP].astype(str).str.lower() == str(resource_group).lower())
        | (df[COL_TAGS].astype(str).str.lower().str.contains(vm_name.lower(), na=False))
    )
    return df.loc[service_mask & scope_mask].copy()


# ---------------------------------------------------------------------------
# FALLBACK EM NUVEM — usado quando o nome não bate com nada na base local
# ---------------------------------------------------------------------------

def build_subscription_map(df: pd.DataFrame) -> dict:
    """Extrai {SubscriptionName: subscription_id} a partir dos ResourceIds já presentes na base."""
    result = {}
    for _, row in df[[COL_SUBSCRIPTION_NAME, COL_RESOURCE_ID]].dropna().drop_duplicates(subset=[COL_SUBSCRIPTION_NAME]).iterrows():
        m = RESOURCE_ID_SUB_REGEX.search(row[COL_RESOURCE_ID])
        if m:
            result[row[COL_SUBSCRIPTION_NAME]] = m.group(1)
    return result


def list_all_accessible_subscriptions() -> dict:
    """
    Retorna {subscription_id: display_name} de TODAS as subscriptions que a sessão
    logada (az login) enxerga — não só as 5 que aparecem na base local. Isso é
    necessário porque recursos de ASR (Azure Site Recovery) replica, DR, ou
    subscriptions não cobertas pela extração mensal podem estar fora da base local.
    """
    from azure.identity import DefaultAzureCredential
    from azure.mgmt.subscription import SubscriptionClient

    credential = DefaultAzureCredential()
    client = SubscriptionClient(credential)
    return {s.subscription_id: s.display_name for s in client.subscriptions.list()}


def build_full_subscription_index(df: pd.DataFrame) -> dict:
    """
    Combina as subscriptions conhecidas da base local (nomes amigáveis tipo KOF-...-BRA)
    com TODAS as subscriptions acessíveis na Azure. Retorna {subscription_id: nome}.
    Se a listagem via Azure falhar (permissão, sem lib instalada, etc.), cai para
    usar só as subscriptions da base local — com aviso no console.
    """
    id_to_name = {}
    for name, sid in build_subscription_map(df).items():
        id_to_name[sid] = name

    try:
        cloud_subs = call_with_retries(list_all_accessible_subscriptions, "listagem de subscriptions acessíveis")
        for sid, display_name in cloud_subs.items():
            id_to_name.setdefault(sid, display_name)
        print(f"     {len(cloud_subs)} subscriptions acessíveis na Azure (fallback vai buscar em todas)")
    except Exception as e:
        logging.getLogger(LOGGER_NAME).error("Falha ao listar subscriptions acessíveis na Azure: %s\n%s", e, traceback.format_exc())
        print("     [aviso] não foi possível listar todas as subscriptions da Azure — ver log para detalhes")
        print("     fallback em nuvem vai usar só as subscriptions já conhecidas na base local")

    return id_to_name


def cloud_search_resource(subscription_ids: list, name: str) -> list:
    """Busca o recurso pelo nome via Azure Resource Graph, em todas as subscriptions informadas."""
    try:
        from azure.identity import DefaultAzureCredential
        from azure.mgmt.resourcegraph import ResourceGraphClient
        from azure.mgmt.resourcegraph.models import QueryRequest
    except ImportError:
        raise RuntimeError(
            "Biblioteca não instalada. Rode:\n  pip install azure-mgmt-resourcegraph"
        )

    credential = DefaultAzureCredential()
    client = ResourceGraphClient(credential)
    # nome exato primeiro (case-insensitive); Resource Graph usa KQL
    query = f"Resources | where name =~ '{name}'"
    request = QueryRequest(subscriptions=subscription_ids, query=query)
    response = client.resources(request)
    return response.data or []


def get_cloud_cost_per_resource(subscription_id: str, resource_ids: list, months: list) -> dict:
    """
    Consulta a Azure Cost Management API diretamente (não usa o CSV local) e retorna
    {resource_id: {mes: custo}} para os meses informados (formato 'YYYY-MM').

    AVISO: os nomes de coluna retornados pela API variam um pouco entre versões do SDK
    (ex.: 'Cost' vs 'PreTaxCost', 'UsageDate' vs 'BillingMonth'). O parsing abaixo é
    feito por nome (não por posição fixa) pra ser resiliente a isso, mas recomendo
    validar com --input contendo 1 recurso só antes de rodar a lista inteira.
    """
    from datetime import datetime, timedelta
    try:
        from azure.identity import DefaultAzureCredential
        from azure.mgmt.costmanagement import CostManagementClient
        from azure.mgmt.costmanagement.models import (
            QueryDefinition, QueryDataset, QueryAggregation, QueryGrouping,
            QueryFilter, QueryComparisonExpression, QueryTimePeriod,
        )
    except ImportError:
        raise RuntimeError(
            "Biblioteca não instalada. Rode:\n  pip install azure-mgmt-costmanagement"
        )

    if not resource_ids:
        return {}

    credential = DefaultAzureCredential()
    client = CostManagementClient(credential)
    scope = f"/subscriptions/{subscription_id}"

    start_dt = datetime.strptime(months[0] + "-01", "%Y-%m-%d")
    y, m = map(int, months[-1].split("-"))
    end_dt = (datetime(y + 1, 1, 1) if m == 12 else datetime(y, m + 1, 1)) - timedelta(days=1)

    # O Resource Graph devolve o ResourceId com a capitalização "original" do ARM
    # (ex.: Microsoft.Compute), mas o billing normalmente grava tudo em minúsculo.
    # Filtrar em minúsculo e casar de volta ignorando case evita "achou o recurso,
    # mas o custo veio zerado" por mismatch de string.
    lower_to_original = {rid.lower(): rid for rid in resource_ids}

    definition = QueryDefinition(
        type="ActualCost",
        timeframe="Custom",
        time_period=QueryTimePeriod(from_property=start_dt, to=end_dt),
        dataset=QueryDataset(
            granularity="Monthly",
            aggregation={"totalCost": QueryAggregation(name="Cost", function="Sum")},
            grouping=[QueryGrouping(type="Dimension", name="ResourceId")],
            filter=QueryFilter(
                dimensions=QueryComparisonExpression(
                    name="ResourceId", operator="In",
                    values_property=list(lower_to_original.keys())
                )
            ),
        ),
    )

    result = client.query.usage(scope, definition)
    columns = [c.name for c in result.columns]
    rows = result.rows or []

    def col_idx(*fragments):
        for i, c in enumerate(columns):
            if any(frag.lower() in c.lower() for frag in fragments):
                return i
        return None

    idx_cost = col_idx("cost")
    idx_date = col_idx("usagedate", "billingmonth", "date")
    idx_resid = col_idx("resourceid")

    out = {rid: {m: 0.0 for m in months} for rid in resource_ids}
    for row in rows:
        raw_rid = row[idx_resid] if idx_resid is not None else None
        rid = lower_to_original.get(str(raw_rid).lower()) if raw_rid is not None else None
        cost = row[idx_cost] if idx_cost is not None else 0
        raw_date = row[idx_date] if idx_date is not None else None
        digits = re.sub(r"\D", "", str(raw_date)) if raw_date is not None else ""
        month_key = f"{digits[:4]}-{digits[4:6]}" if len(digits) >= 6 else None
        if rid in out and month_key in out[rid]:
            out[rid][month_key] += float(cost or 0)

    return out


# ---------------------------------------------------------------------------
# AGREGAÇÃO DE CUSTO POR MÊS
# ---------------------------------------------------------------------------

def aggregate_by_month(df: pd.DataFrame, resource_ids: list, months: list) -> dict:
    """
    Soma custo por mês pra uma lista de ResourceIds. Comparação SEMPRE case-insensitive:
    o mesmo recurso pode aparecer com capitalização diferente entre o que a API da Azure
    devolve (ex.: 'Microsoft.Compute/virtualMachines') e o que está gravado na base local
    (muitas vezes tudo minúsculo) — comparar exato faz o recurso "sumir" silenciosamente
    (zero de custo mesmo o recurso existindo e tendo gasto real).
    """
    if not resource_ids:
        return {m: 0 for m in months}
    ids_lower = {rid.lower() for rid in resource_ids}
    subset = df[df[COL_RESOURCE_ID].str.lower().isin(ids_lower)]
    result = {}
    for m in months:
        result[m] = subset.loc[subset[COL_MONTH] == m, COL_COST].sum()
    return result


def build_summary_tables(out_df: pd.DataFrame, months: list) -> tuple:
    """
    Agrega os custos totais (coluna '{mes}_Total' de cada mês) por Tipo e por Subscription.
    Retorna (resumo_por_tipo, resumo_por_subscription), cada um já ordenado do maior pro
    menor custo total no período. Não usa as colunas de detalhe (_VM/_Disco/_Rede/_Backup)
    para não contar custo em dobro — só a coluna '_Total' de cada mês.
    """
    total_cols = [f"{m}_Total" for m in months if f"{m}_Total" in out_df.columns]
    if not total_cols or out_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df_calc = out_df.copy()
    if "Status" in df_calc.columns:
        df_calc = df_calc[df_calc["Status"] != STATUS_ERRO]  # linhas de erro não têm custo a somar
    if "_subitem" in df_calc.columns:
        df_calc = df_calc[df_calc["_subitem"] != True]  # subitens (VM_/Disco_/Rede_/Outros_) já estão na linha Total_
    if df_calc.empty:
        return pd.DataFrame(), pd.DataFrame()

    for c in total_cols:
        df_calc[c] = pd.to_numeric(df_calc[c], errors="coerce").fillna(0)

    def summarize(group_col: str) -> pd.DataFrame:
        if group_col not in df_calc.columns:
            return pd.DataFrame()
        grouped = df_calc.groupby(df_calc[group_col].fillna("(não informado)"))
        g = grouped[total_cols].sum().reset_index()
        counts = grouped.size().reset_index(name="Qtd_Recursos")
        g = g.merge(counts, on=group_col)
        g["Total_Periodo"] = g[total_cols].sum(axis=1)
        g = g[[group_col, "Qtd_Recursos"] + total_cols + ["Total_Periodo"]]
        return g.sort_values("Total_Periodo", ascending=False).reset_index(drop=True)

    return summarize("Tipo"), summarize("Subscription")


def append_total_row(df: pd.DataFrame, label_col: str, label: str = "TOTAL",
                      exclude_status: str = None) -> pd.DataFrame:
    """
    Acrescenta uma linha de total no final do DataFrame, somando todas as colunas
    numéricas (colunas de mês/custo). Colunas de texto (Tipo, Subscription, Observações
    etc.) ficam em branco na linha de total, exceto label_col que recebe 'TOTAL'.
    Se exclude_status for informado, linhas com esse valor em 'Status' não entram na soma
    (ex.: linhas de ERRO não têm custo real pra somar).
    """
    if df.empty:
        return df

    calc_df = df
    if exclude_status is not None and "Status" in df.columns:
        calc_df = calc_df[calc_df["Status"] != exclude_status]
    if "_subitem" in calc_df.columns:
        calc_df = calc_df[calc_df["_subitem"] != True]  # subitens já estão contados na linha Total_

    total_row = {col: "" for col in df.columns}
    total_row[label_col] = label
    for col in df.columns:
        if col == label_col:
            continue
        numeric = pd.to_numeric(calc_df[col], errors="coerce")
        if numeric.notna().any():
            total_row[col] = round(numeric.sum(), 2)

    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def write_output_excel(results: list, months: list, output_path: str) -> None:
    """
    Grava o Excel com 2 abas:
      - "Resumo" (primeira aba): título com o período + bloco "RESUMO POR TIPO" +
        bloco "RESUMO POR SUBSCRIPTION", cada um com linha TOTAL.
      - "Detalhe": título com o período + tabela recurso a recurso (cabeçalhos já com
        o mês/ano real, ex.: '2026-04_Total') + linha TOTAL (só linhas com Status=OK).
    """
    out_df = pd.DataFrame(results)
    by_tipo, by_sub = build_summary_tables(out_df, months)

    if not by_tipo.empty:
        by_tipo = append_total_row(by_tipo, "Tipo")
    if not by_sub.empty:
        by_sub = append_total_row(by_sub, "Subscription")
    if not out_df.empty:
        out_df = append_total_row(out_df, "Recurso", exclude_status=STATUS_ERRO)
    if "_subitem" in out_df.columns:
        out_df = out_df.drop(columns=["_subitem"])  # marcador interno — não aparece na planilha

    period_label = f"Período analisado: {months[0]} a {months[-1]}   |   Gerado em {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    # --- aba "Resumo" (por Tipo + por Subscription) ---
    resumo_label_positions = []
    resumo_total_rows = []
    next_row = 2  # linha 0 = título, linha 1 = em branco

    tipo_header_row = sub_header_row = None
    if not by_tipo.empty:
        resumo_label_positions.append((next_row, "RESUMO POR TIPO"))
        next_row += 1
        tipo_header_row = next_row
        next_row += 1 + len(by_tipo)
        resumo_total_rows.append(next_row - 1)
        next_row += 1

    if not by_sub.empty:
        resumo_label_positions.append((next_row, "RESUMO POR SUBSCRIPTION"))
        next_row += 1
        sub_header_row = next_row
        next_row += 1 + len(by_sub)
        resumo_total_rows.append(next_row - 1)
        next_row += 1

    # --- aba "Detalhe" (recurso a recurso) ---
    detalhe_label_positions = [(2, "DETALHE POR RECURSO")]
    main_header_row = 3
    detalhe_total_rows = []
    if not out_df.empty:
        detalhe_total_rows.append(main_header_row + len(out_df))

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if tipo_header_row is not None:
            by_tipo.to_excel(writer, index=False, startrow=tipo_header_row, sheet_name="Resumo")
        if sub_header_row is not None:
            by_sub.to_excel(writer, index=False, startrow=sub_header_row, sheet_name="Resumo")
        out_df.to_excel(writer, index=False, startrow=main_header_row, sheet_name="Detalhe")

    wb = load_workbook(output_path)
    wb._sheets = [wb["Resumo"], wb["Detalhe"]]  # garante "Resumo" como primeira aba
    wb.active = 0

    def style_sheet(ws, n_cols, label_positions, header_rows, total_rows):
        last_col_letter = get_column_letter(max(n_cols, 1))
        ws.merge_cells(f"A1:{last_col_letter}1")
        ws["A1"] = period_label
        ws["A1"].font = Font(bold=True, size=12)
        ws["A1"].alignment = Alignment(horizontal="left")

        for row_0idx, text in label_positions:
            cell = ws.cell(row=row_0idx + 1, column=1, value=text)
            cell.font = Font(bold=True, size=11, italic=True)

        for header_row_0idx in header_rows:
            if header_row_0idx is None:
                continue
            for cell in ws[header_row_0idx + 1]:
                cell.font = Font(bold=True)

        for total_row_0idx in total_rows:
            for cell in ws[total_row_0idx + 1]:
                cell.font = Font(bold=True)

    n_cols_resumo = max(len(by_tipo.columns) if not by_tipo.empty else 0,
                         len(by_sub.columns) if not by_sub.empty else 0)
    style_sheet(wb["Resumo"], n_cols_resumo, resumo_label_positions,
                [tipo_header_row, sub_header_row], resumo_total_rows)
    style_sheet(wb["Detalhe"], len(out_df.columns), detalhe_label_positions,
                [main_header_row], detalhe_total_rows)

    wb.save(output_path)


# ---------------------------------------------------------------------------
# PROCESSAMENTO DE UM RECURSO DA LISTA
# ---------------------------------------------------------------------------

def aggregate_vm_full_cost_local(df: pd.DataFrame, vm_name: str, resource_group: str,
                                  subscription_name: str, months: list) -> tuple:
    """
    Soma o custo de TODO recurso, na mesma Subscription + Resource Group, cujo nome
    contenha o nome da VM como substring — cobre disco, NIC, IP público, backup/restore
    point, snapshot, etc. Não depende do estado atual de anexação via API (que só reflete
    o "agora" — se um disco foi trocado ou algo foi renomeado ao longo dos últimos meses,
    a API não traria o custo histórico dele). Retorna (totais_por_mes, breakdown_por_categoria,
    qtd_de_recursos_distintos).
    """
    name_lower = vm_name.lower()
    mask = (
        (df[COL_RESOURCE_GROUP].astype(str).str.lower() == resource_group.lower())
        & (df[COL_SUBSCRIPTION_NAME] == subscription_name)
        & (df[COL_RESOURCE].astype(str).str.lower().str.contains(re.escape(name_lower), na=False))
    )
    subset = df.loc[mask].copy()
    totals = {m: subset.loc[subset[COL_MONTH] == m, COL_COST].sum() for m in months}

    def categoria(resource_type: str) -> str:
        if is_vm(resource_type):
            return "VM"
        t = (resource_type or "").lower().replace(" ", "")
        if "disks" in t or "disk" == t:
            return "Disco"
        if any(k in t for k in ("networkinterfaces", "publicipaddresses", "virtualnetworks", "loadbalancers", "networksecuritygroups")):
            return "Rede"
        return "Outros"

    subset["_categoria"] = subset[COL_RESOURCE_TYPE].apply(categoria)
    breakdown = {}
    for cat in ["VM", "Disco", "Rede", "Outros"]:
        cat_subset = subset[subset["_categoria"] == cat]
        breakdown[cat] = {m: cat_subset.loc[cat_subset[COL_MONTH] == m, COL_COST].sum() for m in months}

    n_resources = subset[COL_RESOURCE_ID].nunique()
    return totals, breakdown, n_resources


def process_one_candidate(df: pd.DataFrame, display_name: str, info: dict, months: list, use_api: bool) -> list:
    """
    Processa um candidato já resolvido (1 ResourceId específico) e monta a(s) linha(s) de saída.
    Para VM: retorna 5 linhas com o MESMO 'Recurso' (nome puro) e a coluna 'Item' variando —
    'VM', 'Disco', 'Rede', 'Outros' (marcadas internamente como subitem, pra não entrar em
    dobro nos resumos/TOTAL geral) e 'Total' (a que conta pros resumos). Para recurso
    isolado, retorna 1 linha só. Se a consulta à Azure falhar após MAX_RETRIES tentativas,
    propaga a exceção — quem chama esta função decide como registrar o erro (linha mínima
    de status ERRO).
    """
    base = {"Status": STATUS_OK, "Tipo": info["resource_type"],
            "Subscription": info["subscription_name"], "ResourceGroup": info["resource_group"]}

    if is_vm(info["resource_type"]):
        vm_name = info["resource_id"].split("/")[-1]
        totals, breakdown, n_resources = aggregate_vm_full_cost_local(
            df, vm_name, info["resource_group"], info["subscription_name"], months
        )

        rows = []
        for categoria in ["VM", "Disco", "Rede", "Outros"]:
            sub_row = {"Recurso": display_name, "Item": categoria, **base}
            sub_row["_subitem"] = True  # não entra nos resumos/TOTAL geral — já está contado na linha Item=Total
            for m in months:
                sub_row[f"{m}_Total"] = round(breakdown[categoria][m], 2)
            rows.append(sub_row)

        total_row = {"Recurso": display_name, "Item": "Total", **base}
        total_row["Qtd_Recursos"] = n_resources
        total_row["Observacoes"] = (
            f"Custo agregado por nome dentro do Resource Group — {n_resources} recursos "
            f"cujo nome contém '{vm_name}' (VM + disco + rede + backup/outros associados)"
        )
        for m in months:
            total_row[f"{m}_Total"] = round(totals[m], 2)
        rows.append(total_row)
        return rows

    else:
        # Recurso isolado (disco, storage, etc.) — só o próprio ResourceId, 1 linha só
        totals = aggregate_by_month(df, [info["resource_id"]], months)
        row = {"Recurso": display_name, "Item": "Total", **base}
        for m in months:
            row[f"{m}_Total"] = round(totals[m], 2)
        return [row]


def minimal_error_row(display_name: str) -> dict:
    """Linha mínima pra recurso com erro — só nome + status, sem dado parcial (por decisão explícita)."""
    return {"Recurso": display_name, "Status": STATUS_ERRO,
            "Observacoes": f"Falha ao consultar a Azure após {MAX_RETRIES} tentativas — ver log para detalhes"}


def minimal_not_found_row(display_name: str) -> dict:
    return {"Recurso": display_name, "Status": STATUS_NAO_ENCONTRADO,
            "Observacoes": "Não encontrado na base local nem na Azure"}


def process_resource(df: pd.DataFrame, name: str, months: list, use_api: bool,
                      subscription_index: dict = None, cloud_fallback: bool = True,
                      type_group_cache: dict = None) -> list:
    """
    Retorna uma LISTA de linhas de saída. Normalmente 1 linha; se o nome casar com mais
    de um recurso real (ex.: VM com mesmo nome em RGs diferentes — original e restore de
    backup), retorna 1 linha por recurso, cada uma identificada pelo Resource Group.

    Ordem de busca pra cada nome:
      1. Recurso individual na base local — match EXATO
      2. Resource Group inteiro na base local (soma TODOS os recursos daquele RG)
      3. Recurso individual na base local — match PARCIAL (último recurso local; o
         ResourceId sempre contém o nome do RG como substring, por isso o RG precisa
         ser checado ANTES do parcial, senão nomes de RG nunca chegariam no passo 2)
      4. Recurso individual na Azure (Resource Graph + Cost Management), se cloud_fallback
      5. Resource Group inteiro na Azure, se os passos 1-4 não acharam nada

    Toda consulta à Azure tenta até MAX_RETRIES vezes. Se falhar mesmo assim, a linha
    vem só com Status=ERRO (sem dado parcial). Recursos não encontrados em lugar nenhum
    voltam com Status=NAO_ENCONTRADO — quem chama filtra essas linhas antes de gravar
    a planilha final.
    """
    if type_group_cache is None:
        type_group_cache = {}

    # 1. recurso individual na base local — match exato
    info = find_resource_exact(df, name)

    # 2. Resource Group inteiro na base local — só entra aqui se o match exato não achou nada
    if not info["found"]:
        rg_candidates_local = find_resource_group_local(df, name)
        if rg_candidates_local:
            rows = []
            for cand in rg_candidates_local:
                display_name = name if len(rg_candidates_local) == 1 else f"{name} [{cand['subscription_name']}]"
                totals, n_resources = aggregate_resource_group_local(df, cand["resource_group"], cand["subscription_name"], months)
                row = {"Recurso": display_name, "Status": STATUS_OK, "Tipo": "Resource Group",
                       "Subscription": cand["subscription_name"], "ResourceGroup": cand["resource_group"],
                       "Qtd_Recursos": n_resources,
                       "Observacoes": f"Reconhecido como Resource Group — custo agregado de todos os {n_resources} recursos dentro dele"}
                for m in months:
                    row[f"{m}_Total"] = round(totals[m], 2)
                rows.append(row)
            return rows

    # 3. recurso individual na base local — match parcial (só se 1 e 2 não acharam nada)
    if not info["found"]:
        info = find_resource_partial(df, name)

    if info["found"]:
        candidates = info["candidates"]
        rows = []
        for cand in candidates:
            display_name = name if len(candidates) == 1 else f"{name} [{cand['resource_group']}]"
            try:
                cand_rows = process_one_candidate(df, display_name, cand, months, use_api)
                if len(candidates) > 1:
                    for r in cand_rows:
                        if not r.get("_subitem"):  # nota de ambiguidade só na linha que conta (Total_ ou única)
                            r["Observacoes"] = (r.get("Observacoes", "") +
                                                 f" | {len(candidates)} recursos com esse nome encontrados — este é o de {cand['resource_group']}").strip(" |")
            except Exception:
                cand_rows = [minimal_error_row(display_name)]
            rows.extend(cand_rows)
        return rows

    # não achou como recurso individual nem como RG na base local
    if not cloud_fallback or not subscription_index:
        return [minimal_not_found_row(name)]

    # 4. recurso individual na Azure
    print(f"     '{name}' não está na base local — consultando Azure diretamente ({len(subscription_index)} subscriptions)...")
    try:
        cloud_matches = call_with_retries(
            cloud_search_resource, f"Resource Graph para '{name}'",
            list(subscription_index.keys()), name,
        )
    except Exception:
        return [minimal_error_row(name)]

    if cloud_matches:
        rows = []
        for match in cloud_matches:
            resource_id = match.get("id")
            resource_type = (match.get("type") or "")
            resource_group = match.get("resourceGroup")
            subscription_id = match.get("subscriptionId")
            display_name = name if len(cloud_matches) == 1 else f"{name} [{resource_group}]"

            vm_ids, disk_ids, net_ids = [resource_id], [], []
            if is_vm(resource_type):
                try:
                    related = call_with_retries(
                        get_vm_related_resource_ids, f"disco/rede da VM {display_name}",
                        subscription_id, resource_group, resource_id.split("/")[-1],
                    )
                    vm_ids, disk_ids, net_ids = related["VM"], related["Disco"], related["Rede"]
                except Exception:
                    rows.append(minimal_error_row(display_name))
                    continue

            all_ids = vm_ids + disk_ids + net_ids

            row = {"Recurso": display_name, "Status": STATUS_OK, "Tipo": resource_type,
                   "Subscription": subscription_index.get(subscription_id, subscription_id),
                   "ResourceGroup": resource_group,
                   "Observacoes": "Não estava na base local — custo consultado direto na Azure (Cost Management API)"}

            try:
                per_resource = call_with_retries(
                    get_cloud_cost_per_resource, f"custo (Cost Management) de {display_name}",
                    subscription_id, all_ids, months,
                )
                for m in months:
                    row[f"{m}_Total"] = round(sum(per_resource.get(rid, {}).get(m, 0) for rid in all_ids), 2)
                    if is_vm(resource_type):
                        row[f"{m}_VM"] = round(sum(per_resource.get(rid, {}).get(m, 0) for rid in vm_ids), 2)
                        row[f"{m}_Disco"] = round(sum(per_resource.get(rid, {}).get(m, 0) for rid in disk_ids), 2)
                        row[f"{m}_Rede"] = round(sum(per_resource.get(rid, {}).get(m, 0) for rid in net_ids), 2)

                # Custo zerado em recurso do tipo "restorePointCollections" costuma significar
                # que o billing está no Resource Group/vault, não no ResourceId da collection.
                # Busca o custo agregado de TODOS os recursos desse TIPO dentro do mesmo RG
                # (mais preciso que o RG inteiro, isola só os restore points). Como pode haver
                # várias collections no mesmo RG (uma por VM), só a PRIMEIRA linha desse grupo
                # (mesma subscription+RG+tipo) recebe o valor — as demais apontam pra ela, pra
                # não contar o mesmo custo várias vezes na coluna TOTAL.
                total_zerado = all(row.get(f"{m}_Total", 0) == 0 for m in months)
                if total_zerado and "restorepointcollections" in resource_type.lower():
                    cache_key = (subscription_id, resource_group.lower(), resource_type.lower())
                    if cache_key in type_group_cache:
                        row["Observacoes"] = (
                            f"Custo não aparece no ResourceId individual — valor consolidado já "
                            f"calculado na linha '{type_group_cache[cache_key]}' (mesmo Resource "
                            f"Group + tipo, pra não duplicar o total)"
                        )
                    else:
                        try:
                            type_totals = call_with_retries(
                                get_cloud_cost_by_resource_type_in_group,
                                f"custo de restore points no RG {resource_group}",
                                subscription_id, resource_type, resource_group, months,
                            )
                            for m in months:
                                row[f"{m}_Total"] = round(type_totals[m], 2)
                            row["Observacoes"] = (
                                f"Custo não aparece no ResourceId individual — valor agregado de "
                                f"TODOS os recursos do tipo '{resource_type}' no Resource Group "
                                f"'{resource_group}'. Outras linhas do mesmo RG/tipo apontam pra "
                                f"esta em vez de repetir o valor."
                            )
                            type_group_cache[cache_key] = display_name
                        except Exception:
                            row["Observacoes"] += " | Custo veio zerado e o fallback por tipo de recurso também falhou"
            except Exception:
                rows.append(minimal_error_row(display_name))
                continue

            rows.append(row)
        return rows

    # 5. Resource Group inteiro na Azure — só chega aqui se não achou recurso individual em lugar nenhum
    try:
        rg_cloud_matches = call_with_retries(
            cloud_search_resource_group, f"Resource Graph (Resource Group) para '{name}'",
            list(subscription_index.keys()), name,
        )
    except Exception:
        return [minimal_error_row(name)]

    if not rg_cloud_matches:
        return [minimal_not_found_row(name)]

    rows = []
    for match in rg_cloud_matches:
        resource_group = match.get("name")
        subscription_id = match.get("subscriptionId")
        display_name = name if len(rg_cloud_matches) == 1 else f"{name} [{resource_group}]"

        row = {"Recurso": display_name, "Status": STATUS_OK, "Tipo": "Resource Group",
               "Subscription": subscription_index.get(subscription_id, subscription_id),
               "ResourceGroup": resource_group,
               "Observacoes": "Reconhecido como Resource Group na Azure — custo agregado (Cost Management API)"}

        try:
            totals = call_with_retries(
                get_cloud_cost_for_resource_group, f"custo do Resource Group {display_name}",
                subscription_id, resource_group, months,
            )
            for m in months:
                row[f"{m}_Total"] = round(totals[m], 2)
        except Exception:
            rows.append(minimal_error_row(display_name))
            continue

        rows.append(row)

    return rows


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Coletor de custo por recurso (VM completa ou isolado)")
    parser.add_argument("--input", required=True, help="Excel com coluna 'Recurso'")
    parser.add_argument("--output", required=True, help="Excel de saída")
    parser.add_argument("--base", default=BASE_CSV_PATH, help="Caminho do AzureBase_FinOps.csv")
    parser.add_argument("--no-api", action="store_true",
                         help="(mantido por compatibilidade — não tem mais efeito: o custo local de VM já não depende de chamada à Azure API)")
    parser.add_argument("--no-cloud-fallback", action="store_true",
                         help="Não consulta a Azure quando o recurso não é achado na base local")
    parser.add_argument("--force-reprocess", action="store_true",
                         help="Ignora o cache de execuções anteriores e reprocessa a lista inteira do zero")
    args = parser.parse_args()

    log_path = setup_logging()
    print(f"[log] detalhes técnicos desta execução vão para: {log_path}")
    logger = logging.getLogger(LOGGER_NAME)
    logger.info("Execução iniciada. script_version=%s input=%s output=%s base=%s no_api=%s no_cloud_fallback=%s",
                SCRIPT_VERSION, args.input, args.output, args.base, args.no_api, args.no_cloud_fallback)

    print(f"[1/5] Carregando base de custos: {args.base}")
    try:
        df = load_cost_base(args.base)
    except Exception as e:
        logger.error("Falha ao carregar a base de custos: %s\n%s", e, traceback.format_exc())
        print(f"Erro ao carregar a base de custos — ver log para detalhes: {log_path}")
        sys.exit(1)

    months = last_n_months(df, N_MONTHS)
    print(f"[2/5] Meses considerados: {months}")

    subscription_index = build_full_subscription_index(df)

    print(f"[3/5] Lendo lista de recursos: {args.input}")
    if args.input.lower().endswith(".txt"):
        # .txt: um nome de recurso por linha, sem cabeçalho
        with open(args.input, "r", encoding="utf-8-sig") as f:
            resource_names = [line.strip() for line in f if line.strip()]
    else:
        # .xlsx: coluna 'Recurso'
        input_df = pd.read_excel(args.input)
        resource_names = input_df["Recurso"].dropna().astype(str).tolist()

    # --- cache de execuções anteriores: pula quem já deu certo, reprocessa só erro/novo ---
    state_path = state_path_for(args.output)
    state = {} if args.force_reprocess else load_state(state_path)

    to_process = []
    reused_rows = []
    n_reused_ok = n_reused_not_found = 0
    n_stale_version = 0
    for name in resource_names:
        cached = state.get(name)
        if cached and cached.get("months") == months and cached.get("script_version") != SCRIPT_VERSION:
            n_stale_version += 1
            to_process.append(name)
            continue
        if cached and cached.get("months") == months and cached.get("status") in ("ok", "not_found"):
            if cached["status"] == "ok":
                reused_rows.extend(cached.get("rows", []))
                n_reused_ok += 1
            else:
                n_reused_not_found += 1  # settled como não encontrado — não entra na saída
            continue
        to_process.append(name)

    print(f"[4/5] {len(resource_names)} recursos na lista | "
          f"{n_reused_ok} já OK (reaproveitados do cache) | "
          f"{n_reused_not_found} já confirmados como não encontrados (cache) | "
          f"{n_stale_version} com cache de uma versão antiga do script (reprocessando) | "
          f"{len(to_process)} a processar agora")

    use_api = not args.no_api
    cloud_fallback = not args.no_cloud_fallback
    new_state_entries = {}
    type_group_cache = {}  # dedup de custo por tipo+RG (ex.: restore points), só nesta execução
    for name in to_process:
        print(f"  -> processando: {name}")
        rows = process_resource(df, name, months, use_api, subscription_index, cloud_fallback, type_group_cache)

        statuses = {r.get("Status") for r in rows}
        if statuses == {STATUS_NAO_ENCONTRADO}:
            cache_status = "not_found"
            cache_rows = []
        elif STATUS_ERRO in statuses:
            cache_status = "error"
            cache_rows = [r for r in rows if r.get("Status") != STATUS_NAO_ENCONTRADO]
        else:
            cache_status = "ok"
            cache_rows = rows

        new_state_entries[name] = {
            "status": cache_status, "months": months, "rows": cache_rows,
            "script_version": SCRIPT_VERSION,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        }

    # reconstrói o cache só com os nomes que estão na lista atual (não guarda lixo de listas antigas)
    final_state = {}
    for name in resource_names:
        final_state[name] = new_state_entries.get(name) or state.get(name)
    save_state(state_path, final_state)

    # linhas finais = reaproveitadas do cache (status ok) + recém-processadas (ok ou erro) —
    # nunca inclui NAO_ENCONTRADO na planilha final, por decisão explícita
    results = list(reused_rows)
    for name in to_process:
        entry = new_state_entries[name]
        if entry["status"] in ("ok", "error"):
            results.extend(entry["rows"])

    print(f"[5/5] Gravando saída: {args.output}")
    write_output_excel(results, months, args.output)

    n_erro = sum(1 for r in results if r.get("Status") == STATUS_ERRO)
    n_ok = sum(1 for r in results if r.get("Status") == STATUS_OK)
    n_nao_encontrado_total = n_reused_not_found + sum(
        1 for name in to_process if new_state_entries[name]["status"] == "not_found"
    )
    logger.info("Execução concluída. %d linhas na saída (%d OK, %d ERRO), %d recursos não encontrados (excluídos da saída).",
                len(results), n_ok, n_erro, n_nao_encontrado_total)
    print(f"Concluído. {n_ok} OK, {n_erro} com erro, {n_nao_encontrado_total} não encontrados (removidos da saída).")
    if n_erro:
        print(f"(detalhes dos erros em {log_path} — rode de novo mais tarde: só os {n_erro} com erro serão reprocessados)")


if __name__ == "__main__":
    main()
