"""
================================================================================
SISTEMA DE PESQUISA DE PREÇOS DE MERCADO — PREGÃO SRP (LEI 14.133/2021)
APLICATIVO DESKTOP UNIFICADO EM UM ÚNICO ARQUIVO (PySide6 / Qt)
================================================================================
Arquivo único, 100% autossuficiente e independente.
Contém:
- Motor de coleta automatizada e busca comercial com Playwright
- IA de auditoria de preços e embalagens com Gemini 3.1 Flash Lite
- Validação matemática e dupla busca externa de CNPJ da Receita Federal
- Carimbo oficial monocromático de alta legibilidade para processos públicos
- Gestão completa de planilhas Excel/CSV com mapeamento dinâmico
- Geração de modelos e exportação de pacotes completos (Excel + Prints + Índice)
- Interface desktop moderna, corporativa e responsiva em PySide6
================================================================================
"""

import os
import sys
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except:
    pass
import re
import json
import time
import datetime
import collections
import urllib.parse
import io
import shutil
import zipfile
import glob
import subprocess
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from playwright.sync_api import sync_playwright
from ddgs import DDGS
from google import genai
import openpyxl

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QTabWidget, QProgressBar, QComboBox, QSpinBox,
    QRadioButton, QButtonGroup, QFileDialog, QMessageBox, QSplitter,
    QScrollArea, QFrame, QGroupBox, QAbstractItemView
)
from PySide6.QtCore import Qt, QThread, Signal, Slot, QTimer, QSize, QUrl
from PySide6.QtGui import (
    QFont, QIcon, QPixmap, QColor, QDesktopServices, QPalette,
    QBrush, QLinearGradient, QPainter, QImage
)

# ==============================================================================
# CONFIGURAÇÕES E CONSTANTES GLOBAIS DO SISTEMA
# Tenta carregar variáveis de ambiente de arquivo .env local, se existir
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.exists(_env_path):
    try:
        with open(_env_path, "r", encoding="utf-8") as _f:
            for _line in _f:
                _line = _line.strip()
                if _line and not _line.startswith("#") and "=" in _line:
                    _k, _v = _line.split("=", 1)
                    os.environ[_k.strip()] = _v.strip().strip('"').strip("'")
    except Exception:
        pass

API_KEY = os.environ.get("GEMINI_API_KEY", "")
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1ghfazQM5zPPVqQCPT9g8g67NF9-raEV6a0-qcxUjUfs/export?format=csv"
RESULTS_JSON = "cotacoes_resultado.json"
RESULTS_EXCEL = "cotacoes_pesquisa_mercado.xlsx"
PRINTS_DIR = "prints_cotacoes"
QUOTA_FILE = "gemini_quota.json"
STATUS_RUN_FILE = "cotacao_status.json"
LIVE_STATUS_FILE = "cotador_live_status.json"

ACTIVE_SHEET_JSON = "planilha_itens_ativa.json"
ACTIVE_SHEET_EXCEL = "planilha_itens_ativa.xlsx"
ACTIVE_SHEET_CONFIG = "planilha_itens_config.json"

# Dicionário oficial de CNPJs de grandes redes e distribuidores de material elétrico e construção
KNOWN_STORE_CNPJS = {
    "lojadomecanico.com.br": "05.473.473/0001-99",
    "dutramaquinas.com.br": "01.218.428/0001-44",
    "santil.com.br": "60.404.090/0001-63",
    "lojaeletrica.com.br": "17.155.342/0001-21",
    "eletrofrigor.com.br": "12.984.774/0001-60",
    "cematel.com.br": "08.825.986/0001-42",
    "telhanorte.com.br": "45.717.375/0001-87",
    "leroymerlin.com.br": "01.438.784/0001-05",
    "cec.com.br": "01.438.784/0001-05",
    "ferreiracosta.com.br": "10.428.163/0001-38",
    "andorinhaeletrica.com.br": "52.880.828/0001-08",
    "fibrasnetsp.com": "35.729.386/0001-51",
    "unicaserv.com.br": "07.709.614/0001-96",
    "tupan.com.br": "00.279.531/0008-31",
    "lucamateriais.com.br": "36.671.152/0004-06",
    "comelrg.com.br": "04.123.471/0001-48",
    "solucaocabos.com.br": "00.559.915/0001-31",
    "dpr.com.br": "00.422.413/0004-07",
    "dlight.com.br": "10.691.425/0001-48",
    "sveletrica.com": "35.088.657/0001-37",
    "wurth.com.br": "43.648.971/0001-71",
    "eletroferragens.com.br": "28.539.923/0001-60",
    "eletroeste.com.br": "76.495.272/0001-20",
    "avsupply.com.br": "42.129.648/0001-02",
    "eletrosul.com.br": "00.073.957/0001-68",
    "marconloja.com.br": "19.566.306/0001-38",
    "msartefatos.com.br": "18.732.780/0002-00",
    "cassol.com.br": "83.080.051/0001-42",
    "balaroti.com.br": "75.976.249/0001-26",
    "sodimac.com.br": "08.995.127/0001-10",
    "e-letrica.com.br": "13.921.282/0001-08",
    "ilumishop.com.br": "21.033.456/0001-80",
    "lojaferbag.com.br": "05.163.670/0001-30",
    "antferramentas.com.br": "08.852.122/0001-37",
    "fg.com.br": "84.475.268/0001-83",
    "palaciodasferramentas.com.br": "65.313.344/0001-52",
    "anhangueraferramentas.com.br": "00.081.711/0001-70",
    "jnakao.com.br": "04.591.312/0001-28",
    "martinsatacado.com.br": "25.750.366/0001-55",
    "nacionalepi.com.br": "13.689.967/0001-60",
    "lojaneobetel.com.br": "26.337.892/0001-49"
}

# Lista rigorosa de Marketplaces proibidos
BLACKLIST_DOMAINS = [
    "mercadolivre.com", "mercadolivre.com.br",
    "shopee.com", "shopee.com.br",
    "amazon.com", "amazon.com.br",
    "aliexpress.com", "aliexpress.com.br",
    "enjoei.com.br", "olx.com.br",
    "madeiramadeira.com.br",
    "magazineluiza.com.br", "magazinevoce.com.br",
    "americanas.com.br", "submarino.com.br", "shoptime.com.br",
    "casasbahia.com.br", "pontofrio.com.br", "extra.com.br",
    "carrefour.com.br", "elo7.com.br", "shein.com",
    "buscape.com.br", "zoom.com.br", "bondfaro.com.br",
    "facebook.com", "instagram.com", "youtube.com", "tiktok.com",
    "pinterest.com", "twitter.com", "x.com", "reclameaqui.com.br",
    "jusbrasil.com.br", "gov.br", "wikipedia.org"
]

class QuotaExhaustedException(Exception):
    """Exceção levantada quando a cota diária de 500 requisições do Gemini é atingida."""
    pass

class GeminiQuotaManager:
    def __init__(self, rpd_max=500, rpm_max=12, tpm_max=220000, state_file=QUOTA_FILE):
        self.rpd_max = rpd_max
        self.rpm_max = rpm_max
        self.tpm_max = tpm_max
        self.state_file = state_file
        self.call_history = collections.deque()
        self.rpd_used = 0
        self.load_state()

    def load_state(self):
        today = time.strftime('%Y-%m-%d')
        if os.path.exists(self.state_file):
            try:
                with open(self.state_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if data.get('date') == today:
                    self.rpd_used = data.get('rpd_used', 0)
                    return
            except:
                pass
        self.rpd_used = 0

    def save_state(self):
        today = time.strftime('%Y-%m-%d')
        try:
            with open(self.state_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'date': today,
                    'rpd_used': self.rpd_used,
                    'rpd_max': self.rpd_max,
                    'rpd_remaining': max(0, self.rpd_max - self.rpd_used)
                }, f, indent=2)
        except:
            pass

    def acquire(self, estimated_tokens=1500):
        if self.rpd_used >= self.rpd_max:
            raise QuotaExhaustedException("Limite diário de requisições do Gemini atingido (500 RPD).")

        now = time.time()
        while self.call_history and (now - self.call_history[0][0]) > 60:
            self.call_history.popleft()

        if len(self.call_history) >= self.rpm_max:
            wait_s = max(1.0, 60.0 - (now - self.call_history[0][0]) + 0.5)
            time.sleep(wait_s)
            return self.acquire(estimated_tokens)

        current_tpm = sum(t for _, t in self.call_history)
        if current_tpm + estimated_tokens > self.tpm_max:
            wait_s = max(1.0, 60.0 - (now - self.call_history[0][0]) + 0.5)
            time.sleep(wait_s)
            return self.acquire(estimated_tokens)

        if self.call_history:
            elapsed = now - self.call_history[-1][0]
            if elapsed < 4.5:
                time.sleep(4.5 - elapsed)

        now = time.time()
        self.call_history.append((now, estimated_tokens))
        self.rpd_used += 1
        self.save_state()

quota_manager = GeminiQuotaManager()

def get_gemini_client():
    key = os.environ.get("GEMINI_API_KEY", "") or API_KEY
    if not key:
        raise ValueError("Chave de API Gemini não configurada. Defina a variável de ambiente GEMINI_API_KEY ou configure o arquivo .env.")
    return genai.Client(api_key=key)

# ==============================================================================
# VALIDAÇÃO DE CNPJ E PREÇOS
# ==============================================================================
def is_valid_cnpj(cnpj_str):
    nums = re.sub(r'\D', '', str(cnpj_str))
    if len(nums) != 14 or len(set(nums)) == 1:
        return False
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    d1 = sum(int(nums[i]) * w1[i] for i in range(12)) % 11
    d1 = 0 if d1 < 2 else 11 - d1
    if int(nums[12]) != d1:
        return False
    w2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    d2 = sum(int(nums[i]) * w2[i] for i in range(13)) % 11
    d2 = 0 if d2 < 2 else 11 - d2
    return int(nums[13]) == d2

def format_cnpj(raw_cnpj):
    nums = re.sub(r'\D', '', str(raw_cnpj))
    if len(nums) == 14:
        return f"{nums[0:2]}.{nums[2:5]}.{nums[5:8]}/{nums[8:12]}-{nums[12:14]}"
    return str(raw_cnpj)

def extract_valid_cnpjs_from_text(text):
    found = []
    matches = re.findall(r'\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b', text)
    for m in matches:
        if is_valid_cnpj(m) and m not in found:
            found.append(m)
    matches_raw = re.findall(r'CNPJ[\s\:\.\-]*(\d{14})\b', text, re.IGNORECASE)
    for m in matches_raw:
        fmt = format_cnpj(m)
        if is_valid_cnpj(fmt) and fmt not in found:
            found.append(fmt)
    return found

def find_cnpj_with_dual_search(domain, store_title, current_page_text, page, context):
    clean_d = domain.replace("www.", "").lower().strip()
    for k_dom, k_cnpj in KNOWN_STORE_CNPJS.items():
        if k_dom in clean_d or clean_d in k_dom:
            return k_cnpj, "Base Oficial de Redes Elétricas"

    cnpjs = extract_valid_cnpjs_from_text(current_page_text)
    if cnpjs:
        KNOWN_STORE_CNPJS[clean_d] = cnpjs[0]
        return cnpjs[0], "Rodapé / Página do Produto"

    try:
        html_content = page.content()
        cnpjs_html = extract_valid_cnpjs_from_text(html_content)
        if cnpjs_html:
            KNOWN_STORE_CNPJS[clean_d] = cnpjs_html[0]
            return cnpjs_html[0], "Código-fonte / Metadados da Loja"
    except:
        pass

    inst_links = ["quem-somos", "sobre", "institucional", "fale-conosco"]
    for link_kw in inst_links[:2]:
        try:
            el = page.query_selector(f"a[href*='{link_kw}']")
            if el:
                href = el.get_attribute("href")
                if href:
                    if href.startswith("/"):
                        href = f"https://{domain}{href}"
                    p_inst = context.new_page()
                    p_inst.goto(href, timeout=5000, wait_until="domcontentloaded")
                    p_inst.wait_for_timeout(300)
                    inst_text = p_inst.inner_text("body")
                    inst_cnpjs = extract_valid_cnpjs_from_text(inst_text)
                    if not inst_cnpjs:
                        inst_cnpjs = extract_valid_cnpjs_from_text(p_inst.content())
                    p_inst.close()
                    if inst_cnpjs:
                        KNOWN_STORE_CNPJS[clean_d] = inst_cnpjs[0]
                        return inst_cnpjs[0], f"Página Institucional ({link_kw})"
                    break
        except:
            pass

    domain_slug = clean_d.split(".")[0].strip()
    clean_name = store_title.split("-")[0].split("|")[0].split(":")[0].strip()
    external_queries = [
        f"CNPJ {domain_slug}",
        f"CNPJ {clean_d}",
        f"CNPJ {clean_name}"
    ]

    for ext_q in external_queries:
        if not ext_q.strip():
            continue
        try:
            with DDGS() as ddgs:
                res = list(ddgs.text(ext_q, max_results=5))
                for r in res:
                    text_snip = f"{r.get('title', '')} {r.get('body', '')}"
                    external_cnpjs = extract_valid_cnpjs_from_text(text_snip)
                    if external_cnpjs:
                        KNOWN_STORE_CNPJS[clean_d] = external_cnpjs[0]
                        return external_cnpjs[0], f"Dupla Pesquisa Externa ({ext_q})"
        except:
            continue

    return None, "Não Encontrado"

def parse_brazilian_currency(txt):
    if not txt:
        return 0.0
    clean = re.sub(r'[^\d\,\.]', '', str(txt)).strip()
    if not clean:
        return 0.0
    try:
        if ',' in clean:
            parts = clean.split(',')
            inteiro = parts[0].replace('.', '')
            decimal = parts[1][:2]
            return float(f"{inteiro}.{decimal}")
        elif '.' in clean:
            parts = clean.split('.')
            if len(parts[-1]) == 2:
                inteiro = ''.join(parts[:-1])
                decimal = parts[-1]
                return float(f"{inteiro}.{decimal}")
            else:
                return float(clean.replace('.', ''))
        else:
            return float(clean)
    except:
        return 0.0

def extract_structured_prices(page):
    try:
        prices = page.evaluate("""() => {
            let list = [];
            let meta = document.querySelector('meta[property="product:price:amount"], meta[itemprop="price"]');
            if (meta) {
                let v = meta.getAttribute('content') || meta.getAttribute('value');
                if (v && parseFloat(v) > 0) {
                    list.push({origem: 'meta_produto', valor: 'R$ ' + parseFloat(v).toFixed(2).replace('.', ',')});
                }
            }
            for (let el of document.querySelectorAll('[itemprop="price"]')) {
                let v = el.getAttribute('content') || el.innerText || '';
                if (v && v.trim()) {
                    list.push({origem: 'itemprop_price', valor: v.trim()});
                }
            }
            let els = document.querySelectorAll('[class*="spotPrice"], [class*="sellingPrice"], [class*="bestPrice"], [class*="price-current"], [class*="preco-principal"], [class*="price"]');
            for (let el of els) {
                if (el.closest('header, footer, nav, [class*="news"], [class*="banner"], [class*="carousel"], [class*="recommendation"]')) {
                    continue;
                }
                let t = (el.innerText || '').trim();
                if (t.includes('R$') && t.length < 35 && !t.toLowerCase().includes('parcela') && !t.toLowerCase().includes('sem juros') && !t.toLowerCase().includes('frete')) {
                    list.push({origem: 'dom_produto', valor: t});
                    if (list.length >= 8) break;
                }
            }
            return list;
        }""")
        return prices
    except:
        return []

# ==============================================================================
# GESTÃO DE PLANILHA DE ITENS
# ==============================================================================
def normalize_col_name(name):
    import unicodedata
    s = str(name).strip().lower()
    s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8')
    s = re.sub(r'[^a-z0-9]', '', s)
    return s

def detect_column_mapping(df):
    mapping = {
        "lote": None,
        "descricao": None,
        "unidade": None,
        "catmat": None,
        "quantidade": None,
        "grupo": None
    }
    cols = list(df.columns)
    cols_norm = [normalize_col_name(c) for c in cols]

    desc_keywords = [
        "descricaodoitem", "descricaocompleta", "descricaomaterial",
        "descricao", "especificacaodoitem", "especificacao", "especificacoes",
        "material", "produto", "objeto", "itemdescricao", "discriminacao",
        "detalhamento", "desc"
    ]
    for kw in desc_keywords:
        for idx, norm in enumerate(cols_norm):
            if norm == kw or norm.startswith(kw) or kw in norm:
                mapping["descricao"] = cols[idx]
                break
        if mapping["descricao"]:
            break

    if not mapping["descricao"]:
        best_col = None
        max_len = 0
        for col in cols:
            try:
                avg_len = df[col].astype(str).str.len().mean()
                if avg_len > max_len and avg_len > 12:
                    max_len = avg_len
                    best_col = col
            except:
                pass
        if best_col:
            mapping["descricao"] = best_col
        elif len(cols) >= 4:
            mapping["descricao"] = cols[3]
        elif len(cols) >= 1:
            mapping["descricao"] = cols[0]

    lote_keywords = ["lote", "item", "noitem", "nolote", "numlote", "numitem", "numero", "ordem", "seq", "pos", "posicao"]
    for kw in lote_keywords:
        for idx, norm in enumerate(cols_norm):
            if cols[idx] == mapping.get("descricao"):
                continue
            if norm == kw or norm.startswith(kw) or (kw in norm and len(norm) <= len(kw) + 4):
                mapping["lote"] = cols[idx]
                break
        if mapping["lote"]:
            break
    if not mapping["lote"] and len(cols) > 0 and cols[0] != mapping.get("descricao"):
        mapping["lote"] = cols[0]

    unid_keywords = ["unidade", "unid", "un", "und", "medida", "um", "unidademedida", "unidmedida"]
    for kw in unid_keywords:
        for idx, norm in enumerate(cols_norm):
            if cols[idx] in [mapping.get("descricao"), mapping.get("lote")]:
                continue
            if norm == kw or norm.startswith(kw):
                mapping["unidade"] = cols[idx]
                break
        if mapping["unidade"]:
            break
    if not mapping["unidade"] and len(cols) >= 5 and cols[4] not in [mapping.get("descricao"), mapping.get("lote")]:
        mapping["unidade"] = cols[4]

    return mapping

def parse_items_dataframe(df, mapping=None):
    if mapping is None:
        mapping = detect_column_mapping(df)

    col_lote = mapping.get("lote")
    col_desc = mapping.get("descricao")
    col_unid = mapping.get("unidade")

    items = []
    total_rows = len(df)
    empty_desc_rows = 0

    for idx, row in df.iterrows():
        desc = ""
        if col_desc and col_desc in df.columns and pd.notna(row[col_desc]):
            desc = str(row[col_desc]).strip()

        if not desc or desc.lower() in ["nan", "none", "null", ""] or desc.lower() in ["descrição", "descricao", "especificação", "material"]:
            empty_desc_rows += 1
            continue

        lote_val = None
        if col_lote and col_lote in df.columns and pd.notna(row[col_lote]):
            val_str = str(row[col_lote]).strip()
            digits = re.findall(r'\d+', val_str)
            if digits:
                try:
                    lote_val = int(digits[0])
                except:
                    lote_val = None
        if lote_val is None:
            lote_val = len(items) + 1

        unid = "UNID"
        if col_unid and col_unid in df.columns and pd.notna(row[col_unid]):
            u_str = str(row[col_unid]).strip().upper()
            if u_str and u_str not in ["NAN", "NONE", "-", ""]:
                unid = u_str

        items.append({
            "index": len(items),
            "lote": lote_val,
            "catmat": "",
            "descricao_original": desc,
            "unidade": unid,
            "quantidade": 1,
            "grupo": ""
        })

    stats = {
        "total_rows": total_rows,
        "valid_items": len(items),
        "empty_rows": empty_desc_rows,
        "mapping_used": mapping
    }
    return items, stats

def read_sheet_file(file_path):
    fname = file_path.lower()
    if fname.endswith(".csv"):
        with open(file_path, "rb") as f:
            bytes_data = f.read()
        for enc in ["utf-8-sig", "utf-8", "latin1", "cp1252"]:
            for sep in [None, ";", ",", "\t"]:
                try:
                    if sep is None:
                        df = pd.read_csv(io.BytesIO(bytes_data), encoding=enc, sep=None, engine="python")
                    else:
                        df = pd.read_csv(io.BytesIO(bytes_data), encoding=enc, sep=sep)
                    if len(df.columns) >= 1 and len(df) > 0:
                        return df
                except:
                    continue
        return pd.read_csv(io.BytesIO(bytes_data), encoding="latin1", sep=";")
    else:
        return pd.read_excel(file_path)

def save_active_sheet(items, source_name="planilha_carregada.xlsx"):
    with open(ACTIVE_SHEET_JSON, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)

    df_clean = pd.DataFrame([{
        "LOTE": it.get("lote"),
        "DESCRIÇÃO": it.get("descricao_original")
    } for it in items])
    df_clean.to_excel(ACTIVE_SHEET_EXCEL, index=False)

    config = {
        "source_name": source_name,
        "is_custom": True,
        "total_items": len(items),
        "updated_at": time.strftime("%d/%m/%Y %H:%M:%S")
    }
    with open(ACTIVE_SHEET_CONFIG, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    return True

def get_active_sheet_info():
    if os.path.exists(ACTIVE_SHEET_CONFIG) and os.path.exists(ACTIVE_SHEET_JSON):
        try:
            with open(ACTIVE_SHEET_CONFIG, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return {
        "source_name": "Nenhuma planilha carregada",
        "is_custom": False,
        "total_items": 0,
        "updated_at": "--"
    }

def reset_active_sheet_to_default():
    for fpath in [ACTIVE_SHEET_JSON, ACTIVE_SHEET_EXCEL, ACTIVE_SHEET_CONFIG]:
        if os.path.exists(fpath):
            try:
                os.remove(fpath)
            except:
                pass
    return True

def load_sheet_items():
    if os.path.exists(ACTIVE_SHEET_JSON):
        try:
            with open(ACTIVE_SHEET_JSON, "r", encoding="utf-8") as f:
                items = json.load(f)
            if items and len(items) > 0:
                return items
        except:
            pass
    return []

def generate_template_excel(output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Itens para Cotacao"
    ws.views.sheetView[0].showGridLines = True
    headers = ["LOTE", "DESCRIÇÃO"]
    ws.append(headers)

    examples = [
        [1, "ABRAÇADEIRA BAP 1000MM EM AÇO GALVANIZADO A QUENTE COM PARAFUSO E PORCA"],
        [2, "CABO DE COBRE ISOLADO 1KV 16MM² CONDUTOR DE COBRE ANTICHAMA 750V/1KV"],
        [3, "DISJUNTOR TERMOMAGNÉTICO BIPOLAR DIN 32A CURVA C 5KA 220V/380V"],
        [4, "LÂMPADA LED TUBULAR T8 18W 120CM BIVOLT 6500K BRANCO FRIO G13"],
        [5, "FITA ISOLANTE PROFISSIONAL 19MM X 20M ANTICHAMA CLASSE A 750V"],
        [6, "ALICATE UNIVERSAL 8 POL ISOLADO 1000V EM AÇO CROMO VANÁDIO"]
    ]
    for row in examples:
        ws.append(row)

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 85
    wb.save(output_path)
    return output_path

def generate_template_csv(output_path):
    csv_content = (
        "LOTE;DESCRIÇÃO\n"
        "1;ABRAÇADEIRA BAP 1000MM EM AÇO GALVANIZADO A QUENTE COM PARAFUSO E PORCA\n"
        "2;CABO DE COBRE ISOLADO 1KV 16MM² CONDUTOR DE COBRE ANTICHAMA 750V/1KV\n"
        "3;DISJUNTOR TERMOMAGNÉTICO BIPOLAR DIN 32A CURVA C 5KA 220V/380V\n"
        "4;LÂMPADA LED TUBULAR T8 18W 120CM BIVOLT 6500K BRANCO FRIO G13\n"
        "5;FITA ISOLANTE PROFISSIONAL 19MM X 20M ANTICHAMA CLASSE A 750V\n"
        "6;ALICATE UNIVERSAL 8 POL ISOLADO 1000V EM AÇO CROMO VANÁDIO\n"
    )
    with open(output_path, "w", encoding="utf-8-sig") as f:
        f.write(csv_content)
    return output_path

# ==============================================================================
# MOTOR DE BUSCA COMERCIAL E INTELIGENTE
# ==============================================================================
def intelligent_search_queries(desc):
    queries = []
    if " - " in desc:
        parts = desc.split(" - ")
        part_spec = parts[1].strip()
        spec_clean = re.sub(r'[\(\)\[\]\{\}\"\']', ' ', part_spec)
        spec_clean = re.sub(r'(?i)conformidade com.*?$', '', spec_clean)
        spec_clean = re.sub(r'(?i)resolução conama.*?$', '', spec_clean)
        spec_clean = re.sub(r'[^\w\s\d\.\,\-/]', ' ', spec_clean)
        spec_words = [w for w in spec_clean.split() if len(w) > 1]
        if len(spec_words) >= 2:
            real_product = " ".join(spec_words[:5])
            queries.append(f"{real_product} comprar")
            queries.append(f"{' '.join(spec_words[:3])} ferramentas comprar")

    clean = desc.split(";")[0]
    clean = re.sub(r'[\(\)\[\]\{\}\"\']', ' ', clean)
    clean = re.sub(r'(?i)marca de referência.*?$', '', clean)
    clean = re.sub(r'(?i)conformidade com.*?$', '', clean)
    clean = re.sub(r'(?i)resolução conama.*?$', '', clean)
    clean = re.sub(r'(?i)validade mínima.*?$', '', clean)
    clean = re.sub(r'(?i)embalagem original.*?$', '', clean)
    clean = re.sub(r'(?i)com no mínimo', '', clean)
    clean = re.sub(r'(?i)no mínimo', '', clean)
    clean = re.sub(r'[^\w\s\d\.\,\-/]', ' ', clean)
    words = [w for w in clean.split() if len(w) > 1]

    if words:
        clean_q = " ".join(words[:6])
        queries.append(f"{clean_q} comprar")
        queries.append(f"{' '.join(words[:3])} comprar eletrica")

    upper_desc = desc.upper()
    if "ADAPTADOR" in upper_desc and ("SAÍDA" in upper_desc or "SAIDA" in upper_desc or re.search(r'\bTE\b', upper_desc)):
        queries.insert(0, "adaptador 3 saidas benjamim comprar materiais eletricos")
    if "ESTRIBO" in upper_desc:
        queries.insert(0, "armacao 1 estribo galvanizada comprar materiais eletricos")
    if re.search(r'\bBAP\b', upper_desc):
        queries.insert(0, "abracadeira bap poste comprar eletrica")

    return list(dict.fromkeys(queries))

def intelligent_search_query(desc):
    queries = intelligent_search_queries(desc)
    return queries[0] if queries else desc[:30]

def search_candidate_stores(query_text, max_candidates=8):
    query_attempts = intelligent_search_queries(query_text)
    candidates = []
    seen_domains = set()

    for attempt_q in query_attempts:
        if len(candidates) >= 3:
            break
        try:
            with DDGS() as ddgs:
                raw_results = list(ddgs.text(attempt_q, region="br-pt", max_results=15))
                for r in raw_results:
                    href = r.get("href", "")
                    title = r.get("title", "")
                    domain = urllib.parse.urlparse(href).netloc.lower()
                    clean_dom = domain.replace("www.", "")

                    if any(b in domain for b in BLACKLIST_DOMAINS) or "bing.com" in domain or "google.com" in domain:
                        continue
                    if href.endswith(".pdf") or href.endswith(".doc") or href.endswith(".xml"):
                        continue
                    if clean_dom in seen_domains:
                        continue

                    seen_domains.add(clean_dom)
                    candidates.append({
                        "title": title,
                        "url": href,
                        "domain": domain,
                        "snippet": r.get("body", "")
                    })
                    if len(candidates) >= max_candidates:
                        break
        except:
            continue

    def store_priority(cand):
        dom = cand["domain"].replace("www.", "")
        return 0 if any(k in dom for k in KNOWN_STORE_CNPJS) else 1

    candidates.sort(key=store_priority)
    return candidates

def search_expanded_candidates(query_text, exclude_domains=None, max_candidates=12):
    if exclude_domains is None:
        exclude_domains = set()
    clean_q = intelligent_search_query(query_text)
    words = [w for w in clean_q.split() if len(w) > 2]
    core_kw = " ".join(words[:3]) if len(words) >= 3 else clean_q

    attempts = [
        f"{core_kw} distribuidora eletrica preco",
        f"{core_kw} loja virtual materiais construcao",
        f"{core_kw} comprar atacado varejo",
        f"catalogo {core_kw} comprar"
    ]

    candidates = []
    seen = set(exclude_domains)
    for q in attempts:
        if len(candidates) >= max_candidates:
            break
        try:
            with DDGS() as ddgs:
                results = list(ddgs.text(q, region="br-pt", max_results=30))
                for r in results:
                    href = r.get("href", "")
                    domain = urllib.parse.urlparse(href).netloc.lower()
                    clean_dom = domain.replace("www.", "")
                    if any(b in domain for b in BLACKLIST_DOMAINS) or "bing.com" in domain or "google.com" in domain:
                        continue
                    if href.endswith(".pdf") or href.endswith(".doc") or clean_dom in seen:
                        continue
                    seen.add(clean_dom)
                    candidates.append({
                        "title": r.get("title", ""),
                        "url": href,
                        "domain": domain,
                        "snippet": r.get("body", "")
                    })
                    if len(candidates) >= max_candidates:
                        break
        except:
            continue
    return candidates

# ==============================================================================
# AUDITORIA GEMINI E ESTAMPAGEM DE CARIMBO
# ==============================================================================
def analyze_with_gemini(client, item_info, store_title, structured_prices, page_text_snippet, domain):
    quota_manager.acquire(estimated_tokens=1200)

    prompt = f"""
Você é um auditor rigoroso de compras públicas e pesquisas de preços (Lei 14.133/2021).
Sua missão é extrair com PRECISÃO ABSOLUTA o valor comercial do material.

ITEM DA LICITAÇÃO:
- Descrição: {item_info.get('descricao_original')}
- Unidade solicitada: {item_info.get('unidade')}

PÁGINA DO FORNECEDOR ({domain}):
- Título do anúncio: {store_title}
- Preços estruturados detectados: {structured_prices}
- Resumo do texto da página:
\"\"\"{page_text_snippet[:2200]}\"\"\"

REGRAS OBRIGATÓRIAS:
1. Extraia o preço unitário do item em destaque ({store_title}).
2. Atenção a embalagens / lotes (se vier em pacote de N unidades, divida pelo total para obter o valor unitário).
3. No Brasil, R$ 177,00 = 177.00. NUNCA divida por 100 por engano.

Retorne EXCLUSIVAMENTE um JSON:
{{
  "produto_compativel": true,
  "nome_produto": "Nome do produto na loja",
  "fornecedor": "Nome da loja virtual",
  "texto_preco_original": "R$ 177,00",
  "valor_total_anuncio": 177.00,
  "quantidade_embalagem": 1,
  "valor_unitario": 177.00,
  "observacao": "Memória de cálculo do item"
}}
"""
    retries = 3
    for attempt in range(retries):
        try:
            resp = client.models.generate_content(
                model="gemini-3.1-flash-lite",
                contents=prompt
            )
            raw = resp.text.strip()
            if raw.startswith("```json"): raw = raw[7:]
            if raw.startswith("```"): raw = raw[3:]
            if raw.endswith("```"): raw = raw[:-3]
            data = json.loads(raw.strip())

            orig_txt = data.get("texto_preco_original", "")
            if orig_txt:
                parsed_val = parse_brazilian_currency(orig_txt)
                gemini_unit = float(data.get("valor_unitario", 0.0))
                qtd = int(data.get("quantidade_embalagem", 1))
                if qtd == 1 and parsed_val > 10.0 and abs((gemini_unit * 100) - parsed_val) < 1.0:
                    data["valor_unitario"] = parsed_val
                    data["valor_total_anuncio"] = parsed_val

            return data
        except Exception as e:
            if "503" in str(e) or "UNAVAILABLE" in str(e):
                time.sleep(4.0 * (attempt + 1))
                continue
            break

    if structured_prices:
        p_txt = structured_prices[0].get("valor", "")
        val = parse_brazilian_currency(p_txt)
        title_lower = store_title.lower()
        desc_words = [w.lower() for w in item_info.get("descricao_original", "").split() if len(w) > 3 and w.lower() not in ["para", "com", "sem", "material"]]
        is_compat = any(w in title_lower for w in desc_words) if desc_words else False

        if val > 0.1 and is_compat:
            return {
                "produto_compativel": True,
                "nome_produto": store_title,
                "fornecedor": domain,
                "texto_preco_original": p_txt,
                "valor_total_anuncio": val,
                "quantidade_embalagem": 1,
                "valor_unitario": val,
                "observacao": "Preço obtido diretamente da estrutura oficial da página"
            }

    return {"produto_compativel": False, "valor_unitario": 0.0}

def apply_print_friendly_stamp(image_path, item_info, quote_info):
    try:
        img = Image.open(image_path)
        w, h = img.size
        banner_h = 185
        new_img = Image.new("RGB", (w, h + banner_h), color=(255, 255, 255))
        draw = ImageDraw.Draw(new_img)

        try:
            font_title = ImageFont.truetype("arialbd.ttf", 17)
            font_bold = ImageFont.truetype("arialbd.ttf", 13)
            font_text = ImageFont.truetype("arial.ttf", 12)
            font_sub = ImageFont.truetype("arial.ttf", 11)
        except:
            font_title = font_bold = font_text = font_sub = ImageFont.load_default()

        now_str = quote_info.get("data_hora", datetime.datetime.now().strftime("%d/%m/%Y às %H:%M:%S"))
        lote = item_info.get("lote", 1)
        catmat = item_info.get("catmat", "")
        unid = item_info.get("unidade", "UNID")
        desc = item_info.get("descricao_original", "")

        fornecedor = quote_info.get("fornecedor", "Loja Comercial")
        cnpj = quote_info.get("cnpj", "CNPJ Não Identificado")
        cnpj_origem = quote_info.get("cnpj_origem", "Verificado")
        v_unit = quote_info.get("valor_unitario", 0.0)
        v_total = quote_info.get("valor_total_anuncio", v_unit)
        qtd_emb = quote_info.get("quantidade_embalagem", 1)
        obs = quote_info.get("observacao", "")
        url = quote_info.get("url", "")

        draw.rectangle([(8, 8), (w - 8, banner_h - 8)], outline=(0, 0, 0), width=2, fill=(255, 255, 255))
        draw.rectangle([(10, 10), (w - 10, 40)], fill=(245, 247, 250))
        draw.line([(8, 40), (w - 8, 40)], fill=(0, 0, 0), width=1)
        draw.text((22, 16), "COMPROVANTE DE PESQUISA DE PREÇOS DE MERCADO — PREGÃO SRP (LEI 14.133/2021)", fill=(0, 0, 0), font=font_title)

        v_unit_fmt = f"R$ {v_unit:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        txt_lote_cat = f"LOTE {lote:02d}  |  CATMAT: {catmat or 'N/A'}  |  UNIDADE: {unid}"
        draw.text((22, 48), txt_lote_cat, fill=(0, 0, 0), font=font_bold)
        draw.text((w - 520, 48), f"VALOR UNITÁRIO HOMOLOGADO: {v_unit_fmt}", fill=(0, 0, 0), font=font_title)

        desc_cut = desc if len(desc) <= 125 else desc[:122] + "..."
        draw.text((22, 74), f"ITEM (EDITAL): {desc_cut}", fill=(20, 20, 20), font=font_bold)

        fornec_cut = fornecedor if len(fornecedor) <= 35 else fornecedor[:32] + "..."
        draw.text((22, 100), f"FORNECEDOR: {fornec_cut}  |  CNPJ: {cnpj} [{cnpj_origem.upper()}]", fill=(0, 0, 0), font=font_bold)

        if qtd_emb > 1:
            v_tot_fmt = f"R$ {v_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            txt_calc = f"MEMÓRIA DE CÁLCULO: Anúncio com {qtd_emb} un ({v_tot_fmt}) -> {v_unit_fmt} / {unid}."
        else:
            txt_calc = f"DETALHE COMERCIAL: Preço unitário regular ({v_unit_fmt})."
        if obs:
            txt_calc += f" Obs: {obs}"
        if len(txt_calc) > 130:
            txt_calc = txt_calc[:127] + "..."
        draw.text((22, 126), txt_calc, fill=(50, 50, 50), font=font_text)

        url_cut = url if len(url) <= 120 else url[:117] + "..."
        draw.text((22, 150), f"DATA E HORA DA COLETA: {now_str}  |  FONTE: {url_cut}", fill=(70, 70, 70), font=font_sub)

        draw.line([(0, banner_h - 1), (w, banner_h - 1)], fill=(0, 0, 0), width=2)
        new_img.paste(img, (0, banner_h))
        new_img.save(image_path)
        return True
    except:
        return False

def dismiss_popups_and_clean_page(page):
    try:
        page.keyboard.press("Escape")
        page.evaluate("""() => {
            const cookieSelectors = [
                '[class*="adopt"]', '#adopt-controller', '[id*="adopt"]',
                '[class*="cookie"]', '[id*="cookie"]', '[class*="lgpd"]', '[id*="lgpd"]',
                '[class*="privacidade"]', '[class*="consent"]', '.onetrust-pc-dark-filter',
                '#onetrust-banner-sdk', '[id*="onetrust"]'
            ];
            cookieSelectors.forEach(sel => {
                document.querySelectorAll(sel).forEach(el => el.remove());
            });

            document.querySelectorAll('[class*="cep"], [id*="cep"]').forEach(el => {
                if (!el.querySelector('h1') && (el.innerText.includes('Escolha o CEP') || el.innerText.includes('Informe seu CEP') || el.innerText.includes('endereço') || el.className.includes('modal'))) {
                    el.remove();
                }
            });

            const modalSelectors = [
                '.modal-backdrop', '.backdrop', '[class*="backdrop"]', '[class*="overlay"]',
                '[role="dialog"]', '.modal.show', '.modal.fade.show', '[class*="cep-modal"]',
                '[class*="newsletter"]', '[class*="popup"]', '[id*="modal-cep"]', '[class*="modal-cep"]'
            ];
            modalSelectors.forEach(sel => {
                document.querySelectorAll(sel).forEach(el => {
                    if (!el.querySelector('h1') && !el.querySelector('[itemprop="price"]')) {
                        el.remove();
                    }
                });
            });

            document.querySelectorAll('[id*="jivo"], [id*="zsiq"], [class*="whatsapp"], iframe[title*="chat"]').forEach(el => el.remove());
            document.body.style.overflow = 'auto';
            document.documentElement.style.overflow = 'auto';
            document.body.classList.remove('modal-open');
        }""")
        page.wait_for_timeout(500)
    except:
        pass

def cotar_item_core(browser, gemini_client, item_info, status_callback=None, check_stop=None):
    os.makedirs(PRINTS_DIR, exist_ok=True)
    lote = item_info.get("lote")
    desc = item_info.get("descricao_original")

    if status_callback:
        status_callback({
            "current_lote": lote,
            "material": desc,
            "store": "Buscando lojas virtuais...",
            "step": "Identificando e-commerces especializados..."
        })

    candidates = search_candidate_stores(desc, max_candidates=16)
    if not candidates:
        return None

    context = browser.new_context(
        viewport={"width": 1366, "height": 850},
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        locale="pt-BR"
    )
    context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    page = context.new_page()

    best_quote = None

    for idx_cand, cand in enumerate(candidates, start=1):
        if check_stop and check_stop():
            break

        url = cand["url"]
        domain = cand["domain"]

        if status_callback:
            status_callback({
                "current_lote": lote,
                "material": desc,
                "store": domain,
                "step": f"Analisando fornecedor {idx_cand}/{len(candidates)}: {domain}..."
            })

        try:
            page.goto(url, wait_until="domcontentloaded", timeout=11000)
            page.wait_for_timeout(450)
            body_text = page.inner_text("body")

            if any(c in body_text.lower() for c in ["confirme seu acesso", "hcaptcha", "cf-turnstile", "are you a human", "recaptcha"]):
                continue
            if any(k in page.url.lower() for k in ["/identificacao", "/login", "/auth", "/checkout", "/signin"]):
                continue

            cnpj, cnpj_origem = find_cnpj_with_dual_search(domain, cand["title"], body_text, page, context)
            if not cnpj:
                continue

            structured_prices = extract_structured_prices(page)
            if not structured_prices:
                precos_regex = re.findall(r'R\$\s*[\d\.]+(?:,\d{2})', body_text)
                for pr in list(dict.fromkeys(precos_regex))[:4]:
                    structured_prices.append({"origem": "regex", "valor": pr})

            if not structured_prices:
                continue

            analysis = analyze_with_gemini(gemini_client, item_info, cand["title"], structured_prices, body_text, domain)
            v_unit = analysis.get("valor_unitario", 0.0)

            if not analysis.get("produto_compativel") or v_unit <= 0.05:
                continue

            safe_slug = re.sub(r'[^\w\-]', '_', desc[:28]).strip('_')
            now_str = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M:%S")
            print_filename = f"lote_{lote:03d}_{safe_slug}.png"
            print_path = os.path.join(PRINTS_DIR, print_filename)

            dismiss_popups_and_clean_page(page)
            page.screenshot(path=print_path, full_page=False)

            quote_data = {
                "lote": lote,
                "catmat": item_info.get("catmat", ""),
                "descricao_licitacao": desc,
                "unidade": item_info.get("unidade", "UNID"),
                "nome_produto_cotado": analysis.get("nome_produto"),
                "fornecedor": analysis.get("fornecedor", domain),
                "dominio": domain,
                "cnpj": cnpj,
                "cnpj_origem": cnpj_origem,
                "valor_unitario": float(v_unit),
                "valor_total_anuncio": float(analysis.get("valor_total_anuncio", v_unit)),
                "quantidade_embalagem": int(analysis.get("quantidade_embalagem", 1)),
                "observacao": analysis.get("observacao", ""),
                "url": url,
                "data_hora": now_str,
                "arquivo_print": print_path,
                "status": "Cotado e Homologado"
            }

            apply_print_friendly_stamp(print_path, item_info, quote_data)
            best_quote = quote_data
            break
        except:
            continue

    if not best_quote and not (check_stop and check_stop()):
        already_tested = {c["domain"].replace("www.", "") for c in candidates}
        round2 = search_expanded_candidates(desc, exclude_domains=already_tested, max_candidates=10)
        for cand in round2:
            if check_stop and check_stop():
                break
            url = cand["url"]
            domain = cand["domain"]
            if status_callback:
                status_callback({
                    "current_lote": lote,
                    "material": desc,
                    "store": domain,
                    "step": f"Rodada 2: Avaliando distribuidora {domain}..."
                })
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=12000)
                page.wait_for_timeout(450)
                body_text = page.inner_text("body")
                cnpj, cnpj_origem = find_cnpj_with_dual_search(domain, cand["title"], body_text, page, context)
                if not cnpj:
                    continue
                structured_prices = extract_structured_prices(page)
                if not structured_prices:
                    continue
                analysis = analyze_with_gemini(gemini_client, item_info, cand["title"], structured_prices, body_text, domain)
                v_unit = analysis.get("valor_unitario", 0.0)
                if not analysis.get("produto_compativel") or v_unit <= 0.05:
                    continue

                safe_slug = re.sub(r'[^\w\-]', '_', desc[:28]).strip('_')
                now_str = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M:%S")
                print_filename = f"lote_{lote:03d}_{safe_slug}.png"
                print_path = os.path.join(PRINTS_DIR, print_filename)

                dismiss_popups_and_clean_page(page)
                page.screenshot(path=print_path, full_page=False)

                quote_data = {
                    "lote": lote,
                    "catmat": item_info.get("catmat", ""),
                    "descricao_licitacao": desc,
                    "unidade": item_info.get("unidade", "UNID"),
                    "nome_produto_cotado": analysis.get("nome_produto"),
                    "fornecedor": analysis.get("fornecedor", domain),
                    "dominio": domain,
                    "cnpj": cnpj,
                    "cnpj_origem": cnpj_origem,
                    "valor_unitario": float(v_unit),
                    "valor_total_anuncio": float(analysis.get("valor_total_anuncio", v_unit)),
                    "quantidade_embalagem": int(analysis.get("quantidade_embalagem", 1)),
                    "observacao": analysis.get("observacao", ""),
                    "url": url,
                    "data_hora": now_str,
                    "arquivo_print": print_path,
                    "status": "Cotado e Homologado"
                }
                apply_print_friendly_stamp(print_path, item_info, quote_data)
                best_quote = quote_data
                break
            except:
                continue

    try:
        context.close()
    except:
        pass

    return best_quote

def save_and_export_results(all_results):
    def get_sort_key(item):
        val = item.get("lote", 0)
        try:
            return int(val)
        except:
            return 99999

    all_results.sort(key=get_sort_key)
    with open(RESULTS_JSON, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)

    rows = []
    for r in all_results:
        rows.append({
            "Lote": r.get("lote"),
            "CATMAT": r.get("catmat"),
            "Descrição do Item (Licitação)": r.get("descricao_licitacao"),
            "Unidade": r.get("unidade"),
            "Status": r.get("status"),
            "Valor Unitário Cotado (R$)": r.get("valor_unitario"),
            "Fornecedor / Loja": r.get("fornecedor"),
            "CNPJ": r.get("cnpj"),
            "Origem da Validação CNPJ": r.get("cnpj_origem"),
            "Produto Ofertado": r.get("nome_produto_cotado"),
            "Memória de Cálculo": r.get("observacao"),
            "Data e Hora da Pesquisa": r.get("data_hora"),
            "Link da Fonte": r.get("url"),
            "Arquivo Comprovante (Print)": r.get("arquivo_print")
        })

    df = pd.DataFrame(rows)
    df.to_excel(RESULTS_EXCEL, index=False)

def export_bundle_to_folder(dest_folder):
    os.makedirs(dest_folder, exist_ok=True)
    comprovantes_dir = os.path.join(dest_folder, "Comprovantes_Oficiais")
    os.makedirs(comprovantes_dir, exist_ok=True)

    exported_excel = os.path.join(dest_folder, "Pesquisa_de_Precos_Mercado.xlsx")
    if os.path.exists(RESULTS_EXCEL):
        shutil.copy2(RESULTS_EXCEL, exported_excel)

    copied_prints = 0
    if os.path.exists(PRINTS_DIR):
        for img_path in glob.glob(os.path.join(PRINTS_DIR, "*.*")):
            fname = os.path.basename(img_path)
            shutil.copy2(img_path, os.path.join(comprovantes_dir, fname))
            copied_prints += 1

    quotes = []
    if os.path.exists(RESULTS_JSON):
        try:
            with open(RESULTS_JSON, "r", encoding="utf-8") as f:
                quotes = json.load(f)
        except:
            pass

    index_file = os.path.join(dest_folder, "Indice_Pesquisa_Precos.txt")
    with open(index_file, "w", encoding="utf-8") as f:
        f.write("RELATÓRIO OFICIAL DE PESQUISA DE PREÇOS — PREGÃO SRP\n")
        f.write("=" * 70 + "\n")
        f.write(f"Data e Hora da Exportação: {datetime.datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}\n")
        f.write(f"Total de Comprovantes Anexados: {copied_prints}\n\n")
        f.write("RELAÇÃO DE ITENS E VALORES HOMOLOGADOS:\n")
        f.write("-" * 70 + "\n")
        for q in quotes:
            v_str = f"R$ {q.get('valor_unitario'):,.2f}" if q.get('valor_unitario') is not None else "Pendente"
            f.write(f"Lote {q.get('lote'):02d}: {q.get('descricao_licitacao', '')[:65]}\n")
            f.write(f"   Fornecedor: {q.get('fornecedor', '-')} | CNPJ: {q.get('cnpj', '-')} | Valor: {v_str}\n")
            if q.get('arquivo_print'):
                f.write(f"   Arquivo do Comprovante: Comprovantes_Oficiais/{os.path.basename(q.get('arquivo_print'))}\n")
            f.write("-" * 70 + "\n")

    return {
        "dest_folder": dest_folder,
        "excel_path": exported_excel,
        "comprovantes_dir": comprovantes_dir,
        "prints_count": copied_prints,
        "has_excel": os.path.exists(exported_excel)
    }

def create_export_zip_file(dest_zip_path):
    with zipfile.ZipFile(dest_zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        if os.path.exists(RESULTS_EXCEL):
            z.write(RESULTS_EXCEL, arcname="Pesquisa_de_Precos_Mercado.xlsx")
        if os.path.exists(PRINTS_DIR):
            for img_path in glob.glob(os.path.join(PRINTS_DIR, "*.*")):
                fname = os.path.basename(img_path)
                z.write(img_path, arcname=f"Comprovantes_Oficiais/{fname}")
        quotes = []
        if os.path.exists(RESULTS_JSON):
            try:
                with open(RESULTS_JSON, "r", encoding="utf-8") as f:
                    quotes = json.load(f)
            except:
                pass
        summary_txt = "RELATÓRIO OFICIAL DE PESQUISA DE PREÇOS — PREGÃO SRP\n"
        summary_txt += "=" * 70 + "\n"
        for q in quotes:
            v_str = f"R$ {q.get('valor_unitario'):,.2f}" if q.get('valor_unitario') is not None else "Pendente"
            summary_txt += f"Lote {q.get('lote')}: {q.get('descricao_licitacao', '')} - {q.get('fornecedor')} (CNPJ: {q.get('cnpj')}) - {v_str}\n"
        z.writestr("Resumo_Pesquisa_Precos.txt", summary_txt.encode("utf-8"))
    return dest_zip_path

def reset_all_data(backup=True):
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_folder = f"backup_cotacoes_{timestamp}"
    if backup:
        os.makedirs(backup_folder, exist_ok=True)
        if os.path.exists(RESULTS_JSON):
            try: shutil.copy(RESULTS_JSON, os.path.join(backup_folder, RESULTS_JSON))
            except: pass
        if os.path.exists(RESULTS_EXCEL):
            try: shutil.copy(RESULTS_EXCEL, os.path.join(backup_folder, RESULTS_EXCEL))
            except: pass
        if os.path.exists(PRINTS_DIR):
            try:
                prints_backup = os.path.join(backup_folder, PRINTS_DIR)
                shutil.copytree(PRINTS_DIR, prints_backup, dirs_exist_ok=True)
            except: pass

    with open(RESULTS_JSON, "w", encoding="utf-8") as f:
        json.dump([], f, indent=2)

    if os.path.exists(RESULTS_EXCEL):
        try: os.remove(RESULTS_EXCEL)
        except: pass

    if os.path.exists(PRINTS_DIR):
        for img_file in glob.glob(os.path.join(PRINTS_DIR, "*.*")):
            try: os.remove(img_file)
            except: pass

    reset_active_sheet_to_default()
    if os.path.exists(LIVE_STATUS_FILE):
        try:
            with open(LIVE_STATUS_FILE, "w", encoding="utf-8") as f:
                json.dump({}, f)
        except: pass
    if os.path.exists(STATUS_RUN_FILE):
        try: os.remove(STATUS_RUN_FILE)
        except: pass
    return True

# ==============================================================================
# WORKER THREAD PARA PROCESSAMENTO EM SEGUNDO PLANO NO QT
# ==============================================================================
class CotadorWorkerThread(QThread):
    progress_signal = Signal(dict)
    item_finished_signal = Signal(dict)
    batch_finished_signal = Signal(bool, str)

    def __init__(self, start_idx=0, count=1, parent=None):
        super().__init__(parent)
        self.start_idx = start_idx
        self.count = count
        self.stop_requested = False

    def request_stop(self):
        self.stop_requested = True

    def run(self):
        items = load_sheet_items()
        slice_items = items[self.start_idx:self.start_idx + self.count]
        if not slice_items:
            self.batch_finished_signal.emit(False, "Nenhum item selecionado para pesquisa.")
            return

        existing = []
        if os.path.exists(RESULTS_JSON):
            try:
                with open(RESULTS_JSON, "r", encoding="utf-8") as f:
                    existing = json.load(f)
            except:
                existing = []

        gemini_client = get_gemini_client()

        with sync_playwright() as p:
            try:
                browser = p.chromium.launch(
                    headless=True,
                    args=["--disable-blink-features=AutomationControlled"]
                )
            except Exception as e_launch:
                self.batch_finished_signal.emit(False, f"Erro ao iniciar Playwright/Chromium: {e_launch}")
                return

            total_target = len(slice_items)
            for idx, it in enumerate(slice_items, start=1):
                if self.stop_requested:
                    break

                lote_num = it.get("lote")
                desc_item = it.get("descricao_original")

                self.progress_signal.emit({
                    "current_lote": lote_num,
                    "material": desc_item,
                    "store": "Iniciando pesquisa de mercado...",
                    "step": f"Processando item {idx} de {total_target}...",
                    "progress_pct": int(((idx - 1) / max(1, total_target)) * 100)
                })

                try:
                    if not browser.is_connected():
                        browser = p.chromium.launch(
                            headless=True,
                            args=["--disable-blink-features=AutomationControlled"]
                        )

                    def status_cb(st_dict):
                        st_dict["progress_pct"] = int(((idx - 1) / max(1, total_target)) * 100)
                        self.progress_signal.emit(st_dict)

                    res = cotar_item_core(
                        browser,
                        gemini_client,
                        it,
                        status_callback=status_cb,
                        check_stop=lambda: self.stop_requested
                    )

                    if res:
                        existing = [e for e in existing if e.get("lote") != lote_num]
                        existing.append(res)
                        save_and_export_results(existing)
                        self.item_finished_signal.emit(res)
                    else:
                        existing = [e for e in existing if e.get("lote") != lote_num]
                        pending_data = {
                            "lote": lote_num,
                            "catmat": it.get("catmat", ""),
                            "descricao_licitacao": desc_item,
                            "unidade": it.get("unidade", "UNID"),
                            "nome_produto_cotado": "Pendente de Cotação",
                            "fornecedor": "Não Localizado",
                            "dominio": "",
                            "cnpj": "Pendente",
                            "cnpj_origem": "Pendente",
                            "valor_unitario": None,
                            "valor_total_anuncio": None,
                            "quantidade_embalagem": 1,
                            "observacao": "Pesquisa comercial realizada, mas nenhum e-commerce com CNPJ comprovado atendeu à especificação nesta rodada.",
                            "url": "",
                            "data_hora": datetime.datetime.now().strftime("%d/%m/%Y às %H:%M:%S"),
                            "arquivo_print": "",
                            "status": "Pendente de Cotação"
                        }
                        existing.append(pending_data)
                        save_and_export_results(existing)
                        self.item_finished_signal.emit(pending_data)

                except QuotaExhaustedException as q_err:
                    self.batch_finished_signal.emit(False, str(q_err))
                    try: browser.close()
                    except: pass
                    return

                except Exception as item_err:
                    existing = [e for e in existing if e.get("lote") != lote_num]
                    err_data = {
                        "lote": lote_num,
                        "catmat": it.get("catmat", ""),
                        "descricao_licitacao": desc_item,
                        "unidade": it.get("unidade", "UNID"),
                        "nome_produto_cotado": f"Erro: {str(item_err)[:45]}",
                        "fornecedor": "Falha na Coleta",
                        "dominio": "",
                        "cnpj": "Erro",
                        "cnpj_origem": "Erro",
                        "valor_unitario": None,
                        "valor_total_anuncio": None,
                        "quantidade_embalagem": 1,
                        "observacao": f"Erro no processamento: {item_err}",
                        "url": "",
                        "data_hora": datetime.datetime.now().strftime("%d/%m/%Y às %H:%M:%S"),
                        "arquivo_print": "",
                        "status": "Erro / Re-cotar"
                    }
                    existing.append(err_data)
                    save_and_export_results(existing)
                    self.item_finished_signal.emit(err_data)

                time.sleep(1.0)

            try:
                browser.close()
            except:
                pass

        if self.stop_requested:
            self.batch_finished_signal.emit(True, "Pesquisa interrompida com segurança pelo usuário.")
        else:
            self.batch_finished_signal.emit(True, f"Ciclo de pesquisa finalizado com sucesso! ({len(existing)} itens salvos)")

# ==============================================================================
# ESTILOS CORPORATIVOS E MODERNOS EM QSS
# ==============================================================================
CORPORATE_STYLE = """
QMainWindow, QWidget#centralWidget {
    background-color: #f8fafc;
}
QWidget {
    font-family: "Segoe UI", Arial, sans-serif;
    color: #0f172a;
    font-size: 13px;
}
QToolTip {
    background-color: #0f172a;
    color: #f8fafc;
    border: 1px solid #334155;
    padding: 6px 10px;
    border-radius: 5px;
    font-size: 12px;
}
QTabWidget::pane {
    border: 1px solid #cbd5e1;
    background: #ffffff;
    border-radius: 6px;
    top: -1px;
}
QTabBar::tab {
    background: #f1f5f9;
    color: #475569;
    padding: 9px 20px;
    margin-right: 4px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    font-weight: 600;
    font-size: 13px;
    border: 1px solid #e2e8f0;
    border-bottom: none;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: #1e40af;
    border-top: 3px solid #2563eb;
    border-left: 1px solid #cbd5e1;
    border-right: 1px solid #cbd5e1;
    border-bottom: 1px solid #ffffff;
    font-weight: 700;
}
QTabBar::tab:hover:!selected {
    background: #e2e8f0;
    color: #1e293b;
}
QPushButton {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 14px;
    font-weight: 600;
    color: #1e293b;
}
QPushButton:hover {
    background-color: #f8fafc;
    border-color: #94a3b8;
}
QPushButton:pressed {
    background-color: #e2e8f0;
}
QPushButton:disabled {
    background-color: #f1f5f9;
    color: #94a3b8;
    border-color: #e2e8f0;
}
QPushButton#primaryBtn {
    background-color: #2563eb;
    color: #ffffff;
    border: 1px solid #1d4ed8;
}
QPushButton#primaryBtn:hover {
    background-color: #1d4ed8;
}
QPushButton#primaryBtn:pressed {
    background-color: #1e40af;
}
QPushButton#primaryBtn:disabled {
    background-color: #93c5fd;
    color: #ffffff;
    border-color: #93c5fd;
}
QPushButton#dangerBtn {
    background-color: #dc2626;
    color: #ffffff;
    border: 1px solid #b91c1c;
}
QPushButton#dangerBtn:hover {
    background-color: #b91c1c;
}
QPushButton#dangerBtn:pressed {
    background-color: #991b1b;
}
QTableWidget {
    background-color: #ffffff;
    alternate-background-color: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    gridline-color: #f1f5f9;
    selection-background-color: #dbeafe;
    selection-color: #1e40af;
}
QHeaderView::section {
    background-color: #f8fafc;
    padding: 8px 10px;
    border: none;
    border-bottom: 2px solid #cbd5e1;
    border-right: 1px solid #f1f5f9;
    font-weight: 700;
    color: #334155;
    font-size: 12px;
}
QLineEdit, QComboBox, QSpinBox {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 5px 9px;
    color: #0f172a;
}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
    border: 1.5px solid #2563eb;
}
QProgressBar {
    border: 1px solid #cbd5e1;
    border-radius: 5px;
    text-align: center;
    background-color: #e2e8f0;
    font-weight: 700;
    color: #1e293b;
    font-size: 11px;
}
QProgressBar::chunk {
    background-color: #16a34a;
    border-radius: 4px;
}
QGroupBox {
    font-weight: 700;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    margin-top: 14px;
    padding: 14px 10px 10px 10px;
    background-color: #ffffff;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    top: 2px;
    padding: 2px 8px;
    background-color: #ffffff;
    color: #1e3a8a;
    font-size: 12px;
    font-weight: 700;
}
QGroupBox#grpReset {
    font-weight: 700;
    border: 1px solid #fed7aa;
    border-radius: 8px;
    margin-top: 14px;
    padding: 14px 10px 10px 10px;
    background-color: #fffaf5;
}
QGroupBox#grpReset::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    top: 2px;
    padding: 2px 8px;
    background-color: #fffaf5;
    color: #c2410c;
    font-size: 12px;
    font-weight: 700;
}
QFrame#metricCard {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
}
QFrame#metricCard:hover {
    border-color: #93c5fd;
    background-color: #fbfcfe;
}
QFrame#detailCard {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
}
QScrollArea {
    background-color: #f8fafc;
    border: none;
}
QScrollBar:vertical {
    border: none;
    background: #f1f5f9;
    width: 9px;
    margin: 0px;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #cbd5e1;
    min-height: 24px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #94a3b8;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QScrollBar:horizontal {
    border: none;
    background: #f1f5f9;
    height: 9px;
    margin: 0px;
    border-radius: 4px;
}
QScrollBar::handle:horizontal {
    background: #cbd5e1;
    min-width: 24px;
    border-radius: 4px;
}
QScrollBar::handle:horizontal:hover {
    background: #94a3b8;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
}
QSplitter::handle {
    background-color: #e2e8f0;
}
QSplitter::handle:horizontal {
    width: 6px;
}
QSplitter::handle:vertical {
    height: 6px;
}
QSplitter::handle:hover {
    background-color: #3b82f6;
}
"""

# ==============================================================================
# JANELA PRINCIPAL DA APLICAÇÃO DESKTOP QT
# ==============================================================================
class CotadorQtApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Pesquisa de Preços de Mercado — Pregão SRP (Lei 14.133/2021)")
        self.resize(1360, 890)
        self.setMinimumSize(960, 600)
        self.setStyleSheet(CORPORATE_STYLE)

        self.worker = None
        self.current_master_pixmap = None
        self.image_zoom_mode = "width"  # "width", "window", "manual"
        self.image_zoom_factor = 1.0

        self.init_ui()
        self.refresh_all_data()
        self.start_background_web_server()

    def init_ui(self):
        central_widget = QWidget(self)
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(12, 10, 12, 10)
        main_layout.setSpacing(8)

        # 1. CABEÇALHO COMPACTO INSTITUCIONAL
        header_frame = QFrame()
        header_frame.setObjectName("headerBox")
        header_frame.setFixedHeight(54)
        header_frame.setStyleSheet("""
            QFrame#headerBox {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #0f2942, stop:1 #1e3a8a);
                border-radius: 8px;
                border-bottom: 3px solid #2563eb;
            }
        """)
        h_layout = QHBoxLayout(header_frame)
        h_layout.setContentsMargins(14, 0, 14, 0)
        h_layout.setSpacing(12)

        pill_law = QLabel("🏛️ LEI 14.133/2021")
        pill_law.setStyleSheet("""
            background-color: rgba(255, 255, 255, 0.15);
            color: #ffffff;
            font-size: 10px;
            font-weight: 700;
            padding: 3px 8px;
            border-radius: 4px;
            border: 1px solid rgba(255, 255, 255, 0.25);
        """)

        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(1)
        title_vbox.setAlignment(Qt.AlignVCenter)
        title_lbl = QLabel("Sistema de Pesquisa de Preços de Mercado — Pregão SRP")
        title_lbl.setStyleSheet("color: #ffffff; font-size: 15px; font-weight: 700;")
        sub_lbl = QLabel("Instrução Processual de Preços • Coleta Automatizada com Auditoria por IA e Carimbo Oficial")
        sub_lbl.setStyleSheet("color: #93c5fd; font-size: 11px;")
        title_vbox.addWidget(title_lbl)
        title_vbox.addWidget(sub_lbl)

        h_layout.addWidget(pill_law)
        h_layout.addLayout(title_vbox)
        h_layout.addStretch()

        self.btn_toggle_sidebar = QPushButton("◨ Ocultar Painel")
        self.btn_toggle_sidebar.setCursor(Qt.PointingHandCursor)
        self.btn_toggle_sidebar.setToolTip("Oculta ou exibe o painel lateral para maximizar a área de trabalho da tabela")
        self.btn_toggle_sidebar.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 255, 255, 0.15);
                color: #ffffff;
                font-weight: 600;
                font-size: 12px;
                padding: 6px 12px;
                border: 1px solid rgba(255, 255, 255, 0.3);
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.25);
            }
        """)
        self.btn_toggle_sidebar.clicked.connect(self.toggle_sidebar)

        btn_refresh_top = QPushButton("🔄 Atualizar Dados")
        btn_refresh_top.setCursor(Qt.PointingHandCursor)
        btn_refresh_top.setToolTip("Atualiza as cotações, comprovantes e cálculos na tela")
        btn_refresh_top.setStyleSheet("""
            QPushButton {
                background-color: #2563eb;
                color: #ffffff;
                font-weight: 600;
                font-size: 12px;
                padding: 6px 14px;
                border: 1px solid #1d4ed8;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #1d4ed8;
            }
        """)
        btn_refresh_top.clicked.connect(self.refresh_all_data)

        btn_web_version = QPushButton("🌐 Abrir no Navegador")
        btn_web_version.setCursor(Qt.PointingHandCursor)
        btn_web_version.setToolTip("Inicia o servidor web local e abre o sistema no seu navegador padrão")
        btn_web_version.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 255, 255, 0.15);
                color: #ffffff;
                font-weight: 600;
                font-size: 12px;
                padding: 6px 12px;
                border: 1px solid rgba(255, 255, 255, 0.3);
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.25);
            }
        """)
        btn_web_version.clicked.connect(self.action_open_web_server)

        h_layout.addWidget(self.btn_toggle_sidebar)
        h_layout.addWidget(btn_web_version)
        h_layout.addWidget(btn_refresh_top)
        main_layout.addWidget(header_frame)

        # 2. CARDS DE MÉTRICAS (KPI STRIP COMPACTA E RESPONSIVA)
        metrics_layout = QHBoxLayout()
        metrics_layout.setSpacing(8)

        self.card_total = self.create_metric_card("Total na Planilha", "0 itens", "#1e3a8a", "📋")
        self.card_homolog = self.create_metric_card("Itens Homologados", "0 itens", "#166534", "✅")
        self.card_media = self.create_metric_card("Média de Valores", "R$ 0,00", "#0284c7", "💰")
        self.card_prints = self.create_metric_card("Comprovantes Prontos", "0 arquivos", "#4f46e5", "📸")
        self.card_cota = self.create_metric_card("Cota Gemini (Hoje)", "500 restantes", "#7c3aed", "⚡")

        metrics_layout.addWidget(self.card_total)
        metrics_layout.addWidget(self.card_homolog)
        metrics_layout.addWidget(self.card_media)
        metrics_layout.addWidget(self.card_prints)
        metrics_layout.addWidget(self.card_cota)
        main_layout.addLayout(metrics_layout)

        # 3. MONITOR VISUAL EM TEMPO REAL (LIVE STATUS)
        self.live_box = QFrame()
        self.live_box.setStyleSheet("""
            QFrame {
                background-color: #f0fdf4;
                border: 1px solid #86efac;
                border-left: 5px solid #16a34a;
                border-radius: 6px;
                padding: 6px 10px;
            }
        """)
        live_layout = QVBoxLayout(self.live_box)
        live_layout.setContentsMargins(10, 6, 10, 6)
        live_layout.setSpacing(4)

        live_top = QHBoxLayout()
        self.lbl_live_title = QLabel("● PESQUISA DE MERCADO EM ANDAMENTO")
        self.lbl_live_title.setStyleSheet("color: #166534; font-weight: 700; font-size: 12px;")
        self.lbl_live_badge = QLabel("Lote --")
        self.lbl_live_badge.setStyleSheet("background-color: #dcfce7; color: #166534; font-weight: 600; padding: 2px 8px; border-radius: 4px; border: 1px solid #86efac;")
        live_top.addWidget(self.lbl_live_title)
        live_top.addStretch()
        live_top.addWidget(self.lbl_live_badge)

        self.lbl_live_item = QLabel("Aguardando início...")
        self.lbl_live_item.setStyleSheet("font-size: 13px; font-weight: 600; color: #0f2942;")

        self.lbl_live_details = QLabel("Fornecedor: --  |  Etapa: --")
        self.lbl_live_details.setStyleSheet("color: #475569; font-size: 11px;")

        self.live_progress = QProgressBar()
        self.live_progress.setFixedHeight(10)
        self.live_progress.setValue(0)

        live_layout.addLayout(live_top)
        live_layout.addWidget(self.lbl_live_item)
        live_layout.addWidget(self.lbl_live_details)
        live_layout.addWidget(self.live_progress)
        self.live_box.setVisible(False)
        main_layout.addWidget(self.live_box)

        # 4. ÁREA CENTRAL (SPLITTER ENTRE BARRA LATERAL E ABAS)
        self.main_splitter = QSplitter(Qt.Horizontal)

        # PAINEL LATERAL DE CONTROLE DENTRO DE SCROLL AREA COM FUNDO CLARO
        sidebar_widget = QWidget()
        sidebar_widget.setStyleSheet("background-color: #f8fafc;")
        sidebar_layout = QVBoxLayout(sidebar_widget)
        sidebar_layout.setContentsMargins(4, 4, 8, 8)
        sidebar_layout.setSpacing(10)

        # Grupo 1: Pesquisa de Preços
        grp_exec = QGroupBox("▶️ 1. Pesquisa de Preços")
        g_exec_layout = QVBoxLayout(grp_exec)
        g_exec_layout.setSpacing(8)

        self.rb_todos = QRadioButton("Todos os itens da planilha")
        self.rb_faixa = QRadioButton("Faixa de itens específica")
        self.rb_unit = QRadioButton("Lote unitário específico")
        self.rb_todos.setChecked(True)

        self.btn_group_exec = QButtonGroup(self)
        self.btn_group_exec.addButton(self.rb_todos, 1)
        self.btn_group_exec.addButton(self.rb_faixa, 2)
        self.btn_group_exec.addButton(self.rb_unit, 3)
        self.btn_group_exec.idToggled.connect(self.on_exec_mode_changed)

        self.widget_faixa = QWidget()
        w_f_layout = QHBoxLayout(self.widget_faixa)
        w_f_layout.setContentsMargins(0, 0, 0, 0)
        self.spin_inicio = QSpinBox()
        self.spin_inicio.setRange(1, 9999)
        self.spin_inicio.setValue(1)
        self.spin_qtd = QSpinBox()
        self.spin_qtd.setRange(1, 9999)
        self.spin_qtd.setValue(10)
        w_f_layout.addWidget(QLabel("Do item:"))
        w_f_layout.addWidget(self.spin_inicio)
        w_f_layout.addWidget(QLabel("Qtd:"))
        w_f_layout.addWidget(self.spin_qtd)
        self.widget_faixa.setVisible(False)

        self.widget_unit = QWidget()
        w_u_layout = QHBoxLayout(self.widget_unit)
        w_u_layout.setContentsMargins(0, 0, 0, 0)
        self.spin_lote_especifico = QSpinBox()
        self.spin_lote_especifico.setRange(1, 9999)
        self.spin_lote_especifico.setValue(1)
        w_u_layout.addWidget(QLabel("Lote nº:"))
        w_u_layout.addWidget(self.spin_lote_especifico)
        self.widget_unit.setVisible(False)

        self.btn_iniciar = QPushButton("🚀 Iniciar Pesquisa")
        self.btn_iniciar.setObjectName("primaryBtn")
        self.btn_iniciar.setFixedHeight(36)
        self.btn_iniciar.setToolTip("Inicia o motor automatizado de coleta em lojas comerciais e auditoria por IA")
        self.btn_iniciar.clicked.connect(self.action_start_search)

        self.btn_parar = QPushButton("⏹️ Interromper Pesquisa")
        self.btn_parar.setEnabled(False)
        self.btn_parar.setFixedHeight(32)
        self.btn_parar.clicked.connect(self.action_stop_search)

        g_exec_layout.addWidget(self.rb_todos)
        g_exec_layout.addWidget(self.rb_faixa)
        g_exec_layout.addWidget(self.widget_faixa)
        g_exec_layout.addWidget(self.rb_unit)
        g_exec_layout.addWidget(self.widget_unit)
        g_exec_layout.addWidget(self.btn_iniciar)
        g_exec_layout.addWidget(self.btn_parar)
        sidebar_layout.addWidget(grp_exec)

        # Grupo 2: Planilha de Trabalho
        grp_planilha = QGroupBox("📂 2. Planilha de Trabalho")
        g_plan_layout = QVBoxLayout(grp_planilha)
        g_plan_layout.setSpacing(6)
        self.lbl_plan_status = QLabel("Nenhuma planilha carregada")
        self.lbl_plan_status.setStyleSheet("font-weight: 600; color: #0f172a;")
        self.lbl_plan_status.setWordWrap(True)
        self.lbl_plan_date = QLabel("Carregada em: --")
        self.lbl_plan_date.setStyleSheet("color: #64748b; font-size: 11px;")
        
        btn_ir_planilha = QPushButton("📋 Gerenciar / Trocar Planilha")
        btn_ir_planilha.setToolTip("Abre a aba de importação e mapeamento de colunas de itens")
        btn_ir_planilha.clicked.connect(lambda: self.tabs.setCurrentIndex(3))

        self.btn_descarregar = QPushButton("🗑️ Descarregar Planilha")
        self.btn_descarregar.clicked.connect(self.action_reset_sheet)

        g_plan_layout.addWidget(self.lbl_plan_status)
        g_plan_layout.addWidget(self.lbl_plan_date)
        g_plan_layout.addWidget(btn_ir_planilha)
        g_plan_layout.addWidget(self.btn_descarregar)
        sidebar_layout.addWidget(grp_planilha)

        # Grupo 3: Modelos e Exportação
        grp_modelos = QGroupBox("📦 3. Modelos e Exportação")
        g_mod_layout = QVBoxLayout(grp_modelos)
        g_mod_layout.setSpacing(6)

        self.btn_baixar_modelo_excel = QPushButton("📥 Modelo Excel (.xlsx)")
        self.btn_baixar_modelo_excel.clicked.connect(self.action_download_model_excel)

        self.btn_baixar_modelo_csv = QPushButton("📄 Modelo CSV (.csv)")
        self.btn_baixar_modelo_csv.clicked.connect(self.action_download_model_csv)

        btn_ir_export = QPushButton("📁 Ir para Exportação Oficial")
        btn_ir_export.setToolTip("Abre a aba de exportação com planilha consolidada e pasta de comprovantes")
        btn_ir_export.clicked.connect(lambda: self.tabs.setCurrentIndex(2))

        g_mod_layout.addWidget(self.btn_baixar_modelo_excel)
        g_mod_layout.addWidget(self.btn_baixar_modelo_csv)
        g_mod_layout.addWidget(btn_ir_export)
        sidebar_layout.addWidget(grp_modelos)

        # Grupo 4: Manutenção e Segurança (Área de Ação Destrutiva)
        grp_reset = QGroupBox("🚨 4. Manutenção e Segurança")
        grp_reset.setObjectName("grpReset")
        g_rst_layout = QVBoxLayout(grp_reset)
        g_rst_layout.setSpacing(6)
        lbl_rst_info = QLabel("Reinicia os dados para um novo pregão. Cria backup automático com data e hora antes de limpar.")
        lbl_rst_info.setWordWrap(True)
        lbl_rst_info.setStyleSheet("color: #7c2d12; font-size: 11px;")
        
        btn_reset_all = QPushButton("🚨 Limpar Dados e Reiniciar")
        btn_reset_all.setObjectName("dangerBtn")
        btn_reset_all.setToolTip("Ação destrutiva: zera os registros e prints locais após backup de segurança")
        btn_reset_all.clicked.connect(self.action_reset_all)

        g_rst_layout.addWidget(lbl_rst_info)
        g_rst_layout.addWidget(btn_reset_all)
        sidebar_layout.addWidget(grp_reset)

        sidebar_layout.addStretch()

        self.sidebar_scroll = QScrollArea()
        self.sidebar_scroll.setWidgetResizable(True)
        self.sidebar_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.sidebar_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.sidebar_scroll.setFrameShape(QFrame.NoFrame)
        self.sidebar_scroll.setStyleSheet("QScrollArea { background-color: #f8fafc; border: none; }")
        self.sidebar_scroll.setWidget(sidebar_widget)
        self.sidebar_scroll.setMinimumWidth(255)
        self.sidebar_scroll.setMaximumWidth(325)

        self.main_splitter.addWidget(self.sidebar_scroll)

        # ABAS PRINCIPAIS
        self.tabs = QTabWidget()

        # ABA 1: TABELA DE COTAÇÕES COM PAINEL DE DETALHES
        tab_tabela = QWidget()
        t_tab_layout = QVBoxLayout(tab_tabela)
        t_tab_layout.setContentsMargins(10, 10, 10, 10)
        t_tab_layout.setSpacing(8)

        # Barra de Pesquisa e Filtros
        t_filter_layout = QHBoxLayout()
        t_filter_layout.setSpacing(8)
        self.edit_busca_cotacoes = QLineEdit()
        self.edit_busca_cotacoes.setPlaceholderText("🔍 Buscar por material, fornecedor, CNPJ ou nº do lote...")
        self.edit_busca_cotacoes.textChanged.connect(self.filter_cotacoes_table)

        self.combo_filtro_status = QComboBox()
        self.combo_filtro_status.addItems(["Todos os Status", "Apenas Homologados", "Apenas Pendentes"])
        self.combo_filtro_status.currentIndexChanged.connect(self.filter_cotacoes_table)

        self.btn_limpar_filtros = QPushButton("✕ Limpar")
        self.btn_limpar_filtros.setToolTip("Limpa o campo de busca e restaura o filtro de status")
        self.btn_limpar_filtros.clicked.connect(self.action_clear_filters)

        self.lbl_cotacoes_count = QLabel("Exibindo 0 cotações")
        self.lbl_cotacoes_count.setStyleSheet("color: #475569; font-size: 12px; font-weight: 600;")

        btn_atualizar_tab = QPushButton("🔄 Atualizar")
        btn_atualizar_tab.clicked.connect(self.refresh_all_data)

        t_filter_layout.addWidget(self.edit_busca_cotacoes, 5)
        t_filter_layout.addWidget(self.combo_filtro_status, 2)
        t_filter_layout.addWidget(self.btn_limpar_filtros)
        t_filter_layout.addWidget(self.lbl_cotacoes_count)
        t_filter_layout.addWidget(btn_atualizar_tab)
        t_tab_layout.addLayout(t_filter_layout)

        # Splitter Vertical interno da Aba 1 (Tabela em cima, Detalhes embaixo)
        self.tab_inner_splitter = QSplitter(Qt.Vertical)

        # Container da Tabela
        table_container = QWidget()
        tc_layout = QVBoxLayout(table_container)
        tc_layout.setContentsMargins(0, 0, 0, 0)
        tc_layout.setSpacing(0)

        self.table_cotacoes = QTableWidget()
        self.table_cotacoes.setColumnCount(8)
        self.table_cotacoes.setHorizontalHeaderLabels([
            "Lote", "Descrição Técnica do Item", "Situação", "Preço Unitário",
            "Fornecedor", "CNPJ", "Data / Hora", "Fonte"
        ])
        self.table_cotacoes.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_cotacoes.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table_cotacoes.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_cotacoes.setAlternatingRowColors(True)
        self.table_cotacoes.verticalHeader().setDefaultSectionSize(34)
        self.table_cotacoes.verticalHeader().setVisible(False)

        h_cot = self.table_cotacoes.horizontalHeader()
        h_cot.setStretchLastSection(False)
        h_cot.setSectionResizeMode(0, QHeaderView.Interactive)
        h_cot.setSectionResizeMode(1, QHeaderView.Stretch)     # ESTICA COM A JANELA!
        h_cot.setSectionResizeMode(2, QHeaderView.Interactive)
        h_cot.setSectionResizeMode(3, QHeaderView.Interactive)
        h_cot.setSectionResizeMode(4, QHeaderView.Interactive)
        h_cot.setSectionResizeMode(5, QHeaderView.Interactive)
        h_cot.setSectionResizeMode(6, QHeaderView.Interactive)
        h_cot.setSectionResizeMode(7, QHeaderView.Interactive)

        h_cot.resizeSection(0, 50)   # Lote
        h_cot.resizeSection(2, 125)  # Situação
        h_cot.resizeSection(3, 105)  # Preço Unitário
        h_cot.resizeSection(4, 140)  # Fornecedor
        h_cot.resizeSection(5, 135)  # CNPJ
        h_cot.resizeSection(6, 130)  # Data e Hora
        h_cot.resizeSection(7, 105)  # Link da Fonte

        self.table_cotacoes.itemSelectionChanged.connect(self.on_table_cotacoes_selection_changed)
        self.table_cotacoes.itemDoubleClicked.connect(self.on_table_cotacoes_double_clicked)
        tc_layout.addWidget(self.table_cotacoes)
        self.tab_inner_splitter.addWidget(table_container)

        # PAINEL DE DETALHES DO REGISTRO SELECIONADO
        self.detail_panel = QFrame()
        self.detail_panel.setObjectName("detailCard")
        det_main_layout = QVBoxLayout(self.detail_panel)
        det_main_layout.setContentsMargins(14, 10, 14, 10)
        det_main_layout.setSpacing(6)

        det_top_row = QHBoxLayout()
        det_top_row.setSpacing(8)

        self.lbl_det_lote_badge = QLabel("LOTE --")
        self.lbl_det_lote_badge.setStyleSheet("""
            background-color: #1e3a8a;
            color: #ffffff;
            font-weight: 700;
            font-size: 11px;
            padding: 3px 8px;
            border-radius: 4px;
        """)

        self.lbl_det_status_badge = QLabel("Aguardando Seleção")
        self.lbl_det_status_badge.setStyleSheet("""
            background-color: #f1f5f9;
            color: #475569;
            font-weight: 600;
            font-size: 11px;
            padding: 3px 8px;
            border-radius: 4px;
            border: 1px solid #cbd5e1;
        """)

        self.lbl_det_price_box = QLabel("VALOR: --")
        self.lbl_det_price_box.setStyleSheet("""
            background-color: #f0fdf4;
            color: #15803d;
            font-weight: 800;
            font-size: 13px;
            padding: 3px 10px;
            border-radius: 4px;
            border: 1px solid #86efac;
        """)

        self.btn_det_link = QPushButton("🌐 Abrir no Fornecedor")
        self.btn_det_link.setFixedHeight(28)
        self.btn_det_link.setEnabled(False)
        self.btn_det_link.clicked.connect(self.action_open_store_link)

        self.btn_det_print = QPushButton("📸 Ver Comprovante Deste Lote")
        self.btn_det_print.setFixedHeight(28)
        self.btn_det_print.setEnabled(False)
        self.btn_det_print.clicked.connect(self.action_view_print_from_table)

        det_top_row.addWidget(self.lbl_det_lote_badge)
        det_top_row.addWidget(self.lbl_det_status_badge)
        det_top_row.addWidget(self.lbl_det_price_box)
        det_top_row.addStretch()
        det_top_row.addWidget(self.btn_det_link)
        det_top_row.addWidget(self.btn_det_print)
        det_main_layout.addLayout(det_top_row)

        self.lbl_det_desc = QLabel("Selecione um item na tabela para visualizar as especificações completas.")
        self.lbl_det_desc.setWordWrap(True)
        self.lbl_det_desc.setStyleSheet("font-size: 13px; color: #0f172a; font-weight: 600;")
        det_main_layout.addWidget(self.lbl_det_desc)

        det_sub_row = QHBoxLayout()
        self.lbl_det_fornec = QLabel("<b>Fornecedor:</b> --  |  <b>CNPJ:</b> --")
        self.lbl_det_fornec.setStyleSheet("color: #334155; font-size: 12px;")
        self.lbl_det_data = QLabel("<b>Coleta:</b> --")
        self.lbl_det_data.setStyleSheet("color: #64748b; font-size: 11px;")
        det_sub_row.addWidget(self.lbl_det_fornec, 1)
        det_sub_row.addWidget(self.lbl_det_data)
        det_main_layout.addLayout(det_sub_row)

        self.lbl_det_obs = QLabel("<b>Memória do Preço / Auditoria:</b> --")
        self.lbl_det_obs.setWordWrap(True)
        self.lbl_det_obs.setStyleSheet("color: #475569; font-size: 11px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px; padding: 4px 8px;")
        det_main_layout.addWidget(self.lbl_det_obs)

        self.tab_inner_splitter.addWidget(self.detail_panel)
        self.tab_inner_splitter.setStretchFactor(0, 7)
        self.tab_inner_splitter.setStretchFactor(1, 3)
        self.tab_inner_splitter.setSizes([460, 160])

        t_tab_layout.addWidget(self.tab_inner_splitter)

        # Barra inferior de ações contextuais
        btn_bar_cot = QHBoxLayout()
        btn_bar_cot.setSpacing(8)

        self.btn_abrir_link_tab = QPushButton("🌐 Acessar Fonte / Loja")
        self.btn_abrir_link_tab.clicked.connect(self.action_open_store_link)

        self.btn_ver_print_tab = QPushButton("📸 Ver Comprovante Carimbado")
        self.btn_ver_print_tab.clicked.connect(self.action_view_print_from_table)

        self.btn_toggle_details = QPushButton("📋 Ocultar Painel de Detalhes")
        self.btn_toggle_details.clicked.connect(self.toggle_detail_panel)

        lbl_hint = QLabel("💡 <i>Dica: Dê dois cliques em qualquer registro para visualizar o comprovante com carimbo oficial ou acessar o link da loja.</i>")
        lbl_hint.setStyleSheet("color: #64748b; font-size: 11px;")

        btn_bar_cot.addWidget(self.btn_abrir_link_tab)
        btn_bar_cot.addWidget(self.btn_ver_print_tab)
        btn_bar_cot.addWidget(self.btn_toggle_details)
        btn_bar_cot.addStretch()
        btn_bar_cot.addWidget(lbl_hint)
        t_tab_layout.addLayout(btn_bar_cot)

        self.tabs.addTab(tab_tabela, "📊 Tabela de Preços e Fornecedores")

        # ABA 2: COMPROVANTES OFICIAIS COM ZOOM E ROLAGEM DINÂMICA
        tab_prints = QWidget()
        tab_prints.setStyleSheet("background-color: #f8fafc;")
        p_tab_layout = QVBoxLayout(tab_prints)
        p_tab_layout.setContentsMargins(10, 10, 10, 10)
        p_tab_layout.setSpacing(8)

        sel_row = QHBoxLayout()
        sel_row.setSpacing(6)
        
        sel_row.addWidget(QLabel("<b>Lote:</b>"))
        self.combo_prints_lotes = QComboBox()
        self.combo_prints_lotes.currentIndexChanged.connect(self.on_print_lote_changed)
        sel_row.addWidget(self.combo_prints_lotes, 3)

        self.btn_nav_prev = QPushButton("◀ Anterior")
        self.btn_nav_prev.clicked.connect(self.nav_print_prev)
        self.btn_nav_next = QPushButton("Próximo ▶")
        self.btn_nav_next.clicked.connect(self.nav_print_next)
        sel_row.addWidget(self.btn_nav_prev)
        sel_row.addWidget(self.btn_nav_next)

        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setFrameShadow(QFrame.Sunken)
        sel_row.addWidget(sep)

        self.btn_zoom_out = QPushButton("🔍 -")
        self.btn_zoom_out.setToolTip("Diminuir Zoom (Ctrl + Scroll para baixo)")
        self.btn_zoom_out.clicked.connect(lambda: self.zoom_image(0.85))

        self.btn_zoom_reset = QPushButton("100%")
        self.btn_zoom_reset.setToolTip("Tamanho Real 1:1")
        self.btn_zoom_reset.clicked.connect(self.zoom_reset)

        self.btn_zoom_in = QPushButton("🔍 +")
        self.btn_zoom_in.setToolTip("Aumentar Zoom (Ctrl + Scroll para cima)")
        self.btn_zoom_in.clicked.connect(lambda: self.zoom_image(1.15))

        self.btn_zoom_width = QPushButton("📐 Largura")
        self.btn_zoom_width.setToolTip("Ajustar imagem à largura visível da janela")
        self.btn_zoom_width.clicked.connect(lambda: self.set_zoom_mode("width"))

        self.btn_zoom_fit = QPushButton("🖥️ Janela")
        self.btn_zoom_fit.setToolTip("Ajustar imagem completa dentro da janela")
        self.btn_zoom_fit.clicked.connect(lambda: self.set_zoom_mode("window"))

        self.lbl_zoom_pct = QLabel("100%")
        self.lbl_zoom_pct.setStyleSheet("font-weight: 700; color: #2563eb; min-width: 40px;")

        sel_row.addWidget(self.btn_zoom_out)
        sel_row.addWidget(self.btn_zoom_reset)
        sel_row.addWidget(self.btn_zoom_in)
        sel_row.addWidget(self.btn_zoom_width)
        sel_row.addWidget(self.btn_zoom_fit)
        sel_row.addWidget(self.lbl_zoom_pct)

        sel_row.addStretch()

        self.btn_abrir_img_externa_top = QPushButton("🖼️ Abrir Externo")
        self.btn_abrir_img_externa_top.setToolTip("Abrir imagem no visualizador de fotos padrão do Windows")
        self.btn_abrir_img_externa_top.clicked.connect(self.action_open_image_external)
        sel_row.addWidget(self.btn_abrir_img_externa_top)

        p_tab_layout.addLayout(sel_row)

        print_splitter = QSplitter(Qt.Horizontal)

        self.scroll_img = QScrollArea()
        self.scroll_img.setWidgetResizable(True)
        self.scroll_img.setAlignment(Qt.AlignCenter)
        self.lbl_image_display = QLabel("Nenhum comprovante selecionado.")
        self.lbl_image_display.setAlignment(Qt.AlignCenter)
        self.lbl_image_display.setStyleSheet("background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 4px;")
        self.scroll_img.setWidget(self.lbl_image_display)
        self.scroll_img.viewport().installEventFilter(self)
        print_splitter.addWidget(self.scroll_img)

        ficha_scroll = QScrollArea()
        ficha_scroll.setWidgetResizable(True)
        ficha_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        ficha_scroll.setFrameShape(QFrame.NoFrame)
        ficha_scroll.setStyleSheet("background-color: #f8fafc;")
        
        ficha_widget = QWidget()
        ficha_widget.setStyleSheet("background-color: #f8fafc;")
        f_layout = QVBoxLayout(ficha_widget)
        f_layout.setContentsMargins(6, 0, 4, 0)
        grp_ficha = QGroupBox("📋 Ficha Oficial do Item Homologado")
        g_f_layout = QVBoxLayout(grp_ficha)
        g_f_layout.setSpacing(10)

        self.lbl_ficha_lote = QLabel("<b>Lote:</b> --")
        self.lbl_ficha_lote.setStyleSheet("font-size: 14px; color: #1e3a8a;")

        self.lbl_ficha_desc = QLabel("<b>Material:</b> --")
        self.lbl_ficha_desc.setWordWrap(True)

        box_valor = QFrame()
        box_valor.setStyleSheet("background-color: #f0fdf4; border: 1px solid #86efac; border-radius: 6px; padding: 8px;")
        bv_layout = QVBoxLayout(box_valor)
        bv_layout.setContentsMargins(8, 6, 8, 6)
        bv_title = QLabel("VALOR UNITÁRIO HOMOLOGADO:")
        bv_title.setStyleSheet("font-size: 10px; font-weight: 700; color: #166534;")
        self.lbl_ficha_valor = QLabel("R$ 0,00")
        self.lbl_ficha_valor.setStyleSheet("font-size: 18px; color: #15803d; font-weight: 800;")
        bv_layout.addWidget(bv_title)
        bv_layout.addWidget(self.lbl_ficha_valor)

        self.lbl_ficha_fornec = QLabel("<b>Fornecedor:</b> --")
        self.lbl_ficha_fornec.setWordWrap(True)
        self.lbl_ficha_cnpj = QLabel("<b>CNPJ:</b> --")
        self.lbl_ficha_data = QLabel("<b>Data da Coleta:</b> --")
        self.lbl_ficha_obs = QLabel("<b>Memória de Cálculo:</b> --")
        self.lbl_ficha_obs.setWordWrap(True)
        self.lbl_ficha_obs.setStyleSheet("color: #475569; font-size: 12px; background: #ffffff; padding: 6px; border-radius: 4px; border: 1px solid #e2e8f0;")

        self.btn_abrir_link_ficha = QPushButton("🌐 Acessar Loja / Anúncio Original")
        self.btn_abrir_link_ficha.setObjectName("primaryBtn")
        self.btn_abrir_link_ficha.setFixedHeight(34)
        self.btn_abrir_link_ficha.clicked.connect(self.action_open_store_link_ficha)

        self.btn_abrir_img_externa = QPushButton("🖼️ Abrir Imagem no Windows")
        self.btn_abrir_img_externa.setFixedHeight(32)
        self.btn_abrir_img_externa.clicked.connect(self.action_open_image_external)

        g_f_layout.addWidget(self.lbl_ficha_lote)
        g_f_layout.addWidget(self.lbl_ficha_desc)
        g_f_layout.addWidget(box_valor)
        g_f_layout.addWidget(self.lbl_ficha_fornec)
        g_f_layout.addWidget(self.lbl_ficha_cnpj)
        g_f_layout.addWidget(self.lbl_ficha_data)
        g_f_layout.addWidget(self.lbl_ficha_obs)
        g_f_layout.addWidget(self.btn_abrir_link_ficha)
        g_f_layout.addWidget(self.btn_abrir_img_externa)
        g_f_layout.addStretch()

        f_layout.addWidget(grp_ficha)
        ficha_scroll.setWidget(ficha_widget)
        print_splitter.addWidget(ficha_scroll)

        print_splitter.setStretchFactor(0, 7)
        print_splitter.setStretchFactor(1, 3)
        print_splitter.setSizes([750, 320])
        print_splitter.splitterMoved.connect(lambda pos, idx: self.on_print_splitter_moved())

        p_tab_layout.addWidget(print_splitter)
        self.tabs.addTab(tab_prints, "📸 Comprovantes com Carimbo Oficial")

        # ABA 3: EXPORTAR PLANILHA E COMPROVANTES
        scroll_tab_export = QScrollArea()
        scroll_tab_export.setWidgetResizable(True)
        scroll_tab_export.setFrameShape(QFrame.NoFrame)
        scroll_tab_export.setStyleSheet("background-color: #f8fafc;")

        tab_export = QWidget()
        tab_export.setStyleSheet("background-color: #f8fafc;")
        e_tab_layout = QVBoxLayout(tab_export)
        e_tab_layout.setContentsMargins(16, 16, 16, 16)
        e_tab_layout.setSpacing(14)

        e_info = QLabel("Salve a planilha Excel formatada e a pasta com todos os comprovantes oficiais prontos para impressão em qualquer diretório do seu computador.")
        e_info.setStyleSheet("color: #334155; font-size: 13px; font-weight: 500;")
        e_info.setWordWrap(True)
        e_tab_layout.addWidget(e_info)

        grp_exp_box = QGroupBox("📁 Destino da Exportação no Computador")
        g_exp_layout = QVBoxLayout(grp_exp_box)
        g_exp_layout.setSpacing(10)

        quick_row = QHBoxLayout()
        quick_row.setSpacing(8)
        btn_q_desktop = QPushButton("🖥️ Área de Trabalho")
        btn_q_downloads = QPushButton("📥 Downloads")
        btn_q_docs = QPushButton("📁 Meus Documentos")
        btn_q_browse = QPushButton("📂 Procurar Outra Pasta...")
        quick_row.addWidget(btn_q_desktop)
        quick_row.addWidget(btn_q_downloads)
        quick_row.addWidget(btn_q_docs)
        quick_row.addWidget(btn_q_browse)
        g_exp_layout.addLayout(quick_row)

        user_home = os.path.expanduser("~")
        self.default_export_folder = os.path.join(user_home, "Desktop", "Cotacao_Pregao_SRP")

        btn_q_desktop.clicked.connect(lambda: self.set_export_path(os.path.join(user_home, "Desktop", "Cotacao_Pregao_SRP")))
        btn_q_downloads.clicked.connect(lambda: self.set_export_path(os.path.join(user_home, "Downloads", "Cotacao_Pregao_SRP")))
        btn_q_docs.clicked.connect(lambda: self.set_export_path(os.path.join(user_home, "Documents", "Cotacao_Pregao_SRP")))
        btn_q_browse.clicked.connect(self.action_browse_export_folder)

        dest_row = QHBoxLayout()
        self.edit_export_folder = QLineEdit(self.default_export_folder)
        dest_row.addWidget(QLabel("<b>Caminho Completo:</b>"))
        dest_row.addWidget(self.edit_export_folder, 1)
        g_exp_layout.addLayout(dest_row)

        e_tab_layout.addWidget(grp_exp_box)

        grp_content = QGroupBox("📦 Conteúdo do Pacote Oficial de Exportação")
        g_c_layout = QVBoxLayout(grp_content)
        g_c_layout.setSpacing(8)
        g_c_layout.addWidget(QLabel("📊 <b>Planilha Excel Oficial:</b> <code>Pesquisa_de_Precos_Mercado.xlsx</code> ordenada estritamente por lote, com memória de cálculo e CNPJ."))
        g_c_layout.addWidget(QLabel("📸 <b>Pasta de Comprovantes:</b> Subpasta <code>Comprovantes_Oficiais/</code> contendo os prints monocromáticos de alta fidelidade prontos para impressão."))
        g_c_layout.addWidget(QLabel("📄 <b>Relatório de Índice:</b> Arquivo <code>Indice_Pesquisa_Precos.txt</code> com o resumo analítico e checklist de conformidade."))
        e_tab_layout.addWidget(grp_content)

        exp_btns_layout = QHBoxLayout()
        exp_btns_layout.setSpacing(10)
        self.btn_exec_export = QPushButton("💾 Exportar Planilha e Comprovantes para Esta Pasta")
        self.btn_exec_export.setObjectName("primaryBtn")
        self.btn_exec_export.setFixedHeight(42)
        self.btn_exec_export.clicked.connect(self.action_export_bundle)

        self.btn_export_zip = QPushButton("📦 Salvar Pacote Compactado (.ZIP)")
        self.btn_export_zip.setFixedHeight(42)
        self.btn_export_zip.clicked.connect(self.action_export_zip)

        exp_btns_layout.addWidget(self.btn_exec_export, 6)
        exp_btns_layout.addWidget(self.btn_export_zip, 4)
        e_tab_layout.addLayout(exp_btns_layout)
        e_tab_layout.addStretch()

        scroll_tab_export.setWidget(tab_export)
        self.tabs.addTab(scroll_tab_export, "📁 Exportar Planilha e Comprovantes")

        # ABA 4: GESTÃO DA PLANILHA DE ITENS
        tab_itens = QWidget()
        tab_itens.setStyleSheet("background-color: #f8fafc;")
        i_tab_layout = QVBoxLayout(tab_itens)
        i_tab_layout.setContentsMargins(10, 10, 10, 10)
        i_tab_layout.setSpacing(8)

        grp_upload = QGroupBox("📤 Carregar / Substituir Planilha de Itens do Pregão")
        g_up_layout = QVBoxLayout(grp_upload)
        g_up_layout.setSpacing(8)

        up_btn_row = QHBoxLayout()
        self.btn_selecionar_planilha = QPushButton("📂 Selecionar Arquivo de Planilha (Excel .xlsx/.xls ou CSV .csv)...")
        self.btn_selecionar_planilha.setObjectName("primaryBtn")
        self.btn_selecionar_planilha.setFixedHeight(34)
        self.btn_selecionar_planilha.clicked.connect(self.action_select_sheet_file)
        self.lbl_selected_file_name = QLabel("Nenhum arquivo selecionado.")
        self.lbl_selected_file_name.setStyleSheet("color: #64748b; font-weight: 500;")
        up_btn_row.addWidget(self.btn_selecionar_planilha)
        up_btn_row.addWidget(self.lbl_selected_file_name, 1)
        g_up_layout.addLayout(up_btn_row)

        self.widget_mapping = QWidget()
        m_layout = QHBoxLayout(self.widget_mapping)
        m_layout.setContentsMargins(0, 4, 0, 0)
        self.combo_col_desc = QComboBox()
        self.combo_col_lote = QComboBox()
        m_layout.addWidget(QLabel("<b>Coluna de DESCRIÇÃO:</b>"))
        m_layout.addWidget(self.combo_col_desc, 2)
        m_layout.addWidget(QLabel("<b>Coluna de LOTE:</b>"))
        m_layout.addWidget(self.combo_col_lote, 2)
        self.btn_ativar_planilha = QPushButton("🚀 Ativar Esta Planilha")
        self.btn_ativar_planilha.setObjectName("primaryBtn")
        self.btn_ativar_planilha.setFixedHeight(32)
        self.btn_ativar_planilha.clicked.connect(self.action_activate_selected_sheet)
        m_layout.addWidget(self.btn_ativar_planilha)
        self.widget_mapping.setVisible(False)
        g_up_layout.addWidget(self.widget_mapping)

        i_tab_layout.addWidget(grp_upload)

        filt_itens_row = QHBoxLayout()
        filt_itens_row.setSpacing(8)
        self.edit_busca_itens = QLineEdit()
        self.edit_busca_itens.setPlaceholderText("🔍 Filtrar itens por especificação técnica ou número do lote...")
        self.edit_busca_itens.textChanged.connect(self.filter_itens_table)
        
        self.lbl_itens_count = QLabel("Exibindo 0 itens")
        self.lbl_itens_count.setStyleSheet("color: #64748b; font-size: 12px; font-weight: 600;")
        
        filt_itens_row.addWidget(self.edit_busca_itens, 4)
        filt_itens_row.addWidget(self.lbl_itens_count, 1)
        i_tab_layout.addLayout(filt_itens_row)

        self.table_itens_ativos = QTableWidget()
        self.table_itens_ativos.setColumnCount(3)
        self.table_itens_ativos.setHorizontalHeaderLabels(["Lote", "Descrição Técnica do Material", "Situação"])
        self.table_itens_ativos.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_itens_ativos.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_itens_ativos.setAlternatingRowColors(True)
        self.table_itens_ativos.verticalHeader().setDefaultSectionSize(32)
        self.table_itens_ativos.verticalHeader().setVisible(False)

        h_itens = self.table_itens_ativos.horizontalHeader()
        h_itens.setStretchLastSection(False)
        h_itens.setSectionResizeMode(0, QHeaderView.Interactive)
        h_itens.setSectionResizeMode(1, QHeaderView.Stretch)
        h_itens.setSectionResizeMode(2, QHeaderView.Interactive)
        h_itens.resizeSection(0, 65)
        h_itens.resizeSection(2, 160)

        i_tab_layout.addWidget(self.table_itens_ativos)

        self.tabs.addTab(tab_itens, "⚙️ Planilha de Itens em Uso")

        self.main_splitter.addWidget(self.tabs)
        self.main_splitter.setStretchFactor(0, 0)
        self.main_splitter.setStretchFactor(1, 1)
        self.main_splitter.setSizes([270, 1050])
        self.main_splitter.setCollapsible(0, True)
        self.main_splitter.setCollapsible(1, False)
        main_layout.addWidget(self.main_splitter, 1)

    # AUXILIARES DE UI E CARDS
    def create_metric_card(self, title, val, color="#1e3a8a", icon="📊"):
        frame = QFrame()
        frame.setObjectName("metricCard")
        frame.setFixedHeight(50)
        l = QVBoxLayout(frame)
        l.setContentsMargins(10, 5, 10, 5)
        l.setSpacing(1)

        t_lbl = QLabel(f"{icon}  {title}")
        t_lbl.setStyleSheet("color: #64748b; font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;")
        v_lbl = QLabel(val)
        v_lbl.setStyleSheet(f"color: {color}; font-size: 16px; font-weight: 700;")
        frame.val_label = v_lbl

        l.addWidget(t_lbl)
        l.addWidget(v_lbl)
        return frame

    def toggle_sidebar(self):
        is_hidden = self.sidebar_scroll.isHidden()
        self.sidebar_scroll.setHidden(not is_hidden)
        self.btn_toggle_sidebar.setText("◧ Mostrar Painel" if not is_hidden else "◨ Ocultar Painel")

    def toggle_detail_panel(self):
        is_hidden = self.detail_panel.isHidden()
        self.detail_panel.setHidden(not is_hidden)
        self.btn_toggle_details.setText("📋 Mostrar Painel de Detalhes" if not is_hidden else "📋 Ocultar Painel de Detalhes")

    def action_clear_filters(self):
        self.edit_busca_cotacoes.clear()
        self.combo_filtro_status.setCurrentIndex(0)
        self.filter_cotacoes_table()

    def on_exec_mode_changed(self, btn_id):
        self.widget_faixa.setVisible(self.rb_faixa.isChecked())
        self.widget_unit.setVisible(self.rb_unit.isChecked())

    def set_export_path(self, path):
        self.edit_export_folder.setText(path)

    def set_zoom_mode(self, mode):
        self.image_zoom_mode = mode
        self.update_image_display()

    def zoom_image(self, factor):
        self.image_zoom_mode = "manual"
        self.image_zoom_factor = max(0.2, min(5.0, self.image_zoom_factor * factor))
        self.update_image_display()

    def zoom_reset(self):
        self.image_zoom_mode = "manual"
        self.image_zoom_factor = 1.0
        self.update_image_display()

    def nav_print_prev(self):
        idx = self.combo_prints_lotes.currentIndex()
        if idx > 0:
            self.combo_prints_lotes.setCurrentIndex(idx - 1)

    def nav_print_next(self):
        idx = self.combo_prints_lotes.currentIndex()
        if idx < self.combo_prints_lotes.count() - 1:
            self.combo_prints_lotes.setCurrentIndex(idx + 1)

    def on_print_splitter_moved(self):
        if hasattr(self, "image_zoom_mode") and self.image_zoom_mode in ("width", "window"):
            QTimer.singleShot(40, self.update_image_display)

    def update_image_display(self):
        if not getattr(self, "current_master_pixmap", None) or self.current_master_pixmap.isNull():
            return
        
        vp_width = max(100, self.scroll_img.viewport().width() - 20)
        vp_height = max(100, self.scroll_img.viewport().height() - 20)
        
        orig_w = self.current_master_pixmap.width()
        orig_h = self.current_master_pixmap.height()
        if orig_w <= 0 or orig_h <= 0:
            return

        if self.image_zoom_mode == "width":
            scaled = self.current_master_pixmap.scaledToWidth(vp_width, Qt.SmoothTransformation)
            self.image_zoom_factor = scaled.width() / orig_w
        elif self.image_zoom_mode == "window":
            scaled = self.current_master_pixmap.scaled(vp_width, vp_height, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.image_zoom_factor = scaled.width() / orig_w
        else:  # "manual"
            target_w = int(orig_w * self.image_zoom_factor)
            target_w = max(100, min(8000, target_w))
            scaled = self.current_master_pixmap.scaledToWidth(target_w, Qt.SmoothTransformation)
            self.image_zoom_factor = scaled.width() / orig_w
            
        self.lbl_image_display.setPixmap(scaled)
        self.lbl_image_display.adjustSize()
        if hasattr(self, "lbl_zoom_pct"):
            pct = int(self.image_zoom_factor * 100)
            self.lbl_zoom_pct.setText(f"{pct}%")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "image_zoom_mode") and self.image_zoom_mode in ("width", "window"):
            QTimer.singleShot(50, self.update_image_display)

    def eventFilter(self, source, event):
        if hasattr(self, "scroll_img") and source == self.scroll_img.viewport():
            if event.type() == event.Type.Wheel:
                if event.modifiers() & Qt.ControlModifier:
                    delta = event.angleDelta().y()
                    if delta > 0:
                        self.zoom_image(1.15)
                    elif delta < 0:
                        self.zoom_image(1 / 1.15)
                    return True
        return super().eventFilter(source, event)

    def on_table_cotacoes_double_clicked(self, item):
        if not item:
            return
        col = item.column()
        if col == 7:
            self.action_open_store_link()
        else:
            self.action_view_print_from_table()

    def show_detail_panel_empty(self):
        if hasattr(self, "lbl_det_lote_badge"):
            self.lbl_det_lote_badge.setText("LOTE --")
            self.lbl_det_status_badge.setText("Nenhum Item Selecionado")
            self.lbl_det_price_box.setText("VALOR: --")
            self.lbl_det_desc.setText("Selecione um item na tabela acima para visualizar as especificações completas, comprovante e memória de cálculo.")
            self.lbl_det_fornec.setText("<b>Fornecedor:</b> --  |  <b>CNPJ:</b> --")
            self.lbl_det_data.setText("<b>Coleta:</b> --")
            self.lbl_det_obs.setText("<b>Memória do Preço / Auditoria:</b> --")
            self.btn_det_link.setEnabled(False)
            self.btn_det_print.setEnabled(False)

    def on_table_cotacoes_selection_changed(self):
        curr_row = self.table_cotacoes.currentRow()
        if curr_row < 0:
            self.show_detail_panel_empty()
            return
        
        lote_item = self.table_cotacoes.item(curr_row, 0)
        if not lote_item:
            self.show_detail_panel_empty()
            return
        
        try:
            target_lote = int(lote_item.text())
        except:
            target_lote = lote_item.text()
            
        quotes = getattr(self, "all_quotes_cache", [])
        matched = [q for q in quotes if q.get("lote") == target_lote]
        if not matched:
            self.show_detail_panel_empty()
            return
            
        q = matched[0]
        lote_num = q.get("lote", 0)
        self.lbl_det_lote_badge.setText(f"LOTE {lote_num:02d}")
        
        is_homolog = q.get("status") == "Cotado e Homologado"
        self.lbl_det_status_badge.setText("✓ Homologado com CNPJ" if is_homolog else "⏳ Pendente")
        if is_homolog:
            self.lbl_det_status_badge.setStyleSheet("background-color: #dcfce7; color: #15803d; font-weight: 700; font-size: 11px; padding: 3px 8px; border-radius: 4px; border: 1px solid #86efac;")
        else:
            self.lbl_det_status_badge.setStyleSheet("background-color: #fef9c3; color: #854d0e; font-weight: 700; font-size: 11px; padding: 3px 8px; border-radius: 4px; border: 1px solid #fde047;")

        v_u = q.get("valor_unitario")
        if v_u is not None:
            v_str = f"R$ {v_u:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self.lbl_det_price_box.setText(f"VALOR HOMOLOGADO: {v_str}")
        else:
            self.lbl_det_price_box.setText("VALOR: PENDENTE")

        desc_full = str(q.get("descricao_licitacao", "--"))
        self.lbl_det_desc.setText(f"<b>Material:</b> {desc_full}")

        forn = str(q.get("fornecedor", "--"))
        cnpj = str(q.get("cnpj", "--"))
        orig = str(q.get("cnpj_origem", "Receita Federal / Base de Redes"))
        prod_cot = str(q.get("nome_produto_cotado", "--"))
        self.lbl_det_fornec.setText(f"<b>Fornecedor:</b> {forn} | <b>CNPJ:</b> {cnpj} <span style='color:#64748b;'>({orig})</span> | <b>Item Cotado:</b> {prod_cot}")

        dt = str(q.get("data_hora", "--"))
        self.lbl_det_data.setText(f"<b>Data/Hora Oficial:</b> {dt}")

        obs = str(q.get("observacao", "--"))
        self.lbl_det_obs.setText(f"<b>Memória do Preço & Auditoria IA:</b> {obs}")

        has_link = bool(q.get("url") and str(q.get("url")).startswith("http"))
        self.btn_det_link.setEnabled(has_link)

        has_print = bool(q.get("arquivo_print") and os.path.exists(q.get("arquivo_print")))
        self.btn_det_print.setEnabled(has_print)

    # CARREGAMENTO E ATUALIZAÇÃO DE DADOS
    def refresh_all_data(self):
        items = load_sheet_items()
        sheet_info = get_active_sheet_info()
        quotes = []
        if os.path.exists(RESULTS_JSON):
            try:
                with open(RESULTS_JSON, "r", encoding="utf-8") as f:
                    quotes = json.load(f)
            except:
                pass

        total_items = len(items)
        homologados = [q for q in quotes if q.get("status") == "Cotado e Homologado"]
        quoted_lotes = {q.get("lote") for q in homologados}

        valid_prices = [float(q.get("valor_unitario")) for q in homologados if q.get("valor_unitario") is not None]
        avg_price = sum(valid_prices) / len(valid_prices) if valid_prices else 0.0

        prints_count = len(glob.glob(os.path.join(PRINTS_DIR, "*.*"))) if os.path.exists(PRINTS_DIR) else 0

        quota_manager.load_state()
        rpd_rem = max(0, quota_manager.rpd_max - quota_manager.rpd_used)

        self.card_total.val_label.setText(f"{total_items} itens")
        pct = (len(quoted_lotes) / total_items * 100) if total_items > 0 else 0.0
        self.card_homolog.val_label.setText(f"{len(quoted_lotes)} itens ({pct:.0f}%)")
        self.card_media.val_label.setText(f"R$ {avg_price:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        self.card_prints.val_label.setText(f"{prints_count} arquivos")
        self.card_cota.val_label.setText(f"{rpd_rem} restantes")

        if total_items > 0:
            self.lbl_plan_status.setText(f"Ativa: {sheet_info.get('source_name', 'Planilha Carregada')}")
            self.lbl_plan_date.setText(f"Carregada em: {sheet_info.get('updated_at', '--')}")
            self.btn_descarregar.setEnabled(True)
            self.btn_iniciar.setEnabled(True)
        else:
            self.lbl_plan_status.setText("Nenhuma planilha carregada")
            self.lbl_plan_date.setText("O sistema inicia zerado.")
            self.btn_descarregar.setEnabled(False)
            self.btn_iniciar.setEnabled(False)

        self.all_quotes_cache = quotes
        self.filter_cotacoes_table()

        self.all_homologados_cache = homologados
        self.combo_prints_lotes.blockSignals(True)
        curr_idx = self.combo_prints_lotes.currentIndex()
        self.combo_prints_lotes.clear()
        if homologados:
            for q in homologados:
                lote = q.get("lote", 0)
                desc = q.get("descricao_licitacao", "")[:50]
                self.combo_prints_lotes.addItem(f"Lote {lote:02d} — {desc}", q)
            if curr_idx >= 0 and curr_idx < len(homologados):
                self.combo_prints_lotes.setCurrentIndex(curr_idx)
            else:
                self.combo_prints_lotes.setCurrentIndex(0)
            self.on_print_lote_changed()
        else:
            self.current_master_pixmap = None
            self.lbl_image_display.setText("Nenhum comprovante homologado até o momento.")
            self.lbl_ficha_lote.setText("<b>Lote:</b> --")
            self.lbl_ficha_desc.setText("<b>Material:</b> --")
            self.lbl_ficha_valor.setText("R$ 0,00")
            self.lbl_ficha_fornec.setText("<b>Fornecedor:</b> --")
            self.lbl_ficha_cnpj.setText("<b>CNPJ:</b> --")
            self.lbl_ficha_data.setText("<b>Data da Coleta:</b> --")
            self.lbl_ficha_obs.setText("<b>Memória de Cálculo:</b> --")
            if hasattr(self, "lbl_zoom_pct"):
                self.lbl_zoom_pct.setText("--")
            if hasattr(self, "btn_nav_prev"):
                self.btn_nav_prev.setEnabled(False)
            if hasattr(self, "btn_nav_next"):
                self.btn_nav_next.setEnabled(False)
        self.combo_prints_lotes.blockSignals(False)

        self.all_items_cache = items
        self.quoted_lotes_set = quoted_lotes
        self.filter_itens_table()

    def filter_cotacoes_table(self):
        query = self.edit_busca_cotacoes.text().strip().lower()
        status_filter = self.combo_filtro_status.currentText() if hasattr(self, "combo_filtro_status") else "Todos os Status"
        quotes = getattr(self, "all_quotes_cache", [])

        filtered = []
        for q in quotes:
            st = str(q.get("status", ""))
            if status_filter == "Apenas Homologados" and st != "Cotado e Homologado":
                continue
            if status_filter == "Apenas Pendentes" and st == "Cotado e Homologado":
                continue

            if not query:
                filtered.append(q)
            else:
                m_desc = str(q.get("descricao_licitacao", "")).lower()
                m_forn = str(q.get("fornecedor", "")).lower()
                m_cnpj = str(q.get("cnpj", "")).lower()
                m_lote = str(q.get("lote", ""))
                if query in m_desc or query in m_forn or query in m_cnpj or query == m_lote:
                    filtered.append(q)

        self.table_cotacoes.setRowCount(len(filtered))
        for row, q in enumerate(filtered):
            lote_num = q.get("lote", 0)
            lote_item = QTableWidgetItem(f"{lote_num:02d}")
            lote_item.setTextAlignment(Qt.AlignCenter)
            lote_item.setToolTip(f"Lote {lote_num}")

            desc_text = str(q.get("descricao_licitacao", ""))
            desc_item = QTableWidgetItem(desc_text)
            desc_item.setToolTip(f"<b>Descrição Técnica Completa:</b><br>{desc_text}")

            is_homolog = q.get("status") == "Cotado e Homologado"
            sit_str = "✓ Homologado" if is_homolog else "⏳ Pendente"
            status_item = QTableWidgetItem(sit_str)
            status_item.setTextAlignment(Qt.AlignCenter)
            if is_homolog:
                status_item.setForeground(QColor("#15803d"))
                status_item.setFont(QFont("Segoe UI", 9, QFont.Bold))
                status_item.setToolTip("Preço coletado e homologado com validação oficial de CNPJ.")
            else:
                status_item.setForeground(QColor("#b45309"))
                status_item.setFont(QFont("Segoe UI", 9, QFont.Bold))
                status_item.setToolTip("Cotação pendente de processamento.")

            v_u = q.get("valor_unitario")
            if v_u is not None:
                v_str = f"R$ {v_u:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                val_item = QTableWidgetItem(v_str)
                val_item.setFont(QFont("Segoe UI", 9, QFont.Bold))
                val_item.setForeground(QColor("#15803d"))
                val_item.setToolTip(f"<b>Preço Unitário:</b> {v_str}<br><b>Total Anúncio:</b> R$ {q.get('valor_total_anuncio', 0):,.2f}<br><b>Embalagem:</b> {q.get('quantidade_embalagem', 1)} unid.")
            else:
                val_item = QTableWidgetItem("—")
                val_item.setForeground(QColor("#94a3b8"))
                val_item.setToolTip("Aguardando cotação de mercado")
            val_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            forn_name = str(q.get("fornecedor", "-"))
            forn_item = QTableWidgetItem(forn_name)
            forn_item.setToolTip(f"<b>Fornecedor:</b> {forn_name}<br><b>Item Ofertado:</b> {q.get('nome_produto_cotado', '-')}")

            cnpj_str = str(q.get("cnpj", "-"))
            cnpj_item = QTableWidgetItem(cnpj_str)
            cnpj_item.setTextAlignment(Qt.AlignCenter)
            cnpj_item.setFont(QFont("Consolas", 9, QFont.Bold))
            cnpj_item.setToolTip(f"<b>CNPJ:</b> {cnpj_str}<br><b>Origem:</b> {q.get('cnpj_origem', 'Receita Federal / Base Oficial')}")

            data_str = str(q.get("data_hora", "-"))
            data_item = QTableWidgetItem(data_str)
            data_item.setTextAlignment(Qt.AlignCenter)
            data_item.setToolTip(f"Data e hora oficial da coleta: {data_str}")

            url_str = str(q.get("url", ""))
            link_item = QTableWidgetItem("🌐 Acessar" if url_str.startswith("http") else "—")
            link_item.setTextAlignment(Qt.AlignCenter)
            if url_str.startswith("http"):
                link_item.setForeground(QColor("#2563eb"))
                link_item.setFont(QFont("Segoe UI", 8, QFont.Bold))
                link_item.setToolTip(f"<b>URL de Origem:</b><br>{url_str}")
            else:
                link_item.setForeground(QColor("#94a3b8"))

            self.table_cotacoes.setItem(row, 0, lote_item)
            self.table_cotacoes.setItem(row, 1, desc_item)
            self.table_cotacoes.setItem(row, 2, status_item)
            self.table_cotacoes.setItem(row, 3, val_item)
            self.table_cotacoes.setItem(row, 4, forn_item)
            self.table_cotacoes.setItem(row, 5, cnpj_item)
            self.table_cotacoes.setItem(row, 6, data_item)
            self.table_cotacoes.setItem(row, 7, link_item)

        homolog_count = len([q for q in filtered if q.get("status") == "Cotado e Homologado"])
        if hasattr(self, "lbl_cotacoes_count"):
            self.lbl_cotacoes_count.setText(f"Exibindo <b>{len(filtered)}</b> de {len(quotes)} cotações ({homolog_count} homologadas)")

        if len(filtered) > 0:
            if self.table_cotacoes.currentRow() < 0:
                self.table_cotacoes.selectRow(0)
            else:
                self.on_table_cotacoes_selection_changed()
        else:
            self.show_detail_panel_empty()

    def filter_itens_table(self):
        query = self.edit_busca_itens.text().strip().lower()
        items = getattr(self, "all_items_cache", [])
        quoted_lotes = getattr(self, "quoted_lotes_set", set())

        filtered = []
        for it in items:
            if not query:
                filtered.append(it)
            else:
                m_desc = str(it.get("descricao_original", "")).lower()
                m_lote = str(it.get("lote", ""))
                if query in m_desc or query == m_lote:
                    filtered.append(it)

        self.table_itens_ativos.setRowCount(len(filtered))
        for row, it in enumerate(filtered):
            lote_num = it.get("lote", 0)
            lote_item = QTableWidgetItem(f"{lote_num:02d}")
            lote_item.setTextAlignment(Qt.AlignCenter)
            lote_item.setToolTip(f"Lote {lote_num}")

            desc_text = str(it.get("descricao_original", ""))
            desc_item = QTableWidgetItem(desc_text)
            desc_item.setToolTip(f"<b>Descrição Técnica do Material:</b><br>{desc_text}")

            is_done = lote_num in quoted_lotes
            sit_str = "✓ Homologado com CNPJ" if is_done else "⏳ Pendente"
            sit_item = QTableWidgetItem(sit_str)
            sit_item.setTextAlignment(Qt.AlignCenter)
            if is_done:
                sit_item.setForeground(QColor("#15803d"))
                sit_item.setFont(QFont("Segoe UI", 9, QFont.Bold))
            else:
                sit_item.setForeground(QColor("#b45309"))
                sit_item.setFont(QFont("Segoe UI", 9, QFont.Bold))

            self.table_itens_ativos.setItem(row, 0, lote_item)
            self.table_itens_ativos.setItem(row, 1, desc_item)
            self.table_itens_ativos.setItem(row, 2, sit_item)

        if hasattr(self, "lbl_itens_count"):
            self.lbl_itens_count.setText(f"Exibindo <b>{len(filtered)}</b> de {len(items)} itens")

    def on_print_lote_changed(self):
        quote_data = self.combo_prints_lotes.currentData()
        if not quote_data:
            return

        lote = quote_data.get("lote", "--")
        lote_str = f"{lote:02d}" if isinstance(lote, int) else str(lote)
        self.lbl_ficha_lote.setText(f"<b>Lote:</b> {lote_str}")
        self.lbl_ficha_desc.setText(f"<b>Material:</b> {quote_data.get('descricao_licitacao', '--')}")
        self.lbl_ficha_fornec.setText(f"<b>Fornecedor:</b> {quote_data.get('fornecedor', '--')}")
        self.lbl_ficha_cnpj.setText(f"<b>CNPJ:</b> {quote_data.get('cnpj', '--')} [{quote_data.get('cnpj_origem', '')}]")
        v_u = quote_data.get("valor_unitario")
        v_str = f"R$ {v_u:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if v_u else "Pendente"
        self.lbl_ficha_valor.setText(v_str)
        self.lbl_ficha_data.setText(f"<b>Data da Coleta:</b> {quote_data.get('data_hora', '--')}")
        self.lbl_ficha_obs.setText(f"<b>Memória de Cálculo:</b> {quote_data.get('observacao', '--')}")

        curr_idx = self.combo_prints_lotes.currentIndex()
        total_prints = self.combo_prints_lotes.count()
        if hasattr(self, "btn_nav_prev"):
            self.btn_nav_prev.setEnabled(curr_idx > 0)
        if hasattr(self, "btn_nav_next"):
            self.btn_nav_next.setEnabled(curr_idx < total_prints - 1)

        print_file = quote_data.get("arquivo_print")
        if print_file and os.path.exists(print_file):
            self.current_master_pixmap = QPixmap(print_file)
            self.update_image_display()
        else:
            self.current_master_pixmap = None
            self.lbl_image_display.setText("Arquivo de comprovante não encontrado no disco.")
            if hasattr(self, "lbl_zoom_pct"):
                self.lbl_zoom_pct.setText("--")
            if hasattr(self, "lbl_zoom_pct"):
                self.lbl_zoom_pct.setText("--")
            if hasattr(self, "lbl_zoom_pct"):
                self.lbl_zoom_pct.setText("--")

    # AÇÕES DO SISTEMA
    def action_start_search(self):
        items = load_sheet_items()
        if not items:
            QMessageBox.warning(self, "Aviso", "Carregue uma planilha de itens na aba 'Planilha de Itens' para iniciar a pesquisa.")
            return

        start_idx = 0
        count = len(items)

        if self.rb_faixa.isChecked():
            start_num = self.spin_inicio.value()
            count = self.spin_qtd.value()
            start_idx = max(0, start_num - 1)
        elif self.rb_unit.isChecked():
            target_lote = self.spin_lote_especifico.value()
            matched = [idx for idx, it in enumerate(items) if it.get("lote") == target_lote]
            if matched:
                start_idx = matched[0]
                count = 1
            else:
                QMessageBox.warning(self, "Lote Não Encontrado", f"O Lote {target_lote} não foi encontrado na planilha ativa.")
                return

        self.btn_iniciar.setEnabled(False)
        self.btn_parar.setEnabled(True)
        self.live_box.setVisible(True)
        self.live_progress.setValue(0)
        self.lbl_live_title.setText("● PESQUISA DE MERCADO EM ANDAMENTO")

        self.worker = CotadorWorkerThread(start_idx=start_idx, count=count, parent=self)
        self.worker.progress_signal.connect(self.on_worker_progress)
        self.worker.item_finished_signal.connect(self.on_worker_item_finished)
        self.worker.batch_finished_signal.connect(self.on_worker_finished)
        self.worker.start()

    def action_stop_search(self):
        if self.worker and self.worker.isRunning():
            self.btn_parar.setEnabled(False)
            self.lbl_live_title.setText("● INTERROMPENDO PESQUISA...")
            self.worker.request_stop()

    @Slot(dict)
    def on_worker_progress(self, p_dict):
        lote = p_dict.get("current_lote", "--")
        mat = p_dict.get("material", "Pesquisando...")
        store = p_dict.get("store", "--")
        step = p_dict.get("step", "--")
        pct = p_dict.get("progress_pct", 0)

        self.lbl_live_badge.setText(f"Lote {lote}")
        self.lbl_live_item.setText(f"📦 Lote {lote}: {mat[:70]}")
        self.lbl_live_details.setText(f"🏢 <b>Loja:</b> {store}  |  🔍 <b>Etapa:</b> {step}")
        self.live_progress.setValue(pct)

    @Slot(dict)
    def on_worker_item_finished(self, q_data):
        self.refresh_all_data()

    @Slot(bool, str)
    def on_worker_finished(self, success, message):
        self.btn_iniciar.setEnabled(True)
        self.btn_parar.setEnabled(False)
        self.live_progress.setValue(100)
        self.lbl_live_title.setText("✓ CICLO DE PESQUISA FINALIZADO")
        self.lbl_live_details.setText(message)
        self.refresh_all_data()
        QMessageBox.information(self, "Pesquisa Finalizada", message)

    def action_reset_all(self):
        reply = QMessageBox.question(
            self,
            "⚠️ Confirmação de Operação Destrutiva",
            "ATENÇÃO: Você está prestes a reiniciar o sistema para um novo pregão.\n\n"
            "Consequências desta operação:\n"
            "• Todas as cotações ativas serão arquivadas.\n"
            "• A pasta local de prints carimbados será limpa para o novo certame.\n"
            "• Um backup de segurança COMPLETO com data e hora será gerado automaticamente no disco antes da exclusão.\n\n"
            "Deseja realmente prosseguir com a limpeza e reinicialização?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            reset_all_data(backup=True)
            self.refresh_all_data()
            QMessageBox.information(
                self,
                "Sistema Reiniciado com Sucesso",
                "O sistema foi zerado e está pronto para um novo certame.\n\n"
                "Um backup de segurança completo foi criado na pasta de backups."
            )

    def start_background_web_server(self):
        import threading
        def _run():
            try:
                import servidor_web
                servidor_web.run_server(host="0.0.0.0", port=5000, open_browser=False)
            except Exception as e:
                print(f"[*] Aviso: Servidor web já ativo ou porta ocupada: {e}")

        t = threading.Thread(target=_run, daemon=True)
        t.start()

    def action_open_web_server(self):
        import webbrowser
        webbrowser.open("http://127.0.0.1:5000/")

    def action_reset_sheet(self):
        reset_active_sheet_to_default()
        self.refresh_all_data()
        QMessageBox.information(self, "Planilha Descarregada", "A planilha ativa foi descarregada.")

    def action_download_model_excel(self):
        dest, _ = QFileDialog.getSaveFileName(self, "Salvar Modelo Excel", "Modelo_Itens_Cotacao.xlsx", "Planilhas Excel (*.xlsx)")
        if dest:
            try:
                generate_template_excel(dest)
                QMessageBox.information(self, "Sucesso", f"Modelo Excel salvo com sucesso em:\n{dest}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao salvar arquivo: {e}")

    def action_download_model_csv(self):
        dest, _ = QFileDialog.getSaveFileName(self, "Salvar Modelo CSV", "Modelo_Itens_Cotacao.csv", "Arquivos CSV (*.csv)")
        if dest:
            try:
                generate_template_csv(dest)
                QMessageBox.information(self, "Sucesso", f"Modelo CSV salvo com sucesso em:\n{dest}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao salvar arquivo: {e}")

    def action_open_store_link(self):
        curr_row = self.table_cotacoes.currentRow()
        if curr_row < 0:
            QMessageBox.information(
                self,
                "Nenhum Registro Selecionado",
                "Selecione uma cotação na tabela para acessar o endereço eletrônico do fornecedor."
            )
            return

        lote_item = self.table_cotacoes.item(curr_row, 0)
        if not lote_item:
            return
        try:
            target_lote = int(lote_item.text())
        except:
            target_lote = lote_item.text()

        quotes = getattr(self, "all_quotes_cache", [])
        matched = [q for q in quotes if q.get("lote") == target_lote]
        url = matched[0].get("url") if matched else None

        if url and url.startswith("http"):
            QDesktopServices.openUrl(QUrl(url))
        else:
            QMessageBox.warning(
                self,
                "Link Não Disponível",
                f"O item do Lote {target_lote} não possui um link de fonte comercial registrado."
            )

    def action_view_print_from_table(self):
        curr_row = self.table_cotacoes.currentRow()
        if curr_row < 0:
            QMessageBox.information(
                self,
                "Nenhum Registro Selecionado",
                "Selecione uma cotação na tabela para visualizar o comprovante oficial com carimbo."
            )
            return

        lote_item = self.table_cotacoes.item(curr_row, 0)
        if not lote_item:
            return
        try:
            target_lote = int(lote_item.text())
        except:
            target_lote = lote_item.text()

        found = False
        for i in range(self.combo_prints_lotes.count()):
            q_data = self.combo_prints_lotes.itemData(i)
            if q_data and q_data.get("lote") == target_lote:
                self.combo_prints_lotes.setCurrentIndex(i)
                self.tabs.setCurrentIndex(1)
                found = True
                break

        if not found:
            QMessageBox.information(
                self,
                "Comprovante Não Localizado",
                f"Ainda não foi gerado comprovante oficial com carimbo para o Lote {target_lote}.\n\n"
                "Inicie a pesquisa no painel lateral para coletar e homologar o item."
            )

    def action_open_store_link_ficha(self):
        quote_data = self.combo_prints_lotes.currentData()
        if quote_data and quote_data.get("url", "").startswith("http"):
            QDesktopServices.openUrl(QUrl(quote_data.get("url")))

    def action_open_image_external(self):
        quote_data = self.combo_prints_lotes.currentData()
        if quote_data and quote_data.get("arquivo_print") and os.path.exists(quote_data.get("arquivo_print")):
            try:
                os.startfile(quote_data.get("arquivo_print"))
            except Exception as e:
                QMessageBox.warning(self, "Aviso", f"Não foi possível abrir no visualizador: {e}")

    def action_browse_export_folder(self):
        chosen = QFileDialog.getExistingDirectory(self, "Selecione a Pasta de Destino")
        if chosen:
            target = os.path.join(chosen, "Cotacao_Pregao_SRP")
            self.edit_export_folder.setText(target)

    def action_export_bundle(self):
        folder = self.edit_export_folder.text().strip()
        if not folder:
            QMessageBox.warning(self, "Aviso", "Informe a pasta de destino.")
            return
        try:
            res = export_bundle_to_folder(folder)
            reply = QMessageBox.information(
                self,
                "Exportação Concluída com Sucesso",
                f"A planilha Excel e {res['prints_count']} comprovantes oficiais foram salvos na pasta:\n\n"
                f"{res['dest_folder']}\n\n"
                "Deseja abrir esta pasta no Windows Explorer agora?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                try:
                    os.startfile(res["dest_folder"])
                except:
                    pass
        except Exception as e:
            QMessageBox.critical(self, "Erro na Exportação", f"Erro: {e}")

    def action_export_zip(self):
        dest_zip, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Pacote Completo Compactado",
            "Pacote_Pesquisa_Precos_Pregao.zip",
            "Arquivo Compactado (*.zip)"
        )
        if dest_zip:
            try:
                create_export_zip_file(dest_zip)
                reply = QMessageBox.information(
                    self,
                    "Pacote ZIP Gerado com Sucesso",
                    f"Arquivo gerado em:\n{dest_zip}\n\nDeseja abrir a pasta onde o arquivo foi salvo?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                if reply == QMessageBox.Yes:
                    try:
                        os.startfile(os.path.dirname(dest_zip))
                    except:
                        pass
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao gerar ZIP: {e}")

    # CARREGAMENTO E ATIVAÇÃO DE NOVA PLANILHA
    def action_select_sheet_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecione a Planilha de Itens",
            "",
            "Planilhas Suportadas (*.xlsx *.xls *.csv);;Excel (*.xlsx *.xls);;CSV (*.csv)"
        )
        if not file_path:
            return

        try:
            df_raw = read_sheet_file(file_path)
            self.current_uploaded_df = df_raw
            self.current_uploaded_path = file_path
            self.lbl_selected_file_name.setText(f"Arquivo selecionado: <b>{os.path.basename(file_path)}</b> ({len(df_raw)} linhas)")

            auto_map = detect_column_mapping(df_raw)
            cols = list(df_raw.columns)
            all_lote_cols = ["(Gerar sequencial 1, 2, 3...)"] + cols

            self.combo_col_desc.clear()
            self.combo_col_lote.clear()

            self.combo_col_desc.addItems(cols)
            if auto_map.get("descricao") in cols:
                self.combo_col_desc.setCurrentText(auto_map.get("descricao"))

            self.combo_col_lote.addItems(all_lote_cols)
            if auto_map.get("lote") in cols:
                self.combo_col_lote.setCurrentText(auto_map.get("lote"))
            else:
                self.combo_col_lote.setCurrentIndex(0)

            self.widget_mapping.setVisible(True)

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Ler Planilha", f"Não foi possível ler o arquivo:\n{e}")

    def action_activate_selected_sheet(self):
        df_raw = getattr(self, "current_uploaded_df", None)
        fpath = getattr(self, "current_uploaded_path", "planilha.xlsx")
        if df_raw is None:
            return

        sel_desc = self.combo_col_desc.currentText()
        sel_lote = self.combo_col_lote.currentText()

        chosen_map = {
            "descricao": sel_desc,
            "lote": sel_lote if sel_lote != "(Gerar sequencial 1, 2, 3...)" else None
        }

        parsed_items, stats = parse_items_dataframe(df_raw, mapping=chosen_map)
        if not parsed_items:
            QMessageBox.warning(self, "Aviso", "Nenhum item válido foi encontrado na planilha com os campos selecionados.")
            return

        save_active_sheet(parsed_items, source_name=os.path.basename(fpath))
        self.widget_mapping.setVisible(False)
        self.lbl_selected_file_name.setText("Nenhum arquivo selecionado.")
        self.refresh_all_data()
        QMessageBox.information(
            self,
            "Planilha Ativada com Sucesso",
            f"A planilha '{os.path.basename(fpath)}' foi ativada com sucesso!\n\n"
            f"Total de {len(parsed_items)} itens prontos para cotação."
        )


# ==============================================================================
# PONTO DE ENTRADA DO ARQUIVO ÚNICO
# ==============================================================================
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Cotador Automatizado para Pregão SRP (Qt / CLI)")
    parser.add_argument("--start", type=int, default=None, help="Índice inicial do item (modo CLI)")
    parser.add_argument("--count", type=int, default=None, help="Quantidade de itens a cotar (modo CLI)")
    parser.add_argument("--reset", action="store_true", help="Limpa todos os dados salvos (modo CLI)")
    parser.add_argument("--cli", action="store_true", help="Força execução no modo terminal/CLI")
    parser.add_argument("--web", "--server", action="store_true", help="Inicia o servidor web local (Flask) no navegador")
    parser.add_argument("--port", type=int, default=5000, help="Porta do servidor web (padrão: 5000)")
    parser.add_argument("--host", default="127.0.0.1", help="Host do servidor web (padrão: 127.0.0.1)")

    args, unknown = parser.parse_known_args()

    if args.reset:
        reset_all_data(backup=True)
    elif args.web:
        import servidor_web
        servidor_web.run_server(host=args.host, port=args.port, open_browser=True)
    elif args.start is not None or args.count is not None or args.cli:
        s_idx = args.start if args.start is not None else 0
        c_qtd = args.count if args.count is not None else 3
        print(f"Executando modo lote CLI: início={s_idx}, quantidade={c_qtd}")
        items = load_sheet_items()
        slice_items = items[s_idx:s_idx + c_qtd]
        gemini_client = get_gemini_client()
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled"])
            existing = []
            if os.path.exists(RESULTS_JSON):
                try:
                    with open(RESULTS_JSON, "r", encoding="utf-8") as f:
                        existing = json.load(f)
                except: pass
            for it in slice_items:
                res = cotar_item_core(browser, gemini_client, it)
                if res:
                    existing = [e for e in existing if e.get("lote") != it.get("lote")]
                    existing.append(res)
                    save_and_export_results(existing)
            browser.close()
    else:
        app = QApplication(sys.argv)
        app.setStyle("Fusion")
        app.setFont(QFont("Segoe UI", 9))
        window = CotadorQtApp()
        window.show()
        sys.exit(app.exec())
