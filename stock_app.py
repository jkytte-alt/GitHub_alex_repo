#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""個人股票管理工具 - 買賣記錄 / 庫存樹狀圖(finviz風格) / 個股分析"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os, threading
import warnings

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl import load_workbook

from tkcalendar import Calendar

import matplotlib, matplotlib.ticker
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.font_manager as fm
from matplotlib.patches import Patch
from matplotlib.lines import Line2D

import squarify
import yfinance as yf
import math
import json as _json
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    from curl_cffi import requests as _cffi_req
    _CFFI_OK = True
except ImportError:
    _CFFI_OK = False

def _cffi_get_json(url, params=None, headers=None, timeout=15):
    """curl_cffi-based JSON GET that handles corporate DNS/firewall issues."""
    _hdrs = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/122.0.0.0',
        'Accept': 'application/json',
    }
    if headers:
        _hdrs.update(headers)
    if _CFFI_OK:
        r = _cffi_req.get(url, params=params, headers=_hdrs, timeout=timeout, verify=False)
        return _json.loads(r.content.decode('utf-8'))
    else:
        import urllib.request, urllib.parse, ssl
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        if params:
            url = url + '?' + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers=_hdrs)
        with urllib.request.urlopen(req, context=ctx, timeout=timeout) as resp:
            return _json.loads(resp.read().decode('utf-8'))

warnings.filterwarnings('ignore')

# ─── 路徑與欄位 ──────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, '股票紀錄.xlsx')
SHEET_NAME = '交易記錄'

COLUMNS = ['日期', '股票代號', '股票名稱', '分類', '買賣',
           '數量(股)', '價格(元)', '手續費(元)', '交易稅(元)', '淨金額(元)']
CATEGORIES = ['定期定額', '波段操作', '爸爸合資']

ETF_LIST = [
    # ── 市值型 ──────────────────────────────────────────────────────────────
    ('0050',   '元大台灣50'),
    ('006208', '富邦台灣50'),
    ('0051',   '元大中型100'),
    ('00733',  '富邦臺灣中小'),
    ('00850',  '元大MSCI台灣'),
    # ── 高股息 / 配息 ────────────────────────────────────────────────────────
    ('0056',   '元大高股息'),
    ('00878',  '國泰永續高股息'),
    ('00919',  '群益台灣精選高息'),
    ('00900',  '富邦特選高股息30'),
    ('00929',  '復華台灣科技優息'),
    ('00934',  '中信成長高股息'),
    ('00940',  '元大台灣價值高息'),
    ('00713',  '元大台灣高息低波'),
    ('00915',  '凱基優選高股息30'),
    ('00918',  '大華優利高填息30'),
    ('00939',  '統一台灣高息動能'),
    ('00943',  '兆豐永續高息等權'),
    ('00948',  '街口臺灣高息ETF'),
    ('00960',  '復華臺灣科技中小'),
    # ── 科技 / 半導體 ────────────────────────────────────────────────────────
    ('0052',   '富邦科技'),
    ('0053',   '元大電子'),
    ('00891',  '中信關鍵半導體'),
    ('00892',  '富邦台灣半導體'),
    ('00881',  '國泰台灣5G+'),
    ('00830',  '國泰費城半導體'),
    ('00912',  '中信台灣智慧50'),
    # ── ESG / 永續 ───────────────────────────────────────────────────────────
    ('00692',  '富邦公司治理'),
    ('00757',  '統一MSCI台灣ESG'),
    ('00896',  '中信綠能及電動車'),
    # ── 槓桿 / 反向 ──────────────────────────────────────────────────────────
    ('00631L', '元大台灣50正2'),
    ('00632R', '元大台灣50反1'),
    ('00671L', '富邦臺灣加權正2'),
    ('00672R', '富邦臺灣加權反1'),
    # ── 海外 ────────────────────────────────────────────────────────────────
    ('00646',  '元大S&P500'),
    ('00864B', '中信美國公債20年'),
    ('00679B', '元大美債20年'),
    ('00772B', '中信高評級公司債'),
]

SECTOR_ZH = {
    'technology':          '科技業',
    'financialServices':   '金融保險業',
    'communication':       '通訊服務',
    'industrials':         '工業',
    'healthcare':          '醫療保健',
    'consumerCyclical':    '非必需消費',
    'consumerDefensive':   '必需消費',
    'energy':              '能源',
    'basicMaterials':      '原物料',
    'realestate':          '不動產',
    'utilities':           '公用事業',
    'cash':                '現金',
    'stock':               '股票',
    'bond':                '債券',
    'other':               '其他',
}

FEE_RATE = 0.001425
TAX_RATE = 0.003

# ─── 深色主題色盤 ─────────────────────────────────────────────────────────────
C_BG      = '#1e1e1e'
C_PANEL   = '#252526'
C_INPUT   = '#2d2d2d'
C_FG      = '#cccccc'
C_FG2     = '#888888'
C_ACCENT  = '#0078d4'
C_BORDER  = '#3e3e3e'
C_BUY_FG  = '#4ec94e'
C_SELL_FG = '#f07070'
C_BUY_BG  = '#1a2e1a'
C_SELL_BG = '#2e1a1a'

# ─── 中文字型 ────────────────────────────────────────────────────────────────
def _setup_font():
    candidates = ['Microsoft JhengHei', 'Microsoft YaHei', 'SimHei', 'Noto Sans CJK TC']
    available  = {f.name for f in fm.fontManager.ttflist}
    for c in candidates:
        if c in available:
            return c
    for c in candidates:
        for name in available:
            if c.lower().replace(' ', '') in name.lower().replace(' ', ''):
                return name
    return 'DejaVu Sans'

CHART_FONT  = _setup_font()
BODY_FONT   = ('Microsoft JhengHei', 10)
HDR_FONT    = ('Microsoft JhengHei', 13, 'bold')
INPUT_FONT  = ('Microsoft JhengHei', 12)        # Tab 1 輸入區較大字體

plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.family']        = CHART_FONT

# ─── Excel 工具 ──────────────────────────────────────────────────────────────
def init_excel():
    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(COLUMNS)
        col_widths = [12, 10, 12, 10, 6, 10, 10, 12, 12, 14]
        for i, cell in enumerate(ws[1], 1):
            cell.font      = Font(bold=True, color='FFFFFF')
            cell.fill      = PatternFill('solid', fgColor='1e5fa0')
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[cell.column_letter].width = col_widths[i - 1]
        wb.save(EXCEL_PATH)


def _migrate_add_category():
    """若舊版 Excel 缺少「分類」欄，自動插入"""
    if not os.path.exists(EXCEL_PATH):
        return
    wb = load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]
    header = [cell.value for cell in ws[1]]
    if '分類' in header:
        return  # already up-to-date

    if '買賣' not in header:
        return

    insert_col = header.index('買賣') + 1   # 1-based column index
    max_col    = ws.max_column

    for r in range(1, ws.max_row + 1):
        for c in range(max_col, insert_col - 1, -1):
            ws.cell(row=r, column=c + 1).value = ws.cell(row=r, column=c).value
        ws.cell(row=r, column=insert_col).value = ('分類' if r == 1 else '未分類')

    wb.save(EXCEL_PATH)


def load_df() -> pd.DataFrame:
    init_excel()
    _migrate_add_category()
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, dtype={'股票代號': str})
        if df.empty or '日期' not in df.columns:
            return pd.DataFrame(columns=COLUMNS)
        df['日期'] = pd.to_datetime(df['日期'])
        if '分類' not in df.columns:
            df.insert(df.columns.get_loc('買賣'), '分類', '未分類')
        return df.sort_values('日期').reset_index(drop=True)
    except Exception as e:
        print(f'[load_df] {e}')
        return pd.DataFrame(columns=COLUMNS)


def save_row(row: dict):
    init_excel()
    _migrate_add_category()
    wb = load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]
    header = [cell.value for cell in ws[1]]
    values = []
    for col in header:
        v = row.get(col, '')
        if isinstance(v, datetime):
            v = v.date()
        values.append(v)
    ws.append(values)
    wb.save(EXCEL_PATH)


def rewrite_excel(df: pd.DataFrame):
    """以 DataFrame 完整重寫 Excel（用於刪除 / 修改後）"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(COLUMNS)
    col_widths = [12, 10, 12, 10, 6, 10, 10, 12, 12, 14]
    for i, cell in enumerate(ws[1], 1):
        cell.font      = Font(bold=True, color='FFFFFF')
        cell.fill      = PatternFill('solid', fgColor='1e5fa0')
        cell.alignment = Alignment(horizontal='center')
        ws.column_dimensions[cell.column_letter].width = col_widths[i - 1]
    for _, row in df.iterrows():
        values = []
        for col in COLUMNS:
            v = row.get(col, '')
            if hasattr(v, 'date') and callable(v.date):
                v = v.date()
            elif pd.isna(v):
                v = ''
            values.append(v)
        ws.append(values)
    wb.save(EXCEL_PATH)

# ─── 庫存計算 ────────────────────────────────────────────────────────────────
def calc_holdings(df: pd.DataFrame) -> dict:
    """Key = (股票代號, 分類)，同一股票不同分類分開計算"""
    holdings: dict = {}
    if df.empty:
        return holdings
    for _, r in df.iterrows():
        code = str(r['股票代號']).strip()
        cat  = str(r.get('分類', '未分類'))
        key  = (code, cat)
        qty, price, fee = float(r['數量(股)']), float(r['價格(元)']), float(r['手續費(元)'])
        if r['買賣'] == '買':
            if key not in holdings:
                holdings[key] = {'name': str(r['股票名稱']),
                                 'qty': 0.0, 'total_cost': 0.0, 'category': cat}
            holdings[key]['qty']        += qty
            holdings[key]['total_cost'] += qty * price + fee
        else:
            # 賣出：優先從同分類扣除；若找不到則退回同代碼任意一個持倉
            h = holdings.get(key)
            if h is None:
                for k, v in holdings.items():
                    if k[0] == code:
                        h = v
                        break
            if h and h['qty'] > 0:
                avg = h['total_cost'] / h['qty']
                h['total_cost'] -= avg * qty
                h['qty']        -= qty
    result = {}
    for key, h in holdings.items():
        if h['qty'] > 0.5:
            h['avg_cost'] = h['total_cost'] / h['qty']
            result[key] = h
    return result

# ─── 報價工具 ────────────────────────────────────────────────────────────────
def get_price(code: str):
    code = str(code).strip()
    for s in ['.TW', '.TWO', '']:
        try:
            hist = yf.Ticker(code + s).history(period='5d')
            if not hist.empty:
                return float(hist['Close'].iloc[-1]), code + s
        except Exception:
            pass
    return None, code


def get_price_on_date(code: str, date_str: str):
    code = str(code).strip()
    try:
        target = datetime.strptime(date_str, '%Y-%m-%d').date()
    except ValueError:
        return None, False, code
    end_str = (datetime.strptime(date_str, '%Y-%m-%d') + timedelta(days=7)).strftime('%Y-%m-%d')
    for s in ['.TW', '.TWO', '']:
        try:
            hist = yf.Ticker(code + s).history(
                start=date_str, end=end_str, auto_adjust=False)
            if hist.empty:
                continue
            close_col = 'Close' if 'Close' in hist.columns else hist.columns[3]
            return float(hist[close_col].iloc[0]), (hist.index[0].date() == target), code + s
        except Exception:
            pass
    return None, False, code


def get_history(code: str, start_date: str) -> 'pd.Series | None':
    """取得從 start_date 到今天的每日收盤價（pd.Series，index 為 DatetimeIndex）"""
    end_str = (datetime.today() + timedelta(days=1)).strftime('%Y-%m-%d')
    for s in ['.TW', '.TWO', '']:
        try:
            hist = yf.Ticker(code + s).history(
                start=start_date, end=end_str, auto_adjust=False)
            if not hist.empty:
                close_col = 'Close' if 'Close' in hist.columns else hist.columns[3]
                return hist[close_col]
        except Exception:
            pass
    return None


def _has_cjk(s: str) -> bool:
    return any('\u4e00' <= c <= '\u9fff' for c in str(s))


# ── 證交所 / 櫃買中心名稱快取 ─────────────────────────────────────────────────
_TWSE_NAMES_CACHE: dict[str, str] = {}
_NAMES_CACHE_PATH = os.path.join(BASE_DIR, '.stock_names_cache.json')

# ── 產業別快取 ────────────────────────────────────────────────────────────────
_TWSE_INDUSTRY_CACHE: dict[str, str] = {}
_INDUSTRY_CACHE_PATH = os.path.join(BASE_DIR, '.stock_industry_cache.json')

# TWSE 產業代碼 → 中文名稱（含 1 位與 2 位格式）
_TWSE_IND_CODE_MAP: dict[str, str] = {
    '1': '水泥工業',  '01': '水泥工業',
    '2': '食品工業',  '02': '食品工業',
    '3': '塑膠工業',  '03': '塑膠工業',
    '4': '紡織纖維',  '04': '紡織纖維',
    '5': '電機機械',  '05': '電機機械',
    '6': '電器電纜',  '06': '電器電纜',
    '8': '玻璃陶瓷',  '08': '玻璃陶瓷',
    '9': '造紙工業',  '09': '造紙工業',
    '10': '鋼鐵工業',
    '11': '橡膠工業',
    '12': '汽車工業',
    '14': '建材營造',
    '15': '航運業',
    '16': '觀光餐旅',
    '17': '金融保險業',
    '18': '貿易百貨',
    '20': '其他業',
    '21': '化學工業',
    '22': '生技醫療業',
    '23': '油電燃氣業',
    '24': '半導體業',
    '25': '電腦及週邊設備業',
    '26': '光電業',
    '27': '通信網路業',
    '28': '電子零組件業',
    '29': '電子通路業',
    '30': '資訊服務業',
    '31': '其他電子業',
    '32': '綠能環保',
    '33': '數位雲端',
    '34': '運動休閒',
    '35': '居家生活',
    '36': '電子工業',
}


def _load_twse_stock_names() -> dict[str, str]:
    """從證交所取得 {代號: 中文名稱} 對照表。
    優先從磁碟快取讀取；成功抓取後寫回磁碟；確保名稱不重置為 {}。
    來源順序: STOCK_DAY_ALL → t51sb01 → 磁碟快取
    """
    global _TWSE_NAMES_CACHE
    if _TWSE_NAMES_CACHE:
        return _TWSE_NAMES_CACHE

    names: dict[str, str] = {}
    twse_hdr = {
        'If-Modified-Since': 'Mon, 26 Jul 1997 05:00:00 GMT',
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache',
    }

    # 方法1: STOCK_DAY_ALL（當日全部上市股票，含ETF）────────────────────────
    try:
        data = _cffi_get_json(
            'https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL',
            headers=twse_hdr, timeout=15)
        for row in data:
            c = str(row.get('Code', '')).strip()
            n = str(row.get('Name', '')).strip()
            if c and n:
                names[c] = n
        if names:
            print(f'[TWSE names] STOCK_DAY_ALL {len(names)} 筆')
    except Exception as e:
        print(f'[TWSE names] STOCK_DAY_ALL 失敗: {type(e).__name__}')

    # 方法2: t51sb01 有價證券清單（備用）────────────────────────────────────
    if not names:
        try:
            data = _cffi_get_json(
                'https://openapi.twse.com.tw/v1/openData/t51sb01',
                headers=twse_hdr, timeout=15)
            for row in data:
                c = str(row.get('有價證券代號', '')).strip()
                n = str(row.get('有價證券名稱', '')).strip()
                if c and n:
                    names[c] = n
            if names:
                print(f'[TWSE names] t51sb01 {len(names)} 筆')
        except Exception as e:
            print(f'[TWSE names] t51sb01 失敗: {type(e).__name__}')

    # 方法2b: 補充 TPEx OTC 股票名稱（TWSE ISIN 查詢頁，Big5 HTML）──────────
    try:
        import re as _re
        _isin_url = 'https://isin.twse.com.tw/isin/C_public.jsp?strMode=4'
        if _CFFI_OK:
            _r = _cffi_req.get(_isin_url, timeout=15, verify=False,
                               headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/122.0.0.0'})
            _html = _r.content.decode('big5', errors='replace')
        else:
            import urllib.request, ssl as _ssl
            _ctx = _ssl.create_default_context()
            _ctx.check_hostname = False
            _ctx.verify_mode = _ssl.CERT_NONE
            with urllib.request.urlopen(_isin_url, context=_ctx, timeout=15) as _resp:
                _html = _resp.read().decode('big5', errors='replace')
        # 每筆格式: <td>6223　台灣聯合光纖</td>  (U+3000 全形空格分隔代號和名稱)
        _tpex_added = 0
        for _code, _name in _re.findall(r'<td>(\d{4,6})\u3000([^<\u3000\r\n]+)', _html):
            _code = _code.strip(); _name = _name.strip()
            if _code and _name and _code not in names:
                names[_code] = _name
                _tpex_added += 1
        if _tpex_added:
            print(f'[TPEX names] ISIN 補充 {_tpex_added} 筆 OTC 名稱')
    except Exception as _e:
        print(f'[TPEX names] ISIN 失敗: {type(_e).__name__}: {_e}')

    # 成功取得 → 寫磁碟快取
    if names:
        try:
            with open(_NAMES_CACHE_PATH, 'w', encoding='utf-8') as f:
                _json.dump(names, f, ensure_ascii=False)
        except Exception:
            pass
    else:
        # 方法3: 讀磁碟快取（上次成功的結果）───────────────────────────────
        try:
            with open(_NAMES_CACHE_PATH, encoding='utf-8') as f:
                names = _json.load(f)
            print(f'[TWSE names] 磁碟快取 {len(names)} 筆')
        except Exception:
            print('[TWSE names] 無法取得名稱資料（網路不通且無快取）')

    _TWSE_NAMES_CACHE = names
    return names


def _fetch_twse_industry_map() -> dict[str, str]:
    """回傳 {股票代號: 產業別} 對照表（TWSE 上市公司基本資料）。
    有磁碟快取，網路失敗時讀快取。
    """
    global _TWSE_INDUSTRY_CACHE
    if _TWSE_INDUSTRY_CACHE:
        return _TWSE_INDUSTRY_CACHE

    result: dict[str, str] = {}
    try:
        data = _cffi_get_json(
            'https://openapi.twse.com.tw/v1/openData/t187ap03_L',
            timeout=15)
        if data:
            first = data[0]
            # 動態找欄位名稱（防 API 改版）
            code_key = next((k for k in ['公司代號', '代號', 'Code'] if k in first), None)
            ind_key  = next((k for k in ['產業別', '行業別', 'Industry'] if k in first), None)
            if code_key and ind_key:
                for row in data:
                    c = str(row.get(code_key, '')).strip()
                    i = str(row.get(ind_key,  '')).strip()
                    if c and i:
                        # 若 API 回傳數字代碼，轉換成中文名稱
                        i = _TWSE_IND_CODE_MAP.get(i, i)
                        result[c] = i
                print(f'[TWSE industry] t187ap03_L {len(result)} 筆')
    except Exception as e:
        print(f'[TWSE industry] 抓取失敗: {type(e).__name__}')

    if result:
        _TWSE_INDUSTRY_CACHE = result
        try:
            with open(_INDUSTRY_CACHE_PATH, 'w', encoding='utf-8') as _f:
                _json.dump(result, _f, ensure_ascii=False)
        except Exception:
            pass
    else:
        try:
            with open(_INDUSTRY_CACHE_PATH, encoding='utf-8') as _f:
                cached = _json.load(_f)
            # 讀取快取後也補做代碼轉換（修正舊版快取）
            result = {c: _TWSE_IND_CODE_MAP.get(i, i) for c, i in cached.items()}
            _TWSE_INDUSTRY_CACHE = result
            print(f'[TWSE industry] 磁碟快取 {len(result)} 筆')
        except Exception:
            pass

    return result


def _fetch_twse_etf_holdings(etf_code: str) -> tuple[list[dict], str]:
    """抓 ETF 完整成分股＋持股比例。
    Returns (holdings, debug_msg)
    方法1: 證交所 etfFundHoldingData（DAILY/MONTHLY）
    方法2: Yahoo Finance topHoldings（正確權重，最多10檔）
    """
    code  = str(etf_code).strip()
    debug = []

    # ── 方法 1: 證交所 etfFundHoldingData（DAILY/MONTHLY，近10天）──────────
    twse_hdr = {
        'Accept':           'application/json, text/javascript, */*; q=0.01',
        'Accept-Language':  'zh-TW,zh;q=0.9',
        'X-Requested-With': 'XMLHttpRequest',
        'Referer': f'https://www.twse.com.tw/zh/ETF/fund/{code}',
    }
    for q_type in ('DAILY', 'MONTHLY'):
        for days_back in range(2, 6):   # 只試近4天；DNS或HTML失敗時快速跳過
            date = (datetime.now() - timedelta(days=days_back)).strftime('%Y%m%d')
            try:
                data = _cffi_get_json(
                    'https://www.twse.com.tw/fund/etfFundHoldingData',
                    params={'fund': code, 'type': q_type, 'date': date},
                    headers=twse_hdr, timeout=5)
                if not isinstance(data, dict):
                    continue
                stat = data.get('stat', '?')
                if stat != 'OK':
                    continue
                fields  = data.get('fields', [])
                records = data.get('data',   [])
                if not records:
                    continue
                holdings = []
                for rec in records:
                    row    = dict(zip(fields, rec))
                    s_code = str(next(
                        (row[k] for k in row if any(t in k for t in ('代號', '代碼'))), ''
                    )).strip()
                    s_name = str(next(
                        (row[k] for k in row if '名稱' in k), ''
                    )).strip()
                    w_raw  = next(
                        (row[k] for k in row if any(t in k for t in ('比例', 'percent', 'Percent'))), '0'
                    )
                    try:
                        weight = float(str(w_raw).replace('%', '').replace(',', ''))
                    except (ValueError, TypeError):
                        weight = 0.0
                    if s_code:
                        holdings.append({'code': s_code, 'name': s_name,
                                         'weight': weight, 'sym': s_code + '.TW'})
                if holdings:
                    debug.append(f'TWSE OK: {len(holdings)} 檔')
                    return holdings, ' | '.join(debug)
            except Exception as e:
                debug.append(f'TWSE({q_type},{date}): {type(e).__name__}')
                continue

    # ── 方法 2: Yahoo Finance topHoldings（正確權重，最多10檔）──────────────
    try:
        from yfinance.data import YfData as _YfData
        _yfdata = _YfData(session=None)
        for suffix in ['.TW', '.TWO']:
            try:
                r = _yfdata.cache_get(
                    url='https://query2.finance.yahoo.com/v10/finance/quoteSummary/' + code + suffix,
                    params={'modules': 'topHoldings'},
                    timeout=15)
                payload = _json.loads(r.content)
                result  = payload.get('quoteSummary', {}).get('result', [{}])
                if not result:
                    continue
                raw_h = result[0].get('topHoldings', {}).get('holdings', [])
                if not raw_h:
                    continue
                holdings = []
                for h in raw_h:
                    sym   = str(h.get('symbol', ''))
                    clean = sym.replace('.TWO', '').replace('.TW', '').strip()
                    name  = str(h.get('holdingName', clean))
                    pct   = h.get('holdingPercent', {})
                    w     = (pct.get('raw', 0) if isinstance(pct, dict) else float(pct or 0)) * 100
                    full  = sym if (sym.endswith('.TW') or sym.endswith('.TWO')) \
                            else clean + '.TW'
                    holdings.append({'code': clean, 'sym': full, 'name': name, 'weight': w})
                if holdings:
                    debug.append(f'Yahoo topHoldings: {len(holdings)} 檔')
                    return holdings, ' | '.join(debug)
            except Exception:
                pass
    except Exception as e:
        debug.append(f'Yahoo err: {type(e).__name__}')

    # ── 方法 3: MoneyDJ 完整成分股（HTML 解析，所有成分股）────────────────────
    import re as _re
    try:
        _mj_url = (f'https://www.moneydj.com/ETF/X/Basic/Basic0007B.xdjhtm'
                   f'?etfid={code}.TW')
        _mj_hdrs = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/122.0.0.0',
            'Accept': 'text/html,application/xhtml+xml',
            'Accept-Language': 'zh-TW,zh;q=0.9',
            'Referer': 'https://www.moneydj.com/',
        }
        if _CFFI_OK:
            _r = _cffi_req.get(_mj_url, headers=_mj_hdrs, timeout=20, verify=False)
            _html = _r.content.decode('latin-1')  # lossless; Big5 page
        else:
            import urllib.request as _ureq, ssl as _ssl
            _ctx = _ssl.create_default_context()
            _ctx.check_hostname = False
            _ctx.verify_mode = _ssl.CERT_NONE
            _req2 = _ureq.Request(_mj_url, headers=_mj_hdrs)
            with _ureq.urlopen(_req2, context=_ctx, timeout=20) as _resp:
                _html = _resp.read().decode('latin-1')

        # Extract code from href="...etfid=XXXX.TW..." then weight from col06
        # HTML: etfid=2330.TW...>NAME</a></td><td ...>63.00</td>
        _pat = _re.compile(
            r"etfid=(\d{4,6})\.TWO?[^'\"]*['\"][^>]*>[^<]*</a>"
            r"</td>\s*<td[^>]*>([\d.]+)</td>",
            _re.I
        )
        _matches = list(_pat.finditer(_html))

        holdings = []
        seen_codes: set = set()
        for _m in _matches:
            _s_code = _m.group(1).strip()
            _w_str  = _m.group(2)
            if _s_code in seen_codes:
                continue
            try:
                _weight = float(_w_str)
            except ValueError:
                _weight = 0.0
            if _weight <= 0 or _weight > 100:
                continue
            # Use TWSE cache for canonical Chinese name
            _name = _TWSE_NAMES_CACHE.get(_s_code, _s_code)
            holdings.append({
                'code': _s_code,
                'name': _name,
                'weight': _weight,
                'sym':  _s_code + '.TW',
            })
            seen_codes.add(_s_code)

        if holdings:
            debug.append(f'MoneyDJ: {len(holdings)} 檔')
            return holdings, ' | '.join(debug)
        else:
            debug.append('MoneyDJ: 0 檔（解析失敗）')
    except Exception as _e:
        debug.append(f'MoneyDJ err: {type(_e).__name__}')

    return [], ' | '.join(debug)


def fetch_etf_data(etf_code: str) -> tuple[str, list[dict], str]:
    """Fetch ETF name and all constituent stocks.
    Primary source: TWSE / FinMind. Fallback: yfinance funds_data.
    Returns (etf_display_name, components, debug_msg).
    components: list of {code, name, weight, price, change_pct}
    """
    code = str(etf_code).strip()

    # ── Step 1: TWSE 名稱表（含 ETF 自身名稱）───────────────────────────────
    name_table = _load_twse_stock_names()
    etf_name   = name_table.get(code, code)

    # ── Step 2: 成分股（TWSE 為主，Yahoo Finance topHoldings 為備）────────────
    raw_holdings, debug_msg = _fetch_twse_etf_holdings(code)
    print(f'[ETF DEBUG] {code}: {debug_msg}')

    if not raw_holdings:
        return etf_name, [], debug_msg

    # ── Step 3: 補齊中文名稱（優先用 TWSE 名稱表）───────────────────────────
    for h in raw_holdings:
        if not h.get('sym'):
            h['sym'] = h['code'] + '.TW'
        tw_name = name_table.get(h['code'])
        if tw_name:
            h['name'] = tw_name
        elif not _has_cjk(h.get('name', '')):
            h['name'] = h['code']

    # ── Step 4: 若持股比例全為 0，改用等權重（防 ZeroDivisionError）──────────
    total_w = sum(h['weight'] for h in raw_holdings)
    if total_w <= 0:
        eq = 100.0 / len(raw_holdings)
        for h in raw_holdings:
            h['weight'] = eq

    # ── Step 5: 批次抓即時股價＋漲跌幅 ────────────────────────────────────────
    all_syms:   list[str]        = [h['sym'] for h in raw_holdings]
    price_map:  dict[str, float] = {}
    change_map: dict[str, float] = {}
    try:
        hist = yf.download(all_syms, period='2d', auto_adjust=False,
                           progress=False, threads=True)
        if not hist.empty:
            if isinstance(hist.columns, pd.MultiIndex):
                close_df = hist['Close']
            elif 'Close' in hist.columns:
                close_df = hist[['Close']].rename(columns={'Close': all_syms[0]})
            else:
                close_df = pd.DataFrame()
            for sym in all_syms:
                if sym in close_df.columns:
                    prices = close_df[sym].dropna()
                    if len(prices) >= 2:
                        price_map[sym]  = float(prices.iloc[-1])
                        change_map[sym] = ((float(prices.iloc[-1]) - float(prices.iloc[-2]))
                                           / float(prices.iloc[-2]) * 100)
                    elif len(prices) == 1:
                        price_map[sym]  = float(prices.iloc[0])
                        change_map[sym] = 0.0
    except Exception:
        pass

    # ── Step 6: 組合結果 ──────────────────────────────────────────────────────
    components = []
    for h in raw_holdings:
        components.append({
            'code':       h['code'],
            'name':       h['name'],
            'weight':     h['weight'],
            'price':      price_map.get(h['sym']),
            'change_pct': change_map.get(h['sym'], 0.0),
        })
    components.sort(key=lambda x: -x['weight'])
    return etf_name, components, debug_msg


def fetch_etf_meta(etf_code: str) -> dict:
    """Fetch ETF AUM, yield, NAV, sector weights, asset allocation from Yahoo Finance.
    Returns dict with keys: total_assets, yield_pct, nav, sector_weights, asset_alloc.
    """
    code = str(etf_code).strip()
    try:
        from yfinance.data import YfData as _YfData
        _yfdata = _YfData(session=None)
        for suffix in ['.TW', '.TWO']:
            try:
                r = _yfdata.cache_get(
                    url=f'https://query2.finance.yahoo.com/v10/finance/quoteSummary/{code}{suffix}',
                    params={'modules': 'summaryDetail,etfProfile'},
                    timeout=15)
                payload = _json.loads(r.content)
                result  = payload.get('quoteSummary', {}).get('result', [{}])
                if not result:
                    continue
                res  = result[0]
                meta: dict = {}
                sd   = res.get('summaryDetail', {})
                for src_key, dst_key in [('totalAssets',         'total_assets'),
                                          ('yield',               'yield_pct'),
                                          ('navPrice',            'nav'),
                                          ('regularMarketPrice',  'price'),
                                          ('previousClose',       'prev_close')]:
                    val = sd.get(src_key, {})
                    meta[dst_key] = val.get('raw') if isinstance(val, dict) else (val or None)
                ep = res.get('etfProfile', {})
                meta['sector_weights'] = ep.get('sectorWeightings', [])
                meta['asset_alloc']    = ep.get('assetAllocations', [])
                if any(v for v in meta.values()):
                    return meta
            except Exception:
                pass
    except Exception:
        pass
    return {}


def get_stock_name(code: str) -> str | None:
    code = str(code).strip()
    for s in ['.TW', '.TWO', '']:
        try:
            info = yf.Ticker(code + s).info
            name = info.get('longName') or info.get('shortName')
            if name:
                return name
        except Exception:
            pass
    return None

# ─── 樹狀圖顏色（finviz 風格）────────────────────────────────────────────────
def pnl_color(pct: float, max_abs: float = 10.0) -> str:
    """台灣慣例：正報酬紅色，負報酬綠色。
    顏色深淺依據 pct 佔 max_abs 的比例線性插值，而非固定門檻。"""
    if max_abs <= 0:
        return '#3a3a3a'
    t = min(abs(pct) / max_abs, 1.0)   # 0.0（接近0%）→ 1.0（極值）
    # 中性灰 → 深紅 / 深綠
    if pct >= 0:
        r = int(0x3a + t * (0xcc - 0x3a))
        g = int(0x3a + t * (0x1a - 0x3a))
        b = int(0x3a + t * (0x1a - 0x3a))
    else:
        r = int(0x3a + t * (0x0e - 0x3a))
        g = int(0x3a + t * (0x7a - 0x3a))
        b = int(0x3a + t * (0x1e - 0x3a))
    return f'#{r:02x}{g:02x}{b:02x}'

# ─── 自製日期選擇器（避免 DateEntry 導覽時自動關閉的 bug）────────────────────
class DatePickerEntry(ttk.Frame):
    """Entry + 📅 按鈕，點擊後開啟獨立 Calendar Toplevel，切換月份/年份不會關閉"""

    def __init__(self, parent, initial_date=None, on_date_selected=None, **_kw):
        super().__init__(parent, style='TFrame')
        self._callback = on_date_selected

        # 初始日期
        d = initial_date or datetime.today()
        self._date = d.date() if isinstance(d, datetime) else d

        self._var = tk.StringVar(value=self._date.strftime('%Y-%m-%d'))

        entry = ttk.Entry(self, textvariable=self._var, width=13)
        entry.pack(side='left', ipady=2)
        entry.bind('<FocusOut>', self._parse_entry)
        entry.bind('<Return>',   self._parse_entry)

        ttk.Button(self, text='📅', width=3,
                   command=self._open_popup).pack(side='left', padx=(3, 0))

    # ── 手動輸入解析 ──────────────────────────────────────────────────────────
    def _parse_entry(self, _=None):
        try:
            self._date = datetime.strptime(self._var.get().strip(), '%Y-%m-%d').date()
        except ValueError:
            self._var.set(self._date.strftime('%Y-%m-%d'))   # 格式錯誤還原

    # ── 開啟日曆彈窗 ──────────────────────────────────────────────────────────
    def _open_popup(self):
        parent_top = self.winfo_toplevel()   # 可能是主視窗或 EditDialog

        popup = tk.Toplevel()
        popup.title('選擇日期')
        popup.configure(bg=C_BG)
        popup.resizable(False, False)
        popup.transient(parent_top)

        # 定位在 Entry 下方
        self.update_idletasks()
        px = self.winfo_rootx()
        py = self.winfo_rooty() + self.winfo_height() + 2
        popup.geometry(f'+{px}+{py}')

        cal = Calendar(
            popup,
            selectmode='day',
            year=self._date.year, month=self._date.month, day=self._date.day,
            date_pattern='yyyy-mm-dd',
            background=C_ACCENT,       foreground='white',
            headersbackground='#1c2333', headersforeground='#8ab4d4',
            normalbackground=C_PANEL,  normalforeground=C_FG,
            weekendbackground='#2a2a2a', weekendforeground=C_FG2,
            othermonthforeground='#555', othermonthbackground=C_BG,
            othermonthweforeground='#444', othermonthwebackground=C_BG,
            selectbackground=C_ACCENT, selectforeground='white',
            bordercolor=C_BORDER,
            font=('Microsoft JhengHei', 10),
        )
        cal.pack(padx=8, pady=8)

        def _confirm(_=None):
            date_str = cal.get_date()
            try:
                self._date = datetime.strptime(date_str, '%Y-%m-%d').date()
            except ValueError:
                pass
            self._var.set(self._date.strftime('%Y-%m-%d'))
            _close()
            if self._callback:
                self._callback()

        def _close():
            popup.grab_release()
            popup.destroy()
            # 若父視窗是 modal Toplevel，歸還 grab
            if isinstance(parent_top, tk.Toplevel):
                try:
                    parent_top.grab_set()
                except Exception:
                    pass

        popup.protocol('WM_DELETE_WINDOW', _close)

        btn_f = ttk.Frame(popup)
        btn_f.pack(pady=(0, 8))
        ttk.Button(btn_f, text='✅  確認', command=_confirm).pack(side='left', padx=5)
        ttk.Button(btn_f, text='✖  取消', command=_close).pack(side='left', padx=5)

        cal.bind('<Double-1>', _confirm)

        # 若父視窗有 grab，先暫時釋放再接管
        if isinstance(parent_top, tk.Toplevel):
            try:
                parent_top.grab_release()
            except Exception:
                pass
        popup.grab_set()
        popup.focus_set()

    # ── 公開介面（與 DateEntry 相容）─────────────────────────────────────────
    def get_date(self):
        """返回 datetime.date"""
        return self._date

    def set_date(self, d):
        if isinstance(d, datetime):
            d = d.date()
        elif isinstance(d, str):
            d = datetime.strptime(d, '%Y-%m-%d').date()
        self._date = d
        self._var.set(d.strftime('%Y-%m-%d'))


# ─── 主視窗 ──────────────────────────────────────────────────────────────────
class StockApp(tk.Tk):

    # sidebar nav items: (icon, label, page_index)
    _NAV_ITEMS = [
        ('🌳', '庫存樹狀圖'),
        ('📊', '個股分析'),
        ('📈', 'ETF分析'),
        ('📝', '買賣紀錄'),
    ]
    _SIDEBAR_EXP = 200   # expanded width px
    _SIDEBAR_COL = 54    # collapsed width px
    _SB_BG       = '#16161e'
    _SB_SEL      = '#1f2d45'
    _SB_HOV      = '#1c2030'

    def __init__(self):
        super().__init__()
        self.title('個人股票管理工具')
        self.geometry('1200x780')
        self.configure(bg=C_BG)
        self._apply_dark_theme()

        # ── 主框架：sidebar + content ─────────────────────────────────────────
        root_frame = tk.Frame(self, bg=C_BG)
        root_frame.pack(fill='both', expand=True)

        # Left sidebar
        self._sb_expanded = True
        self._sidebar = tk.Frame(root_frame, bg=self._SB_BG,
                                 width=self._SIDEBAR_EXP)
        self._sidebar.pack(side='left', fill='y')
        self._sidebar.pack_propagate(False)

        # Thin separator line
        tk.Frame(root_frame, bg='#2a2a3a', width=1).pack(side='left', fill='y')

        # Content area (stacked pages)
        self._content = tk.Frame(root_frame, bg=C_BG)
        self._content.pack(side='left', fill='both', expand=True)

        self.tab1 = tk.Frame(self._content, bg=C_BG)
        self.tab2 = tk.Frame(self._content, bg=C_BG)
        self.tab3 = tk.Frame(self._content, bg=C_BG)
        self.tab4 = tk.Frame(self._content, bg=C_BG)
        for f in (self.tab1, self.tab2, self.tab3, self.tab4):
            f.place(relx=0, rely=0, relwidth=1, relheight=1)

        self._build_sidebar()

        self._vars: dict[str, tk.StringVar] = {}
        self._build_tab1()
        self._build_tab2()
        self._build_tab3()
        self._build_tab4()

        self._current_page = -1
        self._show_page(0)

    # ── Sidebar ───────────────────────────────────────────────────────────────
    def _build_sidebar(self):
        SB = self._SB_BG

        # ── 頂部：Logo + 收合按鈕 ──────────────────────────────────────────
        top = tk.Frame(self._sidebar, bg=SB)
        top.pack(fill='x', padx=6, pady=(12, 6))

        self._sb_title = tk.Label(top, text='  股票管理', bg=SB,
                                   fg='#aecde8',
                                   font=('Microsoft JhengHei', 12, 'bold'),
                                   anchor='w')
        self._sb_title.pack(side='left', fill='x', expand=True)

        self._sb_toggle = tk.Button(top, text='◀', bg=SB, fg='#5a7fa8',
                                     font=('Microsoft JhengHei', 11),
                                     relief='flat', bd=0, cursor='hand2',
                                     activebackground=SB, activeforeground='#aecde8',
                                     command=self._toggle_sidebar)
        self._sb_toggle.pack(side='right', padx=4)

        # Divider
        tk.Frame(self._sidebar, bg='#252535', height=1).pack(fill='x',
                                                               padx=8, pady=4)

        # ── 導覽項目 ────────────────────────────────────────────────────────
        self._nav_rows   = []   # outer frame per item
        self._nav_labels = []   # text Label per item (hidden when collapsed)

        for i, (icon, label) in enumerate(self._NAV_ITEMS):
            row = tk.Frame(self._sidebar, bg=SB, cursor='hand2')
            row.pack(fill='x', padx=6, pady=2)

            icon_lbl = tk.Label(row, text=icon, bg=SB, fg=C_FG,
                                 font=('Segoe UI Emoji', 15),
                                 cursor='hand2', width=2)
            icon_lbl.pack(side='left', padx=(8, 4), pady=8)

            txt_lbl = tk.Label(row, text=label, bg=SB, fg='#9090a0',
                                font=('Microsoft JhengHei', 11),
                                anchor='w', cursor='hand2')
            txt_lbl.pack(side='left', fill='x', expand=True, pady=8)

            for w in (row, icon_lbl, txt_lbl):
                w.bind('<Button-1>', lambda e, n=i: self._show_page(n))
                w.bind('<Enter>',    lambda e, r=row, n=i: self._sb_hover(r, n, True))
                w.bind('<Leave>',    lambda e, r=row, n=i: self._sb_hover(r, n, False))

            self._nav_rows.append((row, icon_lbl, txt_lbl))
            self._nav_labels.append(txt_lbl)

    def _toggle_sidebar(self):
        self._sb_expanded = not self._sb_expanded
        if self._sb_expanded:
            self._sidebar.config(width=self._SIDEBAR_EXP)
            self._sb_title.pack(side='left', fill='x', expand=True)
            for _, _, txt in self._nav_rows:
                txt.pack(side='left', fill='x', expand=True, pady=8)
            self._sb_toggle.config(text='◀')
        else:
            self._sidebar.config(width=self._SIDEBAR_COL)
            self._sb_title.pack_forget()
            for _, _, txt in self._nav_rows:
                txt.pack_forget()
            self._sb_toggle.config(text='▶')

    def _sb_hover(self, row, idx, entering):
        if idx == self._current_page:
            return
        bg = self._SB_HOV if entering else self._SB_BG
        for w in [row] + list(row.winfo_children()):
            w.config(bg=bg)

    def _show_page(self, idx):
        pages = [self.tab2, self.tab3, self.tab4, self.tab1]
        pages[idx].tkraise()

        # Update nav highlight
        for i, (row, icon_lbl, txt_lbl) in enumerate(self._nav_rows):
            if i == idx:
                bg = self._SB_SEL
                fg_txt = C_FG
            else:
                bg = self._SB_BG
                fg_txt = '#9090a0'
            row.config(bg=bg)
            icon_lbl.config(bg=bg)
            txt_lbl.config(bg=bg, fg=fg_txt)

        prev = self._current_page
        self._current_page = idx

        # Side-effects
        if idx == 0 and prev != 0:
            self._draw_treemap()
        elif idx == 1 and prev != 1:
            self._update_stock_list()

    # ── 深色主題 ──────────────────────────────────────────────────────────────
    def _apply_dark_theme(self):
        style = ttk.Style()
        style.theme_use('clam')

        style.configure('.',
            background=C_BG, foreground=C_FG,
            fieldbackground=C_INPUT, troughcolor=C_PANEL,
            bordercolor=C_BORDER, darkcolor=C_PANEL, lightcolor=C_PANEL,
            font=BODY_FONT)

        style.configure('TFrame',     background=C_BG)
        style.configure('TLabel',     background=C_BG, foreground=C_FG)
        style.configure('Hdr.TLabel', background=C_BG, foreground=C_FG, font=HDR_FONT)

        style.configure('TLabelframe',
            background=C_BG, bordercolor=C_BORDER, relief='solid')
        style.configure('TLabelframe.Label',
            background=C_BG, foreground=C_ACCENT, font=BODY_FONT)

        style.configure('TNotebook',
            background=C_BG, bordercolor=C_BORDER, tabmargins=[2, 5, 0, 0])
        style.configure('TNotebook.Tab',
            background=C_PANEL, foreground=C_FG2,
            padding=[14, 6], font=('Microsoft JhengHei', 11, 'bold'))
        style.map('TNotebook.Tab',
            background=[('selected', C_BG),    ('active', '#303030')],
            foreground=[('selected', C_FG),     ('active', C_FG)])

        style.configure('TEntry',
            fieldbackground=C_INPUT, foreground=C_FG,
            insertcolor=C_FG, bordercolor=C_BORDER, relief='flat', padding=4)
        style.map('TEntry',
            fieldbackground=[('focus', '#3a3a3a')],
            bordercolor=[('focus', C_ACCENT)])

        style.configure('TCombobox',
            fieldbackground=C_INPUT, foreground=C_FG,
            background=C_PANEL, arrowcolor=C_FG,
            bordercolor=C_BORDER, relief='flat', padding=4)
        style.map('TCombobox',
            fieldbackground=[('readonly', C_INPUT)],
            foreground=[('readonly', C_FG)],
            selectbackground=[('readonly', C_ACCENT)],
            selectforeground=[('readonly', 'white')])

        style.configure('TButton',
            background=C_ACCENT, foreground='white',
            bordercolor=C_ACCENT, relief='flat', padding=[10, 5])
        style.map('TButton',
            background=[('active', '#1a90e0'), ('pressed', '#005ba1')])

        style.configure('Treeview',
            background=C_PANEL, foreground=C_FG,
            fieldbackground=C_PANEL, rowheight=26,
            font=('Microsoft JhengHei UI', 10),
            bordercolor=C_BORDER, relief='flat')
        style.configure('Treeview.Heading',
            background='#333333', foreground=C_FG,
            relief='flat', bordercolor=C_BORDER)
        style.map('Treeview',
            background=[('selected', C_ACCENT)],
            foreground=[('selected', 'white')])

        style.configure('TScrollbar',
            background=C_PANEL, troughcolor=C_BG,
            arrowcolor=C_FG2, bordercolor=C_BORDER)

        # 買賣 Combobox 顏色（全域小字）
        style.configure('Buy.TCombobox',  foreground=C_BUY_FG,  font=BODY_FONT)
        style.configure('Sell.TCombobox', foreground=C_SELL_FG, font=BODY_FONT)
        style.map('Buy.TCombobox',  fieldbackground=[('readonly', C_BUY_BG)])
        style.map('Sell.TCombobox', fieldbackground=[('readonly', C_SELL_BG)])

        # ── 輸入區專用樣式（INPUT_FONT 12pt）──────────────────────────────────
        style.configure('In.TLabelframe',
            background=C_BG, bordercolor=C_BORDER, relief='solid')
        style.configure('In.TLabelframe.Label',
            background=C_BG, foreground=C_ACCENT, font=INPUT_FONT)

        style.configure('In.TLabel',
            background=C_BG, foreground=C_FG, font=INPUT_FONT)

        style.configure('In.TEntry',
            fieldbackground=C_INPUT, foreground=C_FG,
            insertcolor=C_FG, bordercolor=C_BORDER,
            relief='flat', padding=5, font=INPUT_FONT)
        style.map('In.TEntry',
            fieldbackground=[('focus', '#3a3a3a')],
            bordercolor=[('focus', C_ACCENT)])

        style.configure('In.TCombobox',
            fieldbackground=C_INPUT, foreground=C_FG,
            background=C_PANEL, arrowcolor=C_FG,
            bordercolor=C_BORDER, relief='flat', padding=5, font=INPUT_FONT)
        style.map('In.TCombobox',
            fieldbackground=[('readonly', C_INPUT)],
            foreground=[('readonly', C_FG)],
            selectbackground=[('readonly', C_ACCENT)],
            selectforeground=[('readonly', 'white')])

        style.configure('In.TButton',
            background=C_ACCENT, foreground='white',
            bordercolor=C_ACCENT, relief='flat',
            padding=[12, 6], font=INPUT_FONT)
        style.map('In.TButton',
            background=[('active', '#1a90e0'), ('pressed', '#005ba1')])

        style.configure('Nav.TButton',
            background='#1e1e2e', foreground=C_FG,
            bordercolor='#2a2a3a', relief='flat',
            padding=[12, 7], font=('Microsoft JhengHei', 10))
        style.map('Nav.TButton',
            background=[('active', '#252538'), ('pressed', '#1a253a')],
            foreground=[('active', '#ffffff'), ('pressed', '#ffffff')])

        style.configure('In.TCheckbutton',
            background=C_INPUT, foreground=C_FG,
            font=INPUT_FONT, focuscolor=C_INPUT)
        style.map('In.TCheckbutton',
            background=[('active', C_INPUT)],
            foreground=[('active', C_FG)])

        # 買賣 Combobox（輸入區大字）
        style.configure('InBuy.TCombobox',  foreground=C_BUY_FG,  font=INPUT_FONT)
        style.configure('InSell.TCombobox', foreground=C_SELL_FG, font=INPUT_FONT)
        style.map('InBuy.TCombobox',  fieldbackground=[('readonly', C_BUY_BG)])
        style.map('InSell.TCombobox', fieldbackground=[('readonly', C_SELL_BG)])

    # ── Tab 切換 ──────────────────────────────────────────────────────────────
    def _on_tab(self, event):
        pass  # replaced by _show_page

    # ═══════════════════════════════════════════════════════════════════════════
    # Tab 1：買賣紀錄
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_tab1(self):
        f = self.tab1

        box = ttk.LabelFrame(f, text='  新增交易  ', padding=14, style='In.TLabelframe')
        box.pack(fill='x', padx=14, pady=(12, 6))

        lkw = dict(sticky='e', padx=(12, 4), pady=8)
        ekw = dict(sticky='w', padx=(0, 18), pady=8)

        today = datetime.today()

        # Row 0: 日期 | 股票代號
        ttk.Label(box, text='日期', style='In.TLabel').grid(row=0, column=0, **lkw)
        self._date_entry = DatePickerEntry(
            box, on_date_selected=self._on_date_change)
        self._date_entry.grid(row=0, column=1, **ekw)

        ttk.Label(box, text='股票代號', style='In.TLabel').grid(row=0, column=2, **lkw)
        self._vars['code'] = tk.StringVar()
        self._code_entry = ttk.Combobox(box, textvariable=self._vars['code'],
                                         width=20, style='In.TCombobox')
        self._code_entry.grid(row=0, column=3, **ekw)
        self._code_entry.bind('<FocusOut>',           self._on_code_focusout)
        self._code_entry.bind('<Return>',
            lambda e: (self._on_code_focusout(e), self._name_entry.focus()))
        self._code_entry.bind('<<ComboboxSelected>>', self._on_code_selected)

        # Row 1: 股票名稱 | 分類
        ttk.Label(box, text='股票名稱', style='In.TLabel').grid(row=1, column=0, **lkw)
        self._vars['name'] = tk.StringVar()
        self._name_entry = ttk.Combobox(box, textvariable=self._vars['name'],
                                         width=20, style='In.TCombobox')
        self._name_entry.grid(row=1, column=1, **ekw)
        self._name_entry.bind('<FocusOut>',           self._on_name_focusout)
        self._name_entry.bind('<<ComboboxSelected>>', self._on_name_selected)

        ttk.Label(box, text='投資分類', style='In.TLabel').grid(row=1, column=2, **lkw)
        self._vars['category'] = tk.StringVar(value=CATEGORIES[0])
        ttk.Combobox(box, textvariable=self._vars['category'],
                     values=CATEGORIES, state='readonly', width=18,
                     style='In.TCombobox').grid(row=1, column=3, **ekw)

        # Row 2: 買/賣 | 數量
        ttk.Label(box, text='買 / 賣', style='In.TLabel').grid(row=2, column=0, **lkw)
        self._vars['side'] = tk.StringVar(value='買')
        self._side_combo = ttk.Combobox(
            box, textvariable=self._vars['side'],
            values=['買', '賣'], state='readonly', width=18,
            style='InBuy.TCombobox')
        self._side_combo.grid(row=2, column=1, **ekw)
        self._vars['side'].trace_add('write', self._on_side_change)

        self._qty_label_var = tk.StringVar(value='數量 (股)')
        ttk.Label(box, textvariable=self._qty_label_var,
                  style='In.TLabel').grid(row=2, column=2, **lkw)
        self._vars['qty'] = tk.StringVar()
        _qty_frame = ttk.Frame(box)
        _qty_frame.grid(row=2, column=3, **ekw)
        ttk.Entry(_qty_frame, textvariable=self._vars['qty'],
                  width=14, style='In.TEntry').pack(side='left')
        self._lot_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(_qty_frame, text='整張', variable=self._lot_var,
                        style='In.TCheckbutton',
                        command=self._on_lot_toggle).pack(side='left', padx=(8, 0))

        # Row 3: 成交價格 | 手續費
        ttk.Label(box, text='成交價格 (元)', style='In.TLabel').grid(row=3, column=0, **lkw)
        self._vars['price'] = tk.StringVar()
        _price_frame = ttk.Frame(box)
        _price_frame.grid(row=3, column=1, **ekw)
        ttk.Entry(_price_frame, textvariable=self._vars['price'],
                  width=14, style='In.TEntry').pack(side='left')
        _tick_btn_frame = tk.Frame(_price_frame, bg=C_INPUT)
        _tick_btn_frame.pack(side='left', padx=(4, 0))
        tk.Button(_tick_btn_frame, text='▲', font=('Microsoft JhengHei', 8, 'bold'),
                  bg=C_INPUT, fg=C_FG, activebackground=C_BORDER, relief='flat',
                  bd=0, width=2, cursor='hand2',
                  command=lambda: self._adjust_price(+1)).pack(side='top')
        tk.Button(_tick_btn_frame, text='▼', font=('Microsoft JhengHei', 8, 'bold'),
                  bg=C_INPUT, fg=C_FG, activebackground=C_BORDER, relief='flat',
                  bd=0, width=2, cursor='hand2',
                  command=lambda: self._adjust_price(-1)).pack(side='top')

        ttk.Label(box, text='手續費 (留空自動)', style='In.TLabel').grid(row=3, column=2, **lkw)
        self._vars['fee'] = tk.StringVar()
        ttk.Entry(box, textvariable=self._vars['fee'],
                  width=20, style='In.TEntry').grid(row=3, column=3, **ekw)

        # Row 4: 交易稅（賣出時顯示）
        self._tax_label = ttk.Label(box, text='交易稅 (留空自動)', style='In.TLabel')
        self._tax_label.grid(row=4, column=2, **lkw)
        self._vars['tax'] = tk.StringVar()
        self._tax_entry = ttk.Entry(box, textvariable=self._vars['tax'],
                                    width=20, style='In.TEntry')
        self._tax_entry.grid(row=4, column=3, **ekw)
        # 買入時隱藏交易稅欄位
        if self._vars['side'].get() == '買':
            self._tax_label.grid_remove()
            self._tax_entry.grid_remove()

        # Row 5: 狀態列 + 按鈕
        self._status_var = tk.StringVar(value='')
        ttk.Label(box, textvariable=self._status_var,
                  foreground=C_FG2, font=('Microsoft JhengHei', 10)
                  ).grid(row=5, column=0, columnspan=3, sticky='w', padx=14)

        btn_f = ttk.Frame(box)
        btn_f.grid(row=5, column=3, sticky='e', padx=(0, 4), pady=(6, 0))
        ttk.Button(btn_f, text='✅  新增交易', style='Nav.TButton',
                   command=self._add_tx).pack(side='left', padx=4)
        ttk.Button(btn_f, text='🔄  重新整理', style='Nav.TButton',
                   command=self._refresh_table).pack(side='left', padx=4)

        # ── 歷史紀錄 ──────────────────────────────────────────────────────────
        list_box = ttk.LabelFrame(f, text='  歷史交易紀錄  ', padding=8)
        list_box.pack(fill='both', expand=True, padx=14, pady=(0, 12))

        # 主表格：日期 ~ 淨金額，套用買/賣配色
        self._sort_col = None
        self._sort_asc = True
        self._pnl_gen  = 0
        # 總計 ≤ 755px，確保淨金額不需橫向捲動 (P&L 178 + vsb 17 = 195, 可用 ~955-195=760)
        _main_widths = [84, 66, 78, 70, 44, 64, 72, 70, 68, 86]
        self._tree = ttk.Treeview(list_box, columns=COLUMNS, show='headings', height=11)
        for col, w in zip(COLUMNS, _main_widths):
            self._tree.heading(col, text=col,
                               command=lambda c=col: self._sort_by_col(c))
            self._tree.column(col, width=w, anchor='center', stretch=False)
        self._tree.tag_configure('buy',  foreground=C_BUY_FG)
        self._tree.tag_configure('sell', foreground=C_SELL_FG)
        self._tree.tag_configure('row_even', background='#252526')
        self._tree.tag_configure('row_odd',  background='#202030')

        # 損益子表格：獨立配色，紅底=獲利，綠底=虧損
        _PNL_COLS = ['損益(元)', '損益率(%)']
        _pnl_widths = [92, 86]
        self._tree_pnl = ttk.Treeview(list_box, columns=_PNL_COLS, show='headings',
                                       height=11, selectmode='none')
        for col, w in zip(_PNL_COLS, _pnl_widths):
            self._tree_pnl.heading(col, text=col)
            self._tree_pnl.column(col, width=w, anchor='center', stretch=False,
                                   minwidth=w)
        self._tree_pnl.tag_configure('pnl_gain', background='#6b0000', foreground='#ffffff')
        self._tree_pnl.tag_configure('pnl_loss', background='#0a3d0a', foreground='#ffffff')
        self._tree_pnl.tag_configure('row_even', background='#252526')
        self._tree_pnl.tag_configure('row_odd',  background='#202030')

        def _ysync(*args):
            self._tree.yview(*args)
            self._tree_pnl.yview(*args)
        vsb = ttk.Scrollbar(list_box, orient='vertical', command=_ysync)
        hsb = ttk.Scrollbar(list_box, orient='horizontal', command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree_pnl.configure(yscrollcommand=vsb.set)

        def _on_wheel(event):
            delta = -1 * (event.delta // 120)
            self._tree.yview_scroll(delta, 'units')
            self._tree_pnl.yview_scroll(delta, 'units')
            return 'break'
        self._tree.bind('<MouseWheel>', _on_wheel)
        self._tree_pnl.bind('<MouseWheel>', _on_wheel)

        self._tree.grid(row=0, column=0, sticky='nsew')
        self._tree_pnl.grid(row=0, column=1, sticky='ns')
        vsb.grid(row=0, column=2, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        list_box.rowconfigure(0, weight=1)
        list_box.columnconfigure(0, weight=1)

        # ── 操作按鈕列 ────────────────────────────────────────────────────────
        btn_row = ttk.Frame(list_box)
        btn_row.grid(row=2, column=0, columnspan=3, sticky='w', pady=(6, 2))
        ttk.Button(btn_row, text='✏️  修改選取', style='Nav.TButton',
                   command=self._edit_selected).pack(side='left', padx=(0, 6))
        ttk.Button(btn_row, text='🗑️  刪除選取', style='Nav.TButton',
                   command=self._delete_selected).pack(side='left')

        # 右鍵選單
        self._ctx_menu = tk.Menu(self, tearoff=0, bg=C_PANEL, fg=C_FG,
                                  activebackground=C_ACCENT, activeforeground='white',
                                  font=BODY_FONT, bd=0)
        self._ctx_menu.add_command(label='✏️  修改', command=self._edit_selected)
        self._ctx_menu.add_command(label='🗑️  刪除', command=self._delete_selected)
        self._tree.bind('<Button-3>', self._show_ctx_menu)
        self._tree.bind('<Double-1>', lambda e: self._edit_selected())

        self._refresh_table()
        self._refresh_stock_options()

    # ── 股票代號 / 名稱快選選項 ────────────────────────────────────────────────
    def _refresh_stock_options(self):
        """從歷史紀錄更新 Tab 1 代號 / 名稱的快選清單"""
        df = load_df()
        if df.empty:
            return
        # 依最近使用排序（DataFrame 已按日期排序，取 unique 保留最後出現順序）
        codes = list(dict.fromkeys(df['股票代號'].astype(str).tolist()))[::-1]
        names = list(dict.fromkeys(df['股票名稱'].astype(str).tolist()))[::-1]
        self._code_entry['values'] = codes
        self._name_entry['values'] = names

    # ── 代號 / 名稱下拉選取後的處理 ───────────────────────────────────────────
    def _on_code_selected(self, _=None):
        """從下拉選單選取代號：清空名稱後觸發自動帶入"""
        self._vars['name'].set('')
        self._on_code_focusout()

    def _on_name_selected(self, _=None):
        """從下拉選單選取名稱：清空代號後觸發反向查詢"""
        self._vars['code'].set('')
        self._on_name_focusout()

    # ── 事件 ──────────────────────────────────────────────────────────────────
    def _on_side_change(self, *_):
        s = self._vars['side'].get()
        self._side_combo.configure(
            style='InBuy.TCombobox' if s == '買' else 'InSell.TCombobox')
        if s == '賣':
            self._tax_label.grid()
            self._tax_entry.grid()
        else:
            self._vars['tax'].set('')
            self._tax_label.grid_remove()
            self._tax_entry.grid_remove()

    def _on_lot_toggle(self):
        if self._lot_var.get():
            self._qty_label_var.set('數量 (張)')
        else:
            self._qty_label_var.set('數量 (股)')

    def _on_date_change(self, _=None):
        code = self._vars['code'].get().strip().upper()
        if code:
            self._trigger_price_lookup(code)

    def _on_code_focusout(self, _=None):
        code = self._vars['code'].get().strip().upper()
        if not code:
            return
        self._vars['code'].set(code)
        name = self._hist_name(code)
        if name:
            if not self._vars['name'].get():
                self._vars['name'].set(name)
            self._trigger_price_lookup(code)
        else:
            self._set_status('🔍 查詢股票資訊中…')
            threading.Thread(target=self._bg_name, args=(code,), daemon=True).start()

    def _on_name_focusout(self, _=None):
        name = self._vars['name'].get().strip()
        if not name or self._vars['code'].get().strip():
            return
        code = self._hist_code(name)
        if code:
            self._vars['code'].set(code)
            self._set_status(f'✅ 從歷史找到代號：{code}')
            self._trigger_price_lookup(code)

    # ── 台股跳檔輔助 ─────────────────────────────────────────────────────────
    @staticmethod
    def _tick_size(price: float) -> float:
        return 0.01 if price < 10 else 0.05

    def _adjust_price(self, direction: int):
        """direction: +1 上調一檔, -1 下調一檔
        先對齊到最近的 tick 邊界，再往指定方向移動一檔。
        例：100.02 → 上 → 100.05；100.05 → 上 → 100.10
        """
        try:
            p = float(self._vars['price'].get() or 0)
        except ValueError:
            return
        tick = self._tick_size(p)
        eps = 1e-9
        if direction == 1:
            # 先floor到當前tick，再加一格
            p = math.floor(p / tick + eps) * tick + tick
        else:
            # 先ceil到當前tick，再減一格
            p = math.ceil(p / tick - eps) * tick - tick
        p = round(max(p, tick), 2)
        self._vars['price'].set(f'{p:.2f}')

    def _trigger_price_lookup(self, code: str):
        date_str = self._date_entry.get_date().strftime('%Y-%m-%d')
        self._set_status('🔍 查詢歷史收盤價…')
        threading.Thread(target=self._bg_price,
                         args=(code, date_str), daemon=True).start()

    def _bg_name(self, code: str):
        name = get_stock_name(code)
        def _ui():
            if name:
                if not self._vars['name'].get():
                    self._vars['name'].set(name)
                self._set_status(f'✅ 股票名稱：{name}')
            else:
                self._set_status('⚠ 查無名稱，請手動填入')
            self._trigger_price_lookup(code)
        self.after(0, _ui)

    def _bg_price(self, code: str, date_str: str):
        price, exact, ticker = get_price_on_date(code, date_str)
        def _ui():
            if price is not None:
                self._vars['price'].set(f'{price:.2f}')
                if exact:
                    self._set_status(f'✅ {date_str} 收盤價：{price:.2f} 元（{ticker}）')
                else:
                    self._set_status(
                        f'⚠ {date_str} 非交易日，填入最近收盤價 {price:.2f} 元（{ticker}）')
            else:
                self._set_status('⚠ 查無歷史價格，請手動填入')
        self.after(0, _ui)

    def _hist_name(self, code: str) -> str | None:
        df = load_df()
        rows = df[df['股票代號'].astype(str) == code]
        return str(rows['股票名稱'].iloc[0]) if not rows.empty else None

    def _hist_code(self, name: str) -> str | None:
        df = load_df()
        rows = df[df['股票名稱'].astype(str).str.contains(name, na=False)]
        return str(rows['股票代號'].iloc[0]) if not rows.empty else None

    def _set_status(self, msg: str):
        self.after(0, lambda: self._status_var.set(msg))

    def _add_tx(self):
        try:
            date_str = self._date_entry.get_date().strftime('%Y-%m-%d')
            code     = self._vars['code'].get().strip().upper()
            name     = self._vars['name'].get().strip()
            cat      = self._vars['category'].get()
            side     = self._vars['side'].get()
            qty_s    = self._vars['qty'].get().strip()
            price_s  = self._vars['price'].get().strip()
            fee_s    = self._vars['fee'].get().strip()
            tax_s    = self._vars['tax'].get().strip()

            if not code or not name:
                raise ValueError('請填寫股票代號和名稱')
            if not qty_s or not price_s:
                raise ValueError('請填寫數量和價格')

            tx_date = datetime.strptime(date_str, '%Y-%m-%d')
            qty, price = float(qty_s), float(price_s)
            if self._lot_var.get():
                qty *= 1000          # 整張：1 張 = 1000 股
            if qty <= 0 or price <= 0:
                raise ValueError('數量和價格必須大於 0')

            gross = qty * price
            fee   = float(fee_s) if fee_s else max(20, round(gross * FEE_RATE))
            tax   = float(tax_s) if tax_s else (round(gross * TAX_RATE) if side == '賣' else 0)
            net   = round(gross - fee - tax) if side == '賣' else -round(gross + fee)

            save_row({
                '日期': tx_date, '股票代號': code, '股票名稱': name,
                '分類': cat, '買賣': side,
                '數量(股)': int(qty), '價格(元)': price,
                '手續費(元)': int(fee), '交易稅(元)': tax, '淨金額(元)': net,
            })
            self._refresh_table()
            self._refresh_stock_options()
            messagebox.showinfo('新增成功',
                f'[{cat}]  {side} {code} {name}\n'
                f'{int(qty):,} 股 @ {price:.2f} 元\n'
                f'手續費 {fee:,.0f}  稅 {tax:,}  淨額 {net:,}')
            for k in ('qty', 'price', 'fee', 'tax'):
                self._vars[k].set('')
            self._status_var.set('')

        except ValueError as e:
            messagebox.showerror('輸入錯誤', str(e))
        except Exception as e:
            messagebox.showerror('錯誤', str(e))

    def _show_ctx_menu(self, event):
        item = self._tree.identify_row(event.y)
        if item:
            self._tree.selection_set(item)
            self._ctx_menu.post(event.x_root, event.y_root)

    def _sort_by_col(self, col: str):
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        self._refresh_table()

    def _refresh_table(self):
        for item in self._tree.get_children():
            self._tree.delete(item)
        df = load_df()

        # 同步清空損益子表格
        for item in self._tree_pnl.get_children():
            self._tree_pnl.delete(item)

        # 更新欄位標題排序指示符
        for c in COLUMNS:
            label = c + ('  ▲' if self._sort_asc else '  ▼') if c == self._sort_col else c
            self._tree.heading(c, text=label)

        # 依選取欄位排序（保留 iid = 原始行索引，供 edit/delete 使用）
        if self._sort_col and self._sort_col in df.columns:
            df = df.sort_values(self._sort_col, ascending=self._sort_asc,
                                kind='stable')

        # 股票名稱欄自動寬度
        if not df.empty:
            max_chars = df['股票名稱'].astype(str).str.len().max()
            name_w = max(78, min(160, int(max_chars * 10) + 10))
            self._tree.column('股票名稱', width=name_w)

        for pos, (orig_idx, r) in enumerate(df.iterrows()):
            row_tag = 'row_even' if pos % 2 == 0 else 'row_odd'
            vals = [
                r['日期'].strftime('%Y-%m-%d'),
                r['股票代號'], r['股票名稱'],
                r.get('分類', ''),
                r['買賣'],
                f"{int(r['數量(股)']):,}",
                f"{r['價格(元)']:.2f}",
                f"{int(r['手續費(元)']):,}",
                f"{int(r['交易稅(元)']):,}",
                f"{int(r['淨金額(元)']):,}",
            ]
            side_tag = 'buy' if r['買賣'] == '買' else 'sell'
            self._tree.insert('', 'end', iid=str(orig_idx), values=vals,
                              tags=(side_tag, row_tag))
            self._tree_pnl.insert('', 'end', iid=str(orig_idx), values=('—', '—'),
                                  tags=(row_tag,))

        # 背景執行緒非同步抓即時報價並填入損益
        self._pnl_gen += 1
        gen = self._pnl_gen
        meta = df[['股票代號', '買賣', '價格(元)', '數量(股)']].copy()
        threading.Thread(target=self._bg_fill_pnl,
                         args=(meta, gen), daemon=True).start()

    def _bg_fill_pnl(self, meta_df, gen):
        prices = {}
        for code in meta_df['股票代號'].unique():
            p, _ = get_price(str(code))
            prices[str(code)] = p
        self.after(0, lambda: self._apply_pnl_to_table(meta_df, prices, gen))

    def _apply_pnl_to_table(self, meta_df, prices, gen):
        if gen != self._pnl_gen:
            return  # stale update, discard
        for orig_idx, r in meta_df.iterrows():
            iid = str(orig_idx)
            if not self._tree_pnl.exists(iid):
                continue
            if r['買賣'] != '買':
                continue
            cur = prices.get(str(r['股票代號']))
            if not cur:
                continue
            trade_p = float(r['價格(元)'])
            qty = float(r['數量(股)'])
            pnl_amt = (cur - trade_p) * qty
            pnl_pct = (cur / trade_p - 1) * 100 if trade_p else 0
            sign = '+' if pnl_amt >= 0 else ''
            pnl_tag = 'pnl_gain' if pnl_amt >= 0 else 'pnl_loss'
            self._tree_pnl.item(iid,
                                values=(f'{sign}{pnl_amt:,.0f}', f'{sign}{pnl_pct:.2f}%'),
                                tags=(pnl_tag,))

    def _delete_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo('提示', '請先選取一筆紀錄')
            return
        idx  = int(sel[0])
        vals = self._tree.item(sel[0], 'values')
        date_str, code, name = vals[0], vals[1], vals[2]
        if not messagebox.askyesno('確認刪除',
                f'確定要刪除以下紀錄？\n\n'
                f'{date_str}  {code} {name}  {vals[4]}  {vals[5]}股 @ {vals[6]}元',
                icon='warning'):
            return
        df = load_df()
        df = df.drop(index=idx).reset_index(drop=True)
        rewrite_excel(df)
        self._refresh_table()

    def _edit_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo('提示', '請先選取一筆紀錄')
            return
        idx = int(sel[0])
        df  = load_df()
        row = df.iloc[idx]
        EditDialog(self, idx, row, on_save=self._on_edit_save)

    def _on_edit_save(self, idx: int, new_row: dict):
        df = load_df()
        for col, val in new_row.items():
            df.at[idx, col] = val
        rewrite_excel(df)
        self._refresh_table()
        self._refresh_stock_options()

    # ═══════════════════════════════════════════════════════════════════════════
    # Tab 2：庫存樹狀圖（finviz 風格）
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_tab2(self):
        f = self.tab2

        ctrl = ttk.Frame(f)
        ctrl.pack(fill='x', padx=14, pady=8)
        ttk.Label(ctrl, text='庫存現值分佈', style='Hdr.TLabel').pack(side='left')
        self._tm_status = tk.StringVar(value='切換至此頁面自動更新 | 也可手動點擊更新')
        ttk.Label(ctrl, textvariable=self._tm_status,
                  foreground=C_FG2, font=('Microsoft JhengHei', 9)).pack(side='left', padx=14)
        ttk.Button(ctrl, text='💾  另存圖檔', style='Nav.TButton', command=self._save_treemap).pack(side='right', padx=(4, 0))
        ttk.Button(ctrl, text='🔄  更新報價', style='Nav.TButton', command=self._draw_treemap).pack(side='right')

        # ── 整體庫存摘要列 ────────────────────────────────────────────────────
        summary_bar = tk.Frame(f, bg=C_PANEL)
        summary_bar.pack(fill='x', padx=8, pady=(0, 4))

        def _card(parent, title):
            fr = tk.Frame(parent, bg='#1a1a2e', padx=14, pady=6)
            fr.pack(side='left', fill='x', expand=True, padx=3)
            tk.Label(fr, text=title, bg='#1a1a2e',
                     fg='#6a8faf', font=('Microsoft JhengHei', 8)).pack(anchor='w')
            var = tk.StringVar(value='—')
            lbl = tk.Label(fr, textvariable=var, bg='#1a1a2e',
                           fg=C_FG, font=('Microsoft JhengHei', 11, 'bold'))
            lbl.pack(anchor='w')
            return var, lbl

        self._sv_count,  self._sl_count  = _card(summary_bar, '持股標的')
        self._sv_mktval, self._sl_mktval = _card(summary_bar, '總市值')
        self._sv_cost,   self._sl_cost   = _card(summary_bar, '總成本')
        self._sv_pnl,    self._sl_pnl    = _card(summary_bar, '損益金額')
        self._sv_pnlpct, self._sl_pnlpct = _card(summary_bar, '損益率')

        self._tm_fig = plt.Figure(figsize=(9.5, 5.5), dpi=100, facecolor='#111111')
        self._tm_canvas = FigureCanvasTkAgg(self._tm_fig, master=f)
        w = self._tm_canvas.get_tk_widget()
        w.configure(bg='#111111')
        w.pack(fill='both', expand=True, padx=8, pady=(0, 8))

        self._tm_rects     = []      # stock block rects for hover detection
        self._tm_cat_rects = []      # category rects for hover detection
        self._tm_tooltip   = None    # tooltip Toplevel
        self._tm_ax        = None    # saved axes reference for coord transform
        self._tm_drill_cat = None    # None=全覽, str=鑽入的分類名稱
        self._tm_canvas.mpl_connect('motion_notify_event', self._on_tm_motion)
        self._tm_canvas.mpl_connect('button_press_event',  self._on_tm_click)
        self._tm_canvas.mpl_connect('figure_leave_event',  lambda e: self._hide_tooltip())

    # ── treemap click: 鑽入分類 / 返回全覽 ──────────────────────────────────────
    def _on_tm_click(self, event):
        if event.inaxes is None or self._tm_ax is None:
            return
        xd, yd = event.xdata, event.ydata

        # 鑽入模式：點任意處返回全覽
        if self._tm_drill_cat is not None:
            self._tm_drill_cat = None
            self._draw_treemap()
            return

        # 全覽模式：點分類標題列（header bar）進入鑽入
        for c in self._tm_cat_rects:
            hdr_y = c['cy'] + c['ch'] - c['hdr_h']
            if (c['cx'] <= xd <= c['cx'] + c['cw'] and
                    hdr_y <= yd <= c['cy'] + c['ch']):
                self._tm_drill_cat = c['cat_name']
                self._draw_treemap()
                return

    # ── treemap hover helpers ──────────────────────────────────────────────────
    def _on_tm_motion(self, event):
        if event.inaxes is None or self._tm_ax is None:
            self._hide_tooltip()
            return
        xd, yd = event.xdata, event.ydata
        widget = self._tm_canvas.get_tk_widget()
        sx = widget.winfo_rootx() + int(event.x) + 18
        sy = widget.winfo_rooty() + int(widget.winfo_height() - event.y) + 18

        # 先嘗試股票方塊
        for r in self._tm_rects:
            if r['rx'] <= xd <= r['rx'] + r['rw'] and r['ry'] <= yd <= r['ry'] + r['rh']:
                self._show_tooltip(sx, sy, r)
                return

        # 再嘗試分類區域
        for c in self._tm_cat_rects:
            if c['cx'] <= xd <= c['cx'] + c['cw'] and c['cy'] <= yd <= c['cy'] + c['ch']:
                self._show_cat_tooltip(sx, sy, c)
                return

        self._hide_tooltip()

    def _show_tooltip(self, sx, sy, data):
        if self._tm_tooltip is None or not self._tm_tooltip.winfo_exists():
            tip = tk.Toplevel(self)
            tip.overrideredirect(True)
            tip.attributes('-topmost', True)
            tip.configure(bg='#1e1e1e')
            # thin border frame
            border = tk.Frame(tip, bg='#3e3e3e', padx=1, pady=1)
            border.pack(fill='both', expand=True)
            inner = tk.Frame(border, bg='#252526', padx=10, pady=8)
            inner.pack(fill='both', expand=True)
            self._tm_tip_widgets = {}
            for key in ('title', 'price', 'qty', 'mktval', 'avgcost', 'pnl'):
                lbl = tk.Label(inner, bg='#252526',
                               font=('Microsoft JhengHei', 10), anchor='w')
                lbl.pack(fill='x')
                self._tm_tip_widgets[key] = lbl
            self._tm_tooltip = tip

        w = self._tm_tip_widgets
        code  = data['code']
        name  = data['name']
        cat   = data['category']
        price = data['price']
        qty   = data['qty']
        avg   = data['avg_cost']
        val   = data['value']
        pct   = data['pnl_pct']
        pnl_a = (price - avg) * qty if price else None

        w['title'].config(text=f'{code}  {name}  ({cat})',
                          fg='#8ab4d4', font=('Microsoft JhengHei', 11, 'bold'))
        w['price'].config(text=f'現價：{"—" if not price else f"{price:,.2f} 元"}',
                          fg='#cccccc')
        w['qty'].config(text=f'持股：{qty:,.0f} 股', fg='#cccccc')
        w['mktval'].config(text=f'市值：{val:,.0f} 元', fg='#cccccc')
        w['avgcost'].config(text=f'均價：{avg:,.2f} 元', fg='#cccccc')
        pnl_color_tip = '#f07070' if pct >= 0 else '#4ec94e'
        sign = '+' if pct >= 0 else ''
        pnl_txt = (f'損益：{sign}{pnl_a:,.0f} 元  ({sign}{pct:.2f}%)'
                   if pnl_a is not None else '損益：—')
        w['pnl'].config(text=pnl_txt, fg=pnl_color_tip)

        # keep tooltip on screen
        self._tm_tooltip.geometry(f'+{sx}+{sy}')
        self._tm_tooltip.deiconify()

    def _save_treemap(self):
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            title='儲存樹狀圖',
            defaultextension='.png',
            filetypes=[('PNG 圖片', '*.png'), ('JPEG 圖片', '*.jpg'), ('所有檔案', '*.*')],
            initialfile='treemap.png')
        if path:
            self._tm_fig.savefig(path, dpi=150, bbox_inches='tight',
                                 facecolor=self._tm_fig.get_facecolor())
            self._tm_status.set(f'已儲存：{path}')

    def _save_analysis(self):
        from tkinter import filedialog
        sel = self._stock_var.get().strip()
        default_name = sel.split()[0] if sel and not sel.startswith('──') else 'analysis'
        path = filedialog.asksaveasfilename(
            title='儲存個股分析圖',
            defaultextension='.png',
            filetypes=[('PNG 圖片', '*.png'), ('JPEG 圖片', '*.jpg'), ('所有檔案', '*.*')],
            initialfile=f'{default_name}.png')
        if path:
            self._an_fig.savefig(path, dpi=150, bbox_inches='tight',
                                 facecolor=self._an_fig.get_facecolor())

    def _show_cat_tooltip(self, sx, sy, cat):
        if self._tm_tooltip is None or not self._tm_tooltip.winfo_exists():
            tip = tk.Toplevel(self)
            tip.overrideredirect(True)
            tip.attributes('-topmost', True)
            tip.configure(bg='#1e1e1e')
            border = tk.Frame(tip, bg='#3e3e3e', padx=1, pady=1)
            border.pack(fill='both', expand=True)
            inner = tk.Frame(border, bg='#252526', padx=10, pady=8)
            inner.pack(fill='both', expand=True)
            self._tm_tip_widgets = {}
            for key in ('title', 'price', 'qty', 'mktval', 'avgcost', 'pnl'):
                lbl = tk.Label(inner, bg='#252526',
                               font=('Microsoft JhengHei', 10), anchor='w')
                lbl.pack(fill='x')
                self._tm_tip_widgets[key] = lbl
            self._tm_tooltip = tip

        w    = self._tm_tip_widgets
        sign = '+' if cat['pnl_amt'] >= 0 else ''
        pnl_fg = '#f07070' if cat['pnl_amt'] >= 0 else '#4ec94e'
        w['title'].config(text=cat['cat_name'],
                          fg='#aecde8', font=('Microsoft JhengHei', 11, 'bold'))
        w['price'].config(text=f"持股標的：{cat['count']} 檔", fg='#cccccc')
        w['qty'].config(  text=f"市值：{cat['value']:,.0f} 元",  fg='#cccccc')
        w['mktval'].config(text=f"成本：{cat['cost']:,.0f} 元",  fg='#cccccc')
        w['avgcost'].config(text='', fg='#cccccc')
        w['pnl'].config(
            text=f"損益：{sign}{cat['pnl_amt']:,.0f} 元  ({sign}{cat['pnl_pct']:.2f}%)",
            fg=pnl_fg)
        self._tm_tooltip.geometry(f'+{sx}+{sy}')
        self._tm_tooltip.deiconify()

    def _hide_tooltip(self):
        if self._tm_tooltip and self._tm_tooltip.winfo_exists():
            self._tm_tooltip.withdraw()

    def _draw_treemap(self):
        try:
            self._draw_treemap_impl()
        except Exception as e:
            import traceback; traceback.print_exc()
            self._tm_status.set(f'繪製錯誤：{e}')

    def _draw_treemap_impl(self):
        TREE_BG = '#111111'
        SEP_COL = '#111111'

        df       = load_df()
        holdings = calc_holdings(df)

        self._tm_rects     = []
        self._tm_cat_rects = []
        self._hide_tooltip()
        self._tm_fig.clear()
        self._tm_fig.patch.set_facecolor(TREE_BG)

        # 主畫布（保留底部 5% 給狀態列）
        ax = self._tm_fig.add_axes([0, 0.04, 1, 0.96])
        ax.set_facecolor(TREE_BG)
        ax.set_xlim(0, 100)
        ax.set_ylim(0, 100)
        ax.axis('off')
        self._tm_ax = ax

        # 底部狀態軸
        sax = self._tm_fig.add_axes([0, 0, 1, 0.04])
        sax.set_facecolor('#0d0d0d')
        sax.axis('off')

        if not holdings:
            for sv in (self._sv_count, self._sv_mktval, self._sv_cost,
                       self._sv_pnl, self._sv_pnlpct):
                sv.set('—')
            ax.text(50, 50, '目前沒有庫存', ha='center', va='center',
                    color='#888', fontsize=14, fontfamily=CHART_FONT)
            self._tm_canvas.draw()
            return

        # 依分類分組
        groups: dict[str, list] = {}
        no_price_list: list[str] = []
        total_val = 0.0

        for (code, _cat), h in holdings.items():
            p, _ = get_price(code)
            cat  = h.get('category', '未分類')
            avg  = h['avg_cost']
            if p:
                val = h['qty'] * p
                pnl = (p - avg) / avg * 100
            else:
                val = h['total_cost']
                pnl = 0.0
                no_price_list.append(code)
            total_val += val
            groups.setdefault(cat, []).append({
                'code': code, 'name': h['name'],
                'value': val, 'pnl_pct': pnl, 'price': p,
                'avg_cost': avg, 'qty': h['qty'],
                'total_cost': h['total_cost'],
            })

        # ── 更新摘要列 ────────────────────────────────────────────────────────
        total_cost    = sum(h['total_cost'] for h in holdings.values())
        pnl_amt       = total_val - total_cost
        pnl_pct_total = pnl_amt / total_cost * 100 if total_cost else 0
        pnl_fg        = '#f07070' if pnl_amt >= 0 else '#4ec94e'
        sign          = '+' if pnl_amt >= 0 else ''
        self._sv_count .set(f'{len(holdings)} 檔')
        self._sv_mktval.set(f'{total_val:,.0f} 元')
        self._sv_cost  .set(f'{total_cost:,.0f} 元')
        self._sv_pnl   .set(f'{sign}{pnl_amt:,.0f} 元')
        self._sv_pnlpct.set(f'{sign}{pnl_pct_total:.2f}%')
        self._sl_pnl   .config(fg=pnl_fg)
        self._sl_pnlpct.config(fg=pnl_fg)

        # 動態計算報酬率範圍，作為顏色插值基準
        all_pcts = [s['pnl_pct'] for stocks in groups.values() for s in stocks]
        _max_abs = max((abs(p) for p in all_pcts), default=10.0) or 10.0

        # 各分類依市值降冪排列
        for cat in groups:
            groups[cat].sort(key=lambda x: -x['value'])

        # Pixel conversion factors for font sizing (needed by both drill-down and full view)
        _dpi       = self._tm_fig.dpi
        _fig_w_px  = self._tm_fig.get_size_inches()[0] * _dpi
        _fig_h_px  = self._tm_fig.get_size_inches()[1] * _dpi
        _px_per_ux = _fig_w_px / 100
        _px_per_uy = _fig_h_px * 0.96 / 100
        _pt_per_px = 72 / _dpi
        GAP = 0.35

        # ── 鑽入模式：只顯示單一分類 ──────────────────────────────────────────
        if self._tm_drill_cat and self._tm_drill_cat in groups:
            _drill_name  = self._tm_drill_cat
            _drill_stocks = groups[_drill_name]

            # 返回提示列（頂部 5%）
            BACK_H = 5.5
            back_y = 100 - BACK_H
            ax.add_patch(plt.Rectangle((0, back_y), 100, BACK_H,
                facecolor='#1a253a', edgecolor='none', zorder=10))
            ax.text(2, back_y + BACK_H / 2,
                    '← 點擊任意處返回全覽',
                    ha='left', va='center', color='#6a9fcf',
                    fontsize=9, fontfamily=CHART_FONT, zorder=11)
            _drill_pct = sum(s['value'] for s in _drill_stocks) / total_val * 100 if total_val else 0
            ax.text(50, back_y + BACK_H / 2,
                    f'{_drill_name}  {_drill_pct:.1f}%',
                    ha='center', va='center', color='#aecde8',
                    fontsize=11, fontweight='bold', fontfamily=CHART_FONT, zorder=11)

            # 股票佔滿剩餘區域
            _di_y = 0
            _di_h = back_y
            _di_x = 0
            _di_w = 100
            _stock_vals = [s['value'] for s in _drill_stocks]
            if _stock_vals and _di_h > 0:
                _norm = squarify.normalize_sizes(_stock_vals, _di_w, _di_h)
                _raw  = squarify.squarify(_norm, _di_x, _di_y, _di_w, _di_h)
                _srects = [{'x': r['x'],
                             'y': _di_y + _di_h - (r['y'] - _di_y) - r['dy'],
                             'dx': r['dx'], 'dy': r['dy']} for r in _raw]
                for _stk, _sr in zip(_drill_stocks, _srects):
                    _sx, _sy, _sw, _sh = _sr['x'], _sr['y'], _sr['dx'], _sr['dy']
                    _rx, _ry = _sx + GAP / 2, _sy + GAP / 2
                    _rw, _rh = _sw - GAP, _sh - GAP
                    if _rw <= 0 or _rh <= 0:
                        continue
                    ax.add_patch(plt.Rectangle(
                        (_rx, _ry), _rw, _rh,
                        facecolor=pnl_color(_stk['pnl_pct'], _max_abs),
                        edgecolor=SEP_COL, linewidth=0.4, zorder=1))
                    self._tm_rects.append({
                        'rx': _rx, 'ry': _ry, 'rw': _rw, 'rh': _rh,
                        'code':     _stk['code'],  'name':     _stk['name'],
                        'value':    _stk['value'], 'price':    _stk['price'],
                        'pnl_pct':  _stk['pnl_pct'], 'avg_cost': _stk['avg_cost'],
                        'qty':      _stk['qty'],   'category': _drill_name,
                    })
                    _min_dim = min(_rw, _rh)
                    if _min_dim < 2:
                        continue
                    _rw_px = _rw * _px_per_ux
                    _rh_px = _rh * _px_per_uy
                    _chars  = max(len(_stk['code']), 4)
                    _fs_by_w = 0.58 * _rw_px * _pt_per_px / (_chars * 0.60)
                    _fs_by_h = 0.65 * _rh_px * _pt_per_px / 2.4
                    _fs_code = max(min(_fs_by_w, _fs_by_h, 44), 6)
                    _fs_sub  = max(_fs_code * 0.62, 5)
                    _pnl_str = f"{'+' if _stk['pnl_pct'] >= 0 else ''}{_stk['pnl_pct']:.2f}%"
                    _cx_t   = _rx + _rw / 2
                    _mid_y  = _ry + _rh * 0.50
                    _fs_code_uy = _fs_code / _pt_per_px / _px_per_uy
                    _fs_sub_uy  = _fs_sub  / _pt_per_px / _px_per_uy
                    _gap = (_fs_code_uy + _fs_sub_uy) * 0.55
                    if _min_dim >= 5:
                        ax.text(_cx_t, _mid_y + _gap * 0.5, _stk['code'],
                                ha='center', va='center', color='white',
                                fontsize=_fs_code, fontweight='bold',
                                fontfamily=CHART_FONT, clip_on=True, zorder=4)
                        ax.text(_cx_t, _mid_y - _gap * 0.5, _pnl_str,
                                ha='center', va='center', color='white',
                                fontsize=_fs_sub, fontfamily=CHART_FONT,
                                clip_on=True, zorder=4)
                    else:
                        ax.text(_cx_t, _mid_y, _stk['code'],
                                ha='center', va='center', color='white',
                                fontsize=_fs_code, fontweight='bold',
                                fontfamily=CHART_FONT, clip_on=True, zorder=4)

            # 同步摘要列為此分類數據
            _d_val  = sum(s['value']      for s in _drill_stocks)
            _d_cost = sum(s['total_cost'] for s in _drill_stocks)
            _d_pnl  = _d_val - _d_cost
            _d_pct  = _d_pnl / _d_cost * 100 if _d_cost else 0
            _d_sign = '+' if _d_pnl >= 0 else ''
            _d_fg   = '#f07070' if _d_pnl >= 0 else '#4ec94e'
            self._sv_count .set(f'{len(_drill_stocks)} 檔')
            self._sv_mktval.set(f'{_d_val:,.0f} 元')
            self._sv_cost  .set(f'{_d_cost:,.0f} 元')
            self._sv_pnl   .set(f'{_d_sign}{_d_pnl:,.0f} 元')
            self._sv_pnlpct.set(f'{_d_sign}{_d_pct:.2f}%')
            self._sl_pnl   .config(fg=_d_fg)
            self._sl_pnlpct.config(fg=_d_fg)

            warn = f'  ⚠ {", ".join(no_price_list)} 無即時報價' if no_price_list else ''
            sax.text(0.5, 0.5,
                     f'{_drill_name}：{len(_drill_stocks)} 檔  ·  '
                     f'市值 {_d_val:,.0f} 元{warn}',
                     ha='center', va='center', color='#888',
                     fontsize=8, fontfamily=CHART_FONT, transform=sax.transAxes)
            self._tm_canvas.draw()
            return

        cat_names  = list(groups.keys())
        cat_values = [sum(s['value'] for s in groups[c]) for c in cat_names]

        norm_cats  = squarify.normalize_sizes(cat_values, 100, 100)
        _cat_raw   = squarify.squarify(norm_cats, 0, 0, 100, 100)
        # 翻轉 y：squarify 從 y=0（畫面底部）填起，翻轉後最大方塊在左上
        cat_rects  = [{'x': r['x'], 'y': 100 - r['y'] - r['dy'],
                        'dx': r['dx'], 'dy': r['dy']} for r in _cat_raw]

        for cat_name, crect in zip(cat_names, cat_rects):
            cx, cy, cw, ch = crect['x'], crect['y'], crect['dx'], crect['dy']
            stocks = groups[cat_name]

            # 分類標題列（實色橫條，Finviz 風格）
            # 以像素為單位決定高度，再換算回 data-unit，字體隨之縮放
            _target_hdr_px = 20                          # 希望的標題高度 px
            _max_hdr_frac  = 0.40                        # 最多佔分類高度的 40%
            hdr_h_px = min(_target_hdr_px, ch * _px_per_uy * _max_hdr_frac)
            hdr_h    = hdr_h_px / _px_per_uy
            hdr_y    = cy + ch - hdr_h
            fs_hdr   = min(10, max(6, hdr_h_px * _pt_per_px * 0.62))
            ax.add_patch(plt.Rectangle(
                (cx, hdr_y), cw, hdr_h,
                facecolor='#1a1a2e', edgecolor='none', linewidth=0, zorder=6, clip_on=True))
            cat_pct = sum(s['value'] for s in stocks) / total_val * 100 if total_val else 0
            ax.text(cx + 0.8, hdr_y + hdr_h / 2,
                    f'{cat_name}  {cat_pct:.1f}%',
                    ha='left', va='center', color='#aecde8',
                    fontsize=fs_hdr, fontweight='bold', fontfamily=CHART_FONT,
                    clip_on=True, zorder=7)

            # 分類外框
            ax.add_patch(plt.Rectangle(
                (cx, cy), cw, ch,
                facecolor='none', edgecolor='#2a2a2a', linewidth=2.0, zorder=8))

            # 股票填滿標題列以下的區域
            inner_y = cy
            inner_h = ch - hdr_h
            if inner_h < 1 or not stocks:
                continue

            # 儲存分類 rect 供 hover 偵測
            cat_cost = sum(s['total_cost'] for s in stocks)
            cat_val  = sum(s['value']      for s in stocks)
            self._tm_cat_rects.append({
                'cx': cx, 'cy': cy, 'cw': cw, 'ch': ch,
                'cat_name': cat_name,
                'count':    len(stocks),
                'value':    cat_val,
                'cost':     cat_cost,
                'pnl_amt':  cat_val - cat_cost,
                'pnl_pct':  (cat_val - cat_cost) / cat_cost * 100 if cat_cost else 0,
                'hdr_h':    hdr_h,
            })

            stock_vals = [s['value'] for s in stocks]
            if cw >= inner_h:
                # 橫式容器：squarify 後翻轉 y，最大方塊在左上
                norm_stocks = squarify.normalize_sizes(stock_vals, cw, inner_h)
                _sr_raw = squarify.squarify(norm_stocks, cx, inner_y, cw, inner_h)
                srects = [{'x': r['x'],
                           'y': inner_y + inner_h - (r['y'] - inner_y) - r['dy'],
                           'dx': r['dx'], 'dy': r['dy']} for r in _sr_raw]
            else:
                # 直式容器：轉 90° 後翻轉 y，方塊更接近正方形且左上最大
                norm_stocks = squarify.normalize_sizes(stock_vals, inner_h, cw)
                srects_t = squarify.squarify(norm_stocks, 0, 0, inner_h, cw)
                srects = [{'x': cx + r['y'],
                           'y': inner_y + inner_h - r['x'] - r['dx'],
                           'dx': r['dy'], 'dy': r['dx']} for r in srects_t]

            for stock, sr in zip(stocks, srects):
                sx, sy, sw, sh = sr['x'], sr['y'], sr['dx'], sr['dy']
                rx, ry = sx + GAP / 2, sy + GAP / 2
                rw, rh = sw - GAP, sh - GAP
                if rw <= 0 or rh <= 0:
                    continue

                ax.add_patch(plt.Rectangle(
                    (rx, ry), rw, rh,
                    facecolor=pnl_color(stock['pnl_pct'], _max_abs),
                    edgecolor=SEP_COL, linewidth=0.4, zorder=1))

                # 儲存 rect 供 hover 偵測
                self._tm_rects.append({
                    'rx': rx, 'ry': ry, 'rw': rw, 'rh': rh,
                    'code':     stock['code'],
                    'name':     stock['name'],
                    'value':    stock['value'],
                    'price':    stock['price'],
                    'pnl_pct':  stock['pnl_pct'],
                    'avg_cost': stock['avg_cost'],
                    'qty':      stock['qty'],
                    'category': cat_name,
                })

                # 字體依方塊像素大小計算（方塊越大=占比越高=字越大）
                min_dim = min(rw, rh)
                if min_dim < 2:
                    continue

                rw_px   = rw * _px_per_ux
                rh_px   = rh * _px_per_uy
                chars   = max(len(stock['code']), 4)
                # code font: fill ~58% of block width, or fit 2 lines in 65% of height
                fs_by_w = 0.58 * rw_px * _pt_per_px / (chars * 0.60)
                fs_by_h = 0.65 * rh_px * _pt_per_px / 2.4
                fs_code = max(min(fs_by_w, fs_by_h, 44), 6)
                fs_sub  = max(fs_code * 0.62, 5)

                pnl_str = f"{'+' if stock['pnl_pct'] >= 0 else ''}{stock['pnl_pct']:.2f}%"
                cx_t    = rx + rw / 2
                mid_y   = ry + rh * 0.50

                # Convert font height to data-unit gap for tight line spacing
                fs_code_uy = fs_code / _pt_per_px / _px_per_uy
                fs_sub_uy  = fs_sub  / _pt_per_px / _px_per_uy
                gap        = (fs_code_uy + fs_sub_uy) * 0.55  # spacing between baselines

                if min_dim >= 5:
                    ax.text(cx_t, mid_y + gap * 0.5, stock['code'],
                            ha='center', va='center', color='white',
                            fontsize=fs_code, fontweight='bold',
                            fontfamily=CHART_FONT, clip_on=True, zorder=4)
                    ax.text(cx_t, mid_y - gap * 0.5, pnl_str,
                            ha='center', va='center', color='white',
                            fontsize=fs_sub, fontfamily=CHART_FONT,
                            clip_on=True, zorder=4)
                else:
                    ax.text(cx_t, mid_y, stock['code'],
                            ha='center', va='center', color='white',
                            fontsize=fs_code, fontweight='bold',
                            fontfamily=CHART_FONT, clip_on=True, zorder=4)


        # 狀態文字
        warn = f'  ⚠ {", ".join(no_price_list)} 無即時報價' if no_price_list else ''
        sax.text(0.5, 0.5,
                 f'估值合計 {total_val:,.0f} 元  ·  '
                 f'{datetime.now().strftime("%Y-%m-%d %H:%M")} 更新{warn}',
                 ha='center', va='center', color='#666',
                 fontsize=8, fontfamily=CHART_FONT,
                 transform=sax.transAxes)

        self._tm_status.set(f'更新時間：{datetime.now().strftime("%H:%M:%S")}')
        self._tm_canvas.draw()

    # ═══════════════════════════════════════════════════════════════════════════
    # Tab 3：個股買賣分析
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_tab3(self):
        f = self.tab3

        ctrl = ttk.Frame(f)
        ctrl.pack(fill='x', padx=14, pady=8)
        ttk.Label(ctrl, text='個股買賣分析', style='Hdr.TLabel').pack(side='left')
        ttk.Label(ctrl, text='庫存快選 / 輸入代號：').pack(side='left', padx=(20, 4))
        self._stock_var   = tk.StringVar()
        self._stock_combo = ttk.Combobox(ctrl, textvariable=self._stock_var,
                                          state='normal', width=26)
        self._stock_combo.pack(side='left', padx=4)
        self._stock_combo.bind('<Return>', lambda _: self._draw_analysis())
        ttk.Label(ctrl, text='（↵ 或點按鈕）',
                  foreground=C_FG2, font=('Microsoft JhengHei', 8)).pack(side='left', padx=(0, 6))
        ttk.Button(ctrl, text='📊  繪製圖表', style='Nav.TButton', command=self._draw_analysis).pack(side='left')
        ttk.Button(ctrl, text='💾  另存圖檔', style='Nav.TButton', command=self._save_analysis).pack(side='left', padx=(6, 0))

        self._an_fig    = plt.Figure(figsize=(9.5, 5.5), dpi=100, facecolor=C_BG)
        self._an_canvas = FigureCanvasTkAgg(self._an_fig, master=f)
        w = self._an_canvas.get_tk_widget()
        w.configure(bg=C_BG)
        w.pack(fill='both', expand=True, padx=8, pady=(0, 8))

        self._an_bar_data = []      # list of bar dicts for hover detection
        self._an_ax1      = None    # saved axes ref
        self._an_tooltip  = None    # tooltip Toplevel
        self._an_canvas.mpl_connect('motion_notify_event', self._on_an_motion)
        self._an_canvas.mpl_connect('figure_leave_event',  lambda e: self._hide_an_tooltip())

        self._update_stock_list()

    def _update_stock_list(self):
        df = load_df()
        if df.empty:
            return
        holdings = calc_holdings(df)
        holding_keys = set(holdings.keys())   # set of (code, cat)

        # 取得每個 (代號, 分類) 組合的股票名稱
        combos = (df.groupby(['股票代號', '分類'])['股票名稱']
                    .last().reset_index())

        def _opt(row):
            code = str(row['股票代號'])
            cat  = str(row['分類'])
            name = str(row['股票名稱'])
            return f"{code}  {name}  [{cat}]"

        hold_opts, sold_opts = [], []
        for _, row in combos.iterrows():
            key = (str(row['股票代號']), str(row['分類']))
            (hold_opts if key in holding_keys else sold_opts).append(_opt(row))

        options = hold_opts + (['── 已賣出 ──'] if sold_opts else []) + sold_opts
        self._stock_combo['values'] = options
        if not self._stock_var.get() and hold_opts:
            self._stock_var.set(hold_opts[0])

    # ── 個股分析長條 hover helpers ─────────────────────────────────────────────
    def _on_an_motion(self, event):
        if self._an_ax1 is None or not self._an_bar_data or event.x is None:
            self._hide_an_tooltip()
            return
        # 不依賴 event.inaxes（twinx 時可能回傳 ax2），直接轉換顯示座標到 ax1
        try:
            xd, yd = self._an_ax1.transData.inverted().transform((event.x, event.y))
        except Exception:
            self._hide_an_tooltip()
            return
        found = None
        for b in self._an_bar_data:
            if abs(xd - b['x']) <= b['w'] / 2 and 0 <= yd <= b['price']:
                found = b
                break
        if found is None:
            self._hide_an_tooltip()
            return
        widget = self._an_canvas.get_tk_widget()
        sx = widget.winfo_rootx() + int(event.x) + 18
        sy = widget.winfo_rooty() + int(widget.winfo_height() - event.y) + 18
        self._show_an_tooltip(sx, sy, found)

    def _show_an_tooltip(self, sx, sy, data):
        if self._an_tooltip is None or not self._an_tooltip.winfo_exists():
            tip = tk.Toplevel(self)
            tip.overrideredirect(True)
            tip.attributes('-topmost', True)
            tip.configure(bg='#1e1e1e')
            border = tk.Frame(tip, bg='#3e3e3e', padx=1, pady=1)
            border.pack(fill='both', expand=True)
            inner = tk.Frame(border, bg='#252526', padx=10, pady=8)
            inner.pack(fill='both', expand=True)
            self._an_tip_widgets = {}
            for key in ('title', 'price', 'qty', 'fee', 'tax', 'net', 'avg'):
                lbl = tk.Label(inner, bg='#252526',
                               font=('Microsoft JhengHei', 10), anchor='w')
                lbl.pack(fill='x')
                self._an_tip_widgets[key] = lbl
            self._an_tooltip = tip

        w   = self._an_tip_widgets
        side_color = C_BUY_FG if data['side'] == '買' else C_SELL_FG
        w['title'].config(
            text=f"{data['date']}　{data['side']}　({data['category']})",
            fg=side_color, font=('Microsoft JhengHei', 11, 'bold'))
        w['price'].config(text=f"成交價：{data['price']:,.2f} 元", fg='#cccccc')
        w['qty'].config(  text=f"數　量：{data['qty']:,.0f} 股",   fg='#cccccc')
        w['fee'].config(  text=f"手續費：{data['fee']:,.0f} 元",   fg='#cccccc')
        tax_txt = f"交易稅：{data['tax']:,.0f} 元" if data['tax'] else ''
        w['tax'].config(  text=tax_txt, fg='#cccccc')
        net_sign = '+' if data['net'] >= 0 else ''
        w['net'].config(  text=f"淨金額：{net_sign}{data['net']:,.0f} 元",
                          fg='#4ec94e' if data['net'] >= 0 else '#f07070')
        if not np.isnan(data['avg_cost']):
            w['avg'].config(text=f"當時均價：{data['avg_cost']:,.2f} 元", fg='#8ab4d4')
        else:
            w['avg'].config(text='', fg='#cccccc')

        self._an_tooltip.geometry(f'+{sx}+{sy}')
        self._an_tooltip.deiconify()

    def _hide_an_tooltip(self):
        if self._an_tooltip and self._an_tooltip.winfo_exists():
            self._an_tooltip.withdraw()

    def _draw_analysis(self):
        sel = self._stock_var.get().strip()
        if not sel or sel.startswith('──'):
            messagebox.showinfo('提示', '請輸入或選擇股票代號')
            return

        # 格式："CODE  NAME  [分類]" 或 "CODE  NAME" 或 "CODE"
        cat_filter = None
        if sel.endswith(']') and '[' in sel:
            cat_filter = sel[sel.rfind('[') + 1:-1]
            sel = sel[:sel.rfind('[')].strip()
        code = sel.split()[0].upper()
        df   = load_df()
        sdf  = df[df['股票代號'].astype(str) == code].copy()
        if cat_filter:
            sdf = sdf[sdf['分類'].astype(str) == cat_filter]
        sdf = sdf.sort_values('日期').reset_index(drop=True)

        if sdf.empty:
            messagebox.showinfo('提示', f'找不到 {code} 的交易紀錄\n請確認代號，或先在「買賣紀錄」新增交易。')
            return

        # ── 計算持股均價 ──────────────────────────────────────────────────────
        qty_cum, cost_cum = 0.0, 0.0
        avg_costs = []
        for _, r in sdf.iterrows():
            q, p, fee = float(r['數量(股)']), float(r['價格(元)']), float(r['手續費(元)'])
            if r['買賣'] == '買':
                cost_cum += q * p + fee
                qty_cum  += q
            else:
                if qty_cum > 0:
                    cost_cum -= (cost_cum / qty_cum) * q
                    qty_cum  -= q
            avg_costs.append(cost_cum / qty_cum if qty_cum > 0.5 else np.nan)

        sdf['avg_cost'] = avg_costs

        dates  = sdf['日期'].values
        prices = sdf['價格(元)'].values.astype(float)
        sides  = sdf['買賣'].values
        qtys   = sdf['數量(股)'].values.astype(float)

        # ── 抓歷史收盤價 ──────────────────────────────────────────────────────
        start_str  = pd.to_datetime(dates[0]).strftime('%Y-%m-%d')
        hist_close = get_history(code, start_str)   # pd.Series or None

        # ── 繪圖 ──────────────────────────────────────────────────────────────
        self._an_fig.clear()
        ax1 = self._an_fig.add_subplot(111, facecolor=C_PANEL)
        ax2 = ax1.twinx()

        for spine in ax1.spines.values():
            spine.set_edgecolor(C_BORDER)
        ax1.tick_params(colors=C_FG2)
        ax1.xaxis.label.set_color(C_FG)
        ax1.yaxis.label.set_color(C_FG)

        x_num = mdates.date2num(pd.to_datetime(dates))
        bar_w = max(0.6, (x_num[-1] - x_num[0]) / max(len(x_num) * 2, 10)) if len(x_num) > 1 else 1.5

        # 同日多筆：計算各 bar 的並排 x 位置與寬度
        from collections import defaultdict
        date_groups: dict = defaultdict(list)
        for i, xn in enumerate(x_num):
            date_groups[round(xn, 4)].append(i)

        per_x = np.empty(len(x_num))
        per_w = np.empty(len(x_num))
        for indices in date_groups.values():
            n     = len(indices)
            sub_w = bar_w / n * 0.92          # 留 8% 間隙
            for rank, idx in enumerate(indices):
                per_x[idx] = x_num[idx] - bar_w / 2 + (rank + 0.5) * (bar_w / n)
                per_w[idx] = sub_w

        # 歷史收盤價折線（最底層）
        hist_plotted = False
        if hist_close is not None and not hist_close.empty:
            try:
                hidx = hist_close.index
                if getattr(hidx, 'tz', None) is not None:
                    hidx = hidx.tz_localize(None)
                hx = mdates.date2num(pd.to_datetime(hidx))
                hy = hist_close.values.astype(float)
                ax1.plot(hx, hy, color='#b0bec5', linewidth=1.4,
                         alpha=0.75, zorder=2, label='收盤價走勢')
                hist_plotted = True
            except Exception:
                pass

        # 買賣成交價長條（並排）
        bar_colors = [C_BUY_FG if s == '買' else C_SELL_FG for s in sides]
        ax1.bar(per_x, prices, width=per_w, color=bar_colors, alpha=0.70, zorder=3)

        # 持股均價折線（用原始 x_num 保持連續）
        valid = ~np.isnan(sdf['avg_cost'].values)
        if valid.any():
            ax1.plot(x_num[valid], sdf['avg_cost'].values[valid],
                     color='#64b5f6', linewidth=2.2, marker='o', markersize=5, zorder=5)

        # 即時現價橫線
        cur_price, _ = get_price(code)
        if cur_price:
            ax1.axhline(cur_price, color='#ffa726', linestyle='--', linewidth=1.8, zorder=4)

        # 次軸：交易股數（並排）
        vol_colors = ['#4caf50' if s == '買' else '#f44336' for s in sides]
        ax2.bar(per_x, qtys, width=per_w, color=vol_colors, alpha=0.18, zorder=1)
        ax2.set_ylabel('交易股數', color=C_FG2, fontsize=9)
        ax2.tick_params(axis='y', colors=C_FG2)
        ax2.yaxis.set_major_formatter(
            matplotlib.ticker.FuncFormatter(lambda v, _: f'{int(v):,}'))
        for sp in ax2.spines.values():
            sp.set_edgecolor(C_BORDER)

        ax1.xaxis_date()
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        self._an_fig.autofmt_xdate(rotation=28, ha='right')
        ax1.yaxis.set_major_formatter(
            matplotlib.ticker.FuncFormatter(lambda v, _: f'{v:,.1f}'))
        ax1.grid(axis='y', alpha=0.15, linestyle='--', color=C_FG2)

        # 圖例
        legend_handles = [
            Patch(color=C_BUY_FG,  alpha=0.8, label='買入'),
            Patch(color=C_SELL_FG, alpha=0.8, label='賣出'),
        ]
        if hist_plotted:
            legend_handles.append(
                Line2D([0], [0], color='#b0bec5', linewidth=1.4, label='收盤價走勢'))
        legend_handles.append(
            Line2D([0], [0], color='#64b5f6', linewidth=2.2,
                   marker='o', markersize=5, label='持股均價'))
        if cur_price:
            legend_handles.append(
                Line2D([0], [0], color='#ffa726', linestyle='--',
                       linewidth=1.8, label=f'現價 {cur_price:.1f}'))
        legend = ax1.legend(handles=legend_handles, loc='best',
                            fontsize=9, framealpha=0.6)
        legend.get_frame().set_facecolor(C_PANEL)
        for text in legend.get_texts():
            text.set_color(C_FG)

        ax1.set_ylabel('價格 (元)', fontsize=10, color=C_FG)

        name     = sdf['股票名稱'].iloc[0]
        last_avg = cost_cum / qty_cum if qty_cum > 0.5 else None
        subtitle = ''
        if last_avg:
            pnl_str = (f'{cur_price - last_avg:+.2f} ({(cur_price/last_avg - 1)*100:+.1f}%)'
                       if cur_price else '')
            subtitle = (f'庫存 {qty_cum:,.0f} 股  |  均價 {last_avg:.2f} 元'
                        + (f'  |  損益 {pnl_str}' if pnl_str else ''))

        cat_label = f'  [{cat_filter}]' if cat_filter else ''
        ax1.set_title(f'{code}  {name}{cat_label}\n{subtitle}',
                      fontsize=12, fontweight='bold', pad=8, color=C_FG)

        # 儲存長條資料供 hover 偵測
        self._an_ax1     = ax1
        self._an_bar_data = []
        self._hide_an_tooltip()
        for i in range(len(sdf)):
            row = sdf.iloc[i]
            self._an_bar_data.append({
                'x':        per_x[i],
                'w':        per_w[i],
                'price':    float(row['價格(元)']),
                'side':     str(row['買賣']),
                'date':     str(row['日期']),
                'qty':      float(row['數量(股)']),
                'fee':      float(row['手續費(元)']),
                'tax':      float(row['交易稅(元)']),
                'net':      float(row['淨金額(元)']),
                'avg_cost': sdf['avg_cost'].iloc[i],
                'category': str(row['分類']),
            })

        self._an_fig.tight_layout()
        self._an_canvas.draw()


    # ═══════════════════════════════════════════════════════════════════════════
    # Tab 4：ETF 成分股分析
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_tab4(self):
        f = self.tab4

        # ── 控制列 ────────────────────────────────────────────────────────────
        ctrl = ttk.Frame(f)
        ctrl.pack(fill='x', padx=14, pady=8)
        ttk.Label(ctrl, text='ETF 成分股分析', style='Hdr.TLabel').pack(side='left')

        ttk.Label(ctrl, text='選擇 ETF：').pack(side='left', padx=(20, 4))
        self._etf_var = tk.StringVar()
        # 持有中的 ETF 排在前面，加 ★ 標記
        try:
            _held = set(load_df()['股票代號'].dropna().astype(str).unique())
        except Exception:
            _held = set()
        _held_opts  = [f'{c}  ★{n}' for c, n in ETF_LIST if c in _held]
        _other_opts = [f'{c}  {n}'  for c, n in ETF_LIST if c not in _held]
        etf_options = _held_opts + _other_opts
        self._etf_combo = ttk.Combobox(ctrl, textvariable=self._etf_var,
                                        values=etf_options, state='normal', width=26)
        self._etf_combo.pack(side='left', padx=4)
        def _on_combo_select(_=None):
            sel = self._etf_var.get().strip()
            if sel:
                self._etf_code_var.set(sel.split()[0])
            self._draw_etf_map()

        self._etf_combo.bind('<<ComboboxSelected>>', _on_combo_select)
        self._etf_combo.bind('<Return>', _on_combo_select)

        ttk.Label(ctrl, text='或輸入代號：',
                  foreground=C_FG2, font=('Microsoft JhengHei', 9)).pack(side='left', padx=(10, 4))
        self._etf_code_var = tk.StringVar()
        etf_entry = ttk.Entry(ctrl, textvariable=self._etf_code_var, width=10)
        etf_entry.pack(side='left', padx=4)

        def _on_entry_change(*_):
            typed = self._etf_code_var.get().strip()
            # 找下拉清單中有無對應項目並同步
            for opt in self._etf_combo['values']:
                if opt.split()[0] == typed:
                    self._etf_var.set(opt)
                    return
            self._etf_var.set('')

        self._etf_code_var.trace_add('write', _on_entry_change)
        etf_entry.bind('<Return>', lambda _: self._draw_etf_map())

        ttk.Button(ctrl, text='📈  繪製', style='Nav.TButton',
                   command=self._draw_etf_map).pack(side='left', padx=(6, 0))
        ttk.Button(ctrl, text='💾  另存圖檔', style='Nav.TButton',
                   command=self._save_etf_map).pack(side='right', padx=(4, 0))

        # 狀態列（獨立一行，避免被截斷）
        self._etf_status = tk.StringVar(value='請選擇 ETF 或輸入代號後點擊繪製')
        status_bar = tk.Frame(f, bg=C_BG)
        status_bar.pack(fill='x', padx=14, pady=(0, 2))
        ttk.Label(status_bar, textvariable=self._etf_status,
                  foreground=C_FG2, font=('Microsoft JhengHei', 9),
                  wraplength=1100, justify='left').pack(anchor='w')

        # ── 摘要列（兩行卡片）────────────────────────────────────────────────
        summary_bar = tk.Frame(f, bg=C_PANEL)
        summary_bar.pack(fill='x', padx=8, pady=(0, 4))

        row1 = tk.Frame(summary_bar, bg=C_PANEL)
        row1.pack(fill='x')
        row2 = tk.Frame(summary_bar, bg=C_PANEL)
        row2.pack(fill='x', pady=(2, 0))

        def _card(parent, title):
            fr = tk.Frame(parent, bg='#1a1a2e', padx=14, pady=5)
            fr.pack(side='left', fill='x', expand=True, padx=3)
            tk.Label(fr, text=title, bg='#1a1a2e',
                     fg='#6a8faf', font=('Microsoft JhengHei', 8)).pack(anchor='w')
            var = tk.StringVar(value='—')
            lbl = tk.Label(fr, textvariable=var, bg='#1a1a2e',
                           fg=C_FG, font=('Microsoft JhengHei', 11, 'bold'))
            lbl.pack(anchor='w')
            return var, lbl

        self._etf_sv_name,   self._etf_sl_name   = _card(row1, 'ETF 名稱')
        self._etf_sv_count,  self._etf_sl_count  = _card(row1, '成分股 / 前10大占比')
        self._etf_sv_top,    self._etf_sl_top    = _card(row1, '最大持股')
        self._etf_sv_change, self._etf_sl_change = _card(row1, '加權漲跌幅')
        self._etf_sv_aum,    self._etf_sl_aum    = _card(row2, '基金規模')
        self._etf_sv_yield,  self._etf_sl_yield  = _card(row2, '年化殖利率')
        self._etf_sv_nav,    self._etf_sl_nav    = _card(row2, 'ETF 淨值(NAV)')
        self._etf_sv_alloc,  self._etf_sl_alloc  = _card(row2, '折溢價')

        # ── 可捲動主區域（樹狀圖 + 分析圖）──────────────────────────────────
        sc_outer = tk.Frame(f, bg=C_BG)
        sc_outer.pack(fill='both', expand=True, padx=0, pady=(0, 0))

        vsb = ttk.Scrollbar(sc_outer, orient='vertical')
        vsb.pack(side='right', fill='y')

        self._etf_sc = tk.Canvas(sc_outer, bg='#111111',
                                  yscrollcommand=vsb.set, highlightthickness=0)
        self._etf_sc.pack(side='left', fill='both', expand=True)
        vsb.config(command=self._etf_sc.yview)

        self._etf_inner = tk.Frame(self._etf_sc, bg='#111111')
        _sc_win = self._etf_sc.create_window(0, 0, anchor='nw', window=self._etf_inner)

        def _on_inner_cfg(e):
            self._etf_sc.configure(scrollregion=self._etf_sc.bbox('all'))
        self._etf_inner.bind('<Configure>', _on_inner_cfg)

        def _on_sc_cfg(e):
            self._etf_sc.itemconfig(_sc_win, width=e.width)
        self._etf_sc.bind('<Configure>', _on_sc_cfg)

        def _mw(e):
            self._etf_sc.yview_scroll(-1 if e.delta > 0 else 1, 'units')
        for _w in (self._etf_sc, self._etf_inner):
            _w.bind('<MouseWheel>', _mw)

        # ── K 線圖（上）：水平控制列 + 畫布 ─────────────────────────────
        kline_outer = tk.Frame(self._etf_inner, bg='#111111')
        kline_outer.pack(fill='x', padx=8, pady=(4, 2))
        kline_outer.bind('<MouseWheel>', _mw)

        # ── 水平控制列 ────────────────────────────────────────────────────
        kctrl = tk.Frame(kline_outer, bg='#111111')
        kctrl.pack(fill='x', pady=(0, 4))
        kctrl.bind('<MouseWheel>', _mw)

        _BTN_OFF = dict(bg='#2a2a3e', fg='#8888aa',
                        font=('Microsoft JhengHei', 8, 'bold'),
                        relief='flat', bd=0, padx=8, pady=3,
                        cursor='hand2', activebackground='#3a3a5e',
                        activeforeground='#ffffff')
        _BTN_ON  = dict(bg='#3a5fcd', fg='#ffffff',
                        font=('Microsoft JhengHei', 8, 'bold'),
                        relief='flat', bd=0, padx=8, pady=3,
                        cursor='hand2', activebackground='#4a70e0',
                        activeforeground='#ffffff')

        def _kbtn_style(btn: tk.Button, on: bool) -> None:
            btn.configure(**(_BTN_ON if on else _BTN_OFF))

        # ── 區間 ─────────────────────────────────────────────────────────
        self._kline_period = '3M'
        self._kline_period_btns: dict[str, tk.Button] = {}

        tk.Label(kctrl, text='區間:', bg='#111111', fg='#6a8faf',
                 font=('Microsoft JhengHei', 8, 'bold')).pack(side='left', padx=(4, 2))

        def _set_period(lbl: str) -> None:
            self._kline_period = lbl
            for _l, _b in self._kline_period_btns.items():
                _kbtn_style(_b, _l == lbl)
            self._redraw_etf_kline()

        for _pl in ['1M', '3M', '6M', '1Y', '全部']:
            _b = tk.Button(kctrl, text=_pl, **_BTN_OFF,
                           command=lambda l=_pl: _set_period(l))
            _b.pack(side='left', padx=2)
            self._kline_period_btns[_pl] = _b
        _kbtn_style(self._kline_period_btns['3M'], True)

        # 分隔
        tk.Label(kctrl, text='', bg='#111111', width=1).pack(side='left', padx=4)
        ttk.Separator(kctrl, orient='vertical').pack(side='left', fill='y', pady=2, padx=4)

        # ── 指標 ─────────────────────────────────────────────────────────
        self._kline_ind_state: dict[str, bool] = {
            'MA': True, 'BB': False, 'VOL': False,
            'MACD': True, 'RSI': False, 'KD': False}
        self._kline_ind_btns: dict[str, tk.Button] = {}

        tk.Label(kctrl, text='指標:', bg='#111111', fg='#6a8faf',
                 font=('Microsoft JhengHei', 8, 'bold')).pack(side='left', padx=(4, 2))

        def _toggle_ind(lbl: str) -> None:
            self._kline_ind_state[lbl] = not self._kline_ind_state[lbl]
            _kbtn_style(self._kline_ind_btns[lbl], self._kline_ind_state[lbl])
            self._redraw_etf_kline()

        for _il in ['MA', 'BB', 'VOL', 'MACD', 'RSI', 'KD']:
            _on = self._kline_ind_state[_il]
            _b = tk.Button(kctrl, text=_il,
                           **(_BTN_ON if _on else _BTN_OFF),
                           command=lambda l=_il: _toggle_ind(l))
            _b.pack(side='left', padx=2)
            self._kline_ind_btns[_il] = _b

        # K 線畫布
        self._etf_kline_fig = plt.Figure(figsize=(9.5, 5.3), dpi=100, facecolor='#111111')
        self._etf_kline_canvas = FigureCanvasTkAgg(self._etf_kline_fig, master=kline_outer)
        wk = self._etf_kline_canvas.get_tk_widget()
        wk.configure(bg='#111111')
        wk.pack(fill='both', expand=True)
        wk.bind('<MouseWheel>', _mw)

        # ── 樹狀圖畫布（K 線之後）────────────────────────────────────────
        self._etf_fig    = plt.Figure(figsize=(9.5, 5.3), dpi=100, facecolor='#111111')
        self._etf_canvas = FigureCanvasTkAgg(self._etf_fig, master=self._etf_inner)
        w = self._etf_canvas.get_tk_widget()
        w.configure(bg='#111111')
        w.pack(fill='x', padx=8, pady=(2, 2))
        w.bind('<MouseWheel>', _mw)

        # ── 分析圖畫布（環形圖）────────────────────────────────────────────
        self._etf_info_fig = plt.Figure(figsize=(9.5, 4.2), dpi=100, facecolor='#1a1a2e')
        self._etf_info_canvas = FigureCanvasTkAgg(self._etf_info_fig, master=self._etf_inner)
        w2 = self._etf_info_canvas.get_tk_widget()
        w2.configure(bg='#1a1a2e')
        w2.pack(fill='x', padx=8, pady=(2, 8))
        w2.bind('<MouseWheel>', _mw)

        self._etf_rects         = []
        self._etf_tooltip       = None
        self._etf_ax            = None
        self._etf_info_cid      = None
        self._etf_info_charts   = []
        self._kline_ax          = None
        self._kline_ohlcv       = None
        self._kline_n           = 0
        self._kline_header      = None
        self._kline_vline       = None
        self._current_kline_code = None
        self._current_kline_name = ''
        self._etf_canvas.mpl_connect('motion_notify_event', self._on_etf_motion)
        self._etf_canvas.mpl_connect('figure_leave_event',  lambda e: self._hide_etf_tooltip())
        self._etf_kline_canvas.mpl_connect('motion_notify_event', self._on_kline_hover)

    def _get_etf_code(self) -> str:
        code = self._etf_code_var.get().strip()
        if not code:
            sel = self._etf_var.get().strip()
            if sel:
                code = sel.split()[0]
        return code

    def _draw_etf_map(self):
        code = self._get_etf_code()
        if not code:
            self._etf_status.set('請選擇 ETF 或輸入代號')
            return

        self._etf_status.set(f'載入 {code} 中…')
        self._etf_rects = []
        self._hide_etf_tooltip()
        self._etf_fig.clear()
        self._etf_fig.patch.set_facecolor('#111111')
        ax = self._etf_fig.add_axes([0, 0, 1, 1])
        ax.set_facecolor('#111111')
        ax.axis('off')
        ax.text(0.5, 0.5, f'載入 {code} 成分股中，請稍候…',
                ha='center', va='center', color='#888',
                fontsize=13, fontfamily=CHART_FONT, transform=ax.transAxes)
        self._etf_canvas.draw()

        def _worker():
            try:
                import concurrent.futures
                with concurrent.futures.ThreadPoolExecutor(max_workers=3) as _ex:
                    _f1 = _ex.submit(fetch_etf_data, code)
                    _f2 = _ex.submit(fetch_etf_meta, code)
                    _f3 = _ex.submit(_fetch_twse_industry_map)
                    etf_name, components, dbg = _f1.result()
                    meta    = _f2.result()
                    ind_map = _f3.result()
                self.after(0, lambda: self._draw_etf_map_impl(
                    code, etf_name, components, dbg, meta, ind_map))
            except Exception as e:
                self.after(0, lambda: self._etf_status.set(f'錯誤：{e}'))

        threading.Thread(target=_worker, daemon=True).start()

    def _draw_etf_map_impl(self, code: str, etf_name: str, components: list,
                           debug_msg: str = '', meta: dict = None,
                           ind_map: dict = None):
        if meta is None:
            meta = {}
        TREE_BG = '#111111'
        SEP_COL = '#111111'
        is_partial = 'Yahoo topHoldings' in debug_msg  # only top-10

        self._etf_rects = []
        self._hide_etf_tooltip()
        self._etf_fig.clear()
        self._etf_fig.patch.set_facecolor(TREE_BG)

        ax = self._etf_fig.add_axes([0, 0.04, 1, 0.96])
        ax.set_facecolor(TREE_BG)
        ax.set_xlim(0, 100)
        ax.set_ylim(0, 100)
        ax.axis('off')
        self._etf_ax = ax

        sax = self._etf_fig.add_axes([0, 0, 1, 0.04])
        sax.set_facecolor('#0d0d0d')
        sax.axis('off')

        if not components:
            ax.text(50, 50,
                    f'{code} 無法取得成分股資料\n（請查看下方狀態列的診斷訊息）',
                    ha='center', va='center', color='#888',
                    fontsize=12, fontfamily=CHART_FONT)
            self._etf_canvas.draw()
            self._etf_status.set(f'{code} 無法取得成分股資料（網路連線問題，請稍後再試）')
            return

        # ── 更新摘要列 ────────────────────────────────────────────────────────
        total_w = sum(c['weight'] for c in components) or 1.0
        weighted_chg = sum(c['change_pct'] * c['weight'] for c in components) / total_w
        top_stk = components[0]
        chg_fg   = '#f07070' if weighted_chg >= 0 else '#4ec94e'
        chg_sign = '+' if weighted_chg >= 0 else ''
        short_name = etf_name[:18] if len(etf_name) > 18 else etf_name
        self._etf_sv_name  .set(short_name)
        # 請求3: 若資料不完整, 改顯示前N大占比而非成分股數
        if is_partial:
            top_cov = sum(c['weight'] for c in components)
            self._etf_sv_count.set(f'前{len(components)}大  {top_cov:.1f}%')
        else:
            self._etf_sv_count.set(f'{len(components)} 檔')
        self._etf_sv_top   .set(f'{top_stk["code"]}  {top_stk["weight"]:.2f}%')
        self._etf_sv_change.set(f'{chg_sign}{weighted_chg:.2f}%')
        self._etf_sl_change.config(fg=chg_fg)
        # 規模 / 殖利率 / NAV / 配置
        aum = meta.get('total_assets')
        self._etf_sv_aum.set(
            f'NT$ {aum/1e8:.0f} 億' if aum and aum > 1e8
            else (f'US$ {aum/1e9:.2f} B' if aum else '—'))
        yld = meta.get('yield_pct')
        self._etf_sv_yield.set(f'{yld*100:.2f}%' if yld else '—')
        nav = meta.get('nav')
        self._etf_sv_nav.set(f'{nav:.2f}' if nav else '—')
        # 折溢價 = (市場價格 - NAV) / NAV × 100%
        nav   = meta.get('nav')
        price = meta.get('price') or meta.get('prev_close')
        if nav and price and nav > 0:
            prem = (price - nav) / nav * 100
            sign = '+' if prem >= 0 else ''
            label = '溢價' if prem >= 0 else '折價'
            fg    = '#f07070' if prem >= 0 else '#4ec94e'
            self._etf_sv_alloc.set(f'{sign}{prem:.3f}%（{label}）')
            self._etf_sl_alloc.config(fg=fg)
        else:
            self._etf_sv_alloc.set('—')

        # ── 顏色基準 ──────────────────────────────────────────────────────────
        all_pcts = [c['change_pct'] for c in components]
        _max_abs = max((abs(p) for p in all_pcts), default=10.0) or 10.0

        # ── 計算繪圖轉換係數 ──────────────────────────────────────────────────
        _dpi      = self._etf_fig.dpi
        _fig_w_px = self._etf_fig.get_size_inches()[0] * _dpi
        _fig_h_px = self._etf_fig.get_size_inches()[1] * _dpi
        _px_per_ux = _fig_w_px / 100
        _px_per_uy = _fig_h_px * 0.96 / 100
        _pt_per_px = 72 / _dpi
        GAP = 0.3

        # ── 樹狀圖（依 ETF 權重，無分類）────────────────────────────────────
        weights = [max(c['weight'], 0.001) for c in components]   # 防 ZeroDivision
        norm_w    = squarify.normalize_sizes(weights, 100, 100)
        raw_rects = squarify.squarify(norm_w, 0, 0, 100, 100)
        rects     = [{'x': r['x'], 'y': 100 - r['y'] - r['dy'],
                      'dx': r['dx'], 'dy': r['dy']} for r in raw_rects]

        for comp, sr in zip(components, rects):
            rx = sr['x'] + GAP / 2
            ry = sr['y'] + GAP / 2
            rw = sr['dx'] - GAP
            rh = sr['dy'] - GAP
            if rw <= 0 or rh <= 0:
                continue

            ax.add_patch(plt.Rectangle(
                (rx, ry), rw, rh,
                facecolor=pnl_color(comp['change_pct'], _max_abs),
                edgecolor=SEP_COL, linewidth=0.4, zorder=1))

            self._etf_rects.append({
                'rx': rx, 'ry': ry, 'rw': rw, 'rh': rh,
                'code':       comp['code'],
                'name':       comp['name'],
                'weight':     comp['weight'],
                'price':      comp['price'],
                'change_pct': comp['change_pct'],
            })

            min_dim = min(rw, rh)
            if min_dim < 2:
                continue

            rw_px = rw * _px_per_ux
            rh_px = rh * _px_per_uy
            name_disp = comp['name'][:8] if len(comp['name']) > 8 else comp['name']
            code_disp = comp['code']
            pnl_str   = f"{'+' if comp['change_pct'] >= 0 else ''}{comp['change_pct']:.2f}%"

            chars    = max(len(name_disp), len(code_disp), 4)
            fs_by_w  = 0.55 * rw_px * _pt_per_px / (chars * 0.60)
            fs_by_h  = 0.60 * rh_px * _pt_per_px / 3.8
            fs_name  = max(min(fs_by_w, fs_by_h, 36), 5.5)
            fs_code  = max(fs_name * 0.80, 5)
            fs_pct   = max(fs_name * 0.72, 5)
            cx_t     = rx + rw / 2
            mid_y    = ry + rh * 0.50

            if min_dim >= 6:
                line_gap = fs_name / _pt_per_px / _px_per_uy * 1.15
                ax.text(cx_t, mid_y + line_gap, name_disp,
                        ha='center', va='center', color='white',
                        fontsize=fs_name, fontweight='bold',
                        fontfamily=CHART_FONT, clip_on=True, zorder=4)
                ax.text(cx_t, mid_y, code_disp,
                        ha='center', va='center', color='#cccccc',
                        fontsize=fs_code, fontfamily=CHART_FONT,
                        clip_on=True, zorder=4)
                ax.text(cx_t, mid_y - line_gap, pnl_str,
                        ha='center', va='center', color='white',
                        fontsize=fs_pct, fontfamily=CHART_FONT,
                        clip_on=True, zorder=4)
            elif min_dim >= 4:
                gap = (fs_name / _pt_per_px / _px_per_uy +
                       fs_pct  / _pt_per_px / _px_per_uy) * 0.55
                ax.text(cx_t, mid_y + gap * 0.5, name_disp,
                        ha='center', va='center', color='white',
                        fontsize=fs_name, fontweight='bold',
                        fontfamily=CHART_FONT, clip_on=True, zorder=4)
                ax.text(cx_t, mid_y - gap * 0.5, pnl_str,
                        ha='center', va='center', color='white',
                        fontsize=fs_pct, fontfamily=CHART_FONT,
                        clip_on=True, zorder=4)
            else:
                ax.text(cx_t, mid_y, name_disp,
                        ha='center', va='center', color='white',
                        fontsize=fs_name, fontweight='bold',
                        fontfamily=CHART_FONT, clip_on=True, zorder=4)

        count_txt = (f'前{len(components)}大持股' if is_partial
                     else f'{len(components)} 檔成分股')
        sax.text(0.5, 0.5,
                 f'{etf_name}  ·  {count_txt}  ·  '
                 f'{datetime.now().strftime("%Y-%m-%d %H:%M")} 更新',
                 ha='center', va='center', color='#666',
                 fontsize=8, fontfamily=CHART_FONT, transform=sax.transAxes)

        if 'TWSE OK' in debug_msg:
            src_note = '  ✓ 資料來源：證交所'
        elif 'MoneyDJ' in debug_msg:
            src_note = '  ✓ 資料來源：MoneyDJ（完整成分股）'
        elif is_partial:
            top_cov = sum(c['weight'] for c in components)
            src_note = f'  ⚠ Yahoo Finance（前{len(components)}大，合計 {top_cov:.1f}%）'
        else:
            src_note = ''
        count_label = (f'前{len(components)}大' if is_partial else f'{len(components)} 檔')
        self._etf_status.set(f'{code}  {etf_name}  ·  {count_label}  ·  '
                              f'更新：{datetime.now().strftime("%H:%M:%S")}{src_note}')
        self._etf_canvas.draw()
        self._draw_etf_kline(code, etf_name)
        self._draw_etf_info(components, meta, debug_msg, ind_map or {})

    # ── ETF 分析圖（互動式環形圖 × 2）────────────────────────────────────────
    def _draw_etf_info(self, components: list, meta: dict,
                       debug_msg: str = '', ind_map: dict = None):
        INFO_BG = '#1a1a2e'
        PANEL_BG = '#22223a'

        import colorsys as _cs

        def _grad(n, hue=0.60):
            """n 種同色系漸層色（深→淺）。"""
            if n <= 0:
                return []
            out = []
            for i in range(n):
                t = i / max(n - 1, 1)
                r, g, b = _cs.hsv_to_rgb(hue, 0.85 - 0.30 * t, 0.48 + 0.42 * t)
                out.append(f'#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}')
            return out

        fig = self._etf_info_fig
        fig.clear()
        fig.patch.set_facecolor(INFO_BG)

        # Disconnect previous hover callback
        if self._etf_info_cid is not None:
            try:
                self._etf_info_canvas.mpl_disconnect(self._etf_info_cid)
            except Exception:
                pass
            self._etf_info_cid = None
        self._etf_info_charts = []

        # ── Helper: draw one donut + right-side legend ─────────────────────
        def _donut(pie_ax, leg_ax, vals, lbls, title, hue=0.60):
            """Draw donut in pie_ax, legend in leg_ax. Returns (wedges, center_text)."""
            colors = _grad(len(vals), hue)
            # Pie axes
            pie_ax.set_facecolor(PANEL_BG)
            pie_ax.set_title(title, color='#9ab8d8', fontsize=10,
                             fontfamily=CHART_FONT, pad=5)
            wedges, _ = pie_ax.pie(
                vals, colors=colors,
                wedgeprops={'width': 0.50, 'edgecolor': INFO_BG, 'linewidth': 1.5},
                startangle=90)
            pie_ax.set_aspect('equal')   # 確保是正圓
            # Center text (updated on hover)
            ct = pie_ax.text(0, 0, '', ha='center', va='center',
                             color='white', fontsize=9, fontfamily=CHART_FONT,
                             fontweight='bold', multialignment='center', zorder=10)
            # Legend axes（超過 6 項自動兩欄）
            leg_ax.set_facecolor(PANEL_BG)
            leg_ax.axis('off')
            leg_ax.set_xlim(0, 1)
            n = len(lbls)
            leg_ax.set_ylim(0, 1)
            n_cols   = 2 if n > 6 else 1
            per_col  = (n + n_cols - 1) // n_cols      # ceil(n / n_cols)
            col_w    = 1.0 / n_cols
            row_h    = min(0.13, 0.92 / max(per_col, 1))
            y0       = 0.5 + (per_col - 1) * row_h / 2  # 垂直居中
            for i, (lbl, val) in enumerate(zip(lbls, vals)):
                col_idx = i // per_col
                row_idx = i % per_col
                y       = y0 - row_idx * row_h
                xo      = col_idx * col_w               # 欄偏移
                c       = colors[i % len(colors)]
                # scatter 用螢幕像素單位，不受 axes 長寬比影響，永遠是正圓
                dot_sz  = 55 if n_cols == 2 else 70
                leg_ax.scatter([xo + 0.04], [y], s=dot_sz, c=[c],
                               zorder=5, linewidths=0, clip_on=False)
                fs_name = 8.5 if n_cols == 2 else 9.5
                fs_pct  = 7.5 if n_cols == 2 else 8.5
                leg_ax.text(xo + 0.11, y + 0.018, lbl[:9],
                            ha='left', va='center', color='#e0e0e0',
                            fontsize=fs_name, fontfamily=CHART_FONT,
                            transform=leg_ax.transData)
                leg_ax.text(xo + 0.11, y - 0.018, f'{val:.2f}%',
                            ha='left', va='center', color='#aaaaaa',
                            fontsize=fs_pct, fontfamily=CHART_FONT,
                            transform=leg_ax.transData)
            return wedges, ct

        # ── Chart 1: Top holdings ──────────────────────────────────────────
        top_n  = min(10, len(components))
        top    = components[:top_n]
        others = sum(c['weight'] for c in components[top_n:])
        vals1  = [c['weight'] for c in top]
        lbls1  = [c['name'][:10] for c in top]
        if others > 0.01:
            vals1.append(others)
            lbls1.append('其他')

        ax1  = fig.add_axes([0.01, 0.06, 0.23, 0.88])
        axl1 = fig.add_axes([0.25, 0.06, 0.21, 0.88])
        w1, ct1 = _donut(ax1, axl1, vals1, lbls1, f'前 {top_n} 大持股', hue=0.60)
        self._etf_info_charts.append((ax1, w1, lbls1, vals1, ct1))

        # ── Chart 2: Sector distribution ──────────────────────────────────
        # 方法 1: Yahoo Finance sector_weights（英文 key → 中文）
        sw = meta.get('sector_weights', []) if meta else []
        sector_data: dict = {}
        for item in sw:
            if isinstance(item, dict):
                for k, v in item.items():
                    if v and v > 0.001:
                        sector_data[SECTOR_ZH.get(k, k)] = v * 100

        # 方法 2: 由成分股代號 × 產業別對照表計算（fallback）
        if not sector_data and ind_map and components:
            for comp in components:
                ind = ind_map.get(comp['code'], '其他')
                if not ind:
                    ind = '其他'
                sector_data[ind] = sector_data.get(ind, 0.0) + comp['weight']

        ax2  = fig.add_axes([0.51, 0.06, 0.23, 0.88])
        axl2 = fig.add_axes([0.75, 0.06, 0.23, 0.88])

        if sector_data:
            sorted_s = sorted(sector_data.items(), key=lambda x: -x[1])
            top_s    = sorted_s[:5]
            other_s  = sum(v for _, v in sorted_s[5:])
            lbls2    = [k for k, _ in top_s]
            vals2    = [v for _, v in top_s]
            if other_s > 0.01:
                lbls2.append('其他')
                vals2.append(other_s)
            w2, ct2 = _donut(ax2, axl2, vals2, lbls2, '產業分布', hue=0.45)
            self._etf_info_charts.append((ax2, w2, lbls2, vals2, ct2))
        else:
            ax2.set_facecolor(PANEL_BG)
            ax2.axis('off')
            ax2.set_title('產業分布', color='#9ab8d8', fontsize=9,
                          fontfamily=CHART_FONT, pad=5)
            ax2.text(0.5, 0.5, '產業分布資料\n暫不可用', ha='center', va='center',
                     color='#555', fontsize=9, fontfamily=CHART_FONT,
                     transform=ax2.transAxes)
            axl2.axis('off')
            axl2.set_facecolor(PANEL_BG)

        # ── Hover interaction ──────────────────────────────────────────────
        def _on_info_hover(event):
            if event.inaxes is None:
                return
            changed = False
            for pie_ax, wedges, lbls, vals, ct in self._etf_info_charts:
                if event.inaxes is not pie_ax:
                    continue
                found = False
                for wedge, lbl, val in zip(wedges, lbls, vals):
                    try:
                        hit, _ = wedge.contains(event)
                    except Exception:
                        hit = False
                    if hit:
                        new_txt = f'{lbl}\n{val:.2f}%'
                        if ct.get_text() != new_txt:
                            ct.set_text(new_txt)
                            changed = True
                        found = True
                        break
                if not found and ct.get_text():
                    ct.set_text('')
                    changed = True
            if changed:
                self._etf_info_canvas.draw_idle()

        self._etf_info_cid = self._etf_info_canvas.mpl_connect(
            'motion_notify_event', _on_info_hover)
        self._etf_info_canvas.draw()

    # ── ETF K 線圖 ────────────────────────────────────────────────────────────
    def _redraw_etf_kline(self):
        """指標 toggle 後重新繪製 K 線圖。"""
        if self._current_kline_code:
            self._draw_etf_kline(self._current_kline_code, self._current_kline_name)

    def _draw_etf_kline(self, code: str, etf_name: str):
        """Draw K-line (candlestick) + selectable indicators for ETF."""
        import numpy as np

        CHART_BG = '#111111'
        PANEL_BG = '#1a1a2e'

        # 儲存供 toggle redraw 使用
        self._current_kline_code = code
        self._current_kline_name = etf_name

        fig = self._etf_kline_fig
        fig.clear()
        fig.patch.set_facecolor(CHART_BG)
        self._kline_ax = None
        self._kline_ohlcv = None

        # ── 取得 OHLCV 資料 ───────────────────────────────────────────────
        _period_map = {'1M': '1mo', '3M': '3mo', '6M': '6mo', '1Y': '1y', '全部': '5y'}
        _yf_period = _period_map.get(getattr(self, '_kline_period', '3M'), '3mo')
        ohlcv = None
        for s in ['.TW', '.TWO', '']:
            try:
                t = yf.Ticker(code + s)
                h = t.history(period=_yf_period, auto_adjust=False)
                if not h.empty:
                    ohlcv = h
                    break
            except Exception:
                pass

        if ohlcv is None or ohlcv.empty:
            ax = fig.add_subplot(111, facecolor=PANEL_BG)
            ax.text(0.5, 0.5, f'{code} 無法取得 K 線資料',
                    ha='center', va='center', color='#888',
                    fontsize=11, fontfamily=CHART_FONT, transform=ax.transAxes)
            ax.axis('off')
            self._etf_kline_canvas.draw()
            return

        # ── 讀取指標開關 ──────────────────────────────────────────────────
        ind = dict(getattr(self, '_kline_ind_state', {
            'MA': True, 'BB': False, 'VOL': False,
            'MACD': True, 'RSI': False, 'KD': False}))
        close = ohlcv['Close']
        ohlcv = ohlcv.copy()

        # MA
        if ind.get('MA', True):
            ohlcv['MA5']  = close.rolling(5,  min_periods=1).mean()
            ohlcv['MA10'] = close.rolling(10, min_periods=1).mean()
            ohlcv['MA20'] = close.rolling(20, min_periods=1).mean()

        # Bollinger Bands
        if ind.get('BB', False):
            _bb_m = close.rolling(20, min_periods=5).mean()
            _bb_s = close.rolling(20, min_periods=5).std(ddof=0)
            ohlcv['BB_U'] = _bb_m + 2 * _bb_s
            ohlcv['BB_M'] = _bb_m
            ohlcv['BB_L'] = _bb_m - 2 * _bb_s

        # MACD(12,26,9)
        if ind.get('MACD', True):
            _e12 = close.ewm(span=12, adjust=False).mean()
            _e26 = close.ewm(span=26, adjust=False).mean()
            ohlcv['MACD_L'] = _e12 - _e26
            ohlcv['MACD_S'] = ohlcv['MACD_L'].ewm(span=9, adjust=False).mean()
            ohlcv['MACD_H'] = ohlcv['MACD_L'] - ohlcv['MACD_S']

        # RSI(14)
        if ind.get('RSI', False):
            _d  = close.diff()
            _g  = _d.clip(lower=0)
            _l  = (-_d).clip(lower=0)
            _rs = _g.ewm(com=13, adjust=False).mean() / \
                  _l.ewm(com=13, adjust=False).mean().replace(0, float('nan'))
            ohlcv['RSI'] = 100 - 100 / (1 + _rs)

        # KD(9,3,3)
        if ind.get('KD', False):
            _lo = ohlcv.get('Low',  ohlcv['Close']).rolling(9, min_periods=1).min()
            _hi = ohlcv.get('High', ohlcv['Close']).rolling(9, min_periods=1).max()
            _rk = 100 * (close - _lo) / (_hi - _lo + 1e-10)
            ohlcv['KD_K'] = _rk.ewm(com=2, adjust=False).mean()
            ohlcv['KD_D'] = ohlcv['KD_K'].ewm(com=2, adjust=False).mean()

        n  = len(ohlcv)
        xs = np.arange(n)

        # ── 動態佈局：依啟用的副指標數量分配高度 ────────────────────────
        sub_list = []          # 依顯示順序（價格圖之下由上而下）
        if ind.get('VOL',  False): sub_list.append('VOL')
        if ind.get('MACD', True):  sub_list.append('MACD')
        if ind.get('RSI',  False): sub_list.append('RSI')
        if ind.get('KD',   False): sub_list.append('KD')

        L, W    = 0.07, 0.88
        TOP     = 0.94          # 頂端預留給資訊文字
        BOT     = 0.07          # 底端預留給 X 軸日期
        GAP     = 0.012
        n_sub   = len(sub_list)
        avail   = TOP - BOT

        if n_sub == 0:
            sub_h = 0.0
            p_bot = BOT
            p_h   = avail
        else:
            sub_h = min(0.19, (avail * 0.44) / n_sub)
            p_bot = BOT + n_sub * (sub_h + GAP)
            p_h   = TOP - p_bot

        ax_price = fig.add_axes([L, p_bot, W, max(p_h, 0.25)])

        # 建立副圖軸（由下往上堆疊）
        sub_axes: dict = {}
        _y = BOT
        for _name in reversed(sub_list):
            sub_axes[_name] = fig.add_axes([L, _y, W, sub_h - GAP * 0.5])
            _y += sub_h + GAP

        def _style(ax, xticks=False):
            ax.set_facecolor(PANEL_BG)
            ax.tick_params(colors='#888', labelsize=7, length=2)
            ax.yaxis.tick_right()
            ax.set_xlim(-0.5, n - 0.5)
            for sp in ax.spines.values():
                sp.set_color('#333')
            ax.grid(axis='y', color='#2a2a3a', linewidth=0.4, linestyle='-')
            if not xticks:
                ax.set_xticks([])

        # ── 蠟燭圖 ───────────────────────────────────────────────────────
        cw = 0.55
        for i, (_, row) in enumerate(ohlcv.iterrows()):
            o_  = float(row.get('Open',  row['Close']))
            h_  = float(row.get('High',  row['Close']))
            l_  = float(row.get('Low',   row['Close']))
            c_  = float(row['Close'])
            clr = '#f07070' if c_ >= o_ else '#4ec94e'
            ax_price.plot([i, i], [l_, h_], color=clr, linewidth=0.7, zorder=2)
            ax_price.add_patch(plt.Rectangle(
                (i - cw/2, min(o_, c_)), cw, max(abs(c_ - o_), 0.01),
                facecolor=clr, edgecolor=clr, linewidth=0, zorder=3))

        # ── MA 線（覆蓋於蠟燭圖）────────────────────────────────────────
        _leg_handles = []
        if ind.get('MA', True):
            for col, clr, lbl in [('MA5',  '#f0e44a', 'MA5'),
                                   ('MA10', '#4a9cf0', 'MA10'),
                                   ('MA20', '#f0a04a', 'MA20')]:
                if col in ohlcv:
                    ln, = ax_price.plot(xs, ohlcv[col].values, color=clr,
                                        linewidth=1.1, label=lbl, zorder=4)
                    _leg_handles.append(ln)

        # ── Bollinger Bands ──────────────────────────────────────────────
        if ind.get('BB', False) and 'BB_U' in ohlcv:
            ax_price.plot(xs, ohlcv['BB_U'].values, color='#88aaff',
                          linewidth=0.8, linestyle='--', label='BB上', zorder=4)
            ax_price.plot(xs, ohlcv['BB_M'].values, color='#ffffff',
                          linewidth=0.6, linestyle='--', label='BB中', zorder=4)
            ax_price.plot(xs, ohlcv['BB_L'].values, color='#88aaff',
                          linewidth=0.8, linestyle='--', label='BB下', zorder=4)
            ax_price.fill_between(xs,
                ohlcv['BB_U'].values, ohlcv['BB_L'].values,
                color='#4a9cf0', alpha=0.06, zorder=1)

        # ── 價格 Y 軸自動縮放 ────────────────────────────────────────────
        _price_vals = [ohlcv['High'].dropna().values, ohlcv['Low'].dropna().values]
        if 'BB_U' in ohlcv: _price_vals.append(ohlcv['BB_U'].dropna().values)
        if 'BB_L' in ohlcv: _price_vals.append(ohlcv['BB_L'].dropna().values)
        _all_p = np.concatenate(_price_vals)
        _pmin, _pmax = float(np.nanmin(_all_p)), float(np.nanmax(_all_p))
        _mg = (_pmax - _pmin) * 0.05 or 0.5
        ax_price.set_ylim(_pmin - _mg, _pmax + _mg * 2.0)

        _style(ax_price)
        ax_price.tick_params(labelsize=7.5)

        # MA legend
        if _leg_handles:
            ax_price.legend(handles=_leg_handles, loc='upper right',
                            fontsize=8.5, facecolor='#111111',
                            edgecolor='#333', labelcolor='white',
                            handlelength=1.5, framealpha=0.85, borderpad=0.5)

        # ── 副圖：VOL ────────────────────────────────────────────────────
        if 'VOL' in sub_axes:
            ax_v = sub_axes['VOL']
            vol = ohlcv.get('Volume', None)
            if vol is not None:
                v_clrs = ['#f07070' if ohlcv['Close'].iloc[i] >= ohlcv['Open'].iloc[i]
                          else '#4ec94e' for i in range(n)]
                ax_v.bar(xs, vol.values, color=v_clrs, width=0.7, zorder=2)
            ax_v.set_ylabel('VOL', color='#888', fontsize=7, labelpad=2)
            ax_v.yaxis.set_label_position('left')
            _style(ax_v)

        # ── 副圖：MACD(12,26,9) ──────────────────────────────────────────
        if 'MACD' in sub_axes and 'MACD_L' in ohlcv:
            ax_m = sub_axes['MACD']
            h_clrs = ['#f07070' if v >= 0 else '#4ec94e'
                      for v in ohlcv['MACD_H'].values]
            ax_m.bar(xs, ohlcv['MACD_H'].values, color=h_clrs, width=0.6, zorder=2)
            ax_m.plot(xs, ohlcv['MACD_L'].values, color='#4a9cf0',
                      linewidth=0.9, label='MACD', zorder=3)
            ax_m.plot(xs, ohlcv['MACD_S'].values, color='#f0a04a',
                      linewidth=0.9, label='Signal', zorder=3)
            ax_m.axhline(0, color='#444', linewidth=0.5, linestyle='--')
            ax_m.set_ylabel('MACD', color='#888', fontsize=7, labelpad=2)
            ax_m.yaxis.set_label_position('left')
            _style(ax_m)

        # ── 副圖：RSI(14) ────────────────────────────────────────────────
        if 'RSI' in sub_axes and 'RSI' in ohlcv:
            ax_r = sub_axes['RSI']
            ax_r.plot(xs, ohlcv['RSI'].values, color='#c46af0', linewidth=0.9)
            ax_r.axhline(70, color='#f07070', linewidth=0.5, linestyle='--')
            ax_r.axhline(30, color='#4ec94e', linewidth=0.5, linestyle='--')
            ax_r.set_ylim(0, 100)
            ax_r.set_ylabel('RSI', color='#888', fontsize=7, labelpad=2)
            ax_r.yaxis.set_label_position('left')
            _style(ax_r)

        # ── 副圖：KD(9,3,3) ─────────────────────────────────────────────
        if 'KD' in sub_axes and 'KD_K' in ohlcv:
            ax_k = sub_axes['KD']
            ax_k.plot(xs, ohlcv['KD_K'].values, color='#f0e44a',
                      linewidth=0.9, label='K')
            ax_k.plot(xs, ohlcv['KD_D'].values, color='#4a9cf0',
                      linewidth=0.9, label='D')
            ax_k.axhline(80, color='#f07070', linewidth=0.5, linestyle='--')
            ax_k.axhline(20, color='#4ec94e', linewidth=0.5, linestyle='--')
            ax_k.set_ylim(0, 100)
            ax_k.set_ylabel('KD', color='#888', fontsize=7, labelpad=2)
            ax_k.yaxis.set_label_position('left')
            _style(ax_k)

        # ── X 軸日期：顯示在最底層副圖，若無副圖則在價格圖 ──────────────
        _bottom_ax = sub_axes[sub_list[-1]] if sub_list else ax_price
        step   = max(1, n // 8)
        xtks   = list(range(0, n, step))
        xlbls  = [ohlcv.index[i].strftime('%Y-%m-%d') for i in xtks]
        _bottom_ax.set_xticks(xtks)
        _bottom_ax.set_xticklabels(xlbls, rotation=25, ha='right',
                                    fontsize=7, color='#888')

        # ── 頂部資訊文字 ─────────────────────────────────────────────────
        hdr = fig.text(
            0.07, 0.997,
            '日期：—    開：—    高：—    低：—    收：—    量：—',
            va='top', ha='left', color='#cccccc', fontsize=10,
            fontfamily=CHART_FONT,
            bbox=dict(facecolor='#111111', alpha=0.0, edgecolor='none', pad=1))
        self._kline_header = hdr

        # ── 垂直游標線 ───────────────────────────────────────────────────
        self._kline_vline = ax_price.axvline(
            -999, color='#666', linewidth=0.8, linestyle='--', zorder=5)

        # 儲存供 hover 使用
        self._kline_ax    = ax_price
        self._kline_ohlcv = ohlcv
        self._kline_n     = n

        self._etf_kline_canvas.draw()

    def _on_kline_hover(self, event):
        """K 線圖 hover：更新頂部資訊文字與垂直游標線。"""
        if (self._kline_ax is None
                or self._kline_header is None
                or event.inaxes is not self._kline_ax):
            return
        if event.xdata is None:
            return
        xi = int(round(event.xdata))
        if not (0 <= xi < self._kline_n):
            return

        row      = self._kline_ohlcv.iloc[xi]
        date_str = self._kline_ohlcv.index[xi].strftime('%Y-%m-%d')
        o  = float(row.get('Open',   row['Close']))
        h_ = float(row.get('High',   row['Close']))
        l_ = float(row.get('Low',    row['Close']))
        c  = float(row['Close'])
        vol = int(row.get('Volume', 0)) // 1000

        txt = (f'日期：{date_str}    開：{o:.2f}    高：{h_:.2f}'
               f'    低：{l_:.2f}    收：{c:.2f}    量：{vol}K')
        self._kline_header.set_text(txt)
        if self._kline_vline is not None:
            self._kline_vline.set_xdata([xi, xi])
        self._etf_kline_canvas.draw_idle()

    # ── ETF 樹狀圖 hover ──────────────────────────────────────────────────────
    def _on_etf_motion(self, event):
        if event.inaxes is None or self._etf_ax is None:
            self._hide_etf_tooltip()
            return
        xd, yd = event.xdata, event.ydata
        widget = self._etf_canvas.get_tk_widget()
        sx = widget.winfo_rootx() + int(event.x) + 18
        sy = widget.winfo_rooty() + int(widget.winfo_height() - event.y) + 18
        for r in self._etf_rects:
            if (r['rx'] <= xd <= r['rx'] + r['rw'] and
                    r['ry'] <= yd <= r['ry'] + r['rh']):
                self._show_etf_tooltip(sx, sy, r)
                return
        self._hide_etf_tooltip()

    def _show_etf_tooltip(self, sx, sy, data):
        if self._etf_tooltip is None or not self._etf_tooltip.winfo_exists():
            tip = tk.Toplevel(self)
            tip.overrideredirect(True)
            tip.attributes('-topmost', True)
            tip.configure(bg='#1e1e1e')
            border = tk.Frame(tip, bg='#3e3e3e', padx=1, pady=1)
            border.pack(fill='both', expand=True)
            inner = tk.Frame(border, bg='#252526', padx=10, pady=8)
            inner.pack(fill='both', expand=True)
            self._etf_tip_widgets = {}
            for key in ('title', 'weight', 'price', 'change'):
                lbl = tk.Label(inner, bg='#252526',
                               font=('Microsoft JhengHei', 10), anchor='w')
                lbl.pack(fill='x')
                self._etf_tip_widgets[key] = lbl
            self._etf_tooltip = tip

        w    = self._etf_tip_widgets
        pct  = data['change_pct']
        fg   = '#f07070' if pct >= 0 else '#4ec94e'
        sign = '+' if pct >= 0 else ''
        w['title'].config(text=f'{data["name"]}  {data["code"]}',
                          fg='#8ab4d4', font=('Microsoft JhengHei', 11, 'bold'))
        price_txt = '—' if data['price'] is None else f'{data["price"]:,.2f} 元'
        w['weight'].config(text=f'ETF 占比：{data["weight"]:.4f}%', fg='#cccccc')
        w['price'].config(text=f'現價：{price_txt}', fg='#cccccc')
        w['change'].config(text=f'今日漲跌：{sign}{pct:.2f}%', fg=fg)
        self._etf_tooltip.geometry(f'+{sx}+{sy}')
        self._etf_tooltip.deiconify()

    def _hide_etf_tooltip(self):
        if self._etf_tooltip and self._etf_tooltip.winfo_exists():
            self._etf_tooltip.withdraw()

    def _save_etf_map(self):
        from tkinter import filedialog
        code = self._get_etf_code() or 'etf'
        path = filedialog.asksaveasfilename(
            title='儲存 ETF 樹狀圖',
            defaultextension='.png',
            filetypes=[('PNG 圖片', '*.png'), ('JPEG 圖片', '*.jpg'), ('所有檔案', '*.*')],
            initialfile=f'{code}_treemap.png')
        if path:
            self._etf_fig.savefig(path, dpi=150, bbox_inches='tight',
                                   facecolor=self._etf_fig.get_facecolor())
            self._etf_status.set(f'已儲存：{path}')


# ─── 修改交易對話框 ────────────────────────────────────────────────────────────
class EditDialog(tk.Toplevel):
    """彈出式修改視窗，深色風格，與主視窗一致"""

    def __init__(self, parent: tk.Tk, idx: int, row: pd.Series, on_save):
        super().__init__(parent)
        self._idx     = idx
        self._row     = row
        self._on_save = on_save

        self.title(f'修改交易紀錄  #{idx + 1}')
        self.configure(bg=C_BG)
        self.resizable(False, False)
        self.grab_set()   # modal

        # 置中
        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        w, h   = 500, 390
        self.geometry(f'{w}x{h}+{px + (pw - w)//2}+{py + (ph - h)//2}')

        self._vars: dict[str, tk.StringVar] = {}
        self._build()

    def _build(self):
        r = self._row

        pad = dict(padx=14, pady=5)
        lkw = dict(sticky='e', **pad)
        ekw = dict(sticky='ew', **pad)

        f = ttk.Frame(self, style='TFrame')
        f.pack(fill='both', expand=True, padx=14, pady=12)
        f.columnconfigure(1, weight=1)

        today = datetime.today()

        def add_row(row_idx, label, key, widget_fn):
            ttk.Label(f, text=label).grid(row=row_idx, column=0, **lkw)
            w = widget_fn(f)
            w.grid(row=row_idx, column=1, **ekw)
            return w

        # 日期
        ttk.Label(f, text='日期').grid(row=0, column=0, **lkw)
        orig_date = pd.to_datetime(r['日期'])
        self._date_entry = DatePickerEntry(f, initial_date=orig_date)
        self._date_entry.grid(row=0, column=1, **ekw)

        # 股票代號 / 名稱
        for row_idx, (label, key, default) in enumerate([
            ('股票代號', 'code', str(r['股票代號'])),
            ('股票名稱', 'name', str(r['股票名稱'])),
        ], start=1):
            self._vars[key] = tk.StringVar(value=default)
            add_row(row_idx, label, key,
                    lambda p, k=key: ttk.Entry(p, textvariable=self._vars[k], width=22))

        # 分類 / 買賣
        self._vars['category'] = tk.StringVar(value=str(r.get('分類', CATEGORIES[0])))
        ttk.Label(f, text='投資分類').grid(row=3, column=0, **lkw)
        ttk.Combobox(f, textvariable=self._vars['category'],
                     values=CATEGORIES, state='readonly', width=20
                     ).grid(row=3, column=1, **ekw)

        self._vars['side'] = tk.StringVar(value=str(r['買賣']))
        ttk.Label(f, text='買 / 賣').grid(row=4, column=0, **lkw)
        self._side_cb = ttk.Combobox(f, textvariable=self._vars['side'],
                                      values=['買', '賣'], state='readonly', width=20,
                                      style=('Buy.TCombobox' if r['買賣'] == '買'
                                             else 'Sell.TCombobox'))
        self._side_cb.grid(row=4, column=1, **ekw)
        self._vars['side'].trace_add('write', self._on_side_change)

        # 數量 / 價格 / 手續費
        for row_idx, (label, key, default) in enumerate([
            ('數量 (股)',    'qty',   str(int(r['數量(股)']))),
            ('成交價格 (元)', 'price', str(r['價格(元)'])),
            ('手續費 (元)',  'fee',   str(int(r['手續費(元)']))),
        ], start=5):
            self._vars[key] = tk.StringVar(value=default)
            add_row(row_idx, label, key,
                    lambda p, k=key: ttk.Entry(p, textvariable=self._vars[k], width=22))

        # 自動重算淨額提示
        self._note = ttk.Label(f, text='交易稅與淨金額將依買賣方向自動重新計算',
                               foreground=C_FG2, font=('Microsoft JhengHei', 8))
        self._note.grid(row=8, column=0, columnspan=2, pady=(4, 0))

        # 按鈕
        btn_f = ttk.Frame(f)
        btn_f.grid(row=9, column=0, columnspan=2, pady=(10, 0))
        ttk.Button(btn_f, text='💾  儲存修改', command=self._save).pack(side='left', padx=6)
        ttk.Button(btn_f, text='✖  取消',      command=self.destroy).pack(side='left', padx=6)

    def _on_side_change(self, *_):
        s = self._vars['side'].get()
        self._side_cb.configure(style='Buy.TCombobox' if s == '買' else 'Sell.TCombobox')

    def _save(self):
        try:
            date_str = self._date_entry.get_date().strftime('%Y-%m-%d')
            code     = self._vars['code'].get().strip().upper()
            name     = self._vars['name'].get().strip()
            cat      = self._vars['category'].get()
            side     = self._vars['side'].get()
            qty_s    = self._vars['qty'].get().strip()
            price_s  = self._vars['price'].get().strip()
            fee_s    = self._vars['fee'].get().strip()

            if not code or not name:
                raise ValueError('請填寫股票代號和名稱')
            if not qty_s or not price_s or not fee_s:
                raise ValueError('請填寫數量、價格與手續費')

            tx_date = datetime.strptime(date_str, '%Y-%m-%d')
            qty, price, fee = float(qty_s), float(price_s), float(fee_s)
            if qty <= 0 or price <= 0:
                raise ValueError('數量和價格必須大於 0')

            gross = qty * price
            tax   = round(gross * TAX_RATE) if side == '賣' else 0
            net   = round(gross - fee - tax) if side == '賣' else -round(gross + fee)

            new_row = {
                '日期':      tx_date,
                '股票代號':  code,
                '股票名稱':  name,
                '分類':      cat,
                '買賣':      side,
                '數量(股)':  int(qty),
                '價格(元)':  price,
                '手續費(元)': int(fee),
                '交易稅(元)': tax,
                '淨金額(元)': net,
            }
            self._on_save(self._idx, new_row)
            self.destroy()

        except ValueError as e:
            messagebox.showerror('輸入錯誤', str(e), parent=self)
        except Exception as e:
            messagebox.showerror('錯誤', str(e), parent=self)


# ─── 啟動 ────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    app = StockApp()
    app.mainloop()