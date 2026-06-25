#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gas_iren_import.py — Confronto consumi gas IREN Premium Top PSV
Ignazio Rusconi-Clerici + ChatGPT + Claude

Versione 2.0.0 — Marzo 2026
  • GUI tkinter: selezione cartella base, log scrollabile, bottoni Esegui/Esci
  • Backup automatico: Confronto consumi gas (Iren).bak.YYYYMMDD.xlsx
  • Modalità INCREMENTALE: se l'Excel esiste, carica i dati già presenti e
    aggiunge solo le bollette nuove (chiave: numero_fattura; fallback:
    localita+periodo_label). Le righe esistenti non vengono mai riscritte.
  • Modalità COMPLETA: se l'Excel non esiste, legge tutti i PDF e crea ex novo.

Struttura attesa:
  BASE_DIR/
    Gignese/*.pdf
    PuntaAla/*.pdf
    Confronto consumi gas (Iren).xlsx   ← creato/aggiornato dallo script

Requisiti:
  pip install pdfplumber pandas xlsxwriter openpyxl
"""

# --- IRC shared bootstrap ---
# Rende disponibili i moduli in Python/shared/ senza dipendere da PYTHONPATH.
# Saltato se eseguito da bundle PyInstaller (sys.frozen=True): in quel caso
# i moduli sono gia' inclusi nel bundle.
import sys as _sys
from pathlib import Path as _Path
if not getattr(_sys, 'frozen', False):
    _shared = _Path.home() / "Library/CloudStorage/Dropbox/Documenti_IRC/Python/shared"
    if str(_shared) not in _sys.path:
        _sys.path.insert(0, str(_shared))
# --- end IRC shared bootstrap ---


import os
import re
import shutil
import logging
import warnings
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import List, Optional, Tuple
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from path_widgets import PathVar, PathEntry, log_path

# Silenzia warning pdfminer
logging.getLogger("pdfminer").setLevel(logging.ERROR)
warnings.filterwarnings("ignore", message=".*invalid float value.*")

import pdfplumber
import pandas as pd
import xlsxwriter  # noqa — necessario per pandas ExcelWriter engine

VERSIONE_SCRIPT = "2.1.0"
VERSION = "2.1.0"
NOME_EXCEL = "Confronto consumi gas (Iren).xlsx"
NUM_IT = r"-?\d{1,3}(?:\.\d{3})*(?:,\d+)?"


# ─────────────────────────────────────────────
#  DATA CLASSES
# ─────────────────────────────────────────────

@dataclass
class RigaGas:
    file_pdf: str
    localita: str
    numero_cliente: Optional[str] = None
    numero_contratto: Optional[str] = None
    numero_fattura: Optional[str] = None
    pdr: Optional[str] = None
    data_emissione: Optional[str] = None
    periodo_da: Optional[str] = None
    periodo_a: Optional[str] = None
    periodo_label: Optional[str] = None
    anno: Optional[int] = None
    mese_num: Optional[int] = None
    smc: Optional[float] = None
    psv: Optional[float] = None
    spread_dichiarato: Optional[float] = None
    adeguamento_pcs: Optional[float] = None
    sconto_percentuale: Optional[float] = None
    sconto_unit: Optional[float] = None
    spread_calcolato: Optional[float] = None
    spread_ok: Optional[str] = None
    quota_fissa_vendita: Optional[float] = None
    quota_consumi_vendita: Optional[float] = None
    totale_materia: Optional[float] = None
    spesa_rete_oneri: Optional[float] = None
    accise_iva: Optional[float] = None
    totale_bolletta: Optional[float] = None
    totale_da_pagare: Optional[float] = None
    note_parsing: Optional[str] = None


@dataclass
class LogEntry:
    file_pdf: str
    localita: str
    esito: str
    messaggio: str
    periodo_label: Optional[str] = None
    smc: Optional[float] = None
    totale_bolletta: Optional[float] = None


# ─────────────────────────────────────────────
#  UTILITÀ NUMERICHE / TESTO
# ─────────────────────────────────────────────

def it_to_float(s: Optional[str]) -> Optional[float]:
    if s is None:
        return None
    s = s.strip()
    if not s:
        return None
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def match_first(pattern: str, text: str, flags=0) -> Optional[str]:
    m = re.search(pattern, text, flags)
    if not m:
        return None
    if m.lastindex:
        return m.group(1)
    return m.group(0)


def normalizza_spazi(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


MESI_ITA = {
    "GENNAIO": ("gen", 1), "FEBBRAIO": ("feb", 2), "MARZO": ("mar", 3),
    "APRILE": ("apr", 4), "MAGGIO": ("mag", 5), "GIUGNO": ("giu", 6),
    "LUGLIO": ("lug", 7), "AGOSTO": ("ago", 8), "SETTEMBRE": ("set", 9),
    "OTTOBRE": ("ott", 10), "NOVEMBRE": ("nov", 11), "DICEMBRE": ("dic", 12),
}


def estrai_periodo(text: str) -> Tuple[Optional[str], Optional[str]]:
    for pat in [
        r"\*\* Periodo di riferimento:\s*([0-9]{2} .*? [0-9]{4})\s*-\s*([0-9]{2} .*? [0-9]{4})",
        r"Periodo di riferimento:\s*([0-9]{2} .*? [0-9]{4})\s*-\s*([0-9]{2} .*? [0-9]{4})",
    ]:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip(), m.group(2).strip()
    return None, None


def periodo_label_anno_mese(s: Optional[str]) -> Tuple[Optional[str], Optional[int], Optional[int]]:
    if not s:
        return None, None, None
    m = re.search(r"\d{2}\s+([A-ZÀÈÉÌÒÙ]+)\s+(\d{4})", s.upper())
    if not m:
        return None, None, None
    mese_raw, anno_raw = m.group(1), m.group(2)
    if mese_raw not in MESI_ITA:
        return None, None, None
    abbrev, mese_num = MESI_ITA[mese_raw]
    anno = int(anno_raw)
    return f"{abbrev}-{str(anno)[-2:]}", anno, mese_num


# ─────────────────────────────────────────────
#  ESTRATTORI DAL TESTO PDF
# ─────────────────────────────────────────────

def estrai_numero_cliente(t):    return match_first(r"Numero cliente\s+(\d+)", t)
def estrai_numero_contratto(t):  return match_first(r"Numero contratto\s+(\d+)", t)
def estrai_numero_fattura(t):    return match_first(r"Fattura n\.\s*([\d]+)", t)
def estrai_data_emissione(t):    return match_first(r"Emessa in data:\s*([0-9]{2} .*? [0-9]{4})", t)
def estrai_pdr(t):               return match_first(r"PDR\s+(\d{8,})", t)
def estrai_totale_bolletta(t):   return it_to_float(match_first(r"Totale bolletta\s+(" + NUM_IT + r")\s*€", t))
def estrai_totale_da_pagare(t):  return it_to_float(match_first(r"Totale da pagare\s+(" + NUM_IT + r")\s*€", t))
def estrai_accise_iva(t):        return it_to_float(match_first(r"Accise e IVA\s+(" + NUM_IT + r")\s*€", t))


def estrai_smc(text: str) -> Optional[float]:
    pat = r"Quota per consumi\s+(" + NUM_IT + r")\s*Smc\s*x\s*" + NUM_IT + r"\s*€/Smc"
    return it_to_float(match_first(pat, text))


def estrai_psv_pcs_spread(text: str) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    psv    = match_first(r"PSV dal .*? al .*?\s+(" + NUM_IT + r")\s*€/smc", text, re.IGNORECASE)
    spread = match_first(r"Prezzo Spread\s+(" + NUM_IT + r")\s*€/smc", text, re.IGNORECASE)
    pcs    = match_first(r"Adeguamento Pcs dal .*? al .*?\s+(" + NUM_IT + r")\s*€/smc", text, re.IGNORECASE)
    return it_to_float(psv), it_to_float(spread), it_to_float(pcs)


def estrai_sconto_percentuale(text: str) -> Optional[float]:
    return it_to_float(match_first(r"Sconto del\s+(" + NUM_IT + r")\s*%\s+su[l]? PSV\+SPREAD", text, re.IGNORECASE))


def estrai_sconto_unitario(text: str) -> Optional[float]:
    return it_to_float(match_first(r"Sconto % prezzo Vendita .*? Euro/smc\s+(" + NUM_IT + r")", text, re.IGNORECASE))


def estrai_materia_importi(text: str) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    tot = match_first(r"Totale di spesa dovuto per l'offerta:\s*(" + NUM_IT + r")\s*€", text, re.IGNORECASE)
    qf  = match_first(r"di cui spesa per la quota fissa:\s*(" + NUM_IT + r")\s*€", text, re.IGNORECASE)
    qc  = match_first(r"di cui spesa per la quota consumi:\s*(" + NUM_IT + r")\s*€", text, re.IGNORECASE)
    return it_to_float(tot), it_to_float(qf), it_to_float(qc)


def estrai_spesa_rete_oneri(text: str) -> Optional[float]:
    rete  = match_first(r"SPESA PER IL TRASPORTO E LA GESTIONE DEL CONTATORE\s+Importo\s+€\s*(" + NUM_IT + r")", text, re.IGNORECASE)
    oneri = match_first(r"SPESA PER ONERI DI SISTEMA\s+Importo\s+€\s*(" + NUM_IT + r")", text, re.IGNORECASE)
    r_v, o_v = it_to_float(rete), it_to_float(oneri)
    if r_v is None and o_v is None:
        return None
    return (r_v or 0.0) + (o_v or 0.0)


# ─────────────────────────────────────────────
#  PARSING DI UNA BOLLETTA
# ─────────────────────────────────────────────

def leggi_testo_pdf(path_pdf: str) -> str:
    with pdfplumber.open(path_pdf) as pdf:
        return "\n".join(p.extract_text() or "" for p in pdf.pages)


def parse_bolletta_gas(path_pdf: str, localita: str) -> RigaGas:
    text_raw = leggi_testo_pdf(path_pdf)
    text = normalizza_spazi(text_raw)
    r = RigaGas(file_pdf=os.path.basename(path_pdf), localita=localita)

    r.numero_cliente   = estrai_numero_cliente(text)
    r.numero_contratto = estrai_numero_contratto(text)
    r.numero_fattura   = estrai_numero_fattura(text)
    r.data_emissione   = estrai_data_emissione(text)
    r.pdr              = estrai_pdr(text)

    periodo_da, periodo_a = estrai_periodo(text_raw)
    r.periodo_da = periodo_da
    r.periodo_a  = periodo_a
    label, anno, mese_num = periodo_label_anno_mese(periodo_a or periodo_da)
    r.periodo_label = label
    r.anno      = anno
    r.mese_num  = mese_num

    r.smc = estrai_smc(text)
    if r.smc is None:
        r.smc = 0.0
    r.psv, r.spread_dichiarato, r.adeguamento_pcs = estrai_psv_pcs_spread(text)
    r.sconto_percentuale = estrai_sconto_percentuale(text)
    r.sconto_unit        = estrai_sconto_unitario(text)

    tot_materia, qf, qc = estrai_materia_importi(text)
    r.quota_fissa_vendita    = qf
    r.quota_consumi_vendita  = qc
    r.totale_materia         = tot_materia
    r.spesa_rete_oneri       = estrai_spesa_rete_oneri(text)
    r.accise_iva             = estrai_accise_iva(text)
    r.totale_bolletta        = estrai_totale_bolletta(text)
    r.totale_da_pagare       = estrai_totale_da_pagare(text)

    # Calcolo spread reale
    spread_ok = "ND"
    if r.sconto_unit is not None and r.sconto_percentuale is not None and r.psv is not None:
        perc = r.sconto_percentuale / 100.0
        if abs(perc) > 1e-9:
            sc = -r.sconto_unit / perc - r.psv
            r.spread_calcolato = sc
            spread_ok = "OK" if abs(sc) < 5e-4 else "NO"
    r.spread_ok = spread_ok

    note = []
    if r.quota_fissa_vendita is not None and abs(r.quota_fissa_vendita - 6.0) > 0.01:
        note.append(f"quota_fissa_vendita != 6 (val={r.quota_fissa_vendita})")
    if note:
        r.note_parsing = "; ".join(note)

    return r


# ─────────────────────────────────────────────
#  CHIAVE DI DEDUPLICAZIONE
# ─────────────────────────────────────────────

def chiave_bolletta(numero_fattura: Optional[str], localita: str, periodo_label: Optional[str]) -> str:
    """Chiave univoca: numero_fattura se disponibile, altrimenti localita+periodo_label."""
    if numero_fattura:
        return f"fattura:{numero_fattura}"
    return f"loc:{localita}|periodo:{periodo_label or 'ND'}"


# ─────────────────────────────────────────────
#  BACKUP
# ─────────────────────────────────────────────

def crea_backup(output_xlsx: str, log_fn) -> Optional[str]:
    """Copia l'Excel esistente con suffisso .bak.YYYYMMDD.xlsx (sovrascrive se già presente)."""
    if not os.path.isfile(output_xlsx):
        return None
    oggi = datetime.now().strftime("%Y%m%d")
    base, _ = os.path.splitext(output_xlsx)
    bak = f"{base}.bak.{oggi}.xlsx"
    shutil.copy2(output_xlsx, bak)
    log_fn(f"[BAK] Backup creato: {os.path.basename(bak)}")
    return bak


# ─────────────────────────────────────────────
#  CARICA DATI ESISTENTI DALL'EXCEL
# ─────────────────────────────────────────────

COLONNE_DATI = [
    "file_pdf", "localita", "numero_cliente", "numero_contratto",
    "numero_fattura", "pdr", "data_emissione", "periodo_da", "periodo_a",
    "periodo_label", "anno", "mese_num", "smc", "psv", "spread_dichiarato",
    "adeguamento_pcs", "sconto_percentuale", "sconto_unit", "spread_calcolato",
    "spread_ok", "quota_fissa_vendita", "quota_consumi_vendita", "totale_materia",
    "spesa_rete_oneri", "accise_iva", "totale_bolletta", "totale_da_pagare",
    "note_parsing",
]

def carica_dati_esistenti(output_xlsx: str, log_fn) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Legge i fogli Gignese e PuntaAla dall'Excel esistente e restituisce
    (df_esistente, df_log_esistente).
    """
    if not os.path.isfile(output_xlsx):
        return pd.DataFrame(), pd.DataFrame()
    try:
        # Leggi Gignese — fallisce solo questo foglio se corrotto
        try:
            df_g  = pd.read_excel(output_xlsx, sheet_name="Gignese",   engine="openpyxl")
        except Exception as e:
            log_fn(f"[WARN] Impossibile leggere foglio Gignese dall'Excel: {e}")
            df_g = pd.DataFrame()

        # Leggi PuntaAla — indipendente da Gignese
        try:
            df_pa = pd.read_excel(output_xlsx, sheet_name="PuntaAla",  engine="openpyxl")
        except Exception as e:
            log_fn(f"[WARN] Impossibile leggere foglio PuntaAla dall'Excel: {e}")
            df_pa = pd.DataFrame()

        df_log = pd.DataFrame()
        try:
            df_log = pd.read_excel(output_xlsx, sheet_name="Log", engine="openpyxl")
        except Exception:
            pass

        # Rinomina le colonne "display" tornando ai nomi interni
        RENAME = {
            "Periodo": "periodo_label", "Periodo_da": "periodo_da", "Periodo_a": "periodo_a",
            "Smc": "smc", "PSV_€/Smc": "psv", "PCS_€/Smc": "adeguamento_pcs",
            "Sconto_%_PSV+Spread": "sconto_percentuale", "Sconto_unit_€/Smc": "sconto_unit",
            "Spread_dichiarato_€/Smc": "spread_dichiarato",
            "Spread_calcolato_€/Smc": "spread_calcolato",
            "Spread_ok": "spread_ok",
            "Quota_fissa_vendita_€": "quota_fissa_vendita",
            "Quota_consumi_vendita_€": "quota_consumi_vendita",
            "Totale_materia_€": "totale_materia",
            "Spesa_rete+oneri_€": "spesa_rete_oneri",
            "Accise+IVA_€": "accise_iva",
            "Totale_bolletta_€": "totale_bolletta",
            "File_pdf": "file_pdf",
        }
        df_g.rename(columns=RENAME, inplace=True)
        df_pa.rename(columns=RENAME, inplace=True)
        df_g["localita"]  = "Gignese"
        df_pa["localita"] = "PuntaAla"

        df_esistente = pd.concat([df_g, df_pa], ignore_index=True)

        # anno e mese_num non sono salvati nell'Excel (colonne interne):
        # li ricalcoliamo da periodo_label (formato "mmm-AA", es. "feb-26")
        ABBREV_TO_NUM = {v[0]: v[1] for v in MESI_ITA.values()}

        def _anno_mese(label):
            if not isinstance(label, str) or "-" not in label:
                return None, None
            parts = label.split("-")
            if len(parts) != 2:
                return None, None
            mese_abbr, anno_aa = parts[0].lower(), parts[1]
            mese_num = ABBREV_TO_NUM.get(mese_abbr)
            try:
                anno = 2000 + int(anno_aa)
            except ValueError:
                return None, None
            return anno, mese_num

        if not df_esistente.empty:
            risultati = df_esistente["periodo_label"].map(_anno_mese).tolist()
            df_esistente["anno"]     = [r[0] for r in risultati]
            df_esistente["mese_num"] = [r[1] for r in risultati]
        else:
            df_esistente["anno"]     = []
            df_esistente["mese_num"] = []

        n = len(df_esistente)
        log_fn(f"[CARICA] Trovate {n} righe esistenti nell'Excel.")
        return df_esistente, df_log
    except Exception as e:
        log_fn(f"[WARN] Impossibile leggere l'Excel esistente: {e}")
        return pd.DataFrame(), pd.DataFrame()


# ─────────────────────────────────────────────
#  SCANSIONE CARTELLE (con filtro incrementale)
# ─────────────────────────────────────────────

def trova_pdf_cartella(cartella: str) -> List[str]:
    if not os.path.isdir(cartella):
        return []
    out = [os.path.join(cartella, n) for n in os.listdir(cartella) if n.lower().endswith(".pdf")]
    out.sort()
    return out


def raccogli_bollette_nuove(
    base_dir: str,
    chiavi_esistenti: set,
    log_fn,
) -> Tuple[List[RigaGas], List[LogEntry]]:
    """
    Scansiona Gignese/ e PuntaAla/ e restituisce solo le bollette
    la cui chiave NON è già in chiavi_esistenti.
    """
    righe: List[RigaGas] = []
    logs:  List[LogEntry] = []

    for localita in ["Gignese", "PuntaAla"]:
        cartella = os.path.join(base_dir, localita)
        pdfs = trova_pdf_cartella(cartella)
        if not pdfs:
            msg = f"Nessun PDF trovato in {localita}/"
            log_fn(f"[INFO] {msg}")
            logs.append(LogEntry(f"(nessun file in {localita})", localita, "INFO", msg))
            continue

        for p in pdfs:
            nome = os.path.basename(p)
            log_fn(f"  → {nome} ...", newline=False)
            try:
                riga = parse_bolletta_gas(p, localita)

                if riga.periodo_da is None and riga.periodo_a is None:
                    log_fn(" SKIP (formato non riconosciuto)")
                    logs.append(LogEntry(nome, localita, "SKIP",
                                         "Periodo non estratto - formato non supportato"))
                    continue

                chiave = chiave_bolletta(riga.numero_fattura, localita, riga.periodo_label)
                if chiave in chiavi_esistenti:
                    log_fn(" GIÀ PRESENTE — saltata")
                    logs.append(LogEntry(nome, localita, "SKIP",
                                         "Bolletta già presente nell'Excel",
                                         riga.periodo_label, riga.smc, riga.totale_bolletta))
                    continue

                righe.append(riga)
                chiavi_esistenti.add(chiave)
                log_fn(f" OK  ({riga.periodo_label}, {riga.smc} Smc, {riga.totale_bolletta} €)")
                logs.append(LogEntry(nome, localita, "OK", "Aggiunta",
                                     riga.periodo_label, riga.smc, riga.totale_bolletta))
            except Exception as e:
                log_fn(f" ERRORE: {e}")
                logs.append(LogEntry(nome, localita, "ERRORE", str(e)))

    return righe, logs


# ─────────────────────────────────────────────
#  COSTRUZIONE DATAFRAME
# ─────────────────────────────────────────────

def df_da_righe(righe: List[RigaGas]) -> pd.DataFrame:
    if not righe:
        return pd.DataFrame(columns=COLONNE_DATI)
    df = pd.DataFrame([asdict(r) for r in righe])
    return df


def unisci_e_ordina(df_esistente: pd.DataFrame, df_nuove: pd.DataFrame) -> pd.DataFrame:
    frames = [f for f in [df_esistente, df_nuove] if not f.empty]
    if not frames:
        return pd.DataFrame(columns=COLONNE_DATI)
    df = pd.concat(frames, ignore_index=True)
    sort_cols = [c for c in ["localita", "anno", "mese_num"] if c in df.columns]
    if sort_cols:
        df = df.sort_values(by=sort_cols)
    return df


def logs_to_dataframe(logs: List[LogEntry]) -> pd.DataFrame:
    if not logs:
        return pd.DataFrame()
    return pd.DataFrame([asdict(l) for l in logs])


# ─────────────────────────────────────────────
#  EXCEL — SCRITTURA FOGLI
# ─────────────────────────────────────────────

HDR_COLOR = "#B7DEE8"


def _header_fmt(wb):
    return wb.add_format({"bold": True, "valign": "vcenter",
                          "fg_color": HDR_COLOR, "border": 1})


# Formati numerici per colonna (nome colonna Excel → formato xlsxwriter)
_COL_FORMATS: dict[str, str] = {
    # Smc e derivati
    "Smc":                    "0.00",
    "Smc_Gignese":            "0.00",
    "Smc_PuntaAla":           "0.00",
    "Smc_totale":             "0.00",
    "Delta_Gig_meno_PA":      "0.00",
    "Smc_12m_Gignese":        "0.00",
    "Smc_12m_PuntaAla":       "0.00",
    "Delta_12m_Gig_meno_PA":  "0.00",
    # Prezzi indicizzati: 4 decimali
    "PSV_€/Smc":              "0.0000",
    "PCS_€/Smc":              "0.0000",
    "Sconto_unit_€/Smc":      "0.0000",
    "Spread_dichiarato_€/Smc":"0.0000",
    "Spread_calcolato_€/Smc": "0.0000",
    # Percentuale sconto (valore numerico es. 15.0)
    "Sconto_%_PSV+Spread":    "0.00",
    # Importi €: 2 decimali
    "Quota_fissa_vendita_€":  "0.00",
    "Quota_consumi_vendita_€":"0.00",
    "Totale_materia_€":       "0.00",
    "Spesa_rete+oneri_€":     "0.00",
    "Accise+IVA_€":           "0.00",
    "Totale_bolletta_€":      "0.00",
    "Totale_bolletta_Gignese":"0.00",
    "Totale_bolletta_PuntaAla":"0.00",
}


def _autofit(ws, df, wb):
    fmt_hdr = _header_fmt(wb)
    # cache dei formati numerici per non ricrearli ogni volta
    _fmt_cache: dict[str, object] = {}

    for ci, col in enumerate(df.columns):
        ws.write(0, ci, col, fmt_hdr)
        w = max([len(str(col))] + [len(str(x)) for x in df[col].astype(str)])
        ws.set_column(ci, ci, min(w, 40) + 2)

        # applica formato numerico alle celle dati se la colonna ha un formato
        num_fmt_str = _COL_FORMATS.get(col)
        if num_fmt_str:
            if num_fmt_str not in _fmt_cache:
                _fmt_cache[num_fmt_str] = wb.add_format({"num_format": num_fmt_str})
            num_fmt = _fmt_cache[num_fmt_str]
            for ri, val in enumerate(df[col]):
                if val is not None and not (isinstance(val, float) and (val != val)):
                    try:
                        ws.write(ri + 1, ci, float(val), num_fmt)
                    except (TypeError, ValueError):
                        pass  # testo o None: lascia come scritto da pandas


def scrivi_foglio_localita(writer, df_all: pd.DataFrame, localita: str):
    df_loc = df_all[df_all["localita"] == localita].copy() if not df_all.empty else pd.DataFrame()
    COLS = [
        ("periodo_label", "Periodo"), ("periodo_da", "Periodo_da"), ("periodo_a", "Periodo_a"),
        ("smc", "Smc"), ("psv", "PSV_€/Smc"), ("adeguamento_pcs", "PCS_€/Smc"),
        ("sconto_percentuale", "Sconto_%_PSV+Spread"), ("sconto_unit", "Sconto_unit_€/Smc"),
        ("spread_dichiarato", "Spread_dichiarato_€/Smc"),
        ("spread_calcolato", "Spread_calcolato_€/Smc"), ("spread_ok", "Spread_ok"),
        ("quota_fissa_vendita", "Quota_fissa_vendita_€"),
        ("quota_consumi_vendita", "Quota_consumi_vendita_€"),
        ("totale_materia", "Totale_materia_€"),
        ("spesa_rete_oneri", "Spesa_rete+oneri_€"),
        ("accise_iva", "Accise+IVA_€"),
        ("totale_bolletta", "Totale_bolletta_€"),
        ("file_pdf", "File_pdf"),
    ]
    if df_loc.empty:
        out = pd.DataFrame(columns=[n for _, n in COLS])
    else:
        out = pd.DataFrame({n: df_loc[o] for o, n in COLS if o in df_loc.columns})

    out.to_excel(writer, sheet_name=localita, index=False)
    _autofit(writer.sheets[localita], out, writer.book)


def scrivi_foglio_confronto_mensile(writer, df_all: pd.DataFrame):
    if df_all.empty:
        df = pd.DataFrame(columns=["Periodo", "Smc_Gignese", "Smc_PuntaAla",
                                   "Smc_totale", "Delta_Gig_meno_PA",
                                   "Totale_bolletta_Gignese", "Totale_bolletta_PuntaAla"])
    else:
        grp = df_all.groupby(["localita", "periodo_label"], dropna=True)
        agg = grp.agg(smc=("smc", "sum"), totale_bolletta=("totale_bolletta", "sum"),
                      anno=("anno", "first"), mese_num=("mese_num", "first")).reset_index()
        pivo_smc = agg.pivot(index=["periodo_label", "anno", "mese_num"], columns="localita", values="smc")
        pivo_tot = agg.pivot(index=["periodo_label", "anno", "mese_num"], columns="localita", values="totale_bolletta")
        df = pd.DataFrame(index=pivo_smc.index)
        df["Periodo"]                  = df.index.get_level_values("periodo_label")
        df["Smc_Gignese"]              = pivo_smc.get("Gignese")
        df["Smc_PuntaAla"]             = pivo_smc.get("PuntaAla")
        df["Smc_totale"]               = df["Smc_Gignese"].fillna(0) + df["Smc_PuntaAla"].fillna(0)
        df["Delta_Gig_meno_PA"]        = df["Smc_Gignese"].fillna(0) - df["Smc_PuntaAla"].fillna(0)
        df["Totale_bolletta_Gignese"]  = pivo_tot.get("Gignese")
        df["Totale_bolletta_PuntaAla"] = pivo_tot.get("PuntaAla")
        df = df.sort_values(by=["anno", "mese_num"]).reset_index(drop=True)
    df.to_excel(writer, sheet_name="Confronto Mensile", index=False)
    _autofit(writer.sheets["Confronto Mensile"], df, writer.book)


def scrivi_foglio_confronto_annuale(writer, df_all: pd.DataFrame):
    if df_all.empty:
        df = pd.DataFrame(columns=["Periodo", "Smc_12m_Gignese",
                                   "Smc_12m_PuntaAla", "Delta_12m_Gig_meno_PA"])
    else:
        base = df_all.dropna(subset=["anno", "mese_num", "periodo_label"]).copy()
        agg  = base.groupby(["localita", "anno", "mese_num", "periodo_label"]).agg(
                    smc=("smc", "sum")).reset_index()
        agg  = agg.sort_values(["localita", "anno", "mese_num"])
        agg["Smc_12m"] = (agg.groupby("localita")["smc"]
                           .rolling(12, min_periods=1).sum()
                           .reset_index(level=0, drop=True))
        pivo = agg.pivot(index=["anno", "mese_num", "periodo_label"],
                         columns="localita", values="Smc_12m")
        df = pd.DataFrame(index=pivo.index)
        df["Periodo"]           = df.index.get_level_values("periodo_label")
        df["Smc_12m_Gignese"]   = pivo.get("Gignese")
        df["Smc_12m_PuntaAla"]  = pivo.get("PuntaAla")
        df["Delta_12m_Gig_meno_PA"] = (df["Smc_12m_Gignese"].fillna(0)
                                        - df["Smc_12m_PuntaAla"].fillna(0))
        df = df.reset_index(drop=True)
    df.to_excel(writer, sheet_name="Confronto Annuale", index=False)
    _autofit(writer.sheets["Confronto Annuale"], df, writer.book)


def scrivi_foglio_log(writer, df_log: pd.DataFrame):
    if df_log.empty:
        df_log = pd.DataFrame(columns=["file_pdf", "localita", "esito",
                                        "messaggio", "periodo_label", "smc", "totale_bolletta"])
    df_log.to_excel(writer, sheet_name="Log", index=False)
    _autofit(writer.sheets["Log"], df_log, writer.book)


def salva_excel(df_all: pd.DataFrame, df_log: pd.DataFrame, output_xlsx: str):
    with pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
        scrivi_foglio_localita(writer, df_all, "Gignese")
        scrivi_foglio_localita(writer, df_all, "PuntaAla")
        scrivi_foglio_confronto_mensile(writer, df_all)
        scrivi_foglio_confronto_annuale(writer, df_all)
        scrivi_foglio_log(writer, df_log)


# ─────────────────────────────────────────────
#  LOGICA PRINCIPALE
# ─────────────────────────────────────────────

def esegui(base_dir: str, log_fn):
    """
    Funzione principale richiamata dalla GUI (in un thread separato).
    log_fn(msg, newline=True) scrive nel pannello log.
    """
    output_xlsx = os.path.join(base_dir, NOME_EXCEL)
    log_fn(f"[START] Script v{VERSIONE_SCRIPT}")
    log_fn(f"[DIR]   {base_dir}")

    # 1. Backup se l'Excel esiste
    crea_backup(output_xlsx, log_fn)

    # 2. Carica dati già presenti
    df_esistente, df_log_esistente = carica_dati_esistenti(output_xlsx, log_fn)

    # 3. Costruisci insieme chiavi già presenti
    chiavi_esistenti: set = set()
    if not df_esistente.empty:
        for _, row in df_esistente.iterrows():
            k = chiave_bolletta(
                str(row.get("numero_fattura", "")) if pd.notna(row.get("numero_fattura")) else None,
                str(row.get("localita", "")),
                str(row.get("periodo_label", "")) if pd.notna(row.get("periodo_label")) else None,
            )
            chiavi_esistenti.add(k)
        log_fn(f"[INFO]  Chiavi esistenti: {len(chiavi_esistenti)}")

    # 4. Scansiona PDF e prendi solo le bollette nuove
    log_fn("[SCAN]  Scansione PDF...")
    righe_nuove, logs_nuovi = raccogli_bollette_nuove(base_dir, chiavi_esistenti, log_fn)
    log_fn(f"[INFO]  Bollette nuove trovate: {len(righe_nuove)}")

    # 5. Unisci dati
    df_nuove = df_da_righe(righe_nuove)
    df_all   = unisci_e_ordina(df_esistente, df_nuove)

    # Log: accoda nuovi log a quelli storici
    df_log_nuovi = logs_to_dataframe(logs_nuovi)
    df_log_all   = pd.concat([f for f in [df_log_esistente, df_log_nuovi] if not f.empty],
                              ignore_index=True)

    # 6. Salva
    salva_excel(df_all, df_log_all, output_xlsx)
    n_tot = len(df_all)
    log_fn(f"[OK]    Excel aggiornato: {os.path.basename(output_xlsx)}  ({n_tot} righe totali)")

    # 7. Riepilogo spread
    if not df_all.empty and "spread_ok" in df_all.columns:
        n_ok  = int((df_all["spread_ok"] == "OK").sum())
        n_tot2 = int(len(df_all))
        log_fn(f"[INFO]  Spread OK: {n_ok}/{n_tot2} bollette compatibili con spread=0")

    log_fn("[FINE]  Elaborazione completata.")


# ─────────────────────────────────────────────
#  PERSISTENZA CARTELLA (INI)
# ─────────────────────────────────────────────

import json

CONFIG_DIR  = os.path.join(os.path.expanduser("~"), "Library", "CloudStorage",
                           "Dropbox", "Documenti_IRC", "Python", "_config", "GasIrenImport")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")



def _path_to_cfg(p) -> str:
    """Salva il path relativo alla home, per portabilità tra Mac."""
    try:
        return str(_Path(str(p)).expanduser().resolve().relative_to(_Path.home()))
    except ValueError:
        return str(p)

def _path_from_cfg(s: str) -> str:
    """Ricostruisce il path assoluto: se relativo, prepende _Path.home()."""
    if not s:
        return s
    p = _Path(s)
    if p.is_absolute():
        # retrocompatibilità: path assoluto vecchio stile
        return str(p)
    return str(_Path.home() / p)

def carica_ultima_cartella() -> str:
    if os.path.isfile(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            cartella = _path_from_cfg(data.get("ultima_cartella", ""))
            if cartella and os.path.isdir(cartella):
                return cartella
        except Exception:
            pass
    return os.path.expanduser("~")


def salva_ultima_cartella(cartella: str):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    # Preserva eventuali chiavi già presenti
    data = {}
    if os.path.isfile(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass
    data["ultima_cartella"] = _path_to_cfg(cartella)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ─────────────────────────────────────────────
#  GUI TKINTER
# ─────────────────────────────────────────────

def avvia_gui():
    root = tk.Tk()
    root.title(f"Gas Iren Import  v{VERSIONE_SCRIPT}")
    root.resizable(True, True)

    # ── Riga 1: cartella base ──
    frm_path = ttk.Frame(root, padding=8)
    frm_path.grid(row=0, column=0, sticky="ew")
    root.columnconfigure(0, weight=1)

    ttk.Label(frm_path, text="Cartella base:").grid(row=0, column=0, sticky="w")

    # var_path_full: percorso completo (usato per elaborazione e INI)
    # var_path_display: versione abbreviata mostrata nel campo
    var_path_full = carica_ultima_cartella()

    def abbrevia_percorso(p: str) -> str:
        """Abbrevia da Dropbox/ se presente, altrimenti da ~/"""
        try:
            idx = p.find("Dropbox/")
            if idx != -1:
                return "Dropbox/" + p[idx + len("Dropbox/"):]
            home = os.path.expanduser("~")
            if p.startswith(home):
                return "~" + p[len(home):]
        except Exception:
            pass
        return p

    var_path = PathVar(value=var_path_full)
    entry_path = PathEntry(frm_path, pathvar=var_path)
    entry_path.grid(row=0, column=1, padx=6, sticky="ew")
    frm_path.columnconfigure(1, weight=1)

    def sfoglia():
        nonlocal var_path_full
        start = var_path_full if os.path.isdir(var_path_full) else os.path.expanduser("~")
        d = filedialog.askdirectory(title="Seleziona cartella base (con Gignese/ e PuntaAla/)", initialdir=start)
        if d:
            var_path_full = d
            var_path.set(d)

    ttk.Button(frm_path, text="Sfoglia…", command=sfoglia).grid(row=0, column=2)

    # ── Riga 2: area log ──
    frm_log = ttk.LabelFrame(root, text="Log", padding=6)
    frm_log.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)
    root.rowconfigure(1, weight=1)
    frm_log.columnconfigure(0, weight=1)
    frm_log.rowconfigure(0, weight=1)

    txt_log = scrolledtext.ScrolledText(frm_log, width=80, height=20,
                                         state="disabled", font=("Menlo", 11))
    txt_log.grid(row=0, column=0, sticky="nsew")

    def log_fn(msg: str, newline: bool = True):
        """Thread-safe: se chiamata da un worker thread, marshalla via root.after."""
        def _append():
            txt_log.configure(state="normal")
            txt_log.insert(tk.END, msg + ("\n" if newline else ""))
            txt_log.see(tk.END)
            txt_log.configure(state="disabled")
            # Nessuna chiamata a update() o update_idletasks() — non necessaria
            # e pericolosa se chiamata dal thread UI (rientranza loop eventi)
        if threading.current_thread() is threading.main_thread():
            _append()
        else:
            root.after(0, _append)

    # ── Riga 3: bottoni ──
    frm_btn = ttk.Frame(root, padding=8)
    frm_btn.grid(row=2, column=0, sticky="e")

    btn_esegui = ttk.Button(frm_btn, text="▶  Esegui", width=14)
    btn_esegui.grid(row=0, column=0, padx=6)
    ttk.Button(frm_btn, text="Esci", width=10,
               command=root.destroy).grid(row=0, column=1)

    def on_esegui():
        base_dir = var_path_full
        if not os.path.isdir(base_dir):
            log_fn(f"[ERR] Cartella non valida: {log_path(base_dir)}")
            return
        salva_ultima_cartella(base_dir)
        btn_esegui.configure(state="disabled")
        log_fn("─" * 60)

        def _run():
            try:
                esegui(base_dir, log_fn)
            except Exception as ex:
                log_fn(f"[ERRORE CRITICO] {ex}")
            finally:
                # Re-abilita il pulsante dal thread UI
                root.after(0, lambda: btn_esegui.configure(state="normal"))

        threading.Thread(target=_run, daemon=True).start()

    btn_esegui.configure(command=on_esegui)

    root.minsize(680, 460)
    root.mainloop()


# ─────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    avvia_gui()
