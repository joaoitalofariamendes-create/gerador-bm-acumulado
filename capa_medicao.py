# -*- coding: utf-8 -*-
"""Capa da Medição (Word) para cada PRT do mês: SAIDA\\<pasta da PRT>\\CAPA DA MEDIÇÃO - PRT nnnn.docx.

Modelo = 1ª página do CAPA DA MEDIÇÃO.docx (o resto é apagado). Preenche:
  proposta (I6 da Medição), título (\"Título da PRT\" da proposta em cache_md; senão E6 da Medição),
  período (config_mes), ID / NIQ PW / CC (C6, D6, G6 da Medição; vazio = N/A).
Uso: python capa_medicao.py [PRT...]
"""
import glob
import os
import re
import sys

import openpyxl
import win32com.client as win32

import config_mes as C

MODELO = r"<PASTA_TRABALHO>\Medicoes\CAPA DA MEDIÇÃO.docx"
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache_md")

# o que não sai da Medição nem do cache da proposta (PDF escaneado etc.)
TITULO = {"5043": "PROTEÇÃO DE PARTES MÓVEIS ROTATIVAS",                         # capa do BM 5030
          "5001": "MONTAGEM ELETROMECÂNICA, AUTOMAÇÃO E COMISSIONAMENTO - MONIT. DE GASES E PARTICULADOS"}
PROPOSTA = {"5001": "PROPOSTA_PRT-GSV-DOP-B-25 – 5001 - R.3"}                     # com a revisão (capa do BM 5029)
CAMPOS = {"5028": {"CC": "XXXX.XX.0000.XX00.XX00"}}   # CC da proposta aprovada (Medição em branco)

# textos da 1ª página do modelo (PRT 5022) que são trocados
M_PROPOSTA = "PROPOSTA_PRT-GSV-DOP-B-26 – 5022"
M_TITULO = "FORNECIMENTO DE MÃO DE OBRA INDIRETOS E EQUIPAMENTOS"
M_PERIODO = "06/04/2026 A 05/05/2026"


def titulo_proposta(pasta):
    for md in glob.glob(os.path.join(CACHE, pasta, "*.md")):
        linhas = open(md, encoding="utf-8").read().splitlines()
        for i, l in enumerate(linhas[:-1]):
            if l.strip() == "Título da PRT:" and linhas[i + 1].strip():
                return linhas[i + 1].strip()
    return None


def dados(med):
    ws = openpyxl.load_workbook(med, data_only=True)["Formuário_PRT"]
    prt = re.search(r"(\d{4})\.xlsx$", med).group(1)
    pasta = os.path.basename(os.path.dirname(med))
    m = re.search(r"B-(\d{2})-(\d{4})", str(ws["I6"].value or ""))
    ano = m.group(1) if m else "26"
    d = {"ID": ws["C6"].value, "NIQ PW": ws["D6"].value, "CC": ws["G6"].value}
    d = {k: str(v).strip() if v not in (None, "") else "N/A" for k, v in d.items()}
    d.update(CAMPOS.get(prt, {}))
    tit = TITULO.get(prt) or titulo_proposta(pasta) or str(ws["E6"].value).strip()
    prop = PROPOSTA.get(prt) or f"PROPOSTA_PRT-GSV-DOP-B-{ano} – {prt}"
    return prt, pasta, prop, tit, d


def trocar(doc, de, para):
    """Troca em todo o documento, inclusive caixas de texto (a capa é feita de caixas de texto)."""
    achou = 0
    for story in doc.StoryRanges:
        while story is not None:
            rng = story.Duplicate
            f = rng.Find
            f.ClearFormatting()
            while f.Execute(FindText=de, MatchCase=True, Forward=True, Wrap=0):
                rng.Text = para
                rng.Collapse(0)                      # wdCollapseEnd
                achou += 1
            story = story.NextStoryRange
    if not achou:
        raise RuntimeError(f"não achei no modelo: {de!r}")


def main(filtro):
    per = f"{C.PERIODO_INI:%d/%m/%Y} A {C.PERIODO_FIM:%d/%m/%Y}"
    meds = sorted(glob.glob(os.path.join(C.SAIDA, "PRT*", "Medicao_PRT_*.xlsx")))
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for med in meds:
            prt, pasta, prop, tit, d = dados(med)
            if filtro and prt not in filtro:
                continue
            doc = word.Documents.Open(MODELO, ReadOnly=True, AddToRecentFiles=False)
            try:
                ini2 = doc.GoTo(1, 1, 2).Start          # wdGoToPage, wdGoToAbsolute, página 2
                doc.Range(ini2 - 1, doc.Content.End).Delete()   # -1 leva junto a quebra de página
                trocar(doc, M_PROPOSTA, prop)
                trocar(doc, M_TITULO, tit)
                trocar(doc, M_PERIODO, per)
                for k, v in d.items():
                    trocar(doc, f"{k}: N/A", f"{k}: {v}")
                if doc.ComputeStatistics(2) != 1:       # wdStatisticPages
                    raise RuntimeError(f"PRT {prt}: capa ficou com {doc.ComputeStatistics(2)} páginas")
                dst = os.path.join(C.SAIDA, pasta, f"CAPA DA MEDIÇÃO - PRT {prt}.docx")
                doc.SaveAs2(dst, FileFormat=16)          # wdFormatXMLDocument
                print(f"PRT {prt}: {prop} | {tit} | " + " | ".join(f"{k}: {v}" for k, v in d.items()))
            finally:
                doc.Close(SaveChanges=0)
    finally:
        word.Quit()


if __name__ == "__main__":
    main(set(sys.argv[1:]))
