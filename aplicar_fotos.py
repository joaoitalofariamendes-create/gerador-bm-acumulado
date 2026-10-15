# -*- coding: utf-8 -*-
"""Passo depois do aplicar_bm: aba FOTOS de cada BM de abas em SAIDA com as fotos da PRT no mês.

Fotos: pasta gerada pelo fotos_whatsapp.py (config_mes.FOTOS, subpastas 'PRT nnnn\\dd-mm').
PRT sem fotos no mês: o BM fica como está (listado no relatório).
Uso: python aplicar_fotos.py [PRT...]
"""
import glob
import os
import re
import sys

import config_mes as C
from comum import arquivo
from fotos_bm import main as preencher_fotos


def modelo_fotos():
    """BM do mês anterior com aba FOTOS (a mais comum: 1 página de 15 fotos) — para PRT que não tem a aba."""
    import openpyxl
    for p in sorted(glob.glob(os.path.join(C.FONTE_MEDICAO, "PRT*"))):
        bm = arquivo(p, "bm")
        if not bm:
            continue
        w = openpyxl.load_workbook(bm, read_only=True)
        tem = any(n.strip().upper() == "FOTOS" for n in w.sheetnames)
        w.close()                                   # read_only prende o arquivo
        if tem:
            return bm
    return None


def ocultar_fotos(bm):
    """PRT sem fotos no mês: aba FOTOS visível mostraria as fotos do mês passado — fica oculta."""
    import openpyxl
    w = openpyxl.load_workbook(bm, read_only=True)
    vis = [s.title for s in w.worksheets if s.title.strip().upper().startswith("FOTOS") and s.sheet_state == "visible"]
    w.close()                                       # read_only prende o arquivo
    if not vis:
        return False
    from comum import Excel
    with Excel() as xl:
        wb = xl.abrir(bm)
        next(s for s in wb.Worksheets if s.Visible == -1 and s.Name not in vis).Activate()
        for n in vis:
            wb.Worksheets(n).Visible = 0
        wb.Save()
        wb.Close()
    return True


def main():
    so = [a for a in sys.argv[1:] if re.fullmatch(r"\d{4}", a)]
    fotos = getattr(C, "FOTOS", None)
    rel = [f"FOTOS BM {C.BM_NUMERO} — fotos de {fotos}"]
    for pasta in sorted(glob.glob(os.path.join(C.SAIDA, "PRT*"))):
        prt = re.search(r"(\d{4})", os.path.basename(pasta)).group(1)
        if so and prt not in so:
            continue
        bm = arquivo(pasta, "bm")
        origem = os.path.join(fotos or "", f"PRT {prt}")
        if not bm:
            continue
        tem = [f for p in ("*.jp*g", "*.png", os.path.join("*", "*.jp*g"), os.path.join("*", "*.png"))
               for f in glob.glob(os.path.join(origem, p))]
        if not fotos or not tem:
            rel.append(f"PRT {prt}: sem fotos no mês — aba FOTOS oculta (só o que é do mês fica visível)"
                       if ocultar_fotos(bm) else f"PRT {prt}: sem fotos no mês")
            continue
        try:
            criar = prt in getattr(C, "FOTOS_CRIAR_ABA", set())   # aba FOTOS só é criada para PRT listada
            rel += preencher_fotos(bm, origem, modelo_fotos() if criar else None)
        except (SystemExit, Exception) as e:
            rel.append(f"!! PRT {prt}: {e}")
    with open(os.path.join(C.SAIDA, "relatorio_fotos_bm.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(rel))
    print("\n".join(r for r in rel if not r.startswith("   ")))


if __name__ == "__main__":
    main()
