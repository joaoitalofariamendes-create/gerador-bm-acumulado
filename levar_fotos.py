# -*- coding: utf-8 -*-
"""Leva a aba FOTOS ajustada à mão (layout das fotos) de um BM revisado para o BM gerado de novo — a aba inteira:
imagens, grade, mesclas e impressão. A aba FOTOS não tem fórmulas (conferido); vínculo externo, se aparecer, é avisado.

Uso: python levar_fotos.py <pasta revisada> <pasta nova> PRT [PRT...]
"""
import os
import shutil
import sys
import tempfile

import config_mes as C
from comum import Excel, arquivo, pasta_prt


def fotos(wb):
    return next((s for s in wb.Worksheets if s.Name.strip().upper() == "FOTOS"), None)


def main(rev_dir, novo_dir, prts):
    with Excel() as xl:
        for prt in prts:
            # revisado e novo têm o mesmo nome de arquivo, e o Excel não abre dois de mesmo nome: o revisado vai por cópia
            tmp = tempfile.mkdtemp()
            rev = shutil.copy2(arquivo(pasta_prt(rev_dir, prt, C.PASTA_DA_PRT), "bm"), os.path.join(tmp, "revisado.xlsx"))
            wr = xl.abrir(rev, leitura=True)
            wn = xl.abrir(arquivo(pasta_prt(novo_dir, prt, C.PASTA_DA_PRT), "bm"))
            src, velha = fotos(wr), fotos(wn)
            if src is None:
                print(f"!! {prt}: BM revisado sem aba FOTOS — nada levado")
                wr.Close(False)
                wn.Close(False)
                shutil.rmtree(tmp, ignore_errors=True)
                continue
            ult = wn.Sheets(wn.Sheets.Count)
            vis = ult.Visible
            ult.Visible = -1                       # Copy(After=aba oculta) pula as ocultas: última visível na cópia
            antes = {s.Name for s in wn.Sheets}
            src.Copy(None, ult)
            ult.Visible = vis
            nova = next(s for s in wn.Sheets if s.Name not in antes)
            if velha is not None:
                xl.app.DisplayAlerts = False
                velha.Delete()
            nova.Name = src.Name                   # o nome da revisada ('FOTOS', 'fotos '…)
            # a aba traz nomes definidos antigos (_Toc…, _wrn…) que passam a apontar para o arquivo de origem
            for n in list(wn.Names):
                if "revisado.xlsx" in str(n.RefersTo):
                    n.Delete()
            for l in wn.LinkSources(1) or ():
                if "revisado.xlsx" in l:
                    wn.ChangeLink(l, wn.FullName, 1)
            links = wn.LinkSources(1)
            print(f"{prt}: aba FOTOS levada ({nova.Shapes.Count} imagens)" + (f" — !! vínculo externo: {links}" if links else ""))
            wn.Save()
            wn.Close()
            wr.Close(False)
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3:])
