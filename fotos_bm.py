# -*- coding: utf-8 -*-
"""Preenche a aba FOTOS (relatório fotográfico) do BM de abas com fotos do mês.

Reaproveita a grade que já existe na aba: cada foto antiga dentro da área de impressão vira uma
"vaga" (posição e tamanho); logos do cabeçalho (fotos baixas, < 60 pt) ficam. As fotos antigas saem,
as novas entram nas mesmas vagas, ajustadas à proporção e centralizadas. As fotos são escolhidas
espalhadas pelos dias (uma de cada dia antes de repetir dia) e reduzidas antes de entrar.
Uso: python fotos_bm.py <BM.xlsx> <pasta de fotos da PRT (subpastas por dia)>
"""
import glob
import os
import sys
import tempfile

from PIL import Image, ImageOps

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from comum import Excel

LADO_MAX = 1200          # px no lado maior (qualidade de impressão sem inchar o arquivo)
LOGO_ALTURA = 60         # pt: abaixo disso é logo de cabeçalho, não foto


def escolher(pasta, n):
    """n fotos espalhadas: percorre os dias em rodízio (1ª foto de cada dia, depois a 2ª...)."""
    dias = sorted(d for d in glob.glob(os.path.join(pasta, "*")) if os.path.isdir(d)) + [pasta]  # soltas = 1 "dia"
    filas = [sorted(glob.glob(os.path.join(d, "*.jp*g")) + glob.glob(os.path.join(d, "*.png"))) for d in dias]
    filas = [f for f in filas if f]
    # dentro do dia, pega fotos espaçadas (evita 5 fotos seguidas da mesma cena)
    filas = [f[::max(1, len(f) // 3)] + [x for x in f if x not in f[::max(1, len(f) // 3)]] for f in filas]
    out, i = [], 0
    while len(out) < n and any(filas):
        f = filas[i % len(filas)]
        if f:
            out.append(f.pop(0))
        i += 1
    return out


def reduzir(src, tmp):
    im = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
    im.thumbnail((LADO_MAX, LADO_MAX))
    dst = os.path.join(tmp, os.path.basename(src))
    im.save(dst, "JPEG", quality=80, optimize=True)
    return dst, im.size


def main(bm, pasta, modelo=None):
    """modelo: outro BM com aba FOTOS — se este não tiver, a aba é criada copiando a de lá."""
    log = []
    with Excel() as xl, tempfile.TemporaryDirectory() as tmp:
        wb = xl.abrir(bm)
        aba = next((s for s in wb.Worksheets if s.Name.strip().upper() == "FOTOS"), None)
        if aba is None and modelo:
            wm = xl.abrir(modelo, leitura=True)
            fm = next(s for s in wm.Worksheets if s.Name.strip().upper() == "FOTOS")
            antes = {s.Name for s in wb.Sheets}
            fm.Copy(None, wb.Sheets(wb.Sheets.Count))             # posicional: After=última aba
            wm.Close(False)
            novas = [s for s in wb.Sheets if s.Name not in antes]
            if len(novas) != 1:
                raise SystemExit(f"{bm}: não identifiquei a aba FOTOS copiada ({[s.Name for s in novas]})")
            aba = novas[0]
            ult = wb.Sheets(wb.Sheets.Count)                         # Move(After=aba oculta) pula as ocultas e
            vis_ult = ult.Visible                                    # largava a FOTOS no meio (5028): última
            ult.Visible = -1                                         # aba visível durante o Move
            aba.Move(None, ult)                                      # garante no fim
            ult.Visible = vis_ult
            aba.Name = "FOTOS"
            aba.Visible = -1
            log.append(f"   CRIADA aba FOTOS (cópia da aba do {os.path.basename(modelo)})")
        if aba is None:
            raise SystemExit(f"{bm}: sem aba FOTOS")
        area = aba.Range(aba.PageSetup.PrintArea) if aba.PageSetup.PrintArea else aba.UsedRange
        lim_dir, lim_baixo = area.Left + area.Width, area.Top + area.Height
        vagas, apagar = [], []
        for sh in list(aba.Shapes):
            if sh.Type != 13:                       # só imagens
                continue
            dentro = sh.Left < lim_dir and sh.Top < lim_baixo
            if sh.Height < LOGO_ALTURA:
                continue                            # logo (dentro ou fora da área): nunca mexe
            if dentro:
                vagas.append((sh.Top, sh.Left, sh.Width, sh.Height))
            apagar.append(sh)                       # foto antiga (dentro) ou perdida fora da área
        vagas.sort()
        fotos = escolher(pasta, len(vagas))
        for sh in apagar:
            sh.Delete()
        for (top, left, w, h), f in zip(vagas, fotos):
            arq, (pw, ph) = reduzir(f, tmp)
            esc = min(w / pw, h / ph)
            fw, fh = pw * esc, ph * esc
            aba.Shapes.AddPicture(arq, False, True, left + (w - fw) / 2, top + (h - fh) / 2, fw, fh)
        # sobraram páginas só com vagas vazias (ex.: modelo de 33 vagas para 8 fotos): ficam fora da impressão
        if fotos and len(fotos) < len(vagas) and aba.PageSetup.PrintArea:
            ult = max(v[0] for v in vagas[:len(fotos)])
            quebras = [aba.HPageBreaks(i).Location for i in range(1, aba.HPageBreaks.Count + 1)]
            fim = next((q.Row - 1 for q in quebras if q.Top > ult), None)
            a = aba.Range(aba.PageSetup.PrintArea)
            if fim and fim < a.Row + a.Rows.Count - 1:
                aba.PageSetup.PrintArea = aba.Range(aba.Cells(a.Row, a.Column),
                                                    aba.Cells(fim, a.Column + a.Columns.Count - 1)).Address
                log.append(f"   área de impressão até a linha {fim}: as páginas só com vagas vazias ficam fora do book")
        wb.Save()
        wb.Close()
    log.append(f"{os.path.basename(bm)}: {len(vagas)} vagas, {len(fotos)} fotos inseridas "
               f"({len(apagar)} imagens antigas retiradas, logos mantidos)")
    log += [f"   {os.path.basename(f)}" for f in fotos]
    if len(fotos) < len(vagas):
        log.append(f"   !! só havia {len(fotos)} fotos para {len(vagas)} vagas — as vagas restantes ficaram vazias")
    print("\n".join(log))
    return log


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
