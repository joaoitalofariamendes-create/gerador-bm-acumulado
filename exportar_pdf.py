# -*- coding: utf-8 -*-
"""PDFs do BM para entrega, ao lado de cada .xlsx em SAIDA\\PRT xxxx\\ (mesmo nome, .pdf):
  Medição  -> só a aba Formuário_PRT;  BM Acumulado -> só a aba do Acumulado;
  BM de abas -> a pasta de trabalho inteira (o Excel exporta só as abas visíveis).
Abre só para leitura e fecha sem salvar. Uso: python exportar_pdf.py [PRT...]  (sem PRT = todas)"""
import glob
import os
import re
import sys

import config_mes as C
from comum import Excel, arquivo

XL_PDF = 0
# campo de assinatura (vazio) na posição em que o Elizeu assinou os PDFs da 5001: deslocamento a partir de um texto
# da página (dx, dy da borda inferior esquerda do texto; largura, altura; altura do texto na 5001 — o Excel encolhe
# cada página para caber, então tudo escala pela altura do texto). BM de abas: um em cada página.
# Acumulado: o campo ocupa o quadro 'Emissão Contratada' (de 'Aprovado por' até 'Data:', até a coluna seguinte).
ASSINATURA = {"medicao": ("Fornecedor", -155.4, -35.8, 335.8, 29.9, 5.6, "primeira"),
              "acumulado": ("Emissão", None, None, None, None, None, "ultima"),
              "bm": ("Emissão", -32.4, -44.6, 124.3, 42.0, 5.9, "todas")}


def quadro_emissao(pg, a):
    """Retângulo (x0, y0, x1, y1 em coordenadas PDF) do quadro de assinatura da Contratada no Acumulado."""
    ws = pg.extract_words()
    aprov = [w for w in ws if w["text"].startswith("Aprovado") and w["top"] > a["top"]]
    data = [w for w in ws if w["text"].startswith("Data") and w["top"] > a["top"]]
    linha = min(aprov, key=lambda w: w["top"])["top"]
    cols = sorted(w["x0"] for w in aprov if abs(w["top"] - linha) < 2)
    prox = next((x for x in cols if x > a["x0"] + 20), a["x0"] + 250)
    topo = min(w["bottom"] for w in aprov if abs(w["x0"] - a["x0"]) < 5)
    base = min(w["bottom"] for w in data if abs(w["x0"] - a["x0"]) < 5)
    return a["x0"] - 1, pg.height - base, prox - 2, pg.height - topo


def posicoes(pdf, tipo):
    """[(página, x0, y0, x1, y1)] dos campos de assinatura (coordenadas PDF)."""
    import pdfplumber
    texto, dx, dy, w, h, ref, quais = ASSINATURA[tipo]
    alvos = []
    with pdfplumber.open(pdf) as d:
        for i, pg in enumerate(d.pages):
            achados = [x for x in pg.extract_words() if x["text"].startswith(texto)]
            if not achados:
                continue
            a = max(achados, key=lambda x: x["bottom"])     # o do quadro de assinatura (rodapé), não outro no texto
            if tipo == "acumulado":
                alvos.append((i, *quadro_emissao(pg, a)))
            else:
                s = (a["bottom"] - a["top"]) / ref
                x, y = a["x0"] + dx * s, pg.height - a["bottom"] + dy * s
                alvos.append((i, x, y, x + w * s, y + h * s))
    return alvos[:1] if quais == "primeira" else alvos[-1:] if quais == "ultima" else alvos


# BM de abas: o Foxit assina "todas as páginas" no mesmo ponto -> cada página é DESLOCADA (sem escala) para que o canto
# superior esquerdo do quadro 'Emissão (CONTRATADA)' (borda esquerda x linha abaixo do título) caia no mesmo ponto em
# todas as páginas. O ponto é escolhido por PDF: o mais perto da posição mediana que não tira nada da folha.
MARGEM = 8


def canto_quadro(pg):
    """(x, y) em coordenadas PDF do canto superior esquerdo do quadro de assinatura do BM de abas, ou None.
    Borda esquerda = a linha vertical mais à direita antes de '(CONTRATADA)' (a área de impressão às vezes corta o
    começo do título, e aí a borda visível é a da folha); topo = linha horizontal logo abaixo do título."""
    ws = pg.extract_words()
    em = [w for w in ws if w["text"].startswith("Emissão")]
    if not em:
        return None
    a = max(em, key=lambda w: w["bottom"])
    ct = next((w for w in ws if w["text"].startswith("(CONTRATADA") and abs(w["top"] - a["top"]) < 2), a)
    lim = ct["x0"] if ct is not a else a["x1"]
    hz = [e for e in pg.edges if e["orientation"] == "h" and e["top"] >= a["bottom"] - 0.5
          and e["x0"] <= lim and e["x1"] >= lim and e["top"] - a["bottom"] < 8]
    topo = min(e["top"] for e in hz) if hz else a["bottom"] + 1
    vt = [e["x0"] for e in pg.edges if e["orientation"] == "v" and 0 <= e["x0"] < lim
          and e["top"] <= topo + 3 and e["bottom"] >= topo + 10]
    if not vt:
        return None
    return max(vt), pg.height - topo


def alinhar_quadros(pdf):
    """Desloca as páginas do BM de abas para o quadro de assinatura ficar no mesmo lugar. Devolve (alinhadas, puladas, alvo)."""
    import statistics
    import pdfplumber
    from pypdf import PdfReader, PdfWriter, Transformation

    pags = {}
    with pdfplumber.open(pdf) as d:
        for i, pg in enumerate(d.pages):
            c = canto_quadro(pg)
            if not c:
                continue
            W, H = pg.width, pg.height
            objs = [o for o in pg.chars + pg.rects + pg.lines + pg.images
                    if o["x1"] > 0 and o["x0"] < W and o["bottom"] > 0 and o["top"] < H]
            cx0, cx1 = max(0, min(o["x0"] for o in objs)), min(W, max(o["x1"] for o in objs))
            cy0, cy1 = H - min(H, max(o["bottom"] for o in objs)), H - max(0, min(o["top"] for o in objs))
            # faixa do ponto-alvo que mantém esta página inteira na folha
            pags[i] = (c, (MARGEM - cx0 + c[0], W - MARGEM - cx1 + c[0]), (MARGEM - cy0 + c[1], H - MARGEM - cy1 + c[1]))
    if not pags:
        return 0, [], None
    alvo = []
    for k in (0, 1):
        lo = max(v[1 + k][0] for v in pags.values())
        hi = min(v[1 + k][1] for v in pags.values())
        med = statistics.median(v[0][k] for v in pags.values())
        alvo.append(min(max(med, lo), hi) if lo <= hi else med)
    desloc, pulei = {}, []
    for i, (c, fx, fy) in pags.items():
        if fx[0] - 1e-6 <= alvo[0] <= fx[1] + 1e-6 and fy[0] - 1e-6 <= alvo[1] <= fy[1] + 1e-6:
            desloc[i] = (alvo[0] - c[0], alvo[1] - c[1])
        else:
            pulei.append(i + 1)                       # sairia da folha: fica como está (assinar à mão)
    wr = PdfWriter(clone_from=PdfReader(pdf))
    for i, (tx, ty) in desloc.items():
        wr.pages[i].add_transformation(Transformation().translate(tx, ty))
    wr.write(pdf)
    return len(desloc), pulei, tuple(alvo)


def campos_assinatura(pdf, tipo):
    """Acrescenta campos de assinatura vazios (clicar no Foxit e assinar com o certificado). Devolve quantos."""
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import ArrayObject, DictionaryObject, FloatObject, NameObject, NumberObject, TextStringObject

    if any(v.get("/FT") == "/Sig" for v in (PdfReader(pdf).get_fields() or {}).values()):
        raise RuntimeError(f"{pdf} já tem campo de assinatura (assinado?) — não mexi")
    alvos = posicoes(pdf, tipo)
    if not alvos:
        return 0
    wr = PdfWriter(clone_from=PdfReader(pdf))
    if "/AcroForm" not in wr._root_object:
        wr._root_object[NameObject("/AcroForm")] = DictionaryObject({NameObject("/Fields"): ArrayObject()})
    campos = wr._root_object["/AcroForm"]["/Fields"]
    for n, (i, x, y, x1, y1) in enumerate(alvos):
        pag = wr.pages[i]
        campo = wr._add_object(DictionaryObject({
            NameObject("/Type"): NameObject("/Annot"), NameObject("/Subtype"): NameObject("/Widget"),
            NameObject("/FT"): NameObject("/Sig"), NameObject("/T"): TextStringObject(f"Assinatura_{n}"),
            NameObject("/Rect"): ArrayObject([FloatObject(x), FloatObject(y), FloatObject(x1), FloatObject(y1)]),
            NameObject("/F"): NumberObject(4), NameObject("/P"): pag.indirect_reference}))
        if "/Annots" not in pag:
            pag[NameObject("/Annots")] = ArrayObject()
        pag["/Annots"].append(campo)
        campos.append(campo)
    wr.write(pdf)
    return len(alvos)


def consolidacoes(filtro):
    """PDF de cada Consolidação de Horas (pasta de trabalho inteira, só abas visíveis), sem assinatura."""
    with Excel() as xl:
        for f in sorted(glob.glob(os.path.join(C.SAIDA, "PRT *", "*CONSOLIDA*.xlsx"))):
            if filtro and re.search(r"PRT (\d{4})", f).group(1) not in filtro:
                continue
            if os.path.exists(os.path.join(os.path.dirname(f), "~$" + os.path.basename(f))):
                print(f"!! {os.path.basename(f)} está aberto em outro Excel — pulei")
                continue
            wb = xl.abrir(f, leitura=True)
            try:
                wb.ExportAsFixedFormat(XL_PDF, os.path.splitext(f)[0] + ".pdf", 0, True, False)
            finally:
                wb.Close(SaveChanges=False)
            print(f"{os.path.basename(os.path.dirname(f))}: {os.path.basename(f)[:-5]}.pdf")


def main(filtro):
    pastas = [p for p in sorted(glob.glob(os.path.join(C.SAIDA, "PRT *"))) if os.path.isdir(p)
              and (not filtro or re.search(r"(\d{4})", os.path.basename(p)).group(1) in filtro)]
    with Excel() as xl:
        for p in pastas:
            for tipo in ("medicao", "acumulado", "bm"):
                f = arquivo(p, tipo)
                if not f:
                    continue
                if os.path.exists(os.path.join(p, "~$" + os.path.basename(f))):
                    print(f"!! {os.path.basename(f)} está aberto em outro Excel — pulei")
                    continue
                pdf = os.path.splitext(f)[0] + ".pdf"
                wb = xl.abrir(f, leitura=True)
                try:
                    if tipo == "medicao":
                        alvo = wb.Worksheets("Formuário_PRT")
                    elif tipo == "acumulado":
                        alvo = next(ws for ws in wb.Worksheets if ws.Visible == -1)   # 1ª aba visível
                    else:
                        alvo = wb
                    alvo.ExportAsFixedFormat(XL_PDF, pdf, 0, True, False)
                finally:
                    wb.Close(SaveChanges=False)
                if tipo == "bm":
                    n, pulei, alvo = alinhar_quadros(pdf)
                    print(f"{os.path.basename(p)}: {os.path.basename(pdf)} — {n} página(s) com o quadro alinhado"
                          + (f"  !! não alinhei as páginas {pulei}" if pulei else ""))
                else:
                    n = campos_assinatura(pdf, tipo)
                    print(f"{os.path.basename(p)}: {os.path.basename(pdf)} — {n} campo(s) de assinatura"
                          + ("" if n else "  !! não achei onde assinar"))


if __name__ == "__main__":
    args = set(sys.argv[1:])
    if "--consolidacao" in args:                 # python exportar_pdf.py --consolidacao [PRT...]
        consolidacoes(args - {"--consolidacao"})
    else:
        main(args)
