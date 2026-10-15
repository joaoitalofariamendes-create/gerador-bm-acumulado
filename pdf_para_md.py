# -*- coding: utf-8 -*-
"""Converte PDFs em Markdown (cache_md\\) para serem lidos sem reabrir o PDF.

Rode:  python pdf_para_md.py                 -> todos os PDFs de PRT do Drive (sem BOOK)
       python pdf_para_md.py <pdf|pasta> ... -> só esses (pasta = todos os PDFs dela)
Só reconverte se o PDF mudou. Cada .md traz no topo a PQ lida (código | und | qtd | preço)
e depois o texto por página. PDF digitalizado (sem texto) vira .md com o aviso.
"""
import glob
import os
import sys

import pdfplumber

import config_mes as C
from comum import norm
from escopo_pdf import ler_escopo_pdf

CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache_md")


def pdfs_de_prt():
    return [f for f in glob.glob(os.path.join(C.RAIZ, "MES *", "PRT*", "*.pdf"))
            if norm(os.path.basename(f)).startswith("PRT") and "BOOK" not in norm(os.path.basename(f))]


def destino(pdf):
    pasta = os.path.basename(os.path.dirname(pdf))
    return os.path.join(CACHE, pasta, os.path.splitext(os.path.basename(pdf))[0] + ".md")


def converter(pdf):
    md = destino(pdf)
    # mesmo nome de PDF em outro mês cai no mesmo .md: só vale o cache se a Origem for este PDF
    if (os.path.exists(md) and os.path.getmtime(md) >= os.path.getmtime(pdf)
            and f"- Origem: `{pdf}`" in open(md, encoding="utf-8").read(2000)):
        return md, "em cache"
    with pdfplumber.open(pdf) as doc:
        paginas = [(p.extract_text() or "").strip() for p in doc.pages]
    out = [f"# {os.path.basename(pdf)}", "", f"- Origem: `{pdf}`", f"- Páginas: {len(paginas)}", ""]
    if not any(paginas):
        out.append("**PDF sem texto (digitalizado) — ler o original ou digitar os dados.**")
    else:
        try:
            pq = ler_escopo_pdf(pdf)
        except Exception:
            pq = {}
        if pq:
            out += ["## PQ lida do PDF", "", "| código | und | qtd contratada | preço |", "|---|---|---|---|"]
            out += [f"| {c} | {u} | {q:g} | {p:,.2f} |" for c, (u, q, p) in pq.items()]
            out.append("")
        for i, t in enumerate(paginas, 1):
            out += [f"## Página {i}", "", t or "_(sem texto)_", ""]
    os.makedirs(os.path.dirname(md), exist_ok=True)
    with open(md, "w", encoding="utf-8") as f:
        f.write("\n".join(out))
    return md, "convertido"


def main():
    alvos = []
    for a in sys.argv[1:]:
        alvos += glob.glob(os.path.join(a, "*.pdf")) if os.path.isdir(a) else [a]
    for pdf in alvos or pdfs_de_prt():
        try:
            md, st = converter(pdf)
            print(f"{st:10} {os.path.relpath(md, CACHE)}")
        except Exception as e:
            print(f"ERRO       {pdf}: {e}")


if __name__ == "__main__":
    main()
