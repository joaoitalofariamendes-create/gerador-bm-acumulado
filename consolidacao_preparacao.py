# -*- coding: utf-8 -*-
"""Consolidação de Horas da Preparação de Cargas (contrato 4600000002): uma planilha por equipe (Fixa, Caldeiraria,
Parada) na pasta da equipe em BM Manutenção, a partir da aba Consolidação do Caderno (blocos 'CLIENTE - PREPARAÇÃO DE
CARGAS - <EQUIPE>'). Modelo = a Consolidação de Manutenção de várias abas do mês passado (uma aba por equipe); a deste
mês fica em _arquivo\\consolidacao_varias_abas (modelo do mês seguinte). Mesmas funções da Consolidação das PRTs
(aplicar_consolidacao): datas do ciclo, horas por dia na linha da função (normal / 'HORAS 50 %' = 'H. Extras 75%' do
Caderno, 1ª faixa de HE deste contrato / 100% / noturno), blocos sem hora ocultos, período e data de entrega.
Função ou tipo sem linha na aba: para sem gravar (acrescentar a linha no modelo). Confere dia a dia no fim.
Uso: python consolidacao_preparacao.py [Consolidação de Manutenção do mês passado]   (padrão: a do _arquivo)
"""
import glob
import os
import re
import shutil
import sys
from datetime import timedelta

import config_mes as C
from aplicar_consolidacao import (ARQ, COL_DIA1, TIPO, achar_grupo, ajustar_colunas, data_entrega, estrutura,
                                  exportar_aba, formatar_dias, ler_caderno_consolidacao, limpar_dias, por_rotulo,
                                  referencias, totais_da_linha)
from comum import Excel, data_excel, norm, padronizar_fontes
from conferir_consolidacao import conferir_arquivo

AQUI = os.path.dirname(os.path.abspath(__file__))
PASTA = os.path.join(AQUI, "BM Manutenção")
# equipe no Caderno -> (palavra no nome da aba do modelo, subpasta em BM Manutenção, nome no arquivo)
EQUIPES = {"EQUIPE FIXA": ("FIXA", "Equipe Fixa", "EQUIPE FIXA"),
           "CALDEIRARIA": ("CALDEIRARIA", "Equipe Caldeiraria", "EQUIPE CALDEIRARIA"),
           "PARADA": ("PARADA", "Equipe Parada", "EQUIPE PARADA")}
PERIODO = r"\d\d/\d\d/\d{4} A \d\d/\d\d/\d{4}"


def equipe(titulo):
    """Título do bloco no Caderno -> equipe ('CLIENTE - PREPARAÇÃO DE CARGAS - EQUIPE FIXA' -> 'EQUIPE FIXA'); PRT -> None."""
    return titulo.split(" - ")[-1].strip().upper() if "PREPARACAO DE CARGAS" in norm(titulo) else None


def aba(wb, palavra):
    cands = [ws for ws in wb.Worksheets if palavra in norm(ws.Name) and "MANUTEN" in norm(ws.Name)]
    return ([ws for ws in cands if ws.Visible == -1] or cands or [None])[0]


def preencher(xl, ws, por_func, dias, lg):
    """Horas do ciclo na aba. Devolve as faltas (função/tipo sem linha, dia fora do ciclo) SEM mexer na aba se houver."""
    _, _, _, blocos = estrutura(ws)
    grupos, pedidos, faltas = por_rotulo(blocos), {}, []
    for (func, tipo), por_dia in por_func.items():
        g = achar_grupo(grupos, func)
        r = grupos[g]["linhas"].get(TIPO.get(tipo)) if g else None
        if r is None:
            faltas.append(f"{func} / {tipo} ({sum(por_dia.values()):g} h): sem linha na aba")
            continue
        for d, h in por_dia.items():
            if d not in dias:
                faltas.append(f"{func} / {tipo}: dia {d:%d/%m/%Y} fora do ciclo ({h:g} h)")
            pedidos[(r, d)] = pedidos.get((r, d), 0) + h
        lg.append(f"{func} / {tipo}: {sum(por_dia.values()):g} h -> linha {r} '{str(ws.Cells(r, 2).Value).strip()}'")
    if faltas:
        return faltas
    ajustar_colunas(xl, ws, len(dias), lg)
    ref = referencias(ws)                       # antes de trocar as datas
    soma, _, _, blocos = estrutura(ws)
    for i, c in enumerate(range(COL_DIA1, soma)):
        if not ws.Cells(11, c).HasFormula:
            ws.Cells(11, c).Value = data_excel(dias[i]) if i < len(dias) else None   # número: datetime pelo COM = 03:00
    formatar_dias(ws, ref, dias)
    periodo = f"{C.PERIODO_INI:%d/%m/%Y} A {C.PERIODO_FIM:%d/%m/%Y}"
    for i, row in enumerate(ws.Range("A6:CZ9").Value):
        for j, v in enumerate(row):
            if isinstance(v, str) and re.search(PERIODO, v):
                ws.Cells(6 + i, j + 1).Value = re.sub(PERIODO, periodo, v)
    data_entrega(ws)
    for bl in blocos:
        for r in bl["linhas"].values():
            limpar_dias(ws, r, soma)
    for (r, d), h in pedidos.items():
        cel = ws.Cells(r, COL_DIA1 + dias.index(d))
        cel.Value = h / 24
        if str(cel.NumberFormat) != "[h]:mm":
            cel.NumberFormat = "[h]:mm"
    totais_da_linha(ws)
    ocultos = []
    for bl in blocos:                           # bloco sem hora no mês fica oculto
        usa = any(any(v not in (None, "", 0) for v in ws.Range(ws.Cells(r, COL_DIA1), ws.Cells(r, soma - 1)).Value[0])
                  for r in bl["linhas"].values())
        ws.Rows(f"{bl['ini']}:{bl['fim']}").Hidden = not usa
        if not usa:
            ocultos.append(bl["rot"])
    if ocultos:
        lg.append(f"ocultos (sem hora no mês): {', '.join(ocultos)}")
    return []


def main():
    ini, fim = C.PERIODO_INI.strftime("%d-%m"), C.PERIODO_FIM.strftime("%d-%m-%Y")
    nome = f"BM GSV- CONSOLIDAÇÃO DE HORAS - {ini} a {fim}"
    if len(sys.argv) > 1:
        modelo = sys.argv[1]
    else:
        cands = [f for f in glob.glob(os.path.join(ARQ, "*CONSOLIDA*Manuten*.xlsx")) if nome not in os.path.basename(f)]
        if not cands:
            raise SystemExit("Sem modelo no _arquivo: passe a Consolidação de Manutenção do mês passado")
        modelo = max(cands, key=os.path.getmtime)
    dados, _ = ler_caderno_consolidacao(equipe)
    if set(dados) - set(EQUIPES):
        raise SystemExit(f"!! equipes no Caderno sem aba no script: {set(dados) - set(EQUIPES)}")
    os.makedirs(ARQ, exist_ok=True)
    multi = os.path.join(ARQ, f"{nome} - Manutenção.xlsx")
    shutil.copy2(modelo, multi)
    dias = [C.PERIODO_INI + timedelta(d) for d in range((C.PERIODO_FIM - C.PERIODO_INI).days + 1)]
    log = [f"CONSOLIDAÇÃO DE HORAS — PREPARAÇÃO DE CARGAS — BM {C.BM_NUMERO} ({C.PERIODO_INI:%d/%m} a {C.PERIODO_FIM:%d/%m})",
           f"Modelo: {modelo}"]
    feitas = {}
    with Excel() as xl:
        wb = xl.abrir(multi)
        xl.app.ScreenUpdating = False
        for eq, (palavra, _, _) in EQUIPES.items():
            if eq not in dados:
                log.append(f"\n{eq}: sem horas no ciclo — sem planilha")
                continue
            ws = aba(wb, palavra)
            if ws is None:
                raise SystemExit(f"!! modelo sem aba de {eq}")
            ws.Visible = -1
            lg = []
            faltas = preencher(xl, ws, dados[eq], dias, lg)
            if faltas:
                wb.Close(False)
                raise SystemExit("\n".join([f"!! {eq} (aba '{ws.Name}') — nada gravado:"] + [f"   {x}" for x in faltas]))
            feitas[eq] = ws
            log += [f"\n{eq} [aba '{ws.Name}']"] + [f"   {x}" for x in lg]
        nomes = {ws.Name for ws in feitas.values()}
        for ws in wb.Worksheets:                # aba visível sem hora no mês (ex.: EQUIPE NR 12) fica oculta
            if ws.Visible == -1 and ws.Name not in nomes:
                ws.Visible = 0
                log.append(f"\nAba '{ws.Name}' OCULTA: sem hora neste mês")
        padronizar_fontes(wb, "Consolidação", log)
        saidas = {}
        for eq, ws in feitas.items():
            dst = os.path.join(PASTA, EQUIPES[eq][1], f"{nome} - {EQUIPES[eq][2]}.xlsx")
            if exportar_aba(xl, ws, dst):
                saidas[eq] = dst
            else:
                log.append(f"!! {dst}: aberta no Excel — NÃO regravei (feche e rode de novo)")
        next(ws for ws in wb.Worksheets if ws.Visible == -1).Activate()
        wb.Save()
        wb.Close()
    log.append("\nPLANILHAS E CONFERÊNCIA DIA A DIA CONTRA O CADERNO:")
    for eq, dst in saidas.items():
        err, n_dias, _, n_cel = conferir_arquivo(dst, dados[eq], [])
        log.append(f"   {os.path.relpath(dst, AQUI)}: {'OK' if not err else f'{len(err)} problema(s)'} — "
                   f"{n_dias} dias, {n_cel} célula(s) lançadas")
        log += [f"      !! {x}" for x in err]
    log.append(f"   várias abas (modelo do mês seguinte): {os.path.relpath(multi, AQUI)}")
    txt = "\n".join(log)
    open(os.path.join(ARQ, f"relatorio_consolidacao_preparacao_BM{C.BM_NUMERO}.txt"), "w", encoding="utf-8").write(txt)
    print(txt)


if __name__ == "__main__" and sys.argv[1:] == ["teste"]:
    assert equipe("CLIENTE - PREPARAÇÃO DE CARGAS - EQUIPE FIXA") == "EQUIPE FIXA"
    assert equipe("CLIENTE - PRT_5007 - PERFURATRIZ") is None
    print("ok")
elif __name__ == "__main__":
    main()
