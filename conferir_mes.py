# -*- coding: utf-8 -*-
"""PASSO 5 — confere o que foi gerado em SAIDA (sem gabarito), PRT por PRT:

Medição:      toda linha com quantidade no mês está EXIBIDA e bate com a Base.
BM Acumulado: Acum. Anterior (J) = Acum. Atual (L) do mês passado; No Mês (K) = Base;
              Acum. Atual (L) = J + K; linha com valor no mês não está oculta.
BM:           aba com valor no mês VISÍVEL, sem valor OCULTA; linha do mês com o BM atual
              e a quantidade da Base; quantidade do mês passado foi para
              'ACUMULADO BMS ANTERIORES'.

Rode:  python conferir_mes.py [PRT ...]      (relatório em SAIDA\\conferencia_mes.txt)
"""
import os
import re
import sys
import warnings

import openpyxl

import config_mes as C
from comum import fonte, ADICIONAIS, AJUSTE, arquivo, eh_codigo, ler_acumulado, ler_medicao, norm, num, pasta_prt
from decimal import Decimal, ROUND_HALF_UP

warnings.filterwarnings("ignore")
TOL = 1e-6   # quantidades e R$: iguais, com todas as casas


def arred(x, casas=2):
    """ARRED/ROUND do Excel: meio para cima, sobre 15 dígitos (2,67499999999998 -> 2,675 -> 2,68)."""
    return float(Decimal(format(x, ".15g")).quantize(Decimal(1).scaleb(-casas), ROUND_HALF_UP))


def conferir_formulas(caminho, onde, err):
    """Nenhuma fórmula pode ter número digitado somado/subtraído (ajuste manual escondido)."""
    if not caminho:
        return
    for ws in openpyxl.load_workbook(caminho).worksheets:
        for row in ws.iter_rows():
            for c in row:
                f = c.value
                if isinstance(f, str) and f.startswith("=") and AJUSTE.search(f):
                    tipo = "valor digitado depois da fórmula" if re.search(r"[A-Za-z]", f[1:]) else "conta digitada no lugar de valor"
                    err.append(f"{onde} {ws.title}!{c.coordinate}: {tipo} ({f})")


def decisoes_base(prt):
    """Linhas da Base em que o VALOR do mês difere do Caderno (decisão manual) — só listagem."""
    ws = openpyxl.load_workbook(os.path.join(C.SAIDA, f"Base_do_Mes_BM{C.BM_NUMERO}.xlsx"), data_only=True)["Base"]
    out = []
    for p, cod, mae, desc, _, _, cad, val, *_ in ws.iter_rows(min_row=4, max_col=9, values_only=True):
        if str(p).strip() == prt and eh_codigo(cod) and dif(cad, val) and (num(cad) or num(val)):
            out.append(f"{cod}{' (' + str(mae) + ')' if mae else ''} {str(desc)[:40]}: Caderno {cad} -> medido {val}")
    return out


def ler_base():
    ws = openpyxl.load_workbook(os.path.join(C.SAIDA, f"Base_do_Mes_BM{C.BM_NUMERO}.xlsx"),
                                data_only=True)["Base"]
    out = {}
    for prt, cod, mae, *_, val in ws.iter_rows(min_row=4, max_col=8, values_only=True):
        if prt and eh_codigo(cod):
            k = f"{str(mae).strip()}|{str(cod).strip()}" if mae else str(cod).strip()
            out.setdefault(str(prt).strip(), {})[k] = (out.get(str(prt).strip(), {}).get(k) or 0) + (num(val) or 0)
    return out


def dif(a, b):
    return abs((num(a) or 0) - (num(b) or 0)) > TOL


def conferir_medicao(caminho, base, err):
    linhas = ler_medicao(caminho)
    por_chave = {l["chave"]: l for l in linhas}
    for l in linhas:
        if l["und"] in (None, "", 0):
            continue                   # tabela de equipamentos (3.x): col. J = estimativa de uso
        q = num(l["qtd"]) or 0
        if q and l["oculta"]:
            err.append(f"Medição linha {l['r']} ({l['chave']}) tem {q} mas está OCULTA")
        if "#L" not in l["chave"] and dif(q, base.get(l["chave"])):
            err.append(f"Medição {l['chave']}: {q} ≠ Base {base.get(l['chave']) or 0}")
    for k, v in base.items():
        if v and k not in por_chave:
            err.append(f"Medição: {k} = {v} na Base, sem linha na Medição")


def p_total_acumulado(caminho):
    """(P do TOTAL 1, {chave: P da linha}) — valor do mês no BM Acumulado."""
    from comum import aba_acumulado
    wv = openpyxl.load_workbook(caminho, data_only=True)
    ws = aba_acumulado(wv)
    tot = next((r for r in range(12, ws.max_row + 1)
                if norm(ws.cell(r, 1).value or "") == "TOTAL 1" or str(ws.cell(r, 2).value or "").strip() == "P0"), None)
    linhas = {l["chave"]: num(ws.cell(l["r"], 16).value) or 0 for l in ler_acumulado(caminho)[1] if l["und"] not in (None, "")}
    return (num(ws.cell(tot, 16).value) or 0) if tot else 0, linhas


def conferir_total_medicao(caminho, acumulado, err):
    """Medição: cada 'Custo do Item' cobre o bloco inteiro, o 'Custo Total da PRT' soma todos os blocos e
    o total = valor do mês do BM Acumulado (diferença só do ARRED por linha do Acumulado)."""
    wf = openpyxl.load_workbook(caminho)["Formuário_PRT"]
    wv = openpyxl.load_workbook(caminho, data_only=True)["Formuário_PRT"]
    subs, total, ini = [], None, 12
    for r in range(12, wf.max_row + 1):
        f, rot = str(wf.cell(r, 12).value or ""), norm(f"{wf.cell(r, 11).value or ''} {wf.cell(r, 3).value or ''}")
        if not f.startswith("="):
            continue
        if "CUSTO DO ITEM" in rot:
            m = re.search(r"L\$?(\d+):L\$?(\d+)", f)
            fora = [x for x in range(ini, r) if eh_codigo(wf.cell(x, 3).value) and num(wv.cell(x, 12).value)
                    and not (m and int(m.group(1)) <= x <= int(m.group(2)))]
            if fora:
                err.append(f"Medição: 'Custo do Item' L{r} ({f}) não soma as linhas {fora} (com valor)")
            subs.append(r)
            ini = r + 1
        elif "CUSTO TOTAL" in rot:
            total = r
            falta = [s for s in subs if s not in {int(x) for x in re.findall(r"L\$?(\d+)", f)}]
            if falta:
                err.append(f"Medição: 'Custo Total da PRT' L{r} ({f}) não soma os subtotais {falta}")
            break
    if total and acumulado:
        p_tot, p_lin = p_total_acumulado(acumulado)
        t = num(wv.cell(total, 12).value) or 0
        if abs(t - p_tot) > 0.005 * max(1, sum(1 for v in p_lin.values() if v)) + TOL:
            err.append(f"Medição: total {t:,.2f} ≠ valor do mês do BM Acumulado {p_tot:,.2f} (diferença {t - p_tot:,.2f})")


def conferir_formulas_acumulado(caminho, err):
    """Linhas de item: S/T/U (reajuste total, total acumulado, avanço financeiro) em fórmula; H/I (reajuste,
    ocultas) vazias — número ali entra em total que some G:I."""
    from comum import aba_acumulado
    ws = aba_acumulado(openpyxl.load_workbook(caminho))
    for l in ler_acumulado(caminho)[1]:
        if l["und"] in (None, ""):
            continue
        r = l["r"]
        for col in "STU":
            if not str(ws[f"{col}{r}"].value or "").startswith("="):
                err.append(f"Acumulado {l['chave']} (linha {r}): {col} sem fórmula ({ws[f'{col}{r}'].value!r})")
        for col in "HI":
            if ws[f"{col}{r}"].value not in (None, "", 0):
                err.append(f"Acumulado {l['chave']} (linha {r}): coluna {col} (reajuste, oculta) com {ws[f'{col}{r}'].value!r}")


def conferir_quadros_bm(caminho, acumulado, err):
    """Abas visíveis: 'Executado' = TOTAL EXECUTADO = Acum. Atual (L) do BM Acumulado (abas de item);
    'MEDIDO ANTERIORMENTE' soma todo o histórico; linha do mês com datas sem hora; aba PQ oculta."""
    import datetime as dt
    from aplicar_bm import linhas_m
    L = {l["chave"]: num(l["atual"]) or 0 for l in ler_acumulado(acumulado)[1]} if acumulado else {}
    wf = openpyxl.load_workbook(caminho, read_only=True)
    wv = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    for f_ws, v_ws in zip(wf.worksheets, wv.worksheets):
        nome = f_ws.title.strip()
        if f_ws.sheet_state != "visible":
            continue
        if nome.upper() == "PQ":
            err.append("BM: aba PQ visível (só as abas medidas no mês ficam visíveis)")
            continue
        if not re.fullmatch(r"\d+(\.\d+)+", nome):
            continue
        F = {r: row for r, row in enumerate(f_ws.iter_rows(min_row=1, max_row=170, max_col=26, values_only=True), 1)}
        V = {r: row for r, row in enumerate(v_ws.iter_rows(min_row=1, max_row=170, max_col=26, values_only=True), 1)}
        cel = lambda T, r, c: (T.get(r) or [None] * 26)[c - 1] if len(T.get(r) or []) >= c else None
        q = {norm(cel(F, r, 9) or ""): r for r in range(89, 29, -1) if norm(cel(F, r, 9) or "") in
             ("TOTAL EXECUTADO", "MEDIDO ANTERIORMENTE", "MEDIDO NO MES")}
        cab = next((r for r in range(70, 161) if norm(cel(F, r, 3) or "") == "BM"
                    and norm(cel(F, r, 5) or "").startswith("PERIODO")), None)
        if not cab or len(q) < 3:
            continue
        hist = []
        for r in range(cab + 1, 161):
            c = norm(cel(F, r, 3) or "")
            if c == "ACUMULADO BMS ANTERIORES":
                hist.append(r)
            elif c == "MEDIDO NO MES":
                break
        if hist and not set(hist) <= linhas_m(cel(F, q["MEDIDO ANTERIORMENTE"], 4)):
            err.append(f"BM aba '{nome}': 'Medido anteriormente' {cel(F, q['MEDIDO ANTERIORMENTE'], 4)} não soma "
                       f"o histórico inteiro (linhas {hist[0]}..{hist[-1]})")
        for col in (5, 9):
            d = cel(V, cab + 1, col)
            if isinstance(d, dt.datetime) and d.time() != dt.time():
                err.append(f"BM aba '{nome}': data da linha do mês com hora ({d:%d/%m/%Y %H:%M})")
                break
        exe = next((cel(V, r + 1, 25) for r in range(15, 23) if norm(cel(F, r, 25) or "") == "EXECUTADO"), None)
        if nome in L and exe is not None and abs((num(exe) or 0) - L[nome]) > TOL:
            err.append(f"BM aba '{nome}': Executado {exe} ≠ Acum. Atual do BM Acumulado {L[nome]:g}")
    wf.close()
    wv.close()


def conferir_acumulado(caminho, anterior, base, err):
    wb = openpyxl.load_workbook(caminho, data_only=True)
    from comum import aba_acumulado
    ws = aba_acumulado(wb)
    ant = {l["chave"]: l for l in ler_acumulado(anterior)[1]} if anterior else {}
    vistos = set()
    for l in ler_acumulado(caminho)[1]:
        r, k = l["r"], l["chave"]
        vistos.add(k)
        j, kk, ll = l["anterior"], l["qtd"], l["atual"]
        l_ant = ant[k]["atual"] if k in ant else 0
        if dif(j, l_ant):
            err.append(f"Acumulado {k} (linha {r}): Acum. Anterior {j} ≠ Acum. Atual do mês passado {l_ant}")
        if dif(kk, base.get(k)):
            err.append(f"Acumulado {k} (linha {r}): No Mês {kk} ≠ Base {base.get(k) or 0}")
        if isinstance(ll, (int, float)) and dif(ll, (num(j) or 0) + (num(kk) or 0)):
            err.append(f"Acumulado {k} (linha {r}): Acum. Atual {ll} ≠ J+K")
        if (num(kk) or 0) and ws.row_dimensions[r].hidden:
            err.append(f"Acumulado {k} (linha {r}): tem valor no mês mas está OCULTA")
    for k, v in base.items():
        if v and k not in vistos:
            err.append(f"Acumulado: {k} = {v} na Base, sem linha no Acumulado")
    if anterior:
        conferir_valor_anterior(caminho, anterior, err)
    conferir_valores_acumulado(caminho, err, com_pdf=anterior is None)


def valores_rs(caminho):
    """{chave: (N anterior R$, R acumulado R$)} e (N, R) do TOTAL 1."""
    from comum import aba_acumulado
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = aba_acumulado(wb)
    out = {l["chave"]: tuple(num(ws.cell(l["r"], c).value) or 0 for c in (14, 18, 5, 6))
           for l in ler_acumulado(caminho)[1] if l["und"] not in (None, "")}
    tot = next((r for r in range(12, ws.max_row + 1) if (norm(ws.cell(r, 1).value or "") == "TOTAL 1" or str(ws.cell(r, 2).value or "").strip() == "P0")), None)
    t = tuple(num(ws.cell(tot, c).value) or 0 for c in (14, 18, 7)) if tot else (0, 0, 0)
    return out, t + (num(ws["G10"].value) or 0,)


def conferir_valor_anterior(caminho, anterior, err):
    """R$ 'Anterior' (N) deste mês = R$ 'Acumulado' (R) do mês passado, item a item e no total."""
    novo, tn = valores_rs(caminho)
    velho, tv = valores_rs(anterior)
    for k, (n, _, e, f) in novo.items():
        _, r_ant, e_ant, f_ant = velho.get(k, (0, 0, None, None))
        if e_ant is not None and (abs(e - e_ant) > 1e-9 or abs(f - f_ant) > 1e-9):
            err.append(f"Acumulado {k}: contratado mudou (Quant. {e_ant} -> {e}, Preço {f_ant} -> {f}) — devia repetir o mês passado")
        if abs(n - r_ant) > TOL:
            err.append(f"Acumulado {k}: R$ Anterior {n:,.2f} ≠ R$ Acumulado do mês passado {r_ant:,.2f}")
    for i, nome in ((2, "TOTAL 1"), (3, "TOTAL GERAL (G10)")):
        if abs(tn[i] - tv[i]) > TOL:
            err.append(f"Acumulado {nome}: valor da proposta {tn[i]:,.2f} ≠ BM passado {tv[i]:,.2f} "
                       f"(diferença {tn[i] - tv[i]:,.2f})")
    if abs(tn[0] - tv[1]) > TOL:
        err.append(f"Acumulado TOTAL: R$ Anterior {tn[0]:,.2f} ≠ R$ Acumulado do mês passado {tv[1]:,.2f} "
                   f"(diferença {tn[0] - tv[1]:,.2f})")


def pdf_prt(prt):
    """PDF da PRT aprovada: pasta da fonte; se não tiver (ex.: simulação), a mais nova no Drive."""
    from escopo_pdf import pdf_da_prt
    for raiz in (C.FONTE_MEDICAO, C.FONTE_ACUMULADO):
        try:
            p = pdf_da_prt(pasta_prt(raiz, prt, C.PASTA_DA_PRT), prt)
        except Exception:
            p = None
        if p:
            return p
    import glob
    c = [f for f in glob.glob(os.path.join(C.RAIZ, "*", f"PRT {prt}*", "*.pdf"))
         if norm(os.path.basename(f)).startswith("PRT") and "BOOK" not in norm(os.path.basename(f))]
    return max(c, key=os.path.getmtime) if c else None


def conferir_valores_acumulado(caminho, err, prt=None, com_pdf=True):
    """Colunas de valor: E×F = G, L = J+K, N = J×F, P = F×K, R = N+P; total = valor da PRT (PDF)."""
    wf = openpyxl.load_workbook(caminho)
    from comum import aba_acumulado
    ws = aba_acumulado(wf)
    wv = openpyxl.load_workbook(caminho, data_only=True)[ws.title]
    itens, total_r = [], None
    for r in range(12, ws.max_row + 1):
        if (norm(ws.cell(r, 1).value or "") == "TOTAL 1" or str(ws.cell(r, 2).value or "").strip() == "P0"):
            total_r = r
            break
        cod, und = ws.cell(r, 2).value, ws.cell(r, 4).value
        if eh_codigo(cod) and und not in (None, ""):
            itens.append((r, str(cod).strip()))
    if total_r is None:
        err.append("Acumulado: não achei a linha 'TOTAL 1'")
        return
    v = lambda r, c: num(wv.cell(r, c).value) or 0
    for r, cod in itens:
        for col in "LR":   # G (contratado) pode ser valor: repete o mês passado, conferido em conferir_valor_anterior
            if not str(ws[f"{col}{r}"].value or "").startswith("="):
                err.append(f"Acumulado {cod} (linha {r}): coluna {col} sem fórmula ({ws[f'{col}{r}'].value})")
        # G com fórmula = E×F exato; G digitado vem da proposta com centavos (repete o mês passado,
        # conferido em conferir_valor_anterior): só erra se passar do meio centavo
        fg = str(ws[f"G{r}"].value or "")
        g_digitado = not fg.startswith("=")
        esperado = arred(v(r, 5) * v(r, 6)) if "ROUND" in fg.upper() else v(r, 5) * v(r, 6)
        if abs(v(r, 7) - esperado) > (0.005 if g_digitado else TOL):
            err.append(f"Acumulado {cod} (linha {r}): G {v(r, 7)} ≠ E×F {v(r, 5) * v(r, 6)}")
        if abs(v(r, 16) - arred(v(r, 6) * v(r, 11))) > TOL:
            err.append(f"Acumulado {cod} (linha {r}): valor no mês P {v(r, 16)} ≠ F×K")
    ult = itens[-1][0] if itens else None
    for col in "GNPR":
        f = str(ws[f"{col}{total_r}"].value or "")
        m = re.search(r"SUM\(\w?(\d+):\w?(\d+)\)", f.replace("$", ""))
        if col != "R" and (not m or int(m.group(2)) < ult):
            err.append(f"Acumulado: TOTAL {col}{total_r} = '{f}' não soma até a última linha de item ({ult})")
    prt = prt or re.search(r"(\d{4})", os.path.basename(caminho)).group(1)
    if not com_pdf:
        return                 # contratado vem do mês passado (ajustado à mão): PDF só na 1ª medição
    pdf = pdf_prt(prt)
    if not pdf:
        err.append("Acumulado: PDF da PRT não encontrado — valor total não conferido")
        return
    from escopo_pdf import ler_escopo_pdf
    esc = ler_escopo_pdf(pdf)
    if not esc:
        err.append(f"Acumulado: PDF sem texto ({os.path.basename(pdf)}) — valor total não conferido")
        return
    no_acu = {cod: r for r, cod in itens}
    for cod, (und, q, preco) in esc.items():
        r = no_acu.get(cod)
        if r is None:
            continue           # item da PRT ainda não medido e sem linha: não afeta o que já foi medido
        if abs(v(r, 5) - q) > 1e-6:
            err.append(f"Acumulado {cod}: quantidade contratada E {v(r, 5)} ≠ PRT {q}")
        if abs(v(r, 6) - preco) > 0.005:
            err.append(f"Acumulado {cod}: preço unitário F {v(r, 6)} ≠ PRT {preco}")
    tot_pdf = sum(q * p for _, q, p in esc.values())
    tot_acu = v(total_r, 7)
    if abs(tot_pdf - tot_acu) > 1.0:
        faltam = sorted(set(esc) - set(no_acu))
        err.append(f"Acumulado: valor total G{total_r} = {tot_acu:,.2f} ≠ valor da PRT {tot_pdf:,.2f}"
                   + (f" (itens da PRT sem linha: {', '.join(faltam)})" if faltam else ""))


def resumo_bm(caminho):
    """{nome da aba: (código, visível, bm do mês, qtd do mês, [qtds do histórico])} — leitura rápida."""
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    out = {}
    for ws in wb.worksheets:
        nome = ws.title.strip()
        linhas = {r: row for r, row in enumerate(
            ws.iter_rows(min_row=11, max_row=160, max_col=17, values_only=True), start=11)}
        k11 = str((linhas.get(11) or [None] * 11)[10] or "").strip()
        m = re.match(r"(\d+(?:\.\d+)+)(\s|$)", nome)
        cod = m.group(1) if m else (k11 if eh_codigo(k11) else None)
        if cod is None:
            continue
        cel = lambda r, c: (linhas.get(r) or [None] * 17)[c - 1] if len(linhas.get(r) or []) >= c else None
        cab, hist = None, []
        for r in range(70, 161):
            c = norm(cel(r, 3) or "")
            if c == "BM" and norm(cel(r, 5) or "").startswith("PERIODO"):
                cab = r
            elif c == "ACUMULADO BMS ANTERIORES":
                hist.append(r)
            elif c == "MEDIDO NO MES" and cab:
                break
        if cab is None:
            continue
        out[nome] = (cod, ws.sheet_state == "visible", cel(cab + 1, 3), num(cel(cab + 1, 13)) or 0,
                     [num(cel(r, 13)) for r in hist if num(cel(r, 13))])
    wb.close()
    return out


def conferir_bm(caminho, anterior, base, err, acumulado=None, acumulado_ant=None):
    """acumulado / acumulado_ant: BM Acumulado gerado e o do mês passado. O histórico da aba
    (ACUMULADO BMS ANTERIORES) tem de somar o Acum. anterior (J) do BM Acumulado."""
    novo = resumo_bm(caminho)
    velho = resumo_bm(anterior) if anterior else {}
    j_novo = {l["chave"]: num(l["anterior"]) or 0 for l in ler_acumulado(acumulado)[1]} if acumulado else {}
    j_velho = {l["chave"]: num(l["anterior"]) or 0 for l in ler_acumulado(acumulado_ant)[1]} if acumulado_ant else {}
    mes_ant = {k: v - j_velho.get(k, 0) for k, v in j_novo.items()}   # medido no BM anterior
    he_base = sum(v for k, v in base.items() if "|" in k)
    he_bm = 0
    for nome, (cod, vis, bm, q, h_nov) in novo.items():
        if cod in ADICIONAIS:
            he_bm += q if vis else 0
            esperado = None          # dono da HE vem pela ordem das abas: confere pelo total
        else:
            esperado = base.get(cod) or 0
        if vis and not q:
            err.append(f"BM aba '{nome}': VISÍVEL sem quantidade no mês")
        if not vis and q and bm is not None and str(bm).split(".")[0] == C.BM_NUMERO:
            err.append(f"BM aba '{nome}': tem {q} no BM {C.BM_NUMERO} mas está OCULTA")
        if esperado is not None and esperado and not vis:
            err.append(f"BM aba '{nome}': Base tem {esperado} mas a aba está OCULTA")
        if vis:
            if bm is None or str(bm).split(".")[0] != C.BM_NUMERO:
                err.append(f"BM aba '{nome}': linha do mês com BM {bm} (esperado {C.BM_NUMERO})")
            if esperado is not None and dif(q, esperado):
                err.append(f"BM aba '{nome}': quantidade do mês {q} ≠ Base {esperado}")
            if cod not in ADICIONAIS and cod in j_novo:
                if dif(sum(h_nov), j_novo[cod]):
                    err.append(f"BM aba '{nome}': ACUMULADO BMS ANTERIORES soma {sum(h_nov):g} ≠ Acum. anterior "
                               f"{j_novo[cod]:g} do BM Acumulado ({h_nov})")
            elif nome in velho:
                _, _, bm_ant, q_ant, h_ant = velho[nome]
                if q_ant and str(bm_ant).split(".")[0] != C.BM_NUMERO and h_nov[:len(h_ant) + 1] != h_ant + [q_ant]:
                    err.append(f"BM aba '{nome}': mês passado ({q_ant}) não foi para ACUMULADO BMS ANTERIORES "
                               f"(antes {h_ant}, agora {h_nov})")
        elif (q and str(bm).split(".")[0] != C.BM_NUMERO and mes_ant.get(cod)
              and not dif(q, mes_ant[cod])):
            err.append(f"BM aba '{nome}' (oculta): BM anterior ({q}) parado na linha do mês — não foi para o histórico")
    cods = {v[0] for v in novo.values()}
    for k, v in base.items():
        if v and "|" not in k and k not in cods:
            err.append(f"BM: {k} = {v} na Base, sem aba")
    if dif(he_base, he_bm):
        err.append(f"BM: horas extras nas abas visíveis = {he_bm} ≠ Base {he_base}")


def acertos_contrato():
    """{prt: chaves com o contrato acertado pela proposta válida} do acertos_historico.py do mês ('|1.5.' = toda HE)."""
    import importlib.util
    f = os.path.join(C.SAIDA, "acertos_historico.py")
    if not os.path.exists(f):
        return {}
    spec = importlib.util.spec_from_file_location("acertos_do_mes", f)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    out = {}
    for prt, lista in getattr(m, "ACERTOS_CONTRATO", {}).items():
        for item in lista:
            if item[0] in ("E", "F", "HE"):
                out.setdefault(prt, set()).add(item[1] if item[0] != "HE" else "|1.5.")
    return out


def eh_acerto(e, acert):
    """Aviso de contrato que vem de acerto pela proposta (não é problema; vai listado à parte)."""
    m = re.match(r"Acumulado (\S+): contratado mudou", e)
    if m:
        return m.group(1) in acert or ("|1.5." in acert and "|1.5." in m.group(1))
    return bool(acert) and bool(re.match(r"Acumulado TOTAL( 1| GERAL \(G10\)): valor da proposta", e))


def blocos_rel(txt):
    """{prt: texto do bloco} de um conferencia_mes.txt (cabeçalho e TOTAL ficam de fora)."""
    out, atual = {}, None
    for l in txt.split("\n"):
        m = re.match(r"PRT (\d{4}):", l)
        if m:
            atual = m.group(1)
            out[atual] = [l]
        elif l.startswith("TOTAL:"):
            atual = None
        elif atual:
            out[atual].append(l)
    return {p: "\n".join(ls).rstrip() for p, ls in out.items()}


def juntar(velho, novo):
    """Rodada filtrada: o relatório completo de antes com os blocos das PRTs rodadas trocados; TOTAL refeito."""
    bl = blocos_rel(velho)
    bl.update(blocos_rel(novo))
    n = sum(int(m.group(1)) for b in bl.values() for m in [re.match(r"PRT \d{4}: (\d+) problema", b)] if m)
    return "\n\n".join([novo.split("\n")[0]] + [bl[p] for p in sorted(bl)] + [f"TOTAL: {n} problema(s)"])


def main():
    so = [a for a in sys.argv[1:] if re.fullmatch(r"\d{4}", a)]
    base = ler_base()
    acertados = acertos_contrato()
    rel = [f"CONFERÊNCIA BM {C.BM_NUMERO} — {C.SAIDA}"]
    tot = 0
    for prt in sorted(base):
        if (so and prt not in so) or not any(base[prt].values()):   # sem medição no mês: não há arquivo
            continue
        err = []
        try:
            pasta = pasta_prt(C.SAIDA, prt, C.PASTA_DA_PRT)
            def ant(tipo):
                try:
                    return fonte(prt, tipo)
                except Exception:
                    return None
            b = base[prt]
            for tipo, onde in (("medicao", "Medição"), ("acumulado", "Acumulado"), ("bm", "BM")):
                conferir_formulas(arquivo(pasta, tipo), onde, err)
            conferir_medicao(arquivo(pasta, "medicao"), b, err)
            conferir_total_medicao(arquivo(pasta, "medicao"), arquivo(pasta, "acumulado"), err)
            conferir_acumulado(arquivo(pasta, "acumulado"), ant("acumulado"), b, err)
            conferir_formulas_acumulado(arquivo(pasta, "acumulado"), err)
            if any(b.values()):
                conferir_bm(arquivo(pasta, "bm"), ant("bm"), b, err, arquivo(pasta, "acumulado"), ant("acumulado"))
                conferir_quadros_bm(arquivo(pasta, "bm"), arquivo(pasta, "acumulado"), err)
        except Exception as e:
            err.append(f"não consegui conferir: {e}")
        notas = [e for e in err if eh_acerto(e, acertados.get(prt, set()))]
        err = [e for e in err if e not in notas]
        tot += len(err)
        rel.append(f"\nPRT {prt}: " + ("OK" if not err else f"{len(err)} problema(s)"))
        rel += [f"   {e}" for e in err]
        if notas:
            rel.append(f"   (contrato acertado pela proposta válida — acertos_historico; não é problema: {len(notas)})")
            rel += [f"      {e}" for e in notas]
        dec = decisoes_base(prt)
        if dec:
            rel.append(f"   (decisões manuais — medido ≠ Caderno, conferir se é intencional: {len(dec)})")
            rel += [f"      {d}" for d in dec]
        print("\n".join(rel[-1 - len(err):]), flush=True)
    rel.append(f"\nTOTAL: {tot} problema(s)")
    caminho = os.path.join(C.SAIDA, "conferencia_mes.txt")
    txt = "\n".join(rel)
    if so and os.path.exists(caminho):        # filtrada: as outras PRTs ficam como estavam no relatório completo
        txt = juntar(open(caminho, encoding="utf-8").read(), txt)
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(txt)
    print(rel[-1] + ("" if not so else f"  (relatório completo: {txt.splitlines()[-1]})"))


if __name__ == "__main__":
    if sys.argv[1:] == ["teste"]:
        v = "CAB\n\nPRT 5001: OK\n\nPRT 5007: 2 problema(s)\n   a\n   b\n\nPRT 5008: 1 problema(s)\n   c\n\nTOTAL: 3 problema(s)"
        assert juntar(v, "CAB2\n\nPRT 5007: OK\n\nTOTAL: 0 problema(s)") == \
            "CAB2\n\nPRT 5001: OK\n\nPRT 5007: OK\n\nPRT 5008: 1 problema(s)\n   c\n\nTOTAL: 1 problema(s)"
        assert juntar(v, "CAB\n\nPRT 5029: 1 problema(s)\n   x\n   (decisões manuais — 1)\n      d\n\nTOTAL: 1 problema(s)") \
            .endswith("PRT 5029: 1 problema(s)\n   x\n   (decisões manuais — 1)\n      d\n\nTOTAL: 4 problema(s)")
        a = {"4.2.2.16", "|1.5."}
        assert eh_acerto("Acumulado 4.2.2.4|1.5.2: contratado mudou (Quant. 80 -> 80, Preço 1 -> 2) — x", a)
        assert eh_acerto("Acumulado 4.2.2.16: contratado mudou (Quant. 88 -> 52.8) — x", a)
        assert not eh_acerto("Acumulado 4.2.2.20: contratado mudou (Quant. 88 -> 80) — x", a)
        assert eh_acerto("Acumulado TOTAL 1: valor da proposta 1 ≠ BM passado 2", a)
        assert not eh_acerto("Acumulado TOTAL 1: valor da proposta 1 ≠ BM passado 2", set())
        print("conferir_mes: ok")
    else:
        main()
