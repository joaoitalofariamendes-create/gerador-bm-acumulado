# -*- coding: utf-8 -*-
"""PASSO 3 — atualiza o BM (arquivo de ~200 abas, uma por item) de cada PRT.

Rode:  python aplicar_bm.py [caminho da Base] [PRT ...]
Gera cópia em SAIDA\\<pasta da PRT>\\ a partir do BM do mês anterior.

Em cada aba de item:
  - a linha do mês que passou (BM / período / quantidade) vai para a primeira linha
    livre de "ACUMULADO BMS ANTERIORES" (quantidade + "BM nnnn" na observação, se livre);
  - a linha do mês recebe o BM atual, o período e a quantidade da Base;
  - aba fica visível se o item tem quantidade no mês, oculta se não tem.
Abas de hora extra (1.5.1 / 1.5.2) pertencem à aba de função imediatamente anterior.
Abas que não são de item (PQ, fotos, 1.1.1...) não mudam de visibilidade.
O print da consolidação e as fotos continuam manuais — o relatório lista onde trocar.

Criação (tudo listado no relatório):
  - sem linha livre em ACUMULADO BMS ANTERIORES: copia a última dessas linhas logo abaixo
    e o "MEDIDO ANTERIORMENTE" passa a somar todas;
  - item do mês sem aba: copia a aba irmã mais parecida (mesmo pai / mesma profundidade),
    renomeia para o código, acerta os títulos (A14:A17) e o Previsto (U20, da aba Escopo);
    hora extra (1.5.2) sem aba: copia outra aba 1.5.2 e põe logo depois da função;
    o print da memória que veio da aba copiada é apagado (colocar o print novo);
  - PRT na 1ª medição: BM novo a partir do BM (menor) de outra PRT do mês passado,
    com histórico e Previsto zerados.
"""
import glob
import os
import re
import shutil
import sys

import config_mes as C
from aplicar_base import ler_base, txt_cod
from comum import (fonte, ADICIONAIS, Excel, tirar_ajustes, padronizar_fontes, arquivo, chave_cod, data_excel, eh_codigo,
                   drive_ok, ler_acumulado, ler_escopo, ler_medicao, mostrar_casas, num, norm, ordem, pasta_prt, prt_info,
                   pq_da_medicao)

VISIVEL, OCULTA = -1, 0
XL_DOWN = -4121


def mapa_abas(wb, desc_hh, adic_validos=frozenset()):
    """[(aba, chave)] para as abas de item. desc_hh: {código hh: descrição}.
    Aba de hora extra ('1.5.2 TEC') -> função anterior cuja descrição começa com a
    sigla do nome (TEC -> Técnico...); sem sigla que case, a função hh mais próxima antes."""
    out, anteriores = [], []   # anteriores: códigos hh na ordem das abas
    for ws in wb.Worksheets:
        nome = ws.Name.strip()
        k11 = str(ws.Range("K11").Value or "").strip()
        m = re.match(r"(\d+(?:\.\d+)+)(\s|$)", nome)   # '1.5.2 SOLD' -> 1.5.2 (K11 é fórmula do nome)
        cod = m.group(1) if m else (k11 if eh_codigo(k11) else None)
        if cod is None:
            continue
        if cod in ADICIONAIS:
            siglas = [t for t in norm(nome).split()[1:] if not eh_codigo(t)]
            mae = None
            for sig in reversed(siglas):
                cand = [c for c in reversed(anteriores)
                        if any(p.startswith(sig) for p in norm(desc_hh[c]).split())
                        and (not adic_validos or f"{c}|{cod}" in adic_validos)]
                if cand:
                    mae = cand[0]
                    break
            if mae is None and anteriores:
                mae = anteriores[-1]
            out.append((ws, f"{mae}|{cod}" if mae else None))
        else:
            if cod in desc_hh:
                anteriores.append(cod)
            out.append((ws, cod))
    return out


def ordenar_abas(wb, abas, acu, log):
    """Abas de item na ordem das linhas do BM Acumulado do mês (pedido do cliente).
    Aba sem linha no Acumulado (modelo oculto) fica grudada na aba que vinha antes dela,
    para a hora extra continuar logo depois da sua função (mapa_abas depende disso).
    PQ, 1.1.1 e demais abas que não são de item não mudam de lugar entre si."""
    if not acu:
        return
    linha = {}
    for l in ler_acumulado(acu)[1]:
        linha.setdefault(l["chave"], l["r"])
    itens, rank, chave_de = [], -1, {}   # chave_de: chave da função -> chave de ordenação da aba dela
    for i, (ws, chave) in enumerate(abas):
        if ws.Name.strip() == "1.1.1":
            continue
        if chave is None:                 # hora extra sem função antes (modelo): fica no início, sem função
            k = (-1, i, 0)
        elif chave in linha:
            rank = linha[chave]
            k = (rank, i, 0)
        elif "|" in chave and chave.split("|")[0] in chave_de:
            m = chave_de[chave.split("|")[0]]   # hora extra sem linha: logo depois da aba da sua função
            k = (m[0], m[1], i)
        else:
            k = (rank, i, 0)              # sem linha: grudada na aba que vinha antes
        chave_de[chave] = k
        itens.append((k, ws))
    alvo = [ws for _, ws in sorted(itens, key=lambda t: t[0])]
    if [ws.Name for ws in alvo] == [ws.Name for _, ws in itens]:
        return
    # Move(After=X) pula as abas OCULTAS que vêm depois de X: com abas ocultas no meio a ordem saía
    # embaralhada (5028: bloco invertido; 5007: HE ligada à função errada). Tudo visível durante a ordenação.
    estado = [(s, s.Visible) for s in wb.Worksheets]
    for s, v in estado:
        if v != VISIVEL:
            s.Visible = VISIVEL
    ancora = min((ws for _, ws in itens), key=lambda w: w.Index).Previous   # 1.1.1 / PQ
    for ws in alvo:
        ws.Move(None, ancora)             # posicional (Before, After): por nome o pywin32 ignora
        ancora = ws
    for s, v in estado:
        if v != VISIVEL:
            s.Visible = v
    log.append("   abas de item reordenadas na ordem das linhas do BM Acumulado")


def bloco_historico(ws):
    """Linhas do quadro 'BM / PERÍODO INICIAL...' da 2a página."""
    col_c = ws.Range("C70:C160").Value
    col_e = ws.Range("E70:E160").Value
    cab = hist = medido = None
    hist = []
    for i, ((c,), (e,)) in enumerate(zip(col_c, col_e)):
        r = 70 + i
        c = norm(c or "")
        if c == "BM" and norm(e or "").startswith("PERIODO"):
            cab = r
        elif c == "ACUMULADO BMS ANTERIORES":
            hist.append(r)
        elif c == "MEDIDO NO MES" and cab:
            medido = r
            break
    return cab, hist, medido


def historico(ws, com_mes=False):
    """Soma das linhas 'ACUMULADO BMS ANTERIORES' da aba (com_mes: + a linha do mês passado) — coluna M."""
    cab, hist, _ = bloco_historico(ws)
    if cab is None:
        return 0.0
    return sum(v for v in (ws.Cells(r, 13).Value for r in hist + ([cab + 1] if com_mes else []))
               if isinstance(v, (int, float)))


def nova_linha_hist(xl, ws, hist):
    """Copia a última linha 'ACUMULADO BMS ANTERIORES' logo abaixo dela e faz o
    'MEDIDO ANTERIORMENTE' somar do primeiro ao novo último."""
    ult = hist[-1]
    ws.Rows(ult).Copy()
    ws.Rows(ult + 1).Insert(XL_DOWN)   # argumentos POSICIONAIS: por nome o pywin32 ignora
    xl.app.CutCopyMode = False
    ws.Cells(ult + 1, 13).Value = None
    ws.Cells(ult + 1, 17).Value = None
    col_i = ws.Range("I40:I80").Value
    for i, (v,) in enumerate(col_i):
        if norm(v or "") == "MEDIDO ANTERIORMENTE":
            ws.Cells(40 + i, 4).Formula = f"=SUM(M{hist[0]}:P{ult + 1})"
            break
    return ult + 1


def rolar_aba(ws, valor, log, xl=None, prev=None, so_confirmado=False):
    """prev = (acumulado até o BM anterior, medido no BM anterior, nº do BM anterior), do
    BM Acumulado do mês passado: é ele que diz se a linha do mês vai para o histórico e
    com qual rótulo. so_confirmado: aba sem valor no mês — só rola se o Acumulado confirmar."""
    cab, hist, _ = bloco_historico(ws)
    if cab is None or not hist:
        if not so_confirmado:
            log.append(f"   aba '{ws.Name}': não achei o quadro BM/ACUMULADO — não mexi")
        return False
    cur = cab + 1
    bm_ant = ws.Cells(cur, 3).Value
    q_ant = ws.Cells(cur, 13).Value
    if isinstance(bm_ant, float):
        bm_ant = int(bm_ant)
    if str(bm_ant) == C.BM_NUMERO:
        if so_confirmado:
            return False
        log.append(f"   aba '{ws.Name}': já está no BM {C.BM_NUMERO} — só atualizei a quantidade")
    elif isinstance(q_ant, (int, float)) and q_ant:
        rotulo, rolar = f"BM {bm_ant}", True
        if prev:
            acum, mes_ant, bm_prev = prev
            h = sum(v for v in (ws.Cells(r, 13).Value for r in hist) if isinstance(v, (int, float)))
            if abs(h + q_ant - acum) < 0.005:
                if mes_ant and abs(q_ant - mes_ant) < 0.005 and str(bm_ant) != bm_prev:
                    rotulo = f"BM {bm_prev}"
                    if not so_confirmado:
                        log.append(f"   aba '{ws.Name}': linha do mês dizia BM {bm_ant}, mas pelo Acumulado os "
                                   f"{q_ant:g} foram medidos no BM {bm_prev} — histórico rotulado BM {bm_prev}")
            elif so_confirmado:
                return False
            elif abs(h - acum) < 0.005:
                rolar = False
                log.append(f"   aba '{ws.Name}': linha do mês trazia BM {bm_ant} = {q_ant:g}, que não consta no "
                           f"BM Acumulado (acumulado {acum:g}) — NÃO foi para o histórico")
            else:
                log.append(f"   aba '{ws.Name}': REVISAR — histórico {h:g} + BM {bm_ant} ({q_ant:g}) = {h + q_ant:g}, "
                           f"mas o BM Acumulado diz {acum:g} até o BM {bm_prev}; lancei como estava")
        elif so_confirmado:
            return False
        if rolar:
            livres = [r for r in hist if ws.Cells(r, 13).Value in (None, "")]
            if livres:
                r = livres[0]
            else:
                r = nova_linha_hist(xl, ws, hist)
                log.append(f"   aba '{ws.Name}': CRIADA linha {r} em ACUMULADO BMS ANTERIORES (não havia livre) "
                           f"— 'Medido anteriormente' agora soma M{hist[0]}:P{r}")
            ws.Cells(r, 13).Value = q_ant
            if ws.Cells(r, 17).Value in (None, ""):
                outros = [ws.Cells(x, 17).Value for x in hist if ws.Cells(x, 17).Value not in (None, "")]
                # rótulo no estilo da aba: só o número (5016, 5017…) ou 'BM 5029'
                so_num = outros and all(isinstance(o, (int, float)) for o in outros) and rotulo.split()[-1].isdigit()
                ws.Cells(r, 17).Value = int(rotulo.split()[-1]) if so_num else rotulo
            else:
                log.append(f"   aba '{ws.Name}': observação na linha {r} ocupada — {rotulo} ({q_ant}) lançado sem rótulo")
    elif so_confirmado:
        return False
    # observação do mês passado ao lado da medição deste mês (e texto solto em linha de histórico vazia):
    # apagada — decisão do usuário, 03/10 ("as observações pode apagar, não precisa manter")
    if valor:
        for r in list(range(cur, hist[0])) + [h for h in hist if ws.Cells(h, 13).Value in (None, "")]:
            q = ws.Cells(r, 17).Value
            if isinstance(q, str) and q.strip() and not re.fullmatch(r"BM\s*-?\s*\d{4}", q.strip()):
                ws.Cells(r, 17).Value = None
                log.append(f"   aba '{ws.Name}': observação do mês passado apagada (Q{r}: {q.strip()[:50]})")
    ws.Cells(cur, 3).Value = int(C.BM_NUMERO)
    # número de data do Excel: datetime pelo COM grava 03:00 (fuso)
    ws.Cells(cur, 5).Value = data_excel(C.PERIODO_INI)
    ws.Cells(cur, 9).Value = data_excel(C.PERIODO_FIM)
    ws.Cells(cur, 13).Value = valor or 0
    return True


def linhas_m(f):
    """Linhas da coluna M referenciadas por uma fórmula (intervalos M87:P93 expandidos)."""
    f = str(f or "").replace("$", "")
    rows = set()
    for a, b in re.findall(r"(?<![A-Z])M(\d+):[A-Z]+(\d+)", f):
        rows |= set(range(int(a), int(b) + 1))
    rows |= {int(x) for x in re.findall(r"(?<![A-Z])M(\d+)", re.sub(r"(?<![A-Z])M\d+:[A-Z]+\d+", "", f))}
    return rows


def acertar_quadro(ws):
    """Quadros da aba conferidos pelos RÓTULOS (abas copiadas/modelos do cliente apontavam outra linha):
    'MEDIDO ANTERIORMENTE' soma todo o ACUMULADO BMS ANTERIORES; 'MEDIDO NO MÊS' aponta o mês;
    'Executado' = TOTAL EXECUTADO; HE no modelo 'Executado Anterior / Medição Atual' idem. Devolve as mudanças."""
    f = ws.Range("A1:AF170").Formula
    cel = lambda r, c: f[r - 1][c - 1]
    cab, medido, hist = None, None, []
    for r in range(70, 161):
        c = norm(cel(r, 3) or "")
        if c == "BM" and norm(cel(r, 5) or "").startswith("PERIODO"):
            cab = r
        elif c == "ACUMULADO BMS ANTERIORES" and cab:
            hist.append(r)
        elif c == "MEDIDO NO MES" and cab:
            medido = r
            break
    quadro = {}
    for r in range(30, 91):
        t = norm(cel(r, 9) or "")
        if t in ("TOTAL EXECUTADO", "MEDIDO ANTERIORMENTE", "MEDIDO NO MES") and t not in quadro:
            quadro[t] = r
    mud = []
    if cab and hist and medido and len(quadro) == 3:
        r = quadro["MEDIDO ANTERIORMENTE"]
        if str(cel(r, 4) or "").startswith("=") and not set(hist) <= linhas_m(cel(r, 4)):
            ws.Cells(r, 4).Formula = f"=SUM(M{hist[0]}:P{hist[-1]})"
            mud.append(f"D{r} {cel(r, 4)} -> =SUM(M{hist[0]}:P{hist[-1]})")
        r = quadro["MEDIDO NO MES"]
        if str(cel(r, 4) or "").startswith("=") and not ({medido, cab + 1} & linhas_m(cel(r, 4))):
            ws.Cells(r, 4).Formula = f"=M{medido}"
            mud.append(f"D{r} {cel(r, 4)} -> =M{medido}")
    if len(quadro) == 3:
        for r in range(15, 23):
            for c in (21, 25, 29):
                rot = norm(cel(r, c) or "")
                certo = {"EXECUTADO": quadro["TOTAL EXECUTADO"], "EXECUTADO ANTERIOR": quadro["MEDIDO ANTERIORMENTE"],
                         "MEDICAO ATUAL": quadro["MEDIDO NO MES"]}.get(rot)
                if certo and str(cel(r + 1, c) or "").replace("$", "") != f"=D{certo}":
                    ws.Cells(r + 1, c).Formula = f"=D{certo}"
                    mud.append(f"{ws.Cells(r + 1, c).Address.replace('$', '')} {cel(r + 1, c)!r} -> =D{certo}")
    return mud


def limpar_historico(ws):
    """Aba nova (copiada de outra): linhas do mês, histórico e OBSERVAÇÕES da aba de origem saem
    (ex.: 'DE MAIO /2026, HORAS QUE NÃO ESTAVA...' ia parar na aba nova)."""
    cab, hist, _ = bloco_historico(ws)
    if cab is None:
        return
    for r in range(cab + 1, hist[0] if hist else cab + 2):
        for c in (3, 5, 9, 13, 17):
            if not ws.Cells(r, c).HasFormula:
                ws.Cells(r, c).Value = None
    for r in hist:
        ws.Cells(r, 13).Value = None
        ws.Cells(r, 17).Value = None


def apagar_print_memoria(ws):
    """Apaga as imagens da área MEMÓRIA DE CÁLCULO da 1ª página (vieram da aba copiada)."""
    col_a = [v for (v,) in ws.Range("A1:A40").Value]
    col_i = [v for (v,) in ws.Range("I40:I80").Value]
    ini = next((i + 1 for i, v in enumerate(col_a) if norm(v or "") == "MEMORIA DE CALCULO"), 22)
    fim = next((40 + i for i, v in enumerate(col_i) if norm(v or "") == "TOTAL EXECUTADO"), 60)
    n = 0
    for shp in list(ws.Shapes):
        try:
            if ini < shp.TopLeftCell.Row < fim:
                shp.Delete()
                n += 1
        except Exception:
            pass
    return n


def por_previsto(ws, qtd):
    """PREVISTO (célula abaixo do rótulo 'Previsto'; o layout muda entre abas) = quantidade contratada.
    Sem contrato fica EM BRANCO (decisão do usuário, 03/10: o valor vem com a proposta revisada — não o da aba
    copiada) e o avanço físico / saldo da aba ficam em branco em vez de #DIV/0! ou saldo negativo."""
    f = ws.Range("A1:AF90").Formula
    cel = lambda r, c: f[r - 1][c - 1]
    prev = next(((r + 1, 21) for r in range(15, 23) if norm(cel(r, 21) or "") == "PREVISTO"), None)
    if prev is None:
        return
    ws.Cells(*prev).Value = qtd
    if qtd is not None:
        return
    end = ws.Cells(*prev).Address.replace("$", "")
    for r in range(30, 91):
        if norm(cel(r, 9) or "") == "AVANCO FISICO":
            fa = str(cel(r, 4) or "")
            if fa.startswith("=") and not fa.upper().startswith("=IFERROR"):
                ws.Cells(r, 4).Formula = f'=IFERROR({fa[1:]},"")'
            break
    for r in range(15, 23):
        if norm(cel(r, 29) or "") == "SALDO A MEDIR":
            fs = str(cel(r + 1, 29) or "")
            if fs.startswith("=") and not fs.upper().startswith("=IF("):
                ws.Cells(r + 1, 29).Formula = f'=IF({end}="","",{fs[1:]})'
            break


def cms_abas(wb, pq, log):
    """CMS da aba de item (V11, rótulo 'CMS' em V10) = o da PQ para o código da aba (K11): aba criada como cópia
    de outra herdava o CMS da irmã — mesma regra do código CMS do BM Acumulado (usuário, 13/10/2026)."""
    trocas, sem = [], set()
    for ws in wb.Worksheets:
        if not re.match(r"^\d+(\.\d+)+", ws.Name.strip()) or norm(ws.Range("V10").Value or "") != "CMS":
            continue
        cod = str(ws.Range("K11").Value or "").strip().lstrip("'")
        if not eh_codigo(cod):
            continue
        cms = (pq.get(cod) or (None,) * 4)[3]
        if cms in (None, "", "-"):
            sem.add(cod)
            continue
        atual = ws.Range("V11").Value
        if " ".join(str(atual or "").split()) != " ".join(str(cms).split()):
            trocas.append(f"{ws.Name.strip()} (era {atual!r})")
            ws.Range("V11").Value = cms
    if trocas:
        log.append(f"   BM de abas: CMS acertado pela PQ em {len(trocas)} aba(s): {', '.join(trocas)}")
    if sem:
        log.append(f"   !! BM de abas: PQ sem código CMS para {', '.join(sorted(sem))} — conferir à mão")


def saldo_vermelho(wb, log):
    """'SALDO A MEDIR' de cada aba de item: negativo em vermelho e negrito (formatação condicional, vale também
    para edição à mão) e listado no relatório — mesma regra do saldo do BM Acumulado (usuário, 13/10/2026)."""
    neg = []
    for ws in wb.Worksheets:
        if not re.match(r"^\d+(\.\d+)+", ws.Name.strip()):
            continue
        rot = ws.Range("AC15:AC22").Value
        r = next((15 + i for i, (v,) in enumerate(rot) if norm(v or "") == "SALDO A MEDIR"), None)
        if r is None:
            continue
        cel = ws.Cells(r + 1, 29)
        ws.Calculate()                                # o aplicar_bm edita com cálculo manual
        cel.FormatConditions.Delete()
        fc = cel.FormatConditions.Add(1, 6, "=0")     # xlCellValue, xlLess
        fc.Font.Color = 255
        fc.Font.Bold = True
        if isinstance(cel.Value, (int, float)) and cel.Value < -0.005:
            neg.append(f"{ws.Name.strip()} ({cel.Value:g})")
    if neg:
        log.append(f"   !! BM de abas: SALDO A MEDIR negativo em {', '.join(neg)}")


def criar_aba(xl, wb, abas, chave, qtd, hh, log):
    """Cria a aba do item copiando uma aba irmã. Devolve a aba nova (ou None)."""
    cod, mae = chave_cod(chave)
    nomes = {ws.Name.strip() for ws in wb.Worksheets}
    if mae:
        mae_ws = next((ws for ws, k in abas if k == mae), None)
        # sem aba do mesmo adicional (ex.: nenhuma 1.5.3 no BM): qualquer aba 1.5.x tem o mesmo layout;
        # o código (K11/A15) é trocado depois da cópia
        modelo = (next((ws for ws, k in abas if k and k.endswith("|" + cod)), None)
                  or next((ws for ws, _ in abas if re.match(re.escape(cod) + r"(\s|$)", ws.Name.strip())), None)
                  or next((ws for ws, _ in abas if ws.Name.strip().startswith("1.5.")), None))
        if mae_ws is None or modelo is None:
            log.append(f"   {chave}: sem aba da função {mae} ou sem aba {cod} para copiar — não criei")
            return None
        depois = mae_ws
        i = mae_ws.Index + 1
        while i <= wb.Sheets.Count and wb.Sheets(i).Name.strip().startswith("1.5."):
            depois = wb.Sheets(i)
            i += 1
        palavra = norm(hh.get(mae, "X")).split()[0]
        nome = None
        for n in range(3, len(palavra) + 1):
            if f"{cod} {palavra[:n]}" not in nomes:
                nome = f"{cod} {palavra[:n]}"
                break
        nome = nome or f"{cod} {mae}"
    else:
        itens = [(ws, k) for ws, k in abas if k and "|" not in k and eh_codigo(k)]
        pai = cod.rsplit(".", 1)[0]
        prof = cod.count(".")
        cands = ([x for x in itens if x[1].rsplit(".", 1)[0] == pai]
                 or [x for x in itens if x[1].count(".") == prof and x[1].split(".")[:2] == cod.split(".")[:2]]
                 or [x for x in itens if x[1].count(".") == prof] or itens)
        if not cands:
            log.append(f"   {chave}: nenhuma aba de item para copiar — não criei")
            return None
        modelo = min(cands, key=lambda x: abs(int(x[1].split(".")[-1]) - int(cod.split(".")[-1])))[0]
        antes = [x for x in itens if x[1].split(".")[0] == cod.split(".")[0] and ordem(x[1]) < ordem(cod)]
        depois = max(antes, key=lambda x: ordem(x[1]))[0] if antes else modelo
        # se a aba anterior tem horas extras logo depois, a nova entra depois delas
        i = depois.Index + 1
        while i <= wb.Sheets.Count and wb.Sheets(i).Name.strip().startswith("1.5."):
            depois = wb.Sheets(i)
            i += 1
        nome = cod
    vis = modelo.Visible
    modelo.Visible = VISIVEL
    antes = {s.Name for s in wb.Sheets}
    # Copy(After=X) pula as abas OCULTAS que vêm depois de X; por isso copia ANTES da aba
    # seguinte, deixando-a visível durante a cópia. Argumentos posicionais (Before, After):
    # por nome o pywin32 ignora e o Excel cria uma pasta de trabalho nova.
    if depois.Index < wb.Sheets.Count:
        prox = wb.Sheets(depois.Index + 1)
        vis_prox = prox.Visible
        prox.Visible = VISIVEL
        modelo.Copy(prox)
        prox.Visible = vis_prox
    else:
        modelo.Copy(None, depois)
    modelo.Visible = vis
    # a cópia é a única aba com nome novo (ex.: '4.2.2.25 (2)') — posição/aba ativa não são confiáveis
    novas = [s for s in wb.Sheets if s.Name not in antes]
    if len(novas) != 1:
        log.append(f"   {chave}: a cópia da aba não saiu como esperado — não criei")
        return None
    nova = novas[0]
    nova.Name = nome
    if mae:
        # K11 do modelo '1.5.2' é fórmula do nome da aba: em '1.5.2 ENC' daria '1.5.2 ENC' e o PROCV na PQ
        # daria #N/D. As abas de HE por função do cliente têm o código digitado — igual aqui.
        nova.Range("K11").Value = txt_cod(cod)
        if str(nova.Range("A15").Value or "").strip().startswith("1.5."):   # hierarquia 1.5 / 1.5.x
            nova.Range("A15").Value = txt_cod(cod)
    else:
        partes = cod.split(".")
        for i in range(4):
            nova.Cells(14 + i, 1).Value = txt_cod(".".join(partes[:i + 1])) if i < len(partes) else None
    limpar_historico(nova)
    so_duas_paginas(nova, log)
    por_previsto(nova, qtd)
    n = apagar_print_memoria(nova)
    falta = "" if qtd is not None else " — Previsto em branco (sem quantidade contratada: vem com a proposta revisada)"
    log.append(f"   CRIADA aba '{nome}' (cópia de '{modelo.Name}', depois de '{depois.Name}'), "
               f"Previsto {qtd}{falta}; {n} imagem(ns) do print antigo apagada(s)")
    return nova


def so_duas_paginas(ws, log):
    """Aba criada no mês fica com as 2 páginas do modelo (informações + memória). Página a mais da aba copiada
    (5042 '1.6.1': 3ª página com o BM 5030 dela) é apagada e sai da área de impressão."""
    qb = [ws.HPageBreaks(i).Location.Row for i in range(1, ws.HPageBreaks.Count + 1)]
    _, _, medido = bloco_historico(ws)
    if len(qb) < 2 or not medido or medido >= qb[1]:      # memória da aba indo até a 3ª página: fica
        return
    fim = qb[1] - 1
    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    if ult > fim:
        ws.Rows(f"{fim + 1}:{ult}").Delete()
    if ws.PageSetup.PrintArea:
        ws.PageSetup.PrintArea = re.sub(r"\d+$", str(fim), ws.PageSetup.PrintArea)
    log.append(f"   aba '{ws.Name}': página a mais da aba copiada (linhas {fim + 1}–{ult}) apagada; impressão até a linha {fim}")


def enxugar(xl, wb, abas, manter, log):
    """PRT na 1ª medição (BM copiado de outra PRT): apaga as abas de item fora do escopo e sem
    medição — não têm histórico nesta PRT. Fica a aba referenciada por fórmula da 1.1.1/PQ."""
    refs = set()
    for nome in ("1.1.1", "PQ", "PQ "):
        try:
            ur = wb.Worksheets(nome).UsedRange
        except Exception:
            continue
        for f in (ur.Formula if isinstance(ur.Formula, tuple) else ((ur.Formula,),)):
            for v in f:
                if isinstance(v, str) and v.startswith("="):
                    refs.update(t.strip() for t in re.findall(r"'?([^'!=(,+*/-]+)'?!", v))
    apagar = [ws for ws, k in abas if ws.Name.strip() not in ("1.1.1",) and ws.Name.strip() not in refs
              and (k is None or k not in manter)]
    xl.app.DisplayAlerts = False
    for ws in apagar:
        ws.Delete()
    xl.app.DisplayAlerts = True
    log.append(f"   BM de 1ª medição enxugado: {len(apagar)} abas de item de outra PRT apagadas "
               f"(ficaram {wb.Worksheets.Count} abas)")


def modelo_bm_nova():
    """BM do mês passado com o menor arquivo — base para PRT na 1ª medição."""
    drive_ok(C.FONTE_ACUMULADO)
    cands = [f for f in glob.glob(os.path.join(C.FONTE_ACUMULADO, "PRT*", "*.xls*"))
             if not os.path.basename(f).startswith("~$")
             and norm(os.path.basename(f)).startswith("BM") and "ACUMULADO" not in norm(os.path.basename(f))
             and "CONSOLIDA" not in norm(os.path.basename(f))]
    return min(cands, key=os.path.getsize)


def aplicar_bm(xl, src, med, valores, log, esc=None, prt=None, nova=False):
    esc = esc or {}
    acu = fonte(prt, "acumulado") if prt else None
    prev = {}   # {chave: (acumulado até o BM anterior, medido no BM anterior, nº do BM anterior)}
    if acu:   # Previsto: quantidade contratada do Acumulado quando a aba Escopo não traz
        linhas_acu = ler_acumulado(acu)[1]
        contr = {l["chave"]: l["contratado"] for l in linhas_acu if num(l.get("contratado"))}
        esc = {**contr, **{k: v for k, v in esc.items() if v is not None}}
        m = re.search(r"BM (\d{4})", os.path.dirname(os.path.dirname(acu)))
        if m and not nova:
            val = lambda v: v if isinstance(v, (int, float)) else 0
            prev = {l["chave"]: (val(l["atual"]), val(l["atual"]) - val(l["anterior"]), m.group(1))
                    for l in linhas_acu}
            # item sem linha no Acumulado = nada medido até o BM anterior (ex.: aba que veio de modelo
            # de outra PRT com 'BM 5028 = 31' na linha do mês): não vai para o histórico
            for ws_k in [k for k in valores if k not in prev]:
                prev[ws_k] = (0.0, 0.0, m.group(1))
    if not any(valores.values()):
        log.append("   nenhuma quantidade no mês — BM não gerado")
        return None
    lm = ler_medicao(med) if med else []
    hh = {l["cod"]: str(l["desc"]) for l in lm if str(l["und"]).strip().lower() == "hh"}
    if nova:
        destino = os.path.join(C.SAIDA, os.path.basename(os.path.dirname(med)))
        dst = os.path.join(destino, f"BM - PRT_{prt}.xlsx")
        log.append(f"   BM NOVO (1ª medição) a partir de {os.path.basename(src)}")
    else:
        destino = os.path.join(C.SAIDA, os.path.basename(os.path.dirname(src)))
        # 'BM30_-_PRT_5022.xlsx' na pasta do BM 5030 (renomeado no Drive em 06/10): o nº de BM no nome é de outro mês
        dst = os.path.join(destino, re.sub(r"^BM\s*\d+\s*_", "BM_", os.path.basename(src)))
    os.makedirs(destino, exist_ok=True)
    shutil.copy2(src, dst)
    wb = xl.abrir(dst)
    xl.app.ScreenUpdating = False
    xl.app.Calculation = -4135  # manual durante a edição
    try:
        wb.Worksheets("1.1.1").Range("A11").Value = f"{C.PERIODO_INI:%d/%m/%Y} A {C.PERIODO_FIM:%d/%m/%Y}"
        if nova:
            wb.Worksheets("1.1.1").Range("X8").Value = f"BM_{prt_info(med)[0]}"
    except Exception:
        log.append("   aba '1.1.1' não encontrada — período/nº do BM do cabeçalho não atualizado")
    abas = mapa_abas(wb, hh, ({l["chave"] for l in lm if l["cod"] in ADICIONAIS} | {k for k in valores if "|" in k}))
    if nova:
        for ws, chave in abas:
            limpar_historico(ws)
            por_previsto(ws, esc.get(chave) if chave else None)
    # hora extra sem aba própria: reaproveita aba de HE da mesma sigla (ex.: '1.5.2 SOL' para
    # Soldador Chaparia) que ficou com outra função de mesma sigla sem valor no mês — só se o que a aba já
    # tem medido é o acumulado desta HE no BM Acumulado (5007, BM 5031: '1.5.2 SOL' tem as 95 h do Soldador
    # Chaparia; a HE do TIG, 0 até o BM 5030, ganha aba própria)
    remap = {}
    for chave in [k for k, v in valores.items() if v and "|" in k and k not in {c for _, c in abas}]:
        cod, mae = chave_cod(chave)
        palavra = norm(hh.get(mae, "")).split()[:1]
        for ws, k2 in abas:
            sig = [t for t in norm(ws.Name).split()[1:] if not eh_codigo(t)]
            if (k2 and k2 != chave and k2.endswith("|" + cod) and not valores.get(k2)
                    and ws.Name.strip() not in remap and palavra and sig
                    and any(palavra[0].startswith(s) for s in sig)
                    and abs(historico(ws, com_mes=True) - (prev.get(chave) or (0.0,))[0]) < 0.005):
                remap[ws.Name.strip()] = chave
                log.append(f"   aba '{ws.Name}' (era da função {chave_cod(k2)[1]}) usada para a hora extra de {mae}")
                break
    reaplicar = lambda ab: [(ws, remap.get(ws.Name.strip(), k)) for ws, k in ab]
    abas = reaplicar(abas)
    existentes = {k for _, k in abas if k}
    for chave in sorted((k for k, v in valores.items() if v and k not in existentes),
                        key=lambda k: ("|" in k, ordem(chave_cod(k)[0]))):
        if criar_aba(xl, wb, abas, chave, esc.get(chave), hh, log):
            if (prev.get(chave) or (0.0,))[0]:    # 5007, BM 5031: '1.5.2 CAL' nasceu sem as 118 h já faturadas
                log.append(f"   !! aba de {chave} criada sem histórico, mas o BM Acumulado já tem {prev[chave][0]:g} até o "
                           f"BM anterior — lançar BM a BM no acertos_historico")
            abas = reaplicar(mapa_abas(wb, hh, ({l["chave"] for l in lm if l["cod"] in ADICIONAIS} | {k for k in valores if "|" in k})))
    usadas, trocar_print, roladas_ocultas = set(), [], []
    for ws, chave in abas:
        v = valores.get(chave) if chave else None
        if chave is None:
            if ws.Visible == VISIVEL:
                ws.Visible = OCULTA
                log.append(f"   aba '{ws.Name}': hora extra sem função antes dela e sem uso no mês — ocultei")
            continue
        if not v and rolar_aba(ws, 0, log, xl, prev.get(chave), so_confirmado=True):
            roladas_ocultas.append(ws.Name)   # medido no mês passado, não neste: histórico em dia
        if v:
            # aba de modelo com histórico de outro contrato (5029 '4.2.2.19': BM-008…BM-011 = 883 h) e o BM
            # Acumulado diz 0 até o BM anterior: o histórico sai, como numa aba nova
            h = historico(ws)
            if h and prev.get(chave, (None,))[0] == 0:
                limpar_historico(ws)
                log.append(f"   aba '{ws.Name}': histórico de fora desta PRT ({h:g}) apagado — o BM Acumulado diz 0 até o BM anterior")
            if rolar_aba(ws, v, log, xl, prev.get(chave)):
                usadas.add(chave)
                trocar_print.append(ws.Name)
                if nova:
                    apagar_print_memoria(ws)
            ws.Visible = VISIVEL
        elif ws.Visible == VISIVEL and ws.Name.strip() != "1.1.1":   # 1.1.1 = cabeçalho da PRT
            ws.Visible = OCULTA
    if nova:
        enxugar(xl, wb, abas, set(esc) | usadas, log)
    for chave, v in valores.items():
        if v and chave not in usadas:
            log.append(f"   {chave} = {v}: não consegui lançar no BM (ver linhas acima)")
    abas = reaplicar(mapa_abas(wb, hh, ({l["chave"] for l in lm if l["cod"] in ADICIONAIS} | {k for k in valores if "|" in k})))
    ordenar_abas(wb, abas, arquivo(destino, "acumulado"), log)   # mapa de novo: enxugar pode ter apagado abas
    quadros = []
    for s in wb.Worksheets:
        if re.match(r"^\d+(\.\d+)+", s.Name.strip()) and s.Name.strip() != "1.1.1":
            if acertar_quadro(s):
                quadros.append(s.Name.strip())
    if quadros:
        log.append(f"   fórmulas de Executado / Medido anteriormente / Medido no mês acertadas nas abas: {', '.join(quadros)}")
    # só as abas medidas no mês ficam visíveis (a PQ é base de preços, não vai no book)
    next(s for s in wb.Worksheets if s.Visible == VISIVEL and s.Name.strip().upper() != "PQ").Activate()
    for s in wb.Worksheets:
        if s.Name.strip().upper() == "PQ" and s.Visible == VISIVEL:
            s.Visible = OCULTA
            log.append(f"   aba '{s.Name}' oculta (só as abas medidas no mês ficam visíveis)")
    xl.app.Calculation = -4105
    for aba_ in wb.Worksheets:
        tirar_ajustes(aba_, "BM", log)
    casas_bm(wb, log)
    brancas = [s.Name.strip() for s in wb.Worksheets if s.Visible == VISIVEL
               and re.match(r"^\d+(\.\d+)+", s.Name.strip()) and sem_pagina_em_branco(s)]
    if brancas:
        log.append(f"   página final em branco tirada da impressão nas abas: {', '.join(brancas)}")
    padronizar_fontes(wb, "BM", log)
    if med:
        cms_abas(wb, pq_da_medicao(med), log)
    saldo_vermelho(wb, log)
    wb.Save()
    wb.Close()
    log.append(f"   BM: {len(usadas)} abas atualizadas e visíveis; demais abas de item ocultas")
    if roladas_ocultas:
        log.append(f"   abas ocultas (sem medição no mês) com o último BM levado ao histórico: {', '.join(roladas_ocultas)}")
    if trocar_print:
        log.append(f"   TROCAR PRINT DA CONSOLIDAÇÃO/MEMÓRIA nas abas: {', '.join(trocar_print)}")
    return dst


def sem_pagina_em_branco(ws):
    """Área de impressão que passa da última quebra só com linhas vazias imprime página em branco no book
    (5007 '1.4.2': área até 146, quebra em 146). A área termina antes da quebra — conteúdo e imagens não mudam."""
    cortes = 0
    vazio = lambda de, ate, c1, c2: (not any(de <= sh.TopLeftCell.Row <= ate for sh in ws.Shapes)
                                     and not ws.Application.WorksheetFunction.CountA(ws.Range(ws.Cells(de, c1), ws.Cells(ate, c2))))
    while ws.PageSetup.PrintArea:
        a = ws.Range(ws.PageSetup.PrintArea)
        fim, ult_col = a.Row + a.Rows.Count - 1, a.Column + a.Columns.Count - 1
        qb = [(b.Location.Row, b.Type) for b in (ws.HPageBreaks(i) for i in range(1, ws.HPageBreaks.Count + 1))
              if b.Location.Row <= fim]
        # página do meio só com linhas vazias antes de uma quebra manual (5019 '2.13.7.14': quebra automática
        # em 72 e manual em 73 — página só com a linha 72): as linhas vazias saem
        meio = next(((q1, q2) for (q1, _), (q2, t2) in zip(qb, qb[1:]) if t2 == -4135 and vazio(q1, q2 - 1, a.Column, ult_col)), None)
        if meio:
            ws.Rows(f"{meio[0]}:{meio[1] - 1}").Delete()
        elif qb and vazio(qb[-1][0], fim, a.Column, ult_col):
            ws.PageSetup.PrintArea = ws.Range(ws.Cells(a.Row, a.Column), ws.Cells(qb[-1][0] - 1, ult_col)).Address
        else:
            break
        cortes += 1
    return cortes


def casas_bm(wb, log):
    """Abas visíveis de item: Previsto/Executado, quadro e memória (colunas D, M, U, Y) mostram todas as casas
    do valor — 5,1914 não pode aparecer 5,191 (nada arredondado, nem na tela)."""
    for s in wb.Worksheets:
        if s.Visible == VISIVEL and re.match(r"^\d+(\.\d+)+", s.Name.strip()) and s.Name.strip() != "1.1.1":
            # o PREVISTO fica com o conferir_previsto (ainda pode trocar 571,9845 por 572: não alarga antes)
            u = s.Range("U15:U22").Value
            p = next((15 + i + 1 for i, (v,) in enumerate(u) if norm(str(v or "")) == "PREVISTO"), None)
            feito = mostrar_casas(s, "A15:Z170", (4, 13, 25))
            for faixa in ((f"U15:U{p - 1}", f"U{p + 1}:U170") if p else ("U15:U170",)):
                feito += mostrar_casas(s, faixa, (21,))
            if feito:
                log.append(f"   aba '{s.Name.strip()}': formato mostrava menos casas que o valor — {', '.join(feito)}")


def main():
    args = [a for a in sys.argv[1:] if not re.fullmatch(r"\d{4}", a)]
    so = [a for a in sys.argv[1:] if re.fullmatch(r"\d{4}", a)]
    caminho = args[0] if args else os.path.join(C.SAIDA, f"Base_do_Mes_BM{C.BM_NUMERO}.xlsx")
    base, erros = ler_base(caminho)
    if erros:
        print("A Base tem pendências — nada foi gerado:")
        for e in erros:
            print("  -", e)
        sys.exit(1)
    escopo = ler_escopo(caminho)
    log = []
    with Excel() as xl:
        for prt in sorted(so or base):
            print(f"PRT {prt}...", flush=True)
            log.append(f"\nPRT {prt}")
            bm = fonte(prt, "bm")
            # Medição já gerada no mês (aplicar_base): tem as funções criadas agora (ex.: 4.2.3.4), que
            # nomeiam/ligam as abas de hora extra; sem ela, a do mês anterior
            med = arquivo(pasta_prt(C.SAIDA, prt, C.PASTA_DA_PRT), "medicao") or fonte(prt, "medicao")
            if bm:
                aplicar_bm(xl, bm, med, base.get(prt, {}), log, escopo.get(prt), prt)
            elif med:
                aplicar_bm(xl, modelo_bm_nova(), med, base.get(prt, {}), log, escopo.get(prt), prt, nova=True)
            else:
                log.append("   sem BM do mês anterior nem Medição — não gerado")
            print("\n".join(log[-3:]), flush=True)
    rel = os.path.join(C.SAIDA, "relatorio_bm.txt")
    with open(rel, "w", encoding="utf-8") as f:
        f.write("\n".join(log))
    print("\n".join(log))
    print(f"\nRelatório: {rel}")


if __name__ == "__main__":
    main()
