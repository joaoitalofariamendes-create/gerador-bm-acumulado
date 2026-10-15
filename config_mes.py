# -*- coding: utf-8 -*-
"""Único arquivo que muda de um mês para o outro."""
from datetime import date

RAIZ = r"<DRIVE_BOLETIM>\DOCUMENTOS - MEDIÇÃO\BM MEDIÇÃO\Ano 2026"

BM_NUMERO = "5031"
PERIODO_INI = date(2026, 9, 6)
PERIODO_FIM = date(2026, 10, 5)
CADERNO = r"<PASTA_TRABALHO>\Medicoes\guia_medições\Caderno_de_Apoio - 06-09 a 05-10.xlsx"

# de onde vêm os modelos que serão copiados e preenchidos (normalmente: pasta do mês anterior)
FONTE_MEDICAO = RAIZ + r"\MES 09-2026 - BM 5030"
FONTE_ACUMULADO = RAIZ + r"\MES 09-2026 - BM 5030"

# para onde vão os arquivos gerados (NUNCA escreve na pasta do Drive)
SAIDA = r"<PASTA_TRABALHO>\Medicoes\gerador_bm_acumulado\saida_BM" + BM_NUMERO

# quando o número da PRT aparece em mais de uma pasta (ex.: "PRT 5016-25" e "PRT 5016-26")
PASTA_DA_PRT = {"5016": "PRT 5016-25"}

# PRT na 1ª medição SEM nenhum histórico (nem Acumulado, nem Medição, nem pasta no mês anterior):
# caminho do xlsx da PRT aprovada — dele saem os itens, unidades, quantidades e preços exatos.
PRT_NOVA = {"5019": r"<DRIVE_GESTAO>\CONTRATO-4600000001\PRT - ANO 2026\PRT_5019 - FORN E MONT TUB E EQUIP CASA DE BOMBA\PROPOSTA\PRT-GSV-DOP-B-2026-5019_rD.xlsx"}

# fotos do mês separadas por PRT (saída do fotos_whatsapp.py) — usadas pelo aplicar_fotos.py na aba FOTOS
FOTOS = r"<PASTA_TRABALHO>\Medicoes\fotos_whatsapp_teste"

# PRTs sem aba FOTOS no BM que devem ganhar uma (cópia da de outro BM) — as demais ficam como estão
# (5022 não tem aba FOTOS, por decisão do usuário)
FOTOS_CRIAR_ABA = {"5028", "5019"}

# data de postagem do BM ("Data BM:" do cabeçalho do Acumulado); None = mantém a do mês passado
DATA_BM = date(2026, 10, 6)

# modelo da Consolidação de Horas (várias abas) — a planilha padrão do usuário; None = a da pasta do mês anterior
# (BM 5031: a da pasta BM30 é a revisada — 5042 sem o bloco do Engenheiro Mecânico, cobrado como Supervisor)
MODELO_CONSOLIDACAO = r"<PASTA_TRABALHO>\Medicoes\BM30\BM GSV- CONSOLIDAÇÃO DE HORAS - 06-08 a 05-09-2026 - BM030.xlsx"
