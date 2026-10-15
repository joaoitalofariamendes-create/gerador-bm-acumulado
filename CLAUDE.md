# Gerador de BM (GSV / CLIENTE) — mapa rápido

Leia isto antes de explorar. Respostas curtas, em português.

## Ambiente
- Pasta de trabalho: `<PASTA_TRABALHO>` (desde 09/10/2026; usada pelo usuário e pelo colega — não rodar o
  mesmo script nos dois PCs ao mesmo tempo). A cópia antiga em `C:\Users\GSV\Desktop\Trabalho` não é mais usada.
- PC novo: dois cliques em `_config_claude\INSTALAR.bat` (Git, Python + libs, Claude Code + VS Code, skills, plugins, memória;
  pode rodar de novo). Skill/memória nova no PC do usuário: copiar também para `_config_claude\skills` / `\memoria`.
- Python: `& $env:GSV_PYTHON` (caminho do Python de cada PC, gravado pelo `_config_claude\configurar.ps1`; não há `python` no PATH). Use `$env:PYTHONIOENCODING='utf-8'` e `-W ignore`.
- Excel via COM (`comum.Excel`, DispatchEx). Se um script cair, sobra EXCEL.EXE oculto: matar só os sem janela
  (`Get-Process EXCEL | ? { $_.MainWindowTitle -eq '' } | Stop-Process -Force`), e só se foram abertos pelo script.
- Drive BOLETIM (fonte, SOMENTE LEITURA): `config_mes.RAIZ` → `MES mm-aaaa - BM nnnn\PRT xxxx\`. Nunca gravar nele (nem nos drives de RDO).

## Pastas
Raiz = só o pipeline em uso. `_antigos\` = scripts aposentados (não usar). `_arquivo\` = saídas/simulações de meses
passados. `books_ordenados\` = books (PDF) reordenados localmente — o original do Drive é trocado pelo usuário.
`Pendências.xlsx` = horas fora da proposta (só itens não orçados), com RDOs de origem.

## Configuração do mês
`config_mes.py` é o único arquivo que muda por mês (BM_NUMERO, período, CADERNO, FONTE_MEDICAO/FONTE_ACUMULADO = pasta do mês anterior, SAIDA,
DATA_BM = data de postagem → "Data BM" do Acumulado; BM 5031: 06/10/2026).
Simulação: copiar os `.py` para uma pasta de rascunho com um `config_mes.py` próprio (ex.: BM 5031 → SAIDA `SIMULACAO_BM5031`, fontes = pasta do BM 5030) e rodar de lá.

## Pipeline (ordem)
1. `montar_base.py` — monta `Base_do_Mes_BMnnnn.xlsx` (Caderno de Apoio → quantidades por PRT).
2. `aplicar_consolidacao.py` — Consolidação de Horas: UMA PLANILHA POR PRT (só a aba dela) em `SAIDA\PRT xxxx\`, só PRTs
   com hora; modelo = `config_mes.MODELO_CONSOLIDACAO` (BM 5031: a da pasta `Medicoes\BM30` do usuário); a de várias
   abas vai para `_arquivo\consolidacao_varias_abas` (modelo do mês seguinte, se o Drive não tiver uma). Retroativos
   (aba Retroativos do Caderno = `PENDENCIAS_DIAS` no gerar_caderno) em colunas de data ANTES do ciclo + RDOs nas
   observações; noturno em linha '<FUNÇÃO> ADICIONAL NOTURNO'; data de entrega = DATA_BM. Conferir com
   `conferir_consolidacao.py` (dia a dia contra o Caderno).
3. `aplicar_base.py [PRT...]` — Medição + BM Acumulado → `relatorio_aplicacao.txt`.
4. `aplicar_bm.py [PRT...]` — BM de abas (histórico "ACUMULADO BMS ANTERIORES") → `relatorio_bm.txt`.
   Depois: `acertos_historico.py` (da pasta da simulação/mês; contrato revisado, histórico) → `aplicar_fotos.py` →
   `conferir_previsto.py --corrigir` (Previsto das abas = E do Acumulado; sem contrato fica em branco).
5. `conferir_mes.py [PRT...]` — conferência obrigatória → `conferencia_mes.txt` (ler este arquivo, é curto).
   Confere também: total da Medição = valor do mês (P) do Acumulado, subtotais cobrindo todas as linhas, S/T/U do
   Acumulado em fórmula, Executado das abas = L do Acumulado, histórico inteiro no "Medido anteriormente", datas sem hora, PQ oculta.
Capa (Word): `capa_medicao.py [PRT...]` → `CAPA DA MEDIÇÃO - PRT nnnn.docx` na pasta de cada PRT (1ª página do
`Medicoes\CAPA DA MEDIÇÃO.docx`; ID/NIQ/CC da Medição, título da proposta em cache_md; faltas em TITULO/CAMPOS).
PDFs de entrega: `exportar_pdf.py [PRT...]` (Medição = aba Formuário_PRT; Acumulado = 1ª aba visível; BM de abas = todas as visíveis).
Medição/Acumulado saem com campo de assinatura vazio; no BM de abas cada página é deslocada para o quadro 'Emissão (CONTRATADA)'
cair no mesmo ponto (o Foxit assina todas as páginas no mesmo lugar). Regerar o PDF apaga assinaturas já feitas.
Revisão do histórico: `conferir_historico.py` → `Conferencia_Historico_BMnnnn.xlsx` (aba a aba: histórico = Acum. Anterior,
rolagem do BM passado, J/N = L/R do mês passado). Vale para as abas do book; aba oculta só conta quando voltar a medir.
Antes de confiar na Base: gerar uma Base nova num rascunho (montar_base com SAIDA de teste) e comparar — em 01/10 a Base perdeu a PRT 5043 inteira.
Filtrar PRT reescreve o relatório só com ela: para relatório completo rode sem filtro.
PRT sem medição no mês não gera arquivo. O BM de abas sai com as abas na ordem das linhas do BM Acumulado.
RDOs (somente leitura): `<DRIVE_RDO>\RDO` (xlsx + pdf).

Fotos: `fotos_whatsapp.py <export.zip> [ini] [fim] [saída]` separa as fotos do grupo por PRT/dia (remetente → cadastro →
projeto apontado no dia; sem certeza → "A REVISAR"). `aplicar_fotos.py` (depois do passo 4) preenche a aba FOTOS
de cada BM com fotos espalhadas pelos dias, na grade que já existe (logos ficam). Pasta: `config_mes.FOTOS`.
RDO pelos grupos do WhatsApp (`Medicoes\whatsapp_coletor\`): `processar_dia.py` (aba de fotos + efetivo x apontamento),
`rdo_engenharia.py` / `rdo_preparacao.py` (rascunhos de RDO). O coletor (`coletor.mjs`) roda SÓ no PC do usuário, como tarefa
agendada: nunca iniciar em outro PC (a sessão do WhatsApp é uma só; dois coletores derrubam a conexão).
Depois do passo 4, na simulação: `acertos_historico.py` (acertos pontuais de histórico/R$ Anterior de BMs passados, com motivo).
PRT na 1ª medição SEM histórico: `config_mes.PRT_NOVA = {"nnnn": <xlsx da PRT>}` — itens/qtd/preços exatos da aba Formuário_PRT.
PRT sem Medição no mês passado (5001, 5026, 5036, 5019 no BM 5031): a Medição sai de `SAIDA\_modelos\` — é reaproveitado
(cabeçalho com ID Planview/NIQ/nome/centro de custo/gestor preenchidos). Rodada nova do mesmo BM: manter/copiar o `_modelos`.
Prova de que o gerador reproduz o revisado: rodar tudo do zero num rascunho (mesma Base, `_modelos`, acertos) e
`comparar_saidas.py <revisada> <rascunho>` (célula a célula) — diferença = acerto feito à mão que o gerador não faz.

## Fechamento (rodada oficial depois da simulação revisada)
1. RDOs novos: `Projetos\jovem_aprendiz\extrator_rdo.py` (Apontamentos fechado; nomes "revisar à mão" → `correcoes_confirmadas.json`,
   tirar os PDFs do `historico_rdos.txt` e rodar de novo) → `guia_medições\gerar_caderno.py --ciclo dd/mm/aaaa`.
2. `montar_base.py` → `levar_decisoes.py <Base nova> <Base da simulação>` (leva valores decididos/zeros/linhas à mão;
   lista o que é novo para o usuário decidir). Regras fixas por PRT ficam em `montar_base.REMAPEAR`.
3. SAIDA com `_modelos\` e `acertos_historico.py` da simulação; pipeline (passos 2–5) + `conferir_historico.py`.
4. `levar_fotos.py <simulação> <SAIDA> PRT...` para as abas FOTOS ajustadas à mão; `comparar_saidas.py` contra a simulação
   (diferença esperada = só quantidades do mês).
5. Depois do fechamento (RDO corrigido etc.): remontar a Base e `python SAIDA\refazer_prt.py PRT` (Consolidação + a PRT +
   conferências completas; relatórios filtrados voltam para os completos). PC com pouca memória mata comando em segundo
   plano: rodar `refazer_prt.py PRT --rapido 0-1`, `2-3`, `4-7` (pedaços < 10 min) e depois `conferir_mes.py PRT...` /
   `conferir_historico.py PRT...` (com filtro eles atualizam só os blocos dessas PRTs no relatório completo). Aba nova com quantidade já faturada no
   Acumulado sai sem histórico: o aplicar_bm avisa ("!! aba de … criada sem histórico") — lançar BM a BM no acertos_historico.
   A simulação NÃO é a entrega (o usuário já editou a simulação achando que era a definitiva).

## Arquivos gerais do contrato (BM Geral, Resumo, Template OMS)
Todo mês, além das PRTs: `BM_..._GERAL.xlsx`, `BM_..._Resumo.xls` e `TEMPLATE MEDIÇÃO OMS ...xlsx`, feitos a partir dos do
mês anterior (separação: 5.x = locação, 2.x = estrutura, resto = MO). Scripts do BM 31: `_arquivo\bm_geral_BM5031\`
(trocar caminhos/PRTs/número); detalhes na memória `bm-geral-resumo-template-oms`. Saída: `Medicoes\BM nn_<Mês>`.
Grafo do código: `graphify-out\graph.html` (só os .py da raiz; refazer com `graphify-out\graphify_codigo.py`).

## Preparação de Cargas (contrato 4600000002)
Boletins em `BM Manutenção\<Equipe>\` (boletim + FOTOS.xlsx + Fotos\). Fixa = RDOs do Encarregado A; Caldeiraria = RDOs do Encarregado B
("ATIVIDADES: CALDEIRARIA"); Parada = projeto "- PARADA". Depois do Caderno: `boletins_preparacao.py <nº Fixa> <nº Parada>
<nº Caldeiraria>` (BM 5031: 037 038 039) — horas da aba Medição do Caderno, rolagem POR EQUIPE (J = J+K, N35 = N35+P35 do
boletim do mês passado), N.º/Data/período no cabeçalho, fotos na grade do FOTOS. Backup + relatório em
`_arquivo\backup_boletins_preparacao_BMnnnn` (o mês passado é lido de lá: rodar de novo não rola 2x).
Consolidação de Horas: `consolidacao_preparacao.py [modelo]` — uma planilha por equipe na pasta dela (aba Consolidação do
Caderno; funções do aplicar_consolidacao; 'HORAS 50 %' = 1ª faixa de HE; confere dia a dia no fim). Modelo = a de várias
abas do mês passado (BM 5031: `Medicoes\BM30\...Manutenção.xlsx`; depois a do `_arquivo\consolidacao_varias_abas`).

## Regras de negócio (resumo; detalhes na memória)
- Só se mede o que o Caderno traz (nunca repetir o mês anterior). Item apontado sem linha: criar (mesmo fora da PRT) e listar.
- Fiel ao RDO assinado: HE só se a coluna HE do RDO trouxer (inclusive sáb/dom); vale o Total do RDO. O Caderno classifica
  pelo horário — antes de fechar, `_arquivo\varredura_BM5031_0710\varredura_rdo.py` (HE, totais, pessoas, equipamentos/volumes)
  e conferir RDO alterado no Drive depois da extração (o extrator só lembra o nome do arquivo).
- Nada arredondado: valor cheio na célula (formato pode esconder casas). Totais = soma pura, sem número digitado; divergência vai para o relatório ("!!").
- Scripts param com erro se o arquivo estiver aberto em outro Excel (antes o Save se perdia).
- BM Acumulado: J = L do mês passado; N (R$ Anterior) = R (R$ Acumulado) do mês passado; E/F/G e totais da proposta repetem o mês passado.
- BM Acumulado (desde 13/10, `aplicar_base`): código CMS (col. A) de toda linha de item = o da aba PQ da Medição para o
  código (`cms_da_pq`; item sem CMS na PQ → "!!" para preencher à mão); saldo V/W negativo em vermelho e negrito por
  formatação condicional (vale também para edição à mão) e listado no relatório com "!!". BM de abas (`aplicar_bm`):
  CMS da aba (V11) = o da PQ para o código em K11 (`cms_abas`; aba copiada herdava o da irmã) e "SALDO A MEDIR"
  negativo em vermelho/negrito (`saldo_vermelho`; o relatório lista também abas ocultas — essas não vão para o PDF).
- Reajuste (O/Q/S) sempre zero. HE: preço = preço da função × 2 (1.5.1) ou × 1,75 (1.5.2) — % sobre a função, nunca preço fixo.
- BM de abas: o BM Acumulado do mês passado decide o que vai para o histórico e o rótulo; abas ocultas também rolam.
  Aba de HE ('1.5.2 ME') → função pela sigla, só entre funções com essa HE na Medição do mês (aplicar_bm e conferir_previsto
  iguais; ex.: ME/MEC = Mecânico Montador 4.2.2.10, não o Ajudante de Mecânica 4.2.2.13); fica logo após a aba da função.
  Visíveis SÓ as abas medidas no mês (PQ oculta; FOTOS só com fotos do mês). Datas via COM = número (`data_excel`), nunca datetime (vira 03:00).
- Medição: 6.x usa o preço da PRT (F do Acumulado) — o valor do 6.1 na PQ é o de outra PRT; adicional noturno (1.5.3) = função × %
  (o "Ad. Noturno" do Caderno é medido — TIPO_ADICIONAL; antes só o de pendência).
- Munck não tem hora extra: HE do operador de Munck vira hora normal do caminhão (`comum.HE_DO_OPERADOR`; Base e Consolidação).
- Função sem HE na proposta: a HE vira hora normal da função (`comum.HE_COMO_NORMAL`: 5042 Sup. Mecânica/Téc. Planejamento,
  5029 toda função). Caso novo: perguntar e acrescentar.
- Proposta válida (PDF na pasta da PRT em SAIDA) manda no contrato: E/F das linhas via `ACERTOS_CONTRATO` do
  acertos_historico (tipos E, F, HE = função × fator, codigo, criar, remover); conferir_mes lista esses acertos à parte.
- Só itens orçados na PRT; itens faltantes são criados clonando linha/aba irmã e listados no relatório.
- 5029 desde 28/09/2026: toda hora trabalhada = normal, inclusive sábado (`guia_medições/regras_horas.py` TUDO_NORMAL,
  sem data de fim — perguntar no BM 5032 se continua).

## PDFs
`pdf_para_md.py` gera `cache_md\<PRT>\*.md` (PQ lida + texto por página; só reconverte se o PDF mudou).
Leia o `.md`, não o PDF. Outros PDFs: `pdf_para_md.py <pdf|pasta>`.
