# Gerador de BM — automação de Boletins de Medição

Scripts Python que automatizam o fechamento mensal de **Boletins de Medição (BM)** de um contrato de montagem
eletromecânica: a partir das horas e quantidades apontadas nos RDOs (Relatórios Diários de Obra), geram e conferem as
planilhas de medição de cada projeto (PRT) no Excel.

> Cópia pública e **anonimizada**: nome do cliente, números de contrato, caminhos de rede, nomes de pessoas e valores
> reais foram trocados por marcadores (`CLIENTE`, `<PASTA_TRABALHO>`, `Encarregado A`…). Nenhuma planilha, PDF ou dado
> de medição faz parte do repositório.

## O que faz

| Etapa | Script | Resultado |
|---|---|---|
| 1 | `montar_base.py` | Base do mês: quantidades por PRT a partir do Caderno de Apoio (horas dos RDOs) |
| 2 | `aplicar_consolidacao.py` | Consolidação de Horas por PRT (dia a dia, retroativos, adicional noturno) |
| 3 | `aplicar_base.py` | Medição + BM Acumulado de cada PRT (rolagem do mês anterior, itens novos, saldos) |
| 4 | `aplicar_bm.py` | BM de abas (uma aba por item, histórico de medições anteriores) |
| 5 | `conferir_mes.py` | Conferência obrigatória: totais, fórmulas, histórico, casas decimais |
| — | `capa_medicao.py`, `exportar_pdf.py` | Capa (Word) e PDFs de entrega |
| — | `fotos_whatsapp.py`, `aplicar_fotos.py` | Fotos do grupo de obra separadas por PRT/dia e lançadas na aba FOTOS |
| — | `boletins_preparacao.py`, `consolidacao_preparacao.py` | Boletins por equipe de um segundo contrato |

`config_mes.py` é o único arquivo que muda de um mês para o outro. `comum.py` reúne o acesso ao Excel via COM e as
regras compartilhadas. `CLAUDE.md` é o mapa do projeto e das regras de negócio (escrito para o assistente de IA usado no
desenvolvimento).

## Princípios

- **Nada arredondado:** valor cheio na célula; totais são fórmulas puras, sem número digitado.
- **Fiel ao RDO assinado:** só se mede o que o apontamento traz; hora extra só se a coluna de HE do RDO trouxer.
- **Conferência sempre:** cada etapa gera relatório; divergências saem marcadas com `!!`.
- **Fonte somente leitura:** os arquivos de origem nunca são alterados; tudo é gerado numa pasta de saída.

## Requisitos

Windows com Microsoft Excel e Word (automação COM), Python 3.10+ e
`pip install openpyxl pandas pdfplumber pypdf pypdfium2 pillow pywin32`.
