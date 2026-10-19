# Padronização dos dados da Vulto

Estes scripts aplicam as decisões 1 a 37 do Registro de Decisões da Entrega 3. Os arquivos originais **não são alterados**.

As decisões 38 a 41 ainda estão pendentes e ficaram de fora:
- vendas do app gravadas em centavos;
- vendas canceladas no relatório;
- regra monetária do Vulto+;
- conciliação da Batel com o vulto.db.

## Como rodar

```bash
pip install pandas openpyxl
python rodar_tudo.py ../vulto-legado-back/data saida
python -m pytest
```

Cada domínio também roda sozinho. Exemplo: `python clientes.py ../vulto-legado-back/data`.

## Scripts

| Script | Decisões | Entrada | Saída |
|---|---|---|---|
| `clientes.py` | 1–5 | clientes_fidelidade_FINAL_v3 (1).xlsx | clientes_corrigidos.xlsx |
| `catalogo.py` | 6–13 | catalogo_vulto_2026_REVISADO_final.xlsx | catalogo_corrigido.xlsx |
| `vendas_batel.py` | 14–20 | vendas_batel_MANUAL_marco.xlsx | vendas_batel_corrigidas.xlsx |
| `atendimento.py` | 21–27 | atendimentos_export_whatsapp_insta.csv | atendimento_corrigido.xlsx |
| `inventario.py` | 28–37 | contagem_inventario_set26.xlsx + cadastro mestre do catálogo | inventario_corrigido.xlsx |
| `rodar_tudo.py` | todas | — | os 5 arquivos acima + registro_decisoes_numeros.xlsx + vulto_dadoscorrigidos.db |
| `comum.py` | — | — | regras compartilhadas (CPF, tamanhos, cores, código do CD, EAN, telefone, Excel) |

## Como cada arquivo de saída é organizado

- A primeira aba tem os dados corrigidos. Colunas `*_original` guardam o valor como estava.
- A aba **Ocorrências** tem uma linha por alteração: decisão, tipo, linha de origem, valor antes e depois.
- A aba **Resumo** conta as ocorrências por decisão. Esses números servem para a coluna "Quantos registros afeta".
- O `vulto_dadoscorrigidos.db` tem as mesmas tabelas em SQLite, para o backend importar.

## Parâmetros que o grupo pode ajustar

Cada parâmetro fica no topo do script correspondente:

- `clientes.py`
  - `FAIXAS`: faixas de pontos por categoria.
  - `DOMINIOS`: erros de domínio de e-mail.
- `catalogo.py`
  - `MARGEM_MINIMA`: margem mínima.
  - `DATA_APROVACAO`: data de aprovação dos preços.
  - `REGRAS_FORNECEDOR_NOME`: regras de fornecedor pelo nome da peça. **Conferir com a aba Listas do grupo.**
- `comum.py`
  - `TAMANHO_NUMERICO`: tamanho numérico → letra. Hoje 46 = XG; o PDF também cita GGG.
- `vendas_batel.py`
  - `VENDEDORES` e `PAGAMENTOS`: nomes padronizados de vendedoras e formas de pagamento.
- `atendimento.py`
  - `ASSUNTOS`: as 9 categorias de assunto.
  - `STATUS` e `STATUS_PELA_RESPOSTA`: regras de status.
