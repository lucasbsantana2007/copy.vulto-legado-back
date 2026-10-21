# Padronização dos dados da Vulto (scripts do Arthur)

> **Os números oficiais saem do `migracao/carregar.py`** (ver `migracao/README.md` e `migracao/registro_decisoes.md`).
> Esta pasta guarda os scripts do Arthur, que limpam as planilhas e o `vulto.db`, com o histórico dos commits (a versão
> dele e as correções de 19/10). O `carregar.py` usa daqui só o leitor de códigos (`comum.py`); o atendimento
> passou a vir do export limpo pelo time (`migracao/entrada`). O resto é uma versão paralela: algumas regras foram trocadas pelo time em 15/10 e 19/10
> (frete é receita, categoria com Vip pelo saldo, tamanho em número no SKU de calça e bermuda, vírgula decimal,
> validade de 12 meses para todos), então os números desta pasta não são os oficiais. A numeração das decisões daqui
> também é própria (não é a do `registro_decisoes.md`).

Corrige as planilhas **e o `vulto.db`** antes da migração para a plataforma da Casa Lorenzi, aplica as 69 decisões
do Registro de Decisões e calcula os três números pedidos pela diretoria. Os arquivos originais **não são
alterados**: o relatório antigo da Vulto precisa continuar reproduzível para a comparação "sistema × nosso número".

## Como rodar

```bash
pip install pandas openpyxl pytest
python migracao/padronizacao/rodar_tudo.py          # lê data/ do repositório e grava em migracao/padronizacao/saida
python -m pytest migracao/padronizacao
```

As tabelas originais do `data/vulto.db` nunca mudam (o `carregar.py` só acrescenta tabelas padronizadas), então este
script sempre lê os dados originais. A trava de `rodar_tudo.py` impede rodar em cima de um banco já corrigido.

## O que sai na pasta `saida/`

| Arquivo | Para quê |
|---|---|
| `vulto_corrigido.db` | **O vulto.db consertado**: as mesmas 10 tabelas e colunas do original, só os dados corrigidos. O faturamento e o estoque saem dele com uma consulta simples (o `rodar_tudo.py` confere) |
| `planilhas_e_relatorios.db` | O que nunca existiu no vulto.db: planilhas corrigidas (`planilha_*`, `export_atendimentos`), histórico de estoque, relatórios, Registro de Decisões e o log de cada alteração (`registro_ocorrencias`) |
| `numeros_diretoria.xlsx` | Os três números, cada um com a comparação "sistema da Vulto × nosso número" |
| `registro_decisoes.xlsx` | O Registro de Decisões com as 7 colunas do enunciado (quantidade e impacto calculados) |
| `vendas_corrigidas.xlsx` | Vendas (uma linha por venda, com o motivo de entrar ou não no faturamento) e itens |
| `pontos_vulto_mais.xlsx` | Extrato corrigido do Vulto+, saldo por cliente e passivo |
| `estoque_historico.xlsx` | Movimentações (o histórico), posição em 01/10/2026 e em 17/09/2026, qualidade por loja |
| `clientes_corrigidos.xlsx`, `catalogo_corrigido.xlsx`, `atendimento_corrigido.xlsx`, `inventario_corrigido.xlsx`, `vendas_batel_corrigidas.xlsx` | Cada domínio corrigido |

Em todos os arquivos, as colunas `*_original` guardam o valor como estava e a aba **Ocorrências** tem uma linha
por alteração (decisão, linha de origem, valor antes e depois).

## Ordem do processamento

```
clientes → catálogo → vendas Batel → vendas (vulto.db) → Vulto+ (pontos) → clientes com saldo
         → atendimento (ligado a cliente e pedido) → inventário → estoque com histórico
         → três números + Registro de Decisões + vulto_corrigido.db + planilhas_e_relatorios.db
```

## Scripts

| Script | Decisões | O que faz |
|---|---|---|
| `clientes.py` | 1–5, 64–66, 69 | CPF, duplicados, e-mail (inclusive e-mail de outra pessoa), consentimento (LGPD), clientes que só estavam no sistema |
| `catalogo.py` | 6–13, 68 | Código único por peça, custos (último e médio ponderado), EAN, fornecedor, menor preço entre coleções |
| `vendas_batel.py` | 14–20 | Vendas manuais da Batel (PDV fora do ar em março/26) |
| `vendas.py` | 38, 39, 41–50 | Vendas do vulto.db: centavos do app, status, duplicados, testes, fuso, canal, frete, devoluções, trocas, ship from store, cliente da venda, conciliação da Batel e faturamento |
| `pontos.py` | 4, 40, 51–58 | Extrato do Vulto+ desde o saldo congelado em 01/01/2025, correções, vencimento e passivo |
| `atendimento.py` | 21–27, 67 | Atendimentos padronizados e ligados ao cliente e ao pedido |
| `inventario.py` | 28–37 | Contagem física de set/26 |
| `estoque.py` | 59–63 | Transferências pareadas, histórico reconstruído e posição em 01/10/2026 |
| `banco_corrigido.py` | todas | Gera o `vulto_corrigido.db` (cópia do original com os dados trocados pelos corrigidos) |
| `decisoes.py` | todas | Texto do Registro de Decisões |
| `rodar_tudo.py` | todas | Roda tudo e gera os arquivos |
| `comum.py` | — | Regras compartilhadas: CPF, código da peça, datas com fuso, quem é o cliente, transferência escrita no texto |

## Parâmetros que o grupo pode ajustar

- `pontos.py`: `VALOR_PONTO` (R$ 0,05), `VALIDADE_MESES` / `VALIDADE_MESES_OURO` (12 / 18), `FAIXAS` da categoria.
- `vendas.py`: `RAZAO_CENTAVOS` (20), `STATUS` e `STATUS_FATURA` (o que entra no faturamento).
- `estoque.py`: `DATA_CONTAGEM` por loja e `JANELA_TRANSFERENCIA` (2 dias).
- `catalogo.py`: `MARGEM_MINIMA`, `REGRAS_FORNECEDOR_NOME`.
- `clientes.py`: `DOMINIOS`, `CONSENTIMENTO`, `LOJA_CADASTRO`.
- `comum.py`: `TAMANHO_NUMERICO` (hoje 46 = XG; o PDF também cita GGG).

## Como ficam as tabelas do `vulto_corrigido.db`

| Tabela | O que mudou |
|---|---|
| `vendas` | Valores em reais (centavos do app ÷ 100), frete do app separado, datas no horário de Brasília, canal (`loja`/`site`/`app`) e status (`pago`, `cancelado`, `estornado`, `chargeback`, `aguardando`) padronizados, cliente pelo CPF. Saem pedidos de teste e gravações duplicadas; entram as vendas da Batel que só estavam na planilha |
| `itens_venda` | SKU do catálogo e preço em reais; saem os itens das vendas removidas |
| `compras` | Valor em reais, canal e loja do app corrigidos, cliente pelo CPF |
| `pontos_mov` | Créditos corrigidos e lançamentos que faltavam (saldo da planilha, dobro online, resgates em papel, vencimentos). A soma por CPF (C + AJ + E − D) é o saldo em 01/10/2026 |
| `estoque` | Posição em 01/10/2026 (contagem física + movimentos); as lojas 5 e 6 deixam de ter estoque (as peças estão no CD) |
| `entradas_estoque` | SKU do catálogo, data ISO, custo em número; entram os recebimentos que a loja não anotou |
| `config` | `valor_ponto` = 0.05 |
| `lojas`, `funcionarios` | Sem mudança (as senhas em texto puro não vão para a Lorenzi) |
