# Migração da Vulto — dados padronizados

Tudo o que o time limpou e decidiu sobre os dados da Vulto, carregado no `data/vulto.db` por um comando só.

## Os três números (jan/2025 a set/2026 e posição em 01/10/2026)

| Número | Sistema da Vulto | Nosso número | De onde vem a diferença |
|---|---|---|---|
| Faturamento jan/25–set/26 | R$ 94.580.496,26 (`/relatorio/faturamento`) | **R$ 6.498.071,30** (frete incluído, decisão 4) | Centavos do app (−R$ 87,5 mi), cancelados e não pagos, testes, duplicados, trocas, Batel (decisões 1 e 9 a 17) |
| Estoque em 01/10/2026 | 46.302 peças (tabela `estoque`) | **15.216 peças** (Savassi oficial por modelo: 2.105) | Lojas 5 e 6 copiam o CD, transferências nunca baixadas, saldos negativos (decisões 27 a 37) |
| Passivo de pontos em 01/10/2026 | 8.829.352 pontos | **3.373.251 pontos = R$ 168.662,55** | Créditos ×100 do app, compras canceladas, trocas, resgates anotados só na planilha, vencimento de 12 meses (decisões 18 a 26 e 47 a 50). Em aberto: dobro online (+R$ 31.581,85) e pontos da Batel (+R$ 543,55), decisões 40 e 41 |

Os detalhes de cada número estão nas consultas 3, 5, 8 e 11 do `consultas.sql`.

```bash
pip install -r requirements.txt
python migracao/carregar.py
```

## Os dois bancos

- **`data/vulto.db`**: as tabelas originais da Vulto, intactas, mais as tabelas padronizadas (abaixo). É onde se
  compara "sistema da Vulto × nosso número".
- **`data/vulto_corrigido.db`**: o banco da Vulto **consertado**, com as mesmas 10 tabelas e colunas do original e os
  dados corrigidos pelas decisões do time (decisão 46). Os três números saem direto dele, e o `banco_corrigido.py`
  confere isso a cada rodada. É o banco que vai para a migração.

Os dois são recriados pelo `carregar.py`: ninguém edita à mão.

## O que fica no `vulto.db`

As tabelas originais da Vulto (`lojas`, `vendas`, `itens_venda`, `compras`, `pontos_mov`, `estoque`, `entradas_estoque`, `funcionarios`, `config`) **nunca mudam**: são o "sistema da Vulto" com que os nossos números são comparados. O `carregar.py` apaga e recria só as tabelas abaixo.

| Tabela | O que é |
|---|---|
| `vendas_padronizadas` | Uma linha por venda: data em Brasília, loja, canal, cliente, vendedor, valores em reais, status, tipo e se entra no faturamento (Sim / Não) com o motivo. Inclui as vendas da Batel feitas sem sistema (`BATEL-<Nº>`) |
| `itens_venda_padronizados` | Itens das vendas com o preço em reais (o app gravava centavos) e o SKU do catálogo (o código como veio fica em `sku_original`) |
| `clientes_padronizados` | Um cliente por CPF (texto, com o zero da frente), categoria pela regra nova e saldo de pontos em 01/10/2026 |
| `extrato_pontos` | Cada movimento de pontos, com a situação (mantido, corrigido ou descartado) e o motivo |
| `estoque_01_10` | Estoque em 01/10/2026 por SKU e por local (lojas + CD) |
| `estoque_movimentos_set26` | Vendas e transferências entre a contagem e 01/10 |
| `estoque_savassi_por_modelo` | A Savassi como foi contada (por modelo): o número oficial dela |
| `estoque_historico` | O histórico do estoque, um movimento por linha com o saldo, de 01/01/2025 a 01/10/2026: a posição em qualquer data é a soma até ela (decisões 44 e 45) |
| `produtos_padronizados` | O catálogo: um produto por SKU, com preço, custo, EAN, fornecedor, coleção e status (decisões 38 e 39) |
| `atendimentos_padronizados` | Os atendimentos de WhatsApp, Instagram e e-mail, ligados ao cliente e à venda (decisões 42 e 43) |
| `registro_decisoes` | O registro de decisões (cópia do `registro_decisoes.md`) |

## Conferência

O `carregar.py` termina com o `validar.py`, que confere as tabelas antes de gravar: tabelas originais intactas, SKU de
todo item e de todo estoque no catálogo, histórico do estoque fechando com o número oficial, nenhuma venda de teste ou
não paga no faturamento, CPF único, atendimento ligado só a venda e cliente que existem. Se algo falhar, nada é gravado.
Para conferir o banco como está: `python migracao/validar.py`.

## Pastas

- `migracao/entrada/` — planilhas que o time limpou à mão. **É aqui que se edita.**
  A `clientes_fidelidade_FINAL_v3 (1).xlsx` é a planilha do Vulto+ limpa em 19/10/2026: um cliente por CPF, a aba
  Resgates com o CPF de cada resgate identificado, a aba Pendências (o que falta confirmar com as lojas) e o
  "Relatório de Mudanças". É a fonte dos clientes e dos pontos do `clientes_pontos.py`. Para resolver um resgate
  pendente, digite o CPF na coluna B da aba Resgates e rode o `carregar.py`. O `varredura_resgates.py` refaz a
  busca dos resgates anotados só com o primeiro nome (venda do dia na mesma loja, com o mesmo desconto) e grava o
  resultado em `saida/varredura_resgates.csv` (decisão 48).
  O `clientes atualizado.xlsx` nasceu do botão "Exportar CSV" do sistema da Vulto (`GET /exportar`, rodando o
  sistema localmente), que lê a planilha `clientes_fidelidade_FINAL_v3 (1).xlsx`; o time limpou o CSV à mão. Ao abrir
  esse CSV, o Excel cortou o zero da frente de 80 CPFs (decisão 18). Desde então o `/exportar` sai com o CPF como texto
  (000.000.000-00) e em UTF-8: um export novo pode ser aberto no Excel sem perder o zero. Hoje ele só serve para
  marcar quais clientes tinham ficado de fora dessa limpeza.
- `migracao/consultas.sql` — consultas prontas para ver os resultados no banco (pendências, faturamento por mês, passivo, estoque).
- `migracao/registro_decisoes.md` — o registro de decisões, no formato do brief. Toda decisão nova entra aqui (as abas Decisões das planilhas geradas são lidas daqui).
- `migracao/padronizacao/` — os scripts do Arthur que limpam as planilhas e o vulto.db (com o histórico dos commits).
  O `carregar.py` usa o leitor de códigos (`comum.py`) e a padronização do atendimento (`atendimento.py`); o resto da
  pasta é uma versão paralela, com regras que o time trocou em 15/10 e 19/10 (frete, categoria, tamanho no SKU): fica
  como referência e **não** é a fonte dos números oficiais.
- `migracao/saida/` — o que os scripts geram para conferir no Excel (`clientes atualizado v2.xlsx`, `estoque_vulto_01-10-2026.xlsx`, `vendas_padronizadas.csv`, o log de cada valor alterado e o gráfico das faixas de valor de compra). Não edite aqui: é sobrescrito a cada rodada.

## Editando em grupo

O `vulto.db` é um arquivo binário: o git não junta as edições de duas pessoas nele. Por isso ninguém edita o banco direto:

1. Edite uma planilha de `migracao/entrada/`, o `registro_decisoes.md` ou um script.
2. Rode `python migracao/carregar.py`.
3. Faça o commit (a planilha ou o script que mudou + o `data/vulto.db` + `migracao/saida/`).

Se o `git pull` der conflito no `vulto.db` ou numa planilha de `saida/`, fique com qualquer um dos lados e rode o `carregar.py` de novo: o banco é sempre recalculado a partir das entradas. Conflito numa planilha de `entrada/` é edição de verdade dos dois lados: combinem qual versão fica.
