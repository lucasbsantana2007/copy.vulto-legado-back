# Migração da Vulto — dados padronizados

Tudo o que o time limpou e decidiu sobre os dados da Vulto, carregado no `data/vulto.db` por um comando só.

```bash
pip install -r requirements.txt
python migracao/carregar.py
```

## O que fica no `vulto.db`

As tabelas originais da Vulto (`lojas`, `vendas`, `itens_venda`, `compras`, `pontos_mov`, `estoque`, `entradas_estoque`, `funcionarios`, `config`) **nunca mudam**: são o "sistema da Vulto" com que os nossos números são comparados. O `carregar.py` apaga e recria só as tabelas abaixo.

| Tabela | O que é |
|---|---|
| `vendas_padronizadas` | Uma linha por venda: data em Brasília, loja, canal, cliente, vendedor, valores em reais, status, tipo e se entra no faturamento (Sim / Não / Pendente) com o motivo |
| `itens_venda_padronizados` | Itens das vendas com o preço em reais (o app gravava centavos) |
| `clientes_padronizados` | Um cliente por CPF (texto, com o zero da frente), categoria pela regra nova e saldo de pontos em 01/10/2026 |
| `extrato_pontos` | Cada movimento de pontos, com a situação (mantido, corrigido ou descartado) e o motivo |
| `estoque_01_10` | Estoque em 01/10/2026 por SKU e por local (lojas + CD) |
| `estoque_movimentos_set26` | Vendas e transferências entre a contagem e 01/10 |
| `estoque_savassi_por_modelo` | A Savassi como foi contada (por modelo): o número oficial dela |
| `registro_decisoes` | O registro de decisões (cópia do `registro_decisoes.md`) |

## Pastas

- `migracao/entrada/` — planilhas que o time limpou à mão. **É aqui que se edita.**
- `migracao/consultas.sql` — consultas prontas para ver os resultados no banco (pendências, faturamento por mês, passivo, estoque).
- `migracao/registro_decisoes.md` — o registro de decisões, no formato do brief. Toda decisão nova entra aqui (as abas Decisões das planilhas geradas são lidas daqui).
- `migracao/saida/` — o que os scripts geram para conferir no Excel (`clientes atualizado v2.xlsx`, `estoque_vulto_01-10-2026.xlsx`, `vendas_padronizadas.csv`, o log de cada valor alterado e o gráfico das faixas de valor de compra). Não edite aqui: é sobrescrito a cada rodada.

## Editando em grupo

O `vulto.db` é um arquivo binário: o git não junta as edições de duas pessoas nele. Por isso ninguém edita o banco direto:

1. Edite uma planilha de `migracao/entrada/`, o `registro_decisoes.md` ou um script.
2. Rode `python migracao/carregar.py`.
3. Faça o commit (a planilha ou o script que mudou + o `data/vulto.db` + `migracao/saida/`).

Se o `git pull` der conflito no `vulto.db` ou numa planilha de `saida/`, fique com qualquer um dos lados e rode o `carregar.py` de novo: o banco é sempre recalculado a partir das entradas. Conflito numa planilha de `entrada/` é edição de verdade dos dois lados: combinem qual versão fica.
