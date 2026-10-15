-- Consultas prontas no data/vulto.db (abra o banco no DB Browser for SQLite ou na extensão SQLite Viewer do VS Code).
-- Os valores estão em texto com vírgula ("559,70"): para somar, use cast(replace(valor_total, ',', '.') as real).

-- 1. O que ainda depende de decisão no faturamento (status vazio, trocas, relançamentos da Batel)
select motivo, count(*) as vendas, round(sum(cast(replace(valor_total, ',', '.') as real)), 2) as valor
from vendas_padronizadas
where entra_no_faturamento = 'Pendente'
group by motivo;

-- 2. Por que cada venda entra ou não no faturamento
select entra_no_faturamento,
       case when instr(motivo, ':') > 0 then substr(motivo, 1, instr(motivo, ':') - 1) else motivo end as motivo,
       count(*) as vendas, round(sum(cast(replace(valor_total, ',', '.') as real)), 2) as valor
from vendas_padronizadas
group by 1, 2
order by 1, 3 desc;

-- 3. Faturamento por mês, loja e canal (só o que entra: Sim)
select mes, loja, canal, count(*) as vendas, round(sum(cast(replace(valor_total, ',', '.') as real)), 2) as faturamento
from vendas_padronizadas
where entra_no_faturamento = 'Sim'
group by mes, loja, canal
order by mes, loja;

-- 4. Uma venda, do original ao padronizado (troque o id)
select o.*, p.* from vendas o join vendas_padronizadas p on p.id_venda = o.id where o.id = 6;

-- 5. Passivo de pontos em 01/10/2026 por categoria
select categoria, count(*) as clientes, sum(saldo_valido_01_10_2026) as pontos, round(sum(valor_r), 2) as reais
from clientes_padronizados
group by categoria
order by pontos desc;

-- 6. Um cliente e o extrato de pontos dele (troque o CPF)
select * from clientes_padronizados where cpf = '013.949.436-79';
select * from extrato_pontos where cpf = '013.949.436-79' order by data;

-- 7. Correções feitas nos pontos (o que foi corrigido ou descartado e por quê)
select situacao, motivo, count(*) as movimentos, sum(pontos_considerado - pontos_original) as efeito
from extrato_pontos
where situacao <> 'Mantido'
group by 1, 2
order by efeito;

-- 8. Estoque em 01/10/2026 por local
select sum(pinheiros) as pinheiros, sum(eldorado) as eldorado, sum(savassi) as savassi, sum(batel) as batel,
       sum(cd_cajamar) as cd, sum(total_lojas_cd) as total
from estoque_01_10;

-- 9. Uma peça em todos os locais (troque o SKU)
select * from estoque_01_10 where sku like 'VLT-CAL-001-OLI%';

-- 10. Registro de decisões ainda em aberto
select numero, o_que_encontramos, o_que_decidimos_fazer from registro_decisoes where o_que_decidimos_fazer like '%A decidir%';
