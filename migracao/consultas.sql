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

-- 11. Estoque de uma loja numa data qualquer (troque o local e a data): "quanto tinha na loja X há duas semanas"
select local, sum(qtd) as pecas
from estoque_historico
where local = 'Pinheiros' and data <= '2026-09-17'
group by local;

-- 12. Histórico de uma peça numa loja, movimento a movimento (troque o SKU e o local)
select data, tipo, qtd, saldo, origem from estoque_historico
where sku = 'VLT-CAM-004-BRA-M' and local = 'CD Cajamar'
order by rowid;

-- 13. O que o sistema não registrou: ajuste de inventário por local
select local, sum(qtd) as pecas, count(*) as skus from estoque_historico
where tipo like 'Ajuste de invent%'
group by local;

-- 14. Uma venda com os itens e os produtos (troque o id; as vendas da Batel feitas sem sistema são BATEL-<Nº>)
select v.id_venda, v.data, v.loja, v.valor_total, i.sku, i.sku_original, p.produto, p.cor, p.tamanho, i.qtd, i.preco_unit
from vendas_padronizadas v
join itens_venda_padronizados i on i.venda_id = v.id_venda
left join produtos_padronizados p on p.sku = i.sku
where v.id_venda = 'BATEL-2';

-- 15. Atendimentos de um cliente, com a venda ligada (troque o CPF)
select a.data, a.canal, a.assunto, a.status, a.pedido, a.id_venda, a.venda_confere
from atendimentos_padronizados a
where a.cliente_cpf = '191.281.999-60'
order by substr(a.data, 7, 4) || substr(a.data, 4, 2) || substr(a.data, 1, 2);

-- 10. Registro de decisões ainda em aberto
select numero, o_que_encontramos, o_que_decidimos_fazer from registro_decisoes where o_que_decidimos_fazer like '%A decidir%';
