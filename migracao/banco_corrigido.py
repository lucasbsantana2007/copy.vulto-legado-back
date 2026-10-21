"""data/vulto_corrigido.db: o vulto.db com as MESMAS tabelas e colunas do original, com os dados corrigidos pelas
decisões do time (decisão 46).

O data/vulto.db continua com as tabelas originais intactas (o "sistema da Vulto" da comparação) e as tabelas
padronizadas ao lado. Este arquivo é a versão consertada do banco da Vulto, tabela por tabela, montada a partir das
tabelas padronizadas: quem abrir o vulto_corrigido.db vê o banco da Vulto como ele deveria estar.

O que muda em cada tabela
- vendas: valores em reais com vírgula ("559,70"), frete separado, data em Brasília (AAAA-MM-DD HH:MM:SS), canal pela
  loja (Loja física / Site / App), status padronizado (Concluída, Cancelada, Estornada, Chargeback, Aguardando
  pagamento, Sem status), cliente como CPF 000.000.000-00 (ou o e-mail/telefone normalizado), vendedor pelo cadastro.
  Saem os pedidos de teste, as gravações duplicadas e os relançamentos da Batel que casam com a planilha; entram as
  vendas da planilha manual da Batel (ids novos, depois do último id da Vulto).
- itens_venda: SKU do catálogo, preço em reais; os das vendas que saíram também saem; os da planilha da Batel entram.
- compras: valor em reais (o app gravava centavos), canal pela loja, data em Brasília, cliente normalizado.
- pontos_mov: o extrato corrigido (créditos ×100 divididos, créditos de vendas que não valem descartados, saldo da
  planilha, resgates anotados só no papel, vencimento e saldo que não fica negativo). A soma por CPF (C + AJ + E − D) é
  o saldo válido em 01/10/2026.
- estoque: a posição oficial em 01/10/2026 (contagem física levada até a virada); as lojas 5 e 6 deixam de ter estoque.
- entradas_estoque: SKU do catálogo, data AAAA-MM-DD, custo com vírgula, NF repetida conta uma vez e entram os
  recebimentos que a loja não anotou. O envio por transferência (sem quantidade no original) ganha a quantidade
  escrita no texto, negativa: é saída de quem mandou.
- funcionarios: uma conta por pessoa, com a matrícula nova em codigo (VLT-0101), o papel da plataforma em perfil
  (Lojista, Operador, Administrador), o cargo fora do nome e senha vazia (nenhuma senha antiga migra). Não migram a
  conta de teste e o admin desativado; as contas de função (gerente da Batel, CD, estagiário) ficam desativadas até
  a Vulto dizer quem as usa.
- config: valor_ponto = 0.05.  lojas: sem mudança.

Uso: python migracao/banco_corrigido.py   (depois do carregar.py; o carregar.py já chama no final)
"""

import os
import re
import sqlite3
import sys
from collections import defaultdict
from datetime import datetime
from decimal import Decimal

import pandas as pd

import batel
import caminhos
import produtos
from estoque_01_10 import quantidade_no_texto
from padroniza_vendas import cliente, data_local, reais

DESTINO = os.path.join(caminhos.DADOS, "vulto_corrigido.db")
ORIGINAIS = ["lojas", "compras", "pontos_mov", "config", "vendas", "itens_venda", "estoque", "funcionarios",
             "entradas_estoque"]
CANAL_POR_LOJA = {1: "Loja física", 2: "Loja física", 3: "Loja física", 4: "Loja física", 5: "Site", 6: "App"}
CORTE = "2026-10-01 00:00:00"
INICIO_VALIDADE = "2025-10-01 00:00:00"
VENCIMENTO = "2026-09-30 23:59:59"
LOJA_DO_LOCAL = {"Pinheiros": 1, "Eldorado": 2, "Savassi": 3, "Batel": 4, "CD Cajamar": 7}
TIPO_PONTO = {"Crédito": "C", "Resgate": "D", "Estorno": "E", "Ajuste": "AJ", "Saldo da planilha": "AJ",
              "Resgate (só na planilha)": "D"}
# Decisões 56 a 59: matrícula nova por pessoa. A chave é o id do vulto.db (o crachá 0102 era de duas pessoas); o nome
# só aparece quando muda (as vendas continuam ligadas pelo nome). None = não migra (decisão 58).
FUNCIONARIOS = {
    1: ("VLT-0101", None), 2: ("VLT-0102", None), 3: ("VLT-0104", None), 4: ("VLT-0103", None),
    5: ("VLT-0100", None), 6: ("VLT-0201", None), 7: ("VLT-0202", None), 8: ("VLT-0203", None),
    9: ("VLT-0200", None), 10: ("VLT-0301", None), 11: ("VLT-0302", None), 12: ("VLT-0303", None),
    13: ("VLT-0300", None), 14: ("VLT-0401", None), 15: ("VLT-0402", None), 16: ("VLT-0403", None),
    17: ("VLT-0400", "Gerente Batel"), 18: ("VLT-0701", "Operador CD Cajamar"), 19: ("VLT-0001", "Diego"),
    20: ("VLT-0002", "Paula"), 21: ("VLT-0003", "Marina"), 22: None, 23: ("VLT-0105", "Estagiário Pinheiros"), 24: None,
}
CONTA_DE_FUNCAO = {17, 18, 23}  # desativada até a Vulto informar o nome de quem usa (decisão 58)
PAPEL = {"vendedor": "Lojista", "gerente": "Lojista", "operador": "Operador", "admin": "Administrador"}  # decisão 57


def valor(texto):
    return Decimal(str(texto).replace(",", ".")) if texto not in (None, "") else Decimal(0)


def ref_cliente(ref):
    """cliente_ref como o time padronizou: CPF 000.000.000-00, e-mail em minúsculas, telefone só com dígitos."""
    tipo, cpf, _, contato = cliente(ref)
    return cpf or contato or None


# ---------------------------------------------------------------- vendas e itens


def tabela_vendas(banco, padrao):
    v = pd.read_sql("select * from vendas_padronizadas", banco)
    original = pd.read_sql("select id, criado_em, compra_id from vendas", banco).set_index("id")
    valores = pd.read_sql("select id, total, desconto, frete from vendas", padrao).set_index("id")
    sai = (v.tipo == "Teste") | (v.duplicada_de.astype(str) != "") | v.motivo.str.startswith("Relançamento da venda manual")
    removidas = set(v.loc[sai, "id_venda"].astype(str))
    v = v[~sai]
    planilha, _ = batel.ler()
    momento_batel = {f"BATEL-{r.numero}": r.momento for r in planilha.itertuples()}
    proximo = int(original.index.max()) + 1
    batel_ids = sorted((i for i in v.id_venda.astype(str) if i.startswith("BATEL-")), key=lambda i: int(i.split("-")[1]))
    novo_id = {i: proximo + n for n, i in enumerate(batel_ids)}
    linhas = []
    for r in v.itertuples():
        vid = str(r.id_venda)
        if vid in novo_id:
            momento, total, desconto, frete, compra = momento_batel[vid], r.valor_total, None, None, None
        else:
            o = original.loc[int(vid)]
            momento = data_local(o.criado_em)[0]
            total, desconto, frete = valores.loc[int(vid), ["total", "desconto", "frete"]]
            compra = None if pd.isna(o.compra_id) else int(o.compra_id)
        contato = r.cliente_cpf or r.cliente_contato or None
        linhas.append(dict(id=novo_id.get(vid, int(vid) if vid.isdigit() else vid), pedido_externo=r.pedido_externo or None,
                           canal=r.canal, loja_id=int(r.loja_id), cliente_ref=contato, vendedor=r.vendedor or None,
                           total=total, desconto=desconto, frete=frete, status=r.status,
                           criado_em=momento.strftime("%Y-%m-%d %H:%M:%S"), obs=r.obs or None, compra_id=compra))
    return pd.DataFrame(linhas).sort_values("id"), removidas, novo_id


def tabela_itens(banco, removidas, novo_id):
    i = pd.read_sql("select id, venda_id, sku, descricao, qtd, preco_unit from itens_venda_padronizados", banco)
    i = i[~i.venda_id.astype(str).isin(removidas)].copy()
    i["venda_id"] = i.venda_id.astype(str).map(lambda x: novo_id.get(x, int(x) if x.isdigit() else x))
    return i.sort_values("id")


def tabela_compras(banco, vendas, removidas, pontos):
    """Valor em reais, canal pela loja, data em Brasília. Sai a compra que só existia por um pedido de teste ou
    duplicado (os pontos dela já foram descartados)."""
    c = pd.read_sql("select * from compras", banco)
    v = pd.read_sql("select id, compra_id, canal, total from vendas", banco)
    ligada = v.dropna(subset=["compra_id"]).groupby("compra_id").id.apply(lambda s: {str(x) for x in s})
    usadas = set(vendas.compra_id.dropna().astype(int)) | set(pontos.compra_id.dropna().astype(int))
    sai = {cid for cid, vs in ligada.items() if vs <= removidas and int(cid) not in usadas}
    centavos = {int(cid) for cid, canal, total in zip(v.compra_id, v.canal, v.total)
                if not pd.isna(cid) and str(canal).strip().lower() in ("app", "online") and re.fullmatch(r"-?\d+", str(total).strip())}
    c = c[~c.id.isin(sai)].copy()
    c["valor"] = [reais(valor(x) / (100 if i in centavos or (o == "app" and re.fullmatch(r"-?\d+", str(x).strip())) else 1))
                  for i, x, o in zip(c.id, c.valor, c.origem)]
    c["canal"] = c.loja_id.map(CANAL_POR_LOJA)
    c["criado_em"] = c.criado_em.map(lambda s: data_local(re.sub(r"\.\d+", "", s))[0].strftime("%Y-%m-%d %H:%M:%S"))
    c["cliente_ref"] = c.cliente_ref.map(ref_cliente)
    return c, len(sai)


# ---------------------------------------------------------------- pontos


def tabela_pontos(banco):
    """Extrato do time no formato do pontos_mov, com os lançamentos que fecham o saldo válido em 01/10/2026."""
    ex = pd.read_sql("select * from extrato_pontos", banco)
    obs_original = dict(banco.execute("select id, obs from pontos_mov").fetchall())
    ex = ex[(ex.situacao != "Descartado") & (pd.to_numeric(ex.pontos_considerado) != 0)].copy()
    ex["pontos"] = pd.to_numeric(ex.pontos_considerado).astype(int)
    linhas, novas = [], []
    por_cpf = defaultdict(list)
    for r in ex.itertuples():
        tipo = TIPO_PONTO[r.tipo]
        qtd = -r.pontos if tipo == "D" else r.pontos
        m = re.search(r"compra (\d+)", str(r.compra_venda or ""))
        compra = int(m.group(1)) if m else None
        m_id = re.match(r"vulto\.db:pontos_mov id (\d+)", str(r.origem))
        if m_id:
            linhas.append(dict(id=int(m_id.group(1)), cpf=r.cpf, tipo=tipo, qtd=qtd, compra_id=compra, criado_em=r.data,
                               obs=obs_original.get(int(m_id.group(1)))))
        elif r.tipo == "Saldo da planilha":
            novas.append(dict(cpf=r.cpf, tipo="AJ", qtd=qtd, compra_id=None, criado_em=r.data,
                              obs="saldo da planilha Vulto+ congelado em 01/01/2025 (decisão 19)"))
        else:
            novas.append(dict(cpf=r.cpf, tipo=tipo, qtd=qtd, compra_id=None, criado_em=r.data,
                              obs=f"resgate anotado só no papel: {r.origem} (decisão 25)"))
        por_cpf[r.cpf].append((r.data, r.pontos))
    # mesma conta do clientes_pontos.saldo_em: resgates gastam os pontos mais antigos; o saldo não fica negativo;
    # no corte vencem os pontos ganhos antes de 01/10/2025
    for cpf, movs in por_cpf.items():
        lotes = []
        for data, pts in sorted(movs, key=lambda m: (m[0], m[1] < 0)):
            if pts > 0:
                lotes.append([data, pts])
                continue
            gastar = -pts
            for lote in lotes:
                usa = min(gastar, lote[1])
                lote[1] -= usa
                gastar -= usa
                if not gastar:
                    break
            if gastar:
                novas.append(dict(cpf=cpf, tipo="AJ", qtd=gastar, compra_id=None, criado_em=data,
                                  obs="resgate maior que o saldo: o saldo para em zero (regra do Vulto+)"))
        vencidos = sum(p for d, p in lotes if d < INICIO_VALIDADE)
        if vencidos:
            novas.append(dict(cpf=cpf, tipo="AJ", qtd=-vencidos, compra_id=None, criado_em=VENCIMENTO,
                              obs="pontos vencidos em 01/10/2026: ganhos antes de 01/10/2025 (decisão 22)"))
    proximo = int(banco.execute("select max(id) from pontos_mov").fetchone()[0]) + 1
    novas = pd.DataFrame(novas).sort_values(["criado_em", "cpf"])
    novas.insert(0, "id", range(proximo, proximo + len(novas)))
    return pd.concat([pd.DataFrame(linhas), novas], ignore_index=True).sort_values("id")


# ---------------------------------------------------------------- estoque


def tabela_funcionarios(banco):
    """Uma conta por pessoa com a matrícula nova e o papel da plataforma; nenhuma senha antiga migra (decisões 56 a 59)."""
    f = pd.read_sql("select * from funcionarios", banco)
    if set(f.id) != set(FUNCIONARIOS):
        sys.exit(f"funcionarios: ids sem matrícula definida em FUNCIONARIOS: {sorted(set(f.id) ^ set(FUNCIONARIOS))}")
    linhas = []
    for r in f.itertuples():
        if FUNCIONARIOS[r.id] is None:
            continue
        matricula, nome = FUNCIONARIOS[r.id]
        linhas.append(dict(id=r.id, codigo=matricula, nome=nome or r.nome, loja_id=r.loja_id, senha=None,
                           perfil=PAPEL[r.perfil], ativo=0 if r.id in CONTA_DE_FUNCAO else r.ativo))
    return pd.DataFrame(linhas)


def tabela_estoque(banco):
    e = pd.read_sql("select * from estoque_01_10", banco)
    linhas = [dict(sku=r["sku"], loja_id=loja, qtd=int(r[col]), atualizado_em=CORTE)
              for _, r in e.iterrows() for col, loja in (("pinheiros", 1), ("eldorado", 2), ("savassi", 3),
                                                         ("batel", 4), ("cd_cajamar", 7)) if int(r[col]) > 0]
    return pd.DataFrame(linhas).sort_values(["loja_id", "sku"])


def tabela_entradas(banco):
    resolver = produtos.resolvedor()
    e = pd.read_sql("select * from entradas_estoque order by id", banco)
    e["sku"] = [resolver(s) or s for s in e.sku]
    dia = lambda s: next((datetime.strptime(str(s).strip(), f).strftime("%Y-%m-%d")       # noqa: E731
                          for f in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y") if _bate(s, f)), None)
    e["data"] = e.data.map(dia)
    # vazio no original continua vazio (o pandas lê como NaN, que viraria o texto "NaN")
    e["custo_unit"] = e.custo_unit.map(lambda x: reais(valor(x)) if pd.notna(x) and str(x).strip() else None)
    receb = e.qtd.notna() & e.fornecedor.fillna("").str.strip().str.upper().isin(["", "CD"]) & (e.loja_id != 7)
    e.loc[receb, "fornecedor"] = e.loc[receb, "obs"].map(
        lambda o: next((n.title() for n in ("pinheiros", "eldorado", "savassi", "batel") if f"veio da {n}" in str(o).lower()), "CD"))
    compra = (e.loja_id == 7) & e.usuario.isin(["marina.compras", "joao.fiscal"]) & e.nf.notna()
    repetida = compra & e.duplicated(subset=["sku", "qtd", "nf"], keep="first") & compra.where(compra, False)
    e = e[~repetida].copy()
    # decisão 30: o envio (do CD ou de uma loja) não tem quantidade, ela está no texto ("2 un -> savassi"); sem número,
    # 1 peça. O envio é saída de quem mandou, então a quantidade entra negativa, como a devolução nas vendas
    envio = e.qtd.isna()
    e.loc[envio, "qtd"] = [-(quantidade_no_texto(str(o or "")) or 1) for o in e.loc[envio, "obs"]]
    h = pd.read_sql("select * from estoque_historico where tipo like 'Recebido%' and (tipo like '%só o envio%' or tipo like '%só o CD%')",
                    banco)
    proximo = int(e.id.max()) + 1
    criados = pd.DataFrame(dict(
        id=range(proximo, proximo + len(h)), sku=h.sku.values, loja_id=h.local.map(LOJA_DO_LOCAL).values, qtd=h.qtd.values,
        custo_unit=None, nf=None, fornecedor=h.tipo.map(lambda t: "CD" if "CD" in t else t.split(" de ")[1].split(" (")[0]).values,
        data=h.data.values, obs=[f"recebimento não anotado pela loja ({o}) (decisão 29)" for o in h.origem],
        usuario="migracao"))
    return pd.concat([e, criados], ignore_index=True), int(repetida.sum()), len(criados)


def _bate(texto, formato):
    try:
        datetime.strptime(str(texto).strip(), formato)
        return True
    except ValueError:
        return False


# ---------------------------------------------------------------- principal


def main():
    banco = sqlite3.connect(caminhos.VULTO_DB)
    padrao = sqlite3.connect(caminhos.PADRONIZADO)
    vendas, removidas, novo_id = tabela_vendas(banco, padrao)
    itens = tabela_itens(banco, removidas, novo_id)
    pontos = tabela_pontos(banco)
    compras, compras_fora = tabela_compras(banco, vendas, removidas, pontos)
    estoque = tabela_estoque(banco)
    entradas, nf_repetidas, recebimentos = tabela_entradas(banco)
    config = pd.read_sql("select * from config", banco)
    config.loc[config.chave == "valor_ponto", "valor"] = "0.05"                     # 100 pontos = R$ 5,00
    funcionarios = tabela_funcionarios(banco)
    novas = {"vendas": vendas, "itens_venda": itens, "compras": compras, "pontos_mov": pontos, "estoque": estoque,
             "entradas_estoque": entradas, "config": config, "funcionarios": funcionarios}

    if os.path.exists(DESTINO):
        os.remove(DESTINO)
    destino = sqlite3.connect(DESTINO)
    for nome, sql in banco.execute(f"select name, sql from sqlite_master where type = 'table' and name in "
                                   f"({', '.join('?' * len(ORIGINAIS))})", ORIGINAIS).fetchall():
        destino.execute(sql)                                                       # mesma estrutura do original
    linhas = banco.execute("select * from lojas").fetchall()
    destino.executemany(f"insert into lojas values ({', '.join('?' * len(linhas[0]))})", linhas)
    for nome, df in novas.items():
        colunas = [c[1] for c in destino.execute(f"pragma table_info({nome})")]
        destino.executemany(f"insert into {nome} ({', '.join(colunas)}) values ({', '.join('?' * len(colunas))})",
                            [tuple(None if (x is None or (isinstance(x, float) and x != x)) else
                                   (int(x) if c in ("id", "loja_id", "qtd", "compra_id", "venda_id") and str(x).lstrip("-").isdigit()
                                    else (x.item() if hasattr(x, "item") else x))
                                   for x, c in zip(linha, colunas)) for linha in df[colunas].itertuples(index=False)])
    for (t,) in destino.execute("select name from sqlite_sequence").fetchall():
        destino.execute(f"update sqlite_sequence set seq = (select max(id) from {t}) where name = ?", (t,))
    for t in ("vendas", "itens_venda", "compras", "pontos_mov", "entradas_estoque", "funcionarios"):
        if not destino.execute("select 1 from sqlite_sequence where name = ?", (t,)).fetchone():
            destino.execute("insert into sqlite_sequence (name, seq) select ?, max(id) from " + t, (t,))
    destino.commit()
    conferir(banco, destino)
    destino.execute("vacuum")
    destino.close()
    print(f"vulto_corrigido.db: vendas {len(vendas)} (−{len(removidas)} de teste/duplicadas/relançadas, "
          f"+{len(novo_id)} da Batel) · itens {len(itens)} · compras {len(compras)} (−{compras_fora}) · pontos_mov {len(pontos)} · "
          f"estoque {len(estoque)} · entradas {len(entradas)} (−{nf_repetidas} NF repetida, +{recebimentos} recebimentos)")
    print(f"gravado: {DESTINO}")


def conferir(banco, destino):
    """O banco corrigido sozinho tem que dar os números oficiais, e a estrutura tem que ser a do original."""
    erros = []
    estrutura = lambda c: {s for (s,) in c.execute("select sql from sqlite_master where type = 'table' and name in "  # noqa: E731
                                                  f"({', '.join('?' * len(ORIGINAIS))})", ORIGINAIS)}
    if estrutura(banco) != estrutura(destino):
        erros.append("estrutura diferente da do vulto.db original")
    fat = destino.execute("""select round(sum(cast(replace(total, ',', '.') as real)), 2) from vendas
        where status in ('Concluída', 'Sem status') and not (lower(coalesce(obs, '')) like 'troca%' and cast(replace(total, ',', '.') as real) >= 0)
          and criado_em >= '2025-01-01' and criado_em < '2026-10-01'""").fetchone()[0]
    oficial = banco.execute("""select round(sum(cast(replace(valor_total, ',', '.') as real)), 2) from vendas_padronizadas
        where entra_no_faturamento = 'Sim' and mes between '2025-01' and '2026-09'""").fetchone()[0]
    if abs(fat - oficial) > 0.01:
        erros.append(f"faturamento {fat} ≠ oficial {oficial}")
    pecas = destino.execute("select sum(qtd) from estoque").fetchone()[0]
    oficial_est = banco.execute("select sum(total_lojas_cd) from estoque_01_10").fetchone()[0]
    if pecas != oficial_est:
        erros.append(f"estoque {pecas} ≠ oficial {oficial_est}")
    saldo = dict(destino.execute("select cpf, sum(case when tipo = 'D' then -qtd else qtd end) from pontos_mov group by cpf"))
    valido = dict(banco.execute("select cpf, saldo_valido_01_10_2026 from clientes_padronizados"))
    dif = [c for c in set(saldo) | set(valido) if int(saldo.get(c, 0)) != int(valido.get(c, 0))]
    if dif:
        erros.append(f"pontos_mov não fecha com o saldo válido em {len(dif)} CPFs (ex.: {dif[:3]})")
    for (tabela,) in destino.execute(f"select name from sqlite_master where type = 'table' and name in ({', '.join('?' * len(ORIGINAIS))})",
                                     ORIGINAIS).fetchall():
        for coluna in [c[1] for c in destino.execute(f"pragma table_info({tabela})")]:
            n = destino.execute(f"select count(*) from {tabela} where lower({coluna}) in ('nan', 'none', 'null')").fetchone()[0]
            if n:
                erros.append(f"{tabela}.{coluna}: {n} valores vazios gravados como texto ('NaN', 'None')")
    sem_qtd = destino.execute("select count(*) from entradas_estoque where qtd is null or qtd = 0").fetchone()[0]
    if sem_qtd:
        erros.append(f"entradas_estoque: {sem_qtd} linhas sem quantidade")
    orfaos = destino.execute("select count(*) from itens_venda where venda_id not in (select id from vendas)").fetchone()[0]
    if orfaos:
        erros.append(f"{orfaos} itens sem venda")
    catalogo = {s for (s,) in banco.execute("select sku from produtos_padronizados")}
    fora = {s for (s,) in destino.execute("select distinct sku from itens_venda union select distinct sku from estoque")} - catalogo
    if fora:
        erros.append(f"SKU fora do catálogo: {sorted(fora)[:5]}")
    repetidas = destino.execute("select count(*) from (select codigo from funcionarios group by codigo having count(*) > 1)").fetchone()[0]
    if repetidas or destino.execute("select count(*) from funcionarios where senha is not null").fetchone()[0]:
        erros.append("funcionarios: matrícula repetida ou senha antiga migrada")
    nomes = {n for (n,) in destino.execute("select nome from funcionarios")} | {"Rafael Lima ou Rafaela Couto"}  # decisão 13
    sem_cadastro = {v for (v,) in destino.execute("select distinct vendedor from vendas where vendedor is not null")} - nomes
    if sem_cadastro:
        erros.append(f"vendedor sem cadastro em funcionarios: {sorted(sem_cadastro)[:5]}")
    if erros:
        sys.exit("vulto_corrigido.db: " + "; ".join(erros))
    print(f"  conferência ok: faturamento R$ {fat:,.2f}, estoque {pecas} peças e passivo {sum(saldo.values())} pontos "
          "saem direto do vulto_corrigido.db; estrutura igual à do original")


if __name__ == "__main__":
    main()
