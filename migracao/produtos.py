"""Catálogo da Vulto: um produto por SKU, no padrão de SKU do time, e o tradutor de códigos.

Entradas (só leitura):
- novo-catalogo_vulto_2026_REVISADO_final (1).xlsx (abas das coleções, limpas pelo time, e a aba de 2024)
- vulto.db (códigos usados em itens_venda, estoque e entradas_estoque) e vulto_padronizado.db (preço praticado em
  reais; rode padroniza_valores.py antes)

Regras:
- Peça em mais de uma coleção: vale a coleção mais nova (regra 2 da Marina, a mesma da aba Consulta de preço).
  Atemporais só quando a peça não está em nenhuma coleção.
- SKU com o tamanho em número quando o código do CD tem número (calças e bermudas, decisão 35): VLT-CAL-001-OLI-38.
- Peças da lista de 2024 que aparecem em vendas, estoque ou entradas e não estão nas coleções entram no cadastro
  ("Fora do catálogo 2026"). Preço = último preço vendido, porque a aba de 2024 diz "não usar os preços"; custo e
  EAN da aba de 2024 (decisão 39).

O tradutor (Resolvedor) leva qualquer grafia de código ao SKU do catálogo: o SKU (VLT-CAL-001-OLI-38 ou -P), o
código do CD (CAL001V38) e as grafias do fiscal ('vlt cal 001 verde oliva 38'). Usa o leitor de códigos dos
scripts do Arthur (migracao/padronizacao/comum.py).

Saída: tabela produtos no vulto_padronizado.db (o carregar.py grava como produtos_padronizados).

Uso: python migracao/produtos.py
"""

import os
import sqlite3
import sys

import pandas as pd

import caminhos
from estoque_01_10 import canonico

sys.path.insert(0, os.path.join(caminhos.MIGRACAO, "padronizacao"))
from comum import NOME_COR, ler_codigo_produto, tamanho_padrao  # noqa: E402

CATALOGO = os.path.join(caminhos.ENTRADA, "novo-catalogo_vulto_2026_REVISADO_final (1).xlsx")
CONTAGEM_TIME = os.path.join(caminhos.ENTRADA, "contagem_inventario_set26.xlsx")
COLECOES = ["Verão 25", "Inverno 25", "Verão 26", "Inverno 26"]   # da mais antiga para a mais nova
GENERICO = "DIVERSO"                                              # peça não identificada no PDV (decisão 37)
COLUNAS = ["sku", "codigo_cd", "modelo", "produto", "categoria", "cor", "tamanho", "preco", "custo", "ean", "fornecedor",
           "colecao", "status", "origem_preco", "obs"]


def numero(valor):
    try:
        return round(float(str(valor).replace(",", ".")), 2)
    except ValueError:
        return None


def chave(codigo):
    """(modelo, número, cor, tamanho em letra) de qualquer grafia: é o que identifica a peça."""
    p = ler_codigo_produto(codigo)
    if not p or not p[3] or not tamanho_padrao(p[3]):
        return None
    return p[0], p[1], p[2], tamanho_padrao(p[3])


class Resolvedor:
    """Qualquer código -> SKU do catálogo (ou None)."""

    def __init__(self, produtos):
        self.skus = set(produtos.sku)
        self.por_codigo, self.por_chave = {}, {}
        for r in produtos.itertuples():
            self.por_codigo[r.sku.upper()] = r.sku
            if isinstance(r.codigo_cd, str):
                self.por_codigo[r.codigo_cd.upper()] = r.sku
            if chave(r.sku):
                self.por_chave[chave(r.sku)] = r.sku

    def __call__(self, codigo):
        texto = str(codigo or "").strip().upper()
        if texto in self.por_codigo:
            return self.por_codigo[texto]
        return self.por_chave.get(chave(texto))


def ler_colecoes():
    """Uma linha por SKU, da coleção mais nova (Atemporais só se não estiver em coleção nenhuma). Sem código do CD na
    coleção (peças descontinuadas), vale o código que o time gerou pela regra na aba CATÁLOGO da contagem."""
    lista = pd.read_excel(CONTAGEM_TIME, sheet_name="CATÁLOGO", dtype=str)
    codigo_do_time = dict(zip(lista.SKU, lista["CÓDIGO CD"]))
    escolhido = {}
    for aba in ["Atemporais"] + COLECOES:
        df = pd.read_excel(CATALOGO, sheet_name=aba, header=3, dtype=str).dropna(subset=["SKU"])
        for r in df.itertuples(index=False):
            codigo = r[1] if isinstance(r[1], str) and r[1] else codigo_do_time.get(r.SKU)
            sku = canonico(r.SKU, codigo)
            escolhido[sku] = dict(sku=sku, codigo_cd=codigo, modelo=r.Modelo, produto=r.Produto, categoria=r.Categoria,
                                  cor=r.Cor, tamanho=sku.rsplit("-", 1)[1], preco=numero(r[7]), custo=numero(r[8]),
                                  ean=r.EAN, fornecedor=r.Fornecedor, colecao=aba, status=r.Status,
                                  origem_preco=f"Catálogo, aba {aba}", obs=r.Obs if isinstance(r.Obs, str) else "")
    return escolhido


def codigos_em_uso():
    """Chaves das peças que aparecem em vendas, estoque ou entradas do vulto.db."""
    conn = sqlite3.connect(caminhos.VULTO_DB)
    codigos = {s for t in ("itens_venda", "estoque", "entradas_estoque") for (s,) in conn.execute(f"select distinct sku from {t}")}
    conn.close()
    return {chave(c) for c in codigos} - {None}


def ultimo_preco_vendido():
    """Último preço pago de cada peça (vendas pagas, sem pedidos de teste), em reais, por chave."""
    conn = sqlite3.connect(caminhos.PADRONIZADO)
    linhas = conn.execute(
        "select i.sku, i.preco_unit, v.criado_em from itens_venda i join vendas v on v.id = i.venda_id "
        "where i.qtd > 0 and lower(trim(coalesce(v.status, ''))) in ('pago', 'aprovado', '') "
        "and coalesce(v.pedido_externo, '') not like 'S99%' and coalesce(v.pedido_externo, '') not like 'APP-99%' "
        "and lower(coalesce(v.obs, '')) not like '%teste%'"
    ).fetchall()
    conn.close()
    ultimo = {}
    for codigo, preco, criado in linhas:
        quando = criado.replace("T", " ").rstrip("Z")[:19]
        if "/" in quando[:10]:                                         # 01/03/2025 19:43 -> 2025-03-01 19:43
            d, m, a = quando[:10].split("/")
            quando = f"{a}-{m}-{d}{quando[10:]}"
        k, valor = chave(codigo), numero(preco)
        if k and valor and (k not in ultimo or quando > ultimo[k][0]):
            ultimo[k] = (quando, valor)
    return ultimo


def pecas_2024(escolhido):
    """Peças da lista de 2024 em uso que não estão nas coleções."""
    antiga = pd.read_excel(CATALOGO, sheet_name="PLANILHA ANTIGA NAO USAR", header=1, dtype=str).dropna(subset=["SKU"])
    no_catalogo = {chave(s) for s in escolhido}
    em_uso = codigos_em_uso()
    vendido = ultimo_preco_vendido()
    novas = {}
    for r in antiga.itertuples(index=False):
        k = chave(r.SKU)
        if k is None or k in no_catalogo or k not in em_uso:
            continue
        sku = r.SKU.upper()
        sku = sku[:-3] + "-XP" if sku.endswith("-PP") else sku           # PP = XP (decisão do time)
        preco_2024 = numero(r.PRECO)
        if k in vendido:
            quando, preco = vendido[k]
            if preco_2024 and preco > 20 * preco_2024:                     # venda do app ainda em centavos
                preco = round(preco / 100, 2)
            origem = f"Último preço vendido ({quando[8:10]}/{quando[5:7]}/{quando[:4]})"
        else:
            preco, origem = preco_2024, "Aba 2024 (sem venda desde 2025)"
        novas[sku] = dict(sku=sku, codigo_cd=None, modelo=f"{k[0]}-{k[1]}", produto=str(r.DESCRICAO).title(),
                          categoria=str(r.GRUPO).title(), cor=str(r.COR).title(), tamanho=sku.rsplit("-", 1)[1],
                          preco=preco, custo=numero(r.CUSTO), ean=r.EAN, fornecedor="", colecao="2024 (lista antiga)",
                          status="Fora do catálogo 2026", origem_preco=origem,
                          obs="Peça de 2024 que ainda aparece em vendas, estoque ou entradas")
        no_catalogo.add(k)
    return novas


def contados_sem_cadastro(escolhido):
    """Peça contada no inventário com um SKU que não existe em lista nenhuma (o time manteve como 'não identificado').
    Ela existe fisicamente, então entra no cadastro no padrão VLT-CAT-NNN-COR-TAM, com o nome, a categoria e o preço do
    mesmo modelo (decisão 38)."""
    por_modelo = {}
    for p in escolhido.values():
        por_modelo.setdefault(p["modelo"], p)
    novas = {}
    for aba in ("PINHEIROS", "ELDORADO", "SAVASSI", "BATEL", "CD"):
        df = pd.read_excel(CONTAGEM_TIME, sheet_name=aba, dtype=str).dropna(subset=["SKU"])
        for r in df.itertuples(index=False):
            sku = canonico(r.SKU, r[1])
            k = chave(sku)
            if sku in escolhido or sku in novas or k is None:
                continue
            base = por_modelo.get(f"{k[0]}-{k[1]}")
            novas[sku] = dict(sku=sku, codigo_cd=None, modelo=f"{k[0]}-{k[1]}", produto=base["produto"] if base else "",
                              categoria=base["categoria"] if base else "", cor=NOME_COR.get(k[2], k[2]),
                              tamanho=sku.rsplit("-", 1)[1], preco=base["preco"] if base else None,
                              custo=base["custo"] if base else None, ean=None, fornecedor=base["fornecedor"] if base else "",
                              colecao="", status="Criado na migração (contado sem cadastro)",
                              origem_preco=f"Mesmo modelo ({base['sku']})" if base else "",
                              obs=f"Contado na aba {aba} da contagem de set/26 com um código que não existe no catálogo")
    return novas


def generico():
    return {GENERICO: dict(sku=GENERICO, codigo_cd=None, modelo="", produto="Peça não identificada no PDV", categoria="",
                           cor="", tamanho="", preco=None, custo=None, ean=None, fornecedor="", colecao="",
                           status="Só histórico", origem_preco="", obs="Código genérico do PDV (decisão 37): só para o histórico de vendas")}


def resolvedor():
    """Tradutor a partir da tabela produtos já gravada (para os outros scripts)."""
    conn = sqlite3.connect(caminhos.PADRONIZADO)
    df = pd.read_sql("select * from produtos", conn)
    conn.close()
    return Resolvedor(df)


def main():
    escolhido = ler_colecoes()
    escolhido.update(pecas_2024(escolhido))
    escolhido.update(contados_sem_cadastro(escolhido))
    escolhido.update(generico())
    df = pd.DataFrame(escolhido.values(), columns=COLUNAS).sort_values("sku")
    conn = sqlite3.connect(caminhos.PADRONIZADO)
    df.to_sql("produtos", conn, if_exists="replace", index=False)
    conn.close()
    print(f"{len(df)} produtos (SKUs) | por coleção: {df.colecao.value_counts().to_dict()}")
    print(f"status: {df.status.value_counts().to_dict()} | sem preço: {int(df.preco.isna().sum()) - 1} (fora o DIVERSO)")
    return df


if __name__ == "__main__":
    main()
