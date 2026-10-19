"""
Inventário de setembro/2026 — decisões 28 a 37 do Registro de Decisões.

Entrada : contagem_inventario_set26.xlsx + cadastro mestre do catálogo (saída de catalogo.py)
Saída   : inventario_corrigido.xlsx (Estoque consolidado, uma aba por local, AVARIA PINHEIROS, Ocorrências, Resumo)
"""
import os
import re
from collections import Counter

import pandas as pd

from comum import Registro, chave, digitos, ler_codigo_produto, montar_sku, tamanho_padrao, vazio

ARQUIVO = "contagem_inventario_set26.xlsx"
DATA_POSICAO = "28/09/2026"                                    # decisão 33
FORA_DO_CATALOGO_2026 = "Fora do catálogo 2026"


class Catalogo:
    """Consulta ao cadastro mestre: SKU, código do CD, status e correção de códigos embaralhados."""

    def __init__(self, mestre, skus_2024):
        self.por_sku = mestre.set_index("sku").to_dict("index")
        self.por_cd = dict(zip(mestre.codigo_cd, mestre.sku))
        self.skus_2024 = skus_2024
        self.modelos_2024 = {s[:11] for s in skus_2024}
        self._anag_sku = {}
        for s in self.por_sku:
            self._anag_sku.setdefault(("".join(sorted(s))), []).append(s)
        self._anag_cd = {}
        for c in self.por_cd:
            self._anag_cd.setdefault((c[:3], "".join(sorted(c))), []).append(c)

    def sku_de(self, texto):
        p = ler_codigo_produto(texto)
        if not p or not p[3] or not tamanho_padrao(p[3]):
            return None
        return montar_sku(p[0], p[1], p[2], tamanho_padrao(p[3]))

    def status(self, sku):
        if sku in self.por_sku:
            return self.por_sku[sku]["status"]
        if sku in self.skus_2024 or (sku and sku[:11] in self.modelos_2024):
            return FORA_DO_CATALOGO_2026
        return "Não encontrado"

    def corrigir_sku(self, texto, reg, origem):
        """SKU fora do padrão ou digitado com letras trocadas (VLT-CLA-003 -> VLT-CAL-003, BAR -> BRA)."""
        sku = self.sku_de(texto)
        if sku and (sku in self.por_sku or self.status(sku) == FORA_DO_CATALOGO_2026):
            if sku != texto:
                reg.add(35, "SKU convertido para o padrão do catálogo", origem, "sku", texto, sku)
            return sku
        corrigido = self._desembaralhar(texto)
        if corrigido:
            reg.add(35, "SKU com letras trocadas ou hífen fora do lugar corrigido", origem, "sku", texto, corrigido)
            return corrigido
        reg.add(35, "SKU não encontrado no catálogo", origem, "sku", texto, sku)
        return sku or str(texto).upper()

    def _desembaralhar(self, texto):
        m = re.fullmatch(r"VLT([A-Z]{3})(\d{3})([A-Z]{3})([A-Z0-9]+)", re.sub(r"[-\s]", "", str(texto).upper()))
        if not m:
            return None
        prefixos = {p[4:7] for p in self.por_sku}
        cores = {p.split("-")[3] for p in self.por_sku}

        def ajusta(seg, validos):
            if seg in validos:
                return seg
            achados = [v for v in validos if sorted(v) == sorted(seg)]
            return achados[0] if len(achados) == 1 else None

        pre, cor = ajusta(m.group(1), prefixos), ajusta(m.group(3), cores)
        tam = tamanho_padrao(m.group(4))
        if not (pre and cor and tam):
            return None
        sku = montar_sku(pre, m.group(2), cor, tam)
        return sku if (sku in self.por_sku or self.status(sku) == FORA_DO_CATALOGO_2026) else None

    def sku_do_codigo_cd(self, codigo, reg, origem):
        """Decisão 37: códigos do CD com dígitos invertidos (ex.: TOP030BP -> TOP003BP)."""
        c = str(codigo).strip().upper()
        if c in self.por_cd:
            return self.por_cd[c]
        sku = self.sku_de(c)
        if sku and self.status(sku) == FORA_DO_CATALOGO_2026:
            return sku
        cands = self._anag_cd.get((c[:3], "".join(sorted(c))), []) or \
            [k for k in self.por_cd if k[:6] == c[:6] and k == c[:6] + c[6:].replace("0", "O")]
        if len(cands) == 1:
            reg.add(37, "Código do CD com dígitos invertidos corrigido", origem, "codigo", codigo, cands[0])
            return self.por_cd[cands[0]]
        reg.add(35, "Código do CD não encontrado no catálogo", origem, "codigo", codigo, sku)
        return sku or c


def numero(v):
    """'2 (1 no estoque + 1 vitrine)' -> 2."""
    m = re.match(r"\s*(\d+)", str(v))
    return int(m.group(1)) if m else 0


def ler(pasta, aba, cabecalho):
    return pd.read_excel(os.path.join(pasta, ARQUIVO), sheet_name=aba, header=cabecalho, dtype=str)


def processar(pasta, mestre, skus_2024):
    reg = Registro("Inventário")
    cat = Catalogo(mestre, skus_2024)
    abas = pd.ExcelFile(os.path.join(pasta, ARQUIVO)).sheet_names
    estoque, avaria = [], []

    def item(loja, sku, qtd, origem_qtd, obs=None, **extra):
        estoque.append(dict(loja=loja, sku=sku, qtd=qtd, origem_qtd=origem_qtd, obs=obs, **extra))

    # ---------------------------------------------------------------- Pinheiros (contada em 20/09)
    pin = ler(pasta, next(a for a in abas if a.upper().startswith("PINHEIROS")), 2)
    for i, r in pin.iterrows():
        if vazio(r["SKU"]):
            continue
        origem = f"Pinheiros!{i + 4}"
        sku = cat.corrigir_sku(r["SKU"], reg, origem)
        qtd = numero(r["QTD"])
        if chave(r["OBS"]).startswith("avaria"):                  # decisão 28
            avaria.append(dict(loja="Pinheiros", sku=sku, qtd=qtd, obs=r["OBS"], linha_planilha=i + 4))
            reg.add(28, "Peça com avaria separada do estoque bom", origem, "qtd", qtd, 0, f"{qtd} peça(s)")
            continue
        item("Pinheiros", sku, qtd, "Contado", None if vazio(r["OBS"]) else f"local: {r['OBS']}")

    # ---------------------------------------------------------------- Eldorado (contada por código do CD)
    eld = ler(pasta, "Eldorado", 0)
    for i, r in eld.iterrows():
        if vazio(r["codigo"]):
            continue
        sku = cat.sku_do_codigo_cd(r["codigo"], reg, f"Eldorado!{i + 2}")
        item("Eldorado", sku, numero(r["qtde"]), "Contado", conferente=r["conferente"])

    # ---------------------------------------------------------------- Batel
    bat = ler(pasta, "batel", 0)
    contagens = pd.DataFrame(estoque)
    for i, r in bat.iterrows():
        if vazio(r["sku"]):
            continue
        origem = f"Batel!{i + 2}"
        sku = cat.corrigir_sku(r["sku"], reg, origem)
        qtd = numero(r["quantidade"])
        if "(" in str(r["quantidade"]):
            reg.add(35, "Quantidade com detalhe em texto", origem, "quantidade", r["quantidade"], qtd)
        obs = chave(r["obs"])
        if obs == "faltou contar provador":                       # decisão 29
            reg.add(29, "Provador não contado: considerado zero", origem, "qtd", qtd, qtd,
                    "Estoque conservador: melhor menor do que vender sem a peça")
            item("Batel", sku, qtd, "Contado (provador = 0)", "faltou contar provador")
        elif obs == "recontar":                                   # decisão 30
            ref = contagens[(contagens.sku == sku) & contagens.loja.isin(["Pinheiros", "Eldorado"])].qtd
            estimado = int(round(ref.median())) if len(ref) else qtd
            reg.add(30, "Recontagem não feita: estimado pela mediana", origem, "qtd", qtd, estimado,
                    f"Mediana do mesmo SKU na Pinheiros/Eldorado ({len(ref)} contagem(ns))" if len(ref)
                    else "Sem referência nas outras lojas: mantido o contado")
            item("Batel", sku, estimado, "Estimado (mediana Pinheiros/Eldorado)",
                 f"contado originalmente: {qtd} (recontagem pedida e não feita)")
        else:
            item("Batel", sku, qtd, "Contado")

    # ---------------------------------------------------------------- Savassi (só total por modelo) — decisão 31
    sav = ler(pasta, next(a for a in abas if a.upper().startswith("SAVASSI")), 0)
    base = pd.DataFrame(estoque)
    for i, r in sav.iterrows():
        if vazio(r["MODELO"]):
            continue
        modelo, total = str(r["MODELO"]).strip().upper(), numero(r["QTD TOTAL"])
        prefixo = f"VLT-{modelo}-"
        fonte = base[(base.loja == "Pinheiros") & base.sku.str.startswith(prefixo)]
        de_onde = "Pinheiros"
        if fonte.qtd.sum() == 0:
            fonte = base[base.loja.isin(["Eldorado", "Batel"]) & base.sku.str.startswith(prefixo)]
            de_onde = "Eldorado + Batel"
        pesos = fonte.groupby("sku").qtd.sum()
        if pesos.sum() == 0:
            reg.add(31, "Modelo sem referência para rateio", f"Savassi!{i + 2}", "qtd", total, None, modelo)
            continue
        cotas = pesos / pesos.sum() * total
        inteiros = cotas.astype(int)
        sobra = total - inteiros.sum()                           # maiores restos ficam com as peças que sobram
        for sku in (cotas - inteiros).sort_values(ascending=False).index[:sobra]:
            inteiros[sku] += 1
        for sku, q in inteiros.items():
            if q:
                item("Savassi", sku, int(q), f"Rateado (proporção {de_onde})", f"total do modelo {modelo}: {total}")
        reg.add(31, "Total do modelo rateado por cor e tamanho", f"Savassi!{i + 2}", "qtd", total,
                f"{(inteiros > 0).sum()} SKUs", f"Proporção da {de_onde}")

    # ---------------------------------------------------------------- CD (posição do WMS em 28/09) — decisão 32
    cd = ler(pasta, next(a for a in abas if a.upper().startswith("CD")), 0)
    for i, r in cd.iterrows():
        if vazio(r["SKU"]):
            continue
        sku = cat.corrigir_sku(r["SKU"], reg, f"CD!{i + 2}")
        reservado = numero(r["RESERVADO"])
        item("CD Cajamar", sku, numero(r["SALDO"]), "Contado (WMS)", disponivel=numero(r["DISPONIVEL"]),
             reservado=reservado)
        if reservado:
            reg.add(32, "Reservado mantido no CD", f"CD!{i + 2}", "reservado", reservado, reservado,
                    "Reposição de loja já pedida: fica no saldo do CD, não soma nas lojas")

    # ---------------------------------------------------------------- consolidação
    est = pd.DataFrame(estoque)
    agrupado = (est.groupby(["loja", "sku"], as_index=False)
                .agg(qtd=("qtd", "sum"), origem_qtd=("origem_qtd", "first"),
                     obs=("obs", lambda s: "; ".join(sorted({x for x in s if not vazio(x)})) or None),
                     disponivel=("disponivel", "sum") if "disponivel" in est else ("qtd", "size"),
                     reservado=("reservado", "sum") if "reservado" in est else ("qtd", "size")))
    for (loja, sku), n in est.groupby(["loja", "sku"]).size().items():
        if n > 1:
            reg.add(35, "SKU contado em mais de um local da loja", loja, "sku", sku, n, "Quantidades somadas")
    agrupado.loc[agrupado.loja != "CD Cajamar", ["disponivel", "reservado"]] = None
    agrupado["data_posicao"] = DATA_POSICAO
    agrupado["codigo_cd"] = agrupado.sku.map(lambda s: cat.por_sku.get(s, {}).get("codigo_cd"))
    agrupado["nome"] = agrupado.sku.map(lambda s: cat.por_sku.get(s, {}).get("nome"))
    agrupado["status_catalogo"] = agrupado.sku.map(cat.status)
    agrupado["candidato_liquidacao"] = agrupado.status_catalogo.isin(["Descontinuado", FORA_DO_CATALOGO_2026]) \
        .map({True: "Sim", False: "Não"})                          # decisão 36
    for loja in ["Eldorado", "Savassi", "Batel", "CD Cajamar"]:
        reg.add(34, "Sem registro de avaria", loja, None, None, 0, "Avaria considerada inexistente")
    for _, r in agrupado[agrupado.candidato_liquidacao == "Sim"].iterrows():
        reg.add(36, f"{r['status_catalogo']}: entra como estoque normal", r["loja"], "sku", r["sku"], r["qtd"],
                "Candidato a desconto/liquidação")
    reg.add(33, "Data de posição única", "todas", "data", "Pinheiros 20/09; demais sem data; CD 28/09",
            DATA_POSICAO)
    colunas = ["loja", "sku", "codigo_cd", "nome", "qtd", "origem_qtd", "status_catalogo", "candidato_liquidacao",
               "disponivel", "reservado", "data_posicao", "obs"]
    consolidado = agrupado[colunas].sort_values(["loja", "sku"])
    av = pd.DataFrame(avaria)
    av["codigo_cd"] = av.sku.map(lambda s: cat.por_sku.get(s, {}).get("codigo_cd"))
    resumo = (consolidado.groupby("loja").agg(skus=("sku", "nunique"), pecas=("qtd", "sum")).reset_index())
    resumo.loc[len(resumo)] = ["AVARIA Pinheiros (fora do estoque)", av.sku.nunique(), av.qtd.sum()]
    return dict(consolidado=consolidado, avaria=av, resumo_lojas=resumo), reg


def destaque(linha):
    if linha.get("candidato_liquidacao") == "Sim":
        return ("FCE4D6", "status_catalogo")                      # laranja: candidato a liquidação
    if str(linha.get("origem_qtd")).startswith(("Estimado", "Rateado")):
        return ("FFF2CC", "qtd")
    return None


if __name__ == "__main__":
    import sys
    import catalogo
    from comum import salvar_excel
    pasta = sys.argv[1] if len(sys.argv) > 1 else "../vulto-legado-back/data"
    cat, _ = catalogo.processar(pasta)
    dados, reg = processar(pasta, cat["mestre"], cat["skus_2024"])
    abas = {"Estoque consolidado": dados["consolidado"], "Resumo por loja": dados["resumo_lojas"]}
    for loja in dados["consolidado"].loja.unique():
        abas[loja.upper()] = dados["consolidado"][dados["consolidado"].loja == loja]
    abas.update({"AVARIA PINHEIROS": dados["avaria"], "Ocorrências": reg.df(), "Resumo": reg.resumo()})
    salvar_excel("inventario_corrigido.xlsx", abas, destaques={k: destaque for k in abas if k.isupper()
                                                               or k == "Estoque consolidado"})
    print(reg.resumo().to_string())
