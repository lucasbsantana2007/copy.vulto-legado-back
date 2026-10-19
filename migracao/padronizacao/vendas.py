"""
Vendas do vulto.db (lojas, site e app) e faturamento — decisões 38, 39, 41 a 50 do Registro de Decisões.

Entrada : vulto.db (vendas, itens_venda) + cadastro mestre do catálogo + clientes corrigidos + vendas da Batel
          corrigidas (saída de vendas_batel.py)
Saída   : vendas corrigidas (uma linha por venda, com o motivo de entrar ou não no faturamento), itens com o SKU
          do catálogo, faturamento por mês/loja/canal e a ponte "relatório da Vulto × nosso número".

O vulto.db original não é alterado: o relatório antigo precisa continuar reproduzível para a comparação.
"""
import os
import re
import sqlite3
from datetime import datetime, timedelta

import pandas as pd

from comum import (COR_POR_NOME, DATA_CORTE, INICIO_PERIODO, LOJAS, NOME_COR, Registro, chave, dinheiro,
                   identificar, montar_sku, sku_canonico, tamanho_padrao, vazio, data_vulto)

# decisão 39 — status gravados de 10 jeitos
STATUS = {"pago": "Paga", "aprovado": "Paga", "cancelado": "Cancelada", "cancelada": "Cancelada",
          "estornado": "Estornada", "chargeback": "Chargeback", "aguardando": "Pagamento não confirmado"}
STATUS_FATURA = {"Paga"}
# chargeback: a peça saiu e o dinheiro voltou; estornada/cancelada/não paga: a peça não saiu (ou voltou)
STATUS_BAIXA_ESTOQUE = {"Paga", "Chargeback"}
PRIORIDADE_DUPLICADA = ["Paga", "Chargeback", "Estornada", "Cancelada", "Pagamento não confirmado"]
CANAL = {"S": "Site", "APP": "App"}
LOJA_DO_CANAL = {"Site": 5, "App": 6}
TESTE = re.compile(r"^(S99\d{4}|APP-99\d{5})$")
SFS = re.compile(r"sfs|saiu da loja pinheiros")
RAZAO_CENTAVOS = 20                          # preço gravado ≥ 20x o preço do catálogo = valor em centavos
INICIO_PDV_BATEL, FIM_PDV_BATEL = datetime(2026, 3, 2), datetime(2026, 3, 23)
GENERICOS = {"DIVERSO", "PRODUTO SEM CADASTRO"}             # código genérico do PDV (peça não identificada)
DESCRICAO_GENERICO = "Peça não identificada no PDV"


def canal_da_venda(pedido_externo):
    """Decisão 45: o canal vem do número do pedido (S... = site, APP-... = app); o texto do canal não serve
    ('online' aparece nos dois e há venda do app gravada na loja do site)."""
    if vazio(pedido_externo):
        return "Loja física"
    p = str(pedido_externo).upper()
    return "App" if p.startswith("APP") else "Site" if p.startswith("S") else "Loja física"


def ler(pasta):
    with sqlite3.connect(os.path.join(pasta, "vulto.db")) as c:
        vendas = pd.read_sql("select * from vendas", c)
        itens = pd.read_sql("select * from itens_venda", c)
    return vendas, itens


def relatorio_vulto(vendas, inicio="2025-01-01", fim="2026-09-30"):
    """Reproduz o GET /relatorio/faturamento do sistema da Vulto (main.py), com os mesmos defeitos:
    soma o total de todas as vendas, de qualquer status, e usa a data gravada sem converter o fuso."""
    def dia(s):
        s = str(s).replace("Z", "")
        for f in ["%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
                  "%Y-%m-%d"]:
            try:
                return datetime.strptime(s, f).strftime("%Y-%m-%d")
            except ValueError:
                pass
        return None

    def valor(v):
        try:
            return float(str(v).replace(",", "."))
        except ValueError:
            return 0.0

    v = vendas.assign(dia_sistema=vendas.criado_em.map(dia), valor_sistema=vendas.total.map(valor))
    v["no_relatorio"] = v.dia_sistema.notna() & (v.dia_sistema >= inicio) & (v.dia_sistema <= fim)
    return v


def processar(pasta, mestre, clientes, batel):
    reg = Registro("Vendas")
    vendas, itens = ler(pasta)
    v = relatorio_vulto(vendas)
    preco_catalogo = dict(zip(mestre.sku, mestre.preco))

    # ------------------------------------------------------------------ itens: SKU do catálogo e preço em número
    itens["sku_original"] = itens.sku
    itens["sku"] = itens.sku.map(lambda s: sku_canonico(s) or str(s).strip().upper())
    itens["no_catalogo"] = itens.sku.isin(preco_catalogo)
    itens["preco_unit"] = itens.preco_unit.map(dinheiro)
    itens["razao_catalogo"] = [abs(p) / preco_catalogo[s] if s in preco_catalogo and preco_catalogo[s] else None
                               for s, p in zip(itens.sku, itens.preco_unit)]
    razao = itens.groupby("venda_id").razao_catalogo.median()

    # ------------------------------------------------------------------ canal e loja (decisão 45)
    v["canal_original"] = v.canal
    v["loja_id_original"] = v.loja_id
    v["canal"] = v.pedido_externo.map(canal_da_venda)
    v["loja_id"] = [LOJA_DO_CANAL.get(c, l) for c, l in zip(v.canal, v.loja_id_original)]
    for _, r in v.iterrows():
        if r.loja_id != r.loja_id_original:
            reg.add(45, "Venda do app gravada na loja do site", r.id, "loja_id", r.loja_id_original, r.loja_id,
                    f"Pedido {r.pedido_externo}")
        elif chave(r.canal_original) not in ("e-commerce", "app", "loja fisica"):
            reg.add(45, "Canal escrito fora do padrão", r.id, "canal", r.canal_original, r.canal)
    v["loja"] = v.loja_id.map(LOJAS)

    # ------------------------------------------------------------------ data no horário de Brasília (decisão 44)
    v["data_hora"] = v.criado_em.map(data_vulto)
    for _, r in v[v.criado_em.str.endswith("Z")].iterrows():
        reg.add(44, "Data em UTC convertida para Brasília", r.id, "criado_em", r.criado_em,
                r.data_hora.strftime("%Y-%m-%d %H:%M:%S"),
                "Mudou de dia" if r.data_hora.strftime("%Y-%m-%d") != r.dia_sistema else "")

    # ------------------------------------------------------------------ pedidos de teste (decisão 43)
    v["teste"] = v.pedido_externo.fillna("").str.match(TESTE) | v.obs.fillna("").str.lower().str.contains("teste")
    for _, r in v[v.teste].iterrows():
        reg.add(43, "Pedido de teste", r.id, "pedido_externo", r.pedido_externo, None,
                f"R$ {r.total} — série de QA {str(r.pedido_externo)[:5]}…" if r.pedido_externo else "obs 'teste'")

    # ------------------------------------------------------------------ valores em centavos (decisão 38)
    v["razao_catalogo"] = v.id.map(razao)
    sem_separador = ~v.total.astype(str).str.contains(r"[.,]")
    v["centavos"] = (v.razao_catalogo > RAZAO_CENTAVOS) | (
        v.razao_catalogo.isna() & (v.canal == "App") & sem_separador & (v.valor_sistema.abs() >= 1000)) & ~v.teste
    for col in ("total", "desconto", "frete"):
        v[col + "_num"] = v[col].map(lambda x: dinheiro(x) or 0.0)
        v.loc[v.centavos, col + "_num"] = (v.loc[v.centavos, col + "_num"] / 100).round(2)
    em_centavos = set(v.loc[v.centavos, "id"])
    itens.loc[itens.venda_id.isin(em_centavos), "preco_unit"] = (itens.preco_unit / 100).round(2)
    for _, r in v[v.centavos].iterrows():
        reg.add(38, "Valor gravado em centavos", r.id, "total", r.total, f"{r.total_num:.2f}",
                f"{r.canal}: preço dos itens {r.razao_catalogo:.0f}x o do catálogo" if r.razao_catalogo == r.razao_catalogo
                else f"{r.canal}: total sem vírgula e sem itens no catálogo")

    # ------------------------------------------------------------------ total = itens − desconto + frete (decisão 46)
    soma_itens = itens.assign(s=itens.qtd * itens.preco_unit).groupby("venda_id").s.sum()
    v["itens_valor"] = v.id.map(soma_itens).fillna(0).round(2)
    v["frete_valor"] = v.frete_num
    diferenca = (v.total_num - (v.itens_valor - v.desconto_num + v.frete_num)).round(2)
    embutido = (diferenca > 0.02) & (v.frete_num == 0) & (v.canal != "Loja física")
    v.loc[embutido, "frete_valor"] = diferenca[embutido]
    for _, r in v[embutido].iterrows():
        reg.add(46, "Frete embutido no total (campo frete vazio)", r.id, "frete", r.frete, f"{r.frete_valor:.2f}",
                f"{r.canal}: total {r.total_num:.2f} = itens {r.itens_valor:.2f} − desconto {r.desconto_num:.2f} + frete")
    nao_confere = (diferenca.abs() > 0.02) & ~embutido
    for _, r in v[nao_confere].iterrows():
        reg.add(46, "Total não confere com os itens: vale o total", r.id, "total", f"{r.total_num:.2f}",
                f"{r.total_num:.2f}", f"itens {r.itens_valor:.2f} − desconto {r.desconto_num:.2f} + frete {r.frete_num:.2f}")
    for _, r in v[(v.frete_valor > 0) & ~embutido].iterrows():
        reg.add(46, "Frete do site separado da receita", r.id, "frete", r.frete, f"{r.frete_valor:.2f}")
    v["valor_mercadorias"] = (v.total_num - v.frete_valor).round(2)

    # ------------------------------------------------------------------ status (decisão 39)
    v["status_original"] = v.status
    loja_fisica = v.canal == "Loja física"
    v["status"] = v.status.map(lambda s: STATUS.get(chave(s)))
    v.loc[v.status.isna() & loja_fisica, "status"] = "Paga"
    for _, r in v.iterrows():
        if vazio(r.status_original):
            reg.add(39, "Status vazio no PDV: venda de balcão considerada paga", r.id, "status", None, r.status)
        elif r.status_original != r.status:
            reg.add(39, "Status padronizado", r.id, "status", r.status_original, r.status,
                    "Fora do faturamento" if r.status not in STATUS_FATURA else "")

    # ------------------------------------------------------------------ pedidos duplicados (decisão 42)
    v["duplicada_de"] = None
    com_pedido = v[v.pedido_externo.fillna("") != ""]
    for ped, g in com_pedido.groupby("pedido_externo"):
        if len(g) < 2:
            continue
        g = g.assign(p=g.status.map(PRIORIDADE_DUPLICADA.index)).sort_values(["p", "id"], ascending=[True, False])
        fica = g.iloc[0]
        for _, r in g.iloc[1:].iterrows():
            v.loc[v.id == r.id, "duplicada_de"] = fica.id
            reg.add(42, "Pedido gravado duas vezes", r.id, "pedido_externo", ped, None,
                    f"{r.status} e depois {fica.status} (venda {fica.id} fica); diferença de "
                    f"{(fica.data_hora - r.data_hora).total_seconds() / 60:.0f} min")

    # ------------------------------------------------------------------ tipo: venda, devolução ou troca (47/48)
    obs = v.obs.fillna("").str.lower()
    v["tipo"] = "Venda"
    v.loc[v.total_num < 0, "tipo"] = "Devolução"
    v.loc[obs.str.contains("troca") & (v.total_num >= 0), "tipo"] = "Troca"
    for _, r in v[v.tipo == "Devolução"].iterrows():
        reg.add(47, "Devolução lançada como venda negativa", r.id, "total", r.total, f"{r.total_num:.2f}",
                "Desconta no mês em que foi devolvida")
    for _, r in v[v.tipo == "Troca"].iterrows():
        reg.add(48, "Troca lançada como venda nova", r.id, "total", r.total, "0,00 no faturamento",
                f"R$ {r.valor_mercadorias:.2f} já tinham sido pagos na compra original; a peça nova sai do estoque")

    # ------------------------------------------------------------------ de onde a peça saiu (decisão 49)
    v["ship_from_store"] = obs.str.contains(SFS)
    v["loja_estoque"] = [7 if c != "Loja física" else l for c, l in zip(v.canal, v.loja_id)]
    v.loc[v.ship_from_store, "loja_estoque"] = 1
    for _, r in v[v.ship_from_store].iterrows():
        reg.add(49, "Pedido do site saiu da loja Pinheiros", r.id, "loja_estoque", 7, 1,
                f"obs '{r.obs}': o faturamento fica no site, a baixa de estoque vai para Pinheiros")

    # ------------------------------------------------------------------ cliente da venda (decisão 50)
    por_email = dict(zip(clientes.email.dropna(), clientes.loc[clientes.email.notna(), "cpf"]))
    por_tel = clientes.dropna(subset=["telefone"]).drop_duplicates("telefone", keep=False)
    por_tel = dict(zip(por_tel.telefone, por_tel.cpf))
    cpfs, como = [], []
    for _, r in v.iterrows():
        tipo, val = identificar(r.cliente_ref)
        cpf = val if tipo == "cpf" else por_email.get(val) if tipo == "email" else por_tel.get(val) if tipo == "telefone" else None
        cpfs.append(cpf)
        como.append({"cpf": "CPF", "email": "E-mail do cadastro", "telefone": "Telefone do cadastro"}.get(tipo)
                    if cpf else ("Não identificado" if tipo != "anonimo" else "Sem cliente"))
        if tipo in ("email", "telefone"):
            reg.add(50, f"Cliente pelo {tipo}" + ("" if cpf else ": não encontrado no cadastro"), r.id, "cliente_ref",
                    r.cliente_ref, cpf)
    v["cliente_cpf"], v["cliente_identificado_por"] = cpfs, como

    # ------------------------------------------------------------------ Batel sem PDV (decisões 41 e 71)
    v, extras = conciliar_batel(v, batel, reg)
    itens, skus_criados = itens_batel(v, itens, batel, mestre, reg)

    # ------------------------------------------------------------------ peça vendida com código genérico (decisão 72)
    generico = itens.sku.isin(GENERICOS)
    for r in itens[generico].itertuples():
        reg.add(72, "Item vendido com código genérico do PDV", f"itens_venda {r.id}", "sku/descricao",
                f"{r.sku_original or r.sku} / {r.descricao}", f"DIVERSO / {DESCRICAO_GENERICO}",
                f"Venda {r.venda_id}, R$ {r.preco_unit:.2f}: valor fica no faturamento, sem baixa de estoque")
    itens.loc[generico, ["sku", "descricao"]] = ["DIVERSO", DESCRICAO_GENERICO]
    v["itens_valor"] = v.id.map(itens.assign(s=itens.qtd * itens.preco_unit).groupby("venda_id").s.sum()).fillna(0).round(2)

    # ------------------------------------------------------------------ entra no faturamento?
    v["motivo_fora"] = None
    v.loc[v.status.isin(STATUS_FATURA) == False, "motivo_fora"] = "Status: " + v.status.fillna("?")  # noqa: E712
    v.loc[v.tipo == "Troca", "motivo_fora"] = "Troca (não é venda nova)"
    v.loc[v.duplicada_de.notna(), "motivo_fora"] = "Pedido duplicado"
    v.loc[v.teste, "motivo_fora"] = "Pedido de teste"
    v["conta_faturamento"] = v.motivo_fora.isna()
    v["baixa_estoque"] = v.status.isin(STATUS_BAIXA_ESTOQUE) & ~v.teste & v.duplicada_de.isna()
    v["data"] = v.data_hora.dt.date
    v["mes"] = v.data_hora.dt.strftime("%Y-%m")
    v["no_periodo"] = (v.data_hora >= INICIO_PERIODO) & (v.data_hora < DATA_CORTE)
    v["origem"] = v.origem.fillna("vulto.db")

    itens = itens[["id", "venda_id", "sku", "sku_original", "no_catalogo", "descricao", "qtd", "preco_unit"]]
    colunas = ["id", "pedido_externo", "origem", "data_hora", "data", "mes", "loja_id", "loja", "canal", "tipo",
               "status", "conta_faturamento", "motivo_fora", "valor_mercadorias", "desconto_num", "frete_valor",
               "total_num", "itens_valor", "centavos", "cliente_cpf", "cliente_identificado_por", "vendedor",
               "loja_estoque", "baixa_estoque", "ship_from_store", "duplicada_de", "teste", "obs", "compra_id",
               "no_periodo", "linha_batel", "criado_em", "canal_original", "loja_id_original", "status_original", "total",
               "cliente_ref", "dia_sistema", "valor_sistema", "no_relatorio"]
    fat, pon = faturamento(v), ponte(v)
    nomes = {"desconto_num": "desconto_valor", "total_num": "total_corrigido"}
    v = v.rename(columns=nomes)
    return dict(vendas=v[[nomes.get(c, c) for c in colunas]], itens=itens, extras_batel=extras, faturamento=fat,
                ponte=pon, skus_criados=skus_criados, diversos=itens_diversos(v, itens, mestre)), reg


def conciliar_batel(v, batel, reg):
    """Decisão 41. A Batel ficou sem PDV de 02 a 22/03/2026 e anotou as vendas numa planilha. Quando o sistema
    voltou, a gerente relançou 'as principais' (com a data do relançamento). Cada venda relançada que bate com
    uma linha da planilha (mesmo valor; mesmo CPF quando os dois têm) conta uma vez só, na data real da planilha.
    Linha da planilha sem relançamento entra como venda nova (sem itens: as peças estão escritas à mão)."""
    v["origem"], v["linha_batel"] = None, None
    validas = batel[batel.status == "Concluída"].copy()
    candidatas = v[(v.loja_id == 4) & (v.data_hora >= FIM_PDV_BATEL) & (v.data_hora < datetime(2026, 4, 16))
                   & (v.total_num > 0)].copy()
    candidatas["retroativo"] = candidatas.obs.fillna("").str.contains("retroativo")
    usadas, casadas = set(), {}
    for i, m in validas.iterrows():
        cpf_m = re.sub(r"\D", "", str(m.cpf)) if not vazio(m.cpf) else None
        c = candidatas[(~candidatas.id.isin(usadas)) & ((candidatas.total_num - float(m.valor)).abs() < 0.01)]
        if cpf_m:
            c = c[c.cliente_cpf.isna() | (c.cliente_cpf == cpf_m)]
        if c.empty:
            continue
        escolhida = c.sort_values(["retroativo", "data_hora"], ascending=[False, True]).iloc[0]
        usadas.add(escolhida.id)
        casadas[i] = escolhida.id
        hora = m.hora if not vazio(m.hora) else "12:00"
        nova = datetime.combine(pd.Timestamp(m.data).date(), datetime.strptime(hora[:5], "%H:%M").time())
        v.loc[v.id == escolhida.id, ["data_hora", "origem", "linha_batel"]] = [nova, "vulto.db (relançada; data da planilha)",
                                                                         m.linha_planilha]
        reg.add(41, "Venda relançada casada com a planilha manual", escolhida.id, "criado_em", escolhida.criado_em,
                nova.strftime("%Y-%m-%d %H:%M"), f"Linha {m.linha_planilha} da planilha, R$ {float(m.valor):.2f}: conta uma vez")
    extras = []
    for i, m in batel.iterrows():
        if i in casadas:
            continue
        hora = m.hora if not vazio(m.hora) else "12:00"
        dh = datetime.combine(pd.Timestamp(m.data).date(), datetime.strptime(hora[:5], "%H:%M").time())
        status = "Paga" if m.status == "Concluída" else "Cancelada"
        cpf_m = re.sub(r"\D", "", str(m.cpf)) if not vazio(m.cpf) else None
        extras.append(dict(id=f"BATEL-{m.linha_planilha}", pedido_externo=None, origem="Planilha Batel (PDV fora do ar)",
                           data_hora=dh, loja_id=4, loja=LOJAS[4], canal="Loja física", tipo="Venda", status=status,
                           total_num=float(m.valor), desconto_num=0.0, frete_valor=0.0, valor_mercadorias=float(m.valor),
                           itens_valor=0.0, centavos=False, cliente_cpf=cpf_m,
                           cliente_identificado_por="CPF" if cpf_m else "Sem cliente", vendedor=m.vendedora,
                           loja_estoque=4, ship_from_store=False, duplicada_de=None, teste=False, obs=m.pecas,
                           compra_id=None, criado_em=None, canal_original=None, loja_id_original=4,
                           status_original=m.status, total=m.valor_original, cliente_ref=m.cpf_original,
                           dia_sistema=None, valor_sistema=0.0, no_relatorio=False, linha_batel=m.linha_planilha))
        if status == "Paga":
            reg.add(41, "Venda da planilha sem relançamento: incluída", f"planilha!{m.linha_planilha}", "venda", None,
                    f"R$ {float(m.valor):.2f}", f"{pd.Timestamp(m.data):%d/%m}: {m.pecas}")
    extras = pd.DataFrame(extras)
    if len(extras):
        v = pd.concat([v, extras], ignore_index=True)
    return v, extras


def ler_pecas(texto, mestre):
    """Texto da planilha da Batel -> [(sku, qtd, descricao)]. Ex.: '2x pochete nylon branco U + camisa linho leve
    lilás G' vira 2 Pochete Nylon BRA-U e 1 Camisa Linho Leve LIL-G. O nome do modelo vem do catálogo; a cor e o
    tamanho, do fim do texto."""
    modelos = mestre.drop_duplicates("modelo")
    nomes = sorted(((chave(n), m) for m, n in zip(modelos.modelo, modelos.nome)), key=lambda x: -len(x[0]))
    cores = sorted(COR_POR_NOME, key=len, reverse=True)
    pecas = []
    for parte in re.split(r"\s*\+\s*", str(texto)):
        k, qtd = chave(parte), 1
        m = re.match(r"(\d+)\s*x\s*", k)
        if m:
            qtd, k = int(m.group(1)), k[m.end():]
        achado = next(((n, mod) for n, mod in nomes if k.startswith(n)), None)
        if achado is None:
            pecas.append((None, qtd, parte))
            continue
        resto = k[len(achado[0]):].strip()
        cor = next((COR_POR_NOME[c] for c in cores if resto.startswith(c)), None)
        tam = tamanho_padrao(resto.split()[-1]) if resto else None
        sku = montar_sku(*achado[1].split("-"), cor, tam) if cor and tam else None
        pecas.append((sku, qtd, parte))
    return pecas


def itens_batel(v, itens, batel, mestre, reg):
    """Decisão 71. As vendas da Batel feitas sem sistema descrevem as peças em texto. Cada peça vira um item com o
    SKU do catálogo; se o modelo existe mas aquela cor/tamanho ainda não tem SKU, o SKU é criado no padrão
    VLT-CAT-NNN-COR-TAM. Nas vendas relançadas, só os itens 'DIVERSO - VENDA MANUAL' são trocados."""
    nomes = dict(zip(mestre.sku, mestre.nome))
    precos = dict(zip(mestre.sku, mestre.preco))
    por_linha = batel.set_index("linha_planilha")
    alvo = v[v.linha_batel.notna()]
    proximo = int(itens.id.max()) + 1
    novos, tirar, criados = [], [], []
    for venda in alvo.itertuples():
        atuais = itens[itens.venda_id == venda.id]
        genericos = atuais[atuais.sku.isin(GENERICOS) | atuais.descricao.fillna("").str.contains("VENDA MANUAL")]
        if len(atuais) and genericos.empty:
            continue                                            # relançada com os códigos certos: fica como está
        m = por_linha.loc[int(venda.linha_batel)]
        pecas = ler_pecas(m.pecas, mestre)
        base = sum(q * precos.get(s, 0) for s, q, _ in pecas if s)
        fator = float(m.valor) / base if base else 1
        for i, (sku, qtd, texto) in enumerate(pecas):
            if sku is None:
                sku, desc = "DIVERSO", "Peça não identificada no texto da planilha"
                reg.add(71, "Peça da planilha não reconhecida", f"Batel!{venda.linha_batel}", "pecas", texto, sku)
            elif sku not in nomes:
                criados.append(sku)
                desc = f"{texto.strip()} (SKU criado)"
                reg.add(71, "SKU criado no padrão para peça da planilha", f"Batel!{venda.linha_batel}", "sku", texto, sku)
            else:
                desc = f"{nomes[sku]} {sku.split('-')[3]}"
                reg.add(71, "Peça da planilha reconhecida no catálogo", f"Batel!{venda.linha_batel}", "sku", texto, sku)
            novos.append(dict(id=proximo, venda_id=venda.id, sku=sku, sku_original=None, no_catalogo=sku in nomes,
                              descricao=desc, qtd=qtd, preco_unit=round(precos.get(sku, 0) * fator, 2)))
            proximo += 1
        tirar += list(genericos.id) if len(atuais) else []
        # o valor dos itens fecha com o valor da planilha: a diferença de centavos fica no último item
        feitos = [n for n in novos if n["venda_id"] == venda.id]
        if feitos:
            soma = sum(n["qtd"] * n["preco_unit"] for n in feitos)
            feitos[-1]["preco_unit"] = round(feitos[-1]["preco_unit"] + (float(m.valor) - soma) / feitos[-1]["qtd"], 2)
    for i in tirar:
        reg.add(71, "Item 'DIVERSO - VENDA MANUAL' trocado pelas peças da planilha", f"itens_venda {i}", "sku", "DIVERSO", None)
    itens = pd.concat([itens[~itens.id.isin(tirar)], pd.DataFrame(novos)], ignore_index=True)
    return itens, sorted(set(criados))


def itens_diversos(v, itens, mestre):
    """Itens vendidos com o código genérico do PDV, com os modelos do catálogo que têm o mesmo preço."""
    d = itens[itens.sku.isin(GENERICOS)].merge(
        v[["id", "data_hora", "loja", "status", "obs"]].rename(columns={"id": "venda_id"}), on="venda_id")
    por_preco = mestre.groupby(mestre.preco.round(2)).nome.agg(lambda s: sorted(set(s)))
    d["modelos_com_esse_preco"] = d.preco_unit.round(2).map(lambda p: len(por_preco.get(p, [])))
    d["candidatos"] = d.preco_unit.round(2).map(lambda p: ", ".join(por_preco.get(p, [])[:6]))
    return d[["id", "venda_id", "data_hora", "loja", "status", "sku", "descricao", "qtd", "preco_unit",
              "modelos_com_esse_preco", "candidatos"]].sort_values("data_hora")


def faturamento(v):
    """Faturamento de mercadorias (sem frete), por mês, loja e canal. Devoluções descontam no mês da devolução."""
    f = v[v.conta_faturamento & v.no_periodo]
    tab = (f.assign(vendas=f.valor_mercadorias.where(f.tipo == "Venda", 0),
                    devolucoes=f.valor_mercadorias.where(f.tipo == "Devolução", 0),
                    pedidos=(f.tipo == "Venda").astype(int))
           .groupby(["mes", "loja", "canal"], as_index=False)
           .agg(pedidos=("pedidos", "sum"), vendas=("vendas", "sum"), devolucoes=("devolucoes", "sum"),
                faturamento=("valor_mercadorias", "sum")))
    frete = v[v.no_periodo & v.status.isin(STATUS_FATURA) & ~v.teste & v.duplicada_de.isna()] \
        .groupby(["mes", "loja", "canal"]).frete_valor.sum()
    tab["frete_cobrado_fora"] = [frete.get((m, l, c), 0.0) for m, l, c in zip(tab.mes, tab.loja, tab.canal)]
    return tab.round(2)


def ponte(v):
    """Do número do relatório da Vulto até o nosso, uma decisão por vez (cada passo mostra quanto mudou)."""
    base = v[v.origem.fillna("") != "Planilha Batel (PDV fora do ar)"]
    passos = []
    atual = base.no_relatorio & True
    valor = base.valor_sistema.copy()

    def registra(nome, decisao, mascara, novo_valor, afetadas):
        nonlocal atual, valor
        antes = valor[atual].sum()
        atual, valor = mascara, novo_valor
        passos.append(dict(passo=nome, decisao=decisao, vendas_afetadas=int(afetadas),
                           variacao=round(valor[atual].sum() - antes, 2), resultado=round(valor[atual].sum(), 2)))

    passos.append(dict(passo="Relatório do sistema da Vulto (jan/25 a set/26)", decisao=None,
                       vendas_afetadas=int(atual.sum()), variacao=None, resultado=round(valor[atual].sum(), 2)))
    registra("Tirar pedidos de teste", 43, atual & ~base.teste, valor, (atual & base.teste).sum())
    registra("Valores do app em centavos ÷ 100", 38, atual, base.total_num.where(base.centavos, valor),
             (atual & base.centavos).sum())
    registra("Tirar pedidos duplicados", 42, atual & base.duplicada_de.isna(), valor, (atual & base.duplicada_de.notna()).sum())
    pago = base.status.isin(STATUS_FATURA)
    registra("Tirar cancelados, estornados, chargeback e não pagos", 39, atual & pago, valor, (atual & ~pago).sum())
    registra("Tirar trocas (já pagas na compra original)", 48, atual & (base.tipo != "Troca"), valor,
             (atual & (base.tipo == "Troca")).sum())
    registra("Tirar o frete (não é receita de mercadoria)", 46, atual, base.valor_mercadorias,
             (atual & (base.frete_valor > 0)).sum())
    registra("Datas no horário de Brasília (vendas que mudam de mês nas pontas do período)", 44,
             base.conta_faturamento & base.no_periodo, valor,
             (base.conta_faturamento & (base.no_periodo != base.no_relatorio)).sum())
    extras = v[(v.origem == "Planilha Batel (PDV fora do ar)") & v.conta_faturamento]
    passos.append(dict(passo="Somar vendas da Batel feitas sem sistema (não relançadas)", decisao=41,
                       vendas_afetadas=len(extras), variacao=round(extras.valor_mercadorias.sum(), 2),
                       resultado=round(valor[atual].sum() + extras.valor_mercadorias.sum(), 2)))
    final = v[v.conta_faturamento & v.no_periodo].valor_mercadorias.sum()
    passos.append(dict(passo="Nosso faturamento (mercadorias, líquido de devoluções)", decisao=None,
                       vendas_afetadas=int((v.conta_faturamento & v.no_periodo).sum()), variacao=None,
                       resultado=round(final, 2)))
    return pd.DataFrame(passos)


if __name__ == "__main__":
    import sys

    import catalogo
    import clientes
    import vendas_batel
    from comum import salvar_excel
    pasta = sys.argv[1] if len(sys.argv) > 1 else "../vulto-legado-back/data"
    cat, _ = catalogo.processar(pasta)
    cli, _ = clientes.processar(pasta)
    bat, _ = vendas_batel.processar(pasta)
    dados, reg = processar(pasta, cat["mestre"], cli["clientes"], bat["vendas"])
    salvar_excel("vendas_corrigidas.xlsx", {"Ponte": dados["ponte"], "Faturamento": dados["faturamento"],
                                            "Vendas": dados["vendas"], "Itens": dados["itens"],
                                            "Ocorrências": reg.df(), "Resumo": reg.resumo()})
    print(dados["ponte"].to_string(index=False))
    print(reg.resumo().to_string())
