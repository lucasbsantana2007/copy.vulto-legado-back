"""
Vulto+ (pontos) e passivo em 01/10/2026 — decisões 4 (revisada), 40 e 51 a 58 do Registro de Decisões.

Entrada : vulto.db (pontos_mov, compras) + clientes corrigidos (saldo da planilha, congelado em 01/01/2025) +
          vendas corrigidas (para saber se a compra foi paga) + vendas da Batel + aba Resgates da planilha
Saída   : extrato corrigido (uma linha por movimento), saldo por cliente em 01/10/2026 e a ponte
          "saldo que o sistema mostra × passivo real".

Como o saldo era calculado (aba LEIA ISSO): a coluna Pontos da planilha foi congelada em 01/01/2025 e,
a partir daí, os movimentos ficam no banco (pontos_mov). O saldo é a soma dos dois.
"""
import calendar
import os
import re
import sqlite3
from collections import deque
from datetime import datetime

import pandas as pd

from comum import DATA_CORTE, Registro, chave, data_vulto, digitos, dinheiro, vazio

ARQUIVO = "clientes_fidelidade_FINAL_v3 (1).xlsx"
INICIO_BANCO = datetime(2025, 1, 1)
VALOR_PONTO = 0.05                     # decisão 40: R$ 5,00 a cada 100 pontos desde mar/2025 (aba LEIA ISSO)
VALOR_PONTO_SISTEMA = 0.10             # o que está na tabela config do sistema (nunca foi atualizado)
VALIDADE_MESES = 12                    # decisão 54
VALIDADE_MESES_OURO = 18
FAIXAS = [(0, "Bronze"), (2000, "Prata"), (5000, "Ouro")]   # decisão 4: pontos acumulados nos últimos 12 meses
STATUS_VALIDO = {"Paga"}
MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10,
         "nov": 11, "dez": 12}


def categoria(acumulado_12m):
    nome = FAIXAS[0][1]
    for minimo, n in FAIXAS:
        if acumulado_12m >= minimo:
            nome = n
    return nome


def mais_meses(d, meses):
    m = d.month - 1 + meses
    ano, mes = d.year + m // 12, m % 12 + 1
    return d.replace(year=ano, month=mes, day=min(d.day, calendar.monthrange(ano, mes)[1]))


def parse_pontos_sistema(p):
    """Cópia do parse_pontos do sistema antigo (planilha.py), para reproduzir o saldo que ele mostrava."""
    try:
        if p == "" or p is None or (isinstance(p, float) and p != p):
            return 0
        if isinstance(p, (int, float)):
            return int(p)
        p = str(p).replace("pts", "").replace(" ", "").replace(".", "")
        return int(p)
    except ValueError:
        return 0


def data_resgate(v):
    """Aba Resgates: 'jan/23', '13/01/2023', '2023-02-16 00:00:00', '16/01/25'."""
    s = str(v).strip().lower()
    m = re.fullmatch(r"([a-z]{3})/(\d{2})", s)
    if m and m.group(1) in MESES:
        return datetime(2000 + int(m.group(2)), MESES[m.group(1)], 1)
    for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    return None


def processar(pasta, clientes, vendas, batel):
    reg = Registro("Vulto+")
    with sqlite3.connect(os.path.join(pasta, "vulto.db")) as c:
        mov = pd.read_sql("select * from pontos_mov", c)
        compras = pd.read_sql("select * from compras", c)
    mov["cpf_original"] = mov.cpf
    mov["cpf"] = mov.cpf.map(lambda x: digitos(x).zfill(11))
    for _, r in mov[mov.cpf_original.str.len() < 11].iterrows():
        reg.add(51, "CPF do extrato sem zeros à esquerda", r.id, "cpf", r.cpf_original, r.cpf)
    mov["data"] = mov.criado_em.map(data_vulto)
    mov["origem"] = "pontos_mov " + mov.id.astype(str)

    # ------------------------------------------------------------------ compra -> venda (para saber se foi paga)
    compras["valor_num"] = compras.valor.map(lambda x: dinheiro(x) or 0.0)
    compras["centavos"] = (compras.origem == "app") & ~compras.valor.astype(str).str.contains(r"[.,]")
    compras["valor_reais"] = compras.valor_num.where(~compras.centavos, compras.valor_num / 100)
    compras["online"] = compras.origem.isin(["site", "app"])
    venda = vendas[vendas.compra_id.notna()].drop_duplicates("compra_id").set_index("compra_id")
    mov = mov.merge(compras[["id", "valor_num", "valor_reais", "centavos", "online"]].rename(columns={"id": "compra_id"}),
                    on="compra_id", how="left")
    mov["venda_id"] = mov.compra_id.map(venda.id)
    mov["venda_status"] = mov.compra_id.map(venda.status)
    mov["venda_tipo"] = mov.compra_id.map(venda.tipo)
    mov["venda_invalida"] = mov.compra_id.map(lambda c: motivo_invalida(venda.loc[c]) if c in venda.index else None)

    # ------------------------------------------------------------------ decisão 51: crédito do app calculado em centavos
    cred = mov.tipo == "C"
    bug = cred & mov.centavos.fillna(False) & (
        ((mov.qtd - mov.valor_num.round()).abs() <= 1) | ((mov.qtd - (2 * mov.valor_num).round()).abs() <= 1))
    mov["qtd_corrigida"] = mov.qtd
    mov.loc[bug, "qtd_corrigida"] = (mov.loc[bug, "qtd"] / 100).round().astype(int)
    correcoes = mov[(mov.tipo == "AJ") & mov.obs.fillna("").str.startswith("correcao bug app")]
    ja_corrigido = set()
    for i, r in mov[bug].iterrows():
        esperado = -(r.qtd - r.qtd_corrigida)
        aj = correcoes[(correcoes.cpf == r.cpf) & ((correcoes.qtd - esperado).abs() <= 2) & ~correcoes.index.isin(ja_corrigido)]
        if len(aj):
            ja_corrigido.add(aj.index[0])
            reg.add(51, "Crédito do app em centavos (já corrigido à mão em jun/26)", r.origem, "qtd", r.qtd, r.qtd_corrigida,
                    f"Ajuste manual {aj.iloc[0].origem} ({aj.iloc[0].qtd}) substituído pela correção do crédito")
        else:
            reg.add(51, "Crédito do app em centavos (nunca corrigido)", r.origem, "qtd", r.qtd, r.qtd_corrigida,
                    f"Compra de R$ {r.valor_reais:.2f}: crédito 100x maior")
    for i, r in correcoes.iterrows():
        if i not in ja_corrigido:
            reg.add(51, "Ajuste 'correcao bug app' sem crédito correspondente: mantido", r.origem, "qtd", r.qtd, r.qtd)
    impacto = {51: int((mov.loc[bug, "qtd_corrigida"] - mov.loc[bug, "qtd"]).sum() - mov.loc[list(ja_corrigido), "qtd"].sum())}
    mov = mov.drop(index=list(ja_corrigido))

    # ------------------------------------------------------------------ decisão 52: pontos de compra que não valeu
    invalida = mov.venda_invalida.notna() & mov.tipo.isin(["C", "E"])
    for _, r in mov[invalida & (mov.tipo == "C")].iterrows():
        tem_estorno = ((mov.compra_id == r.compra_id) & (mov.tipo == "E")).any()
        reg.add(52, "Crédito de compra " + r.venda_invalida + (" (estorno manual também sai)" if tem_estorno else ""),
                r.origem, "qtd", r.qtd_corrigida, 0, f"Venda {r.venda_id}: {r.venda_status}")
    impacto[52] = -int(mov.loc[invalida & (mov.tipo == "C"), "qtd_corrigida"].sum() +
                       mov.loc[invalida & (mov.tipo == "E"), "qtd_corrigida"].sum())
    mov = mov[~invalida]

    # ------------------------------------------------------------------ decisão 53: online vale dobro
    cred = mov.tipo == "C"
    razao = mov.qtd_corrigida / mov.valor_reais.where(mov.valor_reais >= 1)
    simples = cred & mov.online.fillna(False) & razao.between(0.8, 1.2)
    dobro = mov[simples].assign(tipo="DOBRO_DEVIDO", obs="dobro online não creditado (decisão 53)")
    for _, r in dobro.iterrows():
        reg.add(53, "Compra online creditada em dobro só pela metade", r.origem, "qtd", r.qtd_corrigida,
                2 * r.qtd_corrigida, f"Compra de R$ {r.valor_reais:.2f}")

    # ------------------------------------------------------------------ extrato corrigido
    nomes = {"C": "CREDITO", "D": "RESGATE", "E": "ESTORNO", "AJ": "AJUSTE"}
    sinal = {"C": 1, "D": -1, "E": 1, "AJ": 1}                # E já vem negativo; D vem positivo
    linhas = [dict(cpf=r.cpf, data=r.data, tipo=nomes[r.tipo], qtd=int(sinal[r.tipo] * r.qtd_corrigida),
                   origem=r.origem, compra_id=r.compra_id, obs=r.obs) for r in mov.itertuples()]
    linhas += [dict(cpf=r.cpf, data=r.data, tipo="DOBRO_DEVIDO", qtd=int(r.qtd_corrigida), origem=r.origem,
                    compra_id=r.compra_id, obs=r.obs) for r in dobro.itertuples()]

    # saldo congelado da planilha (decisões 1 a 3 já aplicadas: CPF completo, duplicados somados, formato do número)
    for r in clientes.itertuples():
        if r.pontos:
            linhas.append(dict(cpf=r.cpf, data=INICIO_BANCO, tipo="SALDO_PLANILHA", qtd=int(r.pontos),
                               origem=f"planilha linha {r.linha_planilha}", compra_id=None,
                               obs="saldo congelado em 01/01/2025"))

    papel = resgates_em_papel(pasta, clientes, mov, reg)
    batel_cred = batel_sem_credito(batel, vendas, mov, reg)
    linhas += papel + batel_cred
    impacto.update({53: int(dobro.qtd_corrigida.sum()), 55: sum(x["qtd"] for x in papel),
                    58: sum(x["qtd"] for x in batel_cred)})
    extrato = pd.DataFrame(linhas).sort_values(["cpf", "data", "qtd"], ascending=[True, True, False])
    extrato = extrato[extrato.data < DATA_CORTE]

    # ------------------------------------------------------------------ categoria (decisão 4 revisada)
    ini_12m = mais_meses(DATA_CORTE, -12)
    ganho = extrato[extrato.tipo.isin(["CREDITO", "DOBRO_DEVIDO", "CREDITO_BATEL"]) & (extrato.data >= ini_12m)] \
        .groupby("cpf").qtd.sum()
    cats = {cpf: categoria(ganho.get(cpf, 0)) for cpf in extrato.cpf.unique()}

    # ------------------------------------------------------------------ vencimento e saldo (decisões 54 e 56)
    saldos, eventos = [], []
    for cpf, g in extrato.groupby("cpf", sort=False):
        ev = []
        saldos.append(dict(cpf=cpf, categoria=cats[cpf], acumulado_12m=int(ganho.get(cpf, 0)),
                           **simular(g, VALIDADE_MESES_OURO if cats[cpf] == "Ouro" else VALIDADE_MESES, ev)))
        eventos += [dict(e, cpf=cpf) for e in ev]
    saldos = pd.DataFrame(saldos)
    eventos = pd.DataFrame(eventos, columns=["cpf", "data", "tipo", "qtd"])
    for r in saldos[saldos.resgatado_sem_saldo > 0].itertuples():
        reg.add(56, "Resgate maior que o saldo", r.cpf, "saldo", -r.resgatado_sem_saldo, 0,
                "O desconto já foi dado; o saldo não fica negativo (regra do programa)")
    for r in saldos[saldos.vencidos > 0].itertuples():
        reg.add(54, f"Pontos vencidos ({VALIDADE_MESES_OURO if r.categoria == 'Ouro' else VALIDADE_MESES} meses)",
                r.cpf, "saldo", r.saldo_sem_vencimento, r.saldo, f"{r.vencidos} pontos venceram")

    # ------------------------------------------------------------------ o que o sistema da Vulto mostraria
    saldos = saldos.merge(saldo_do_sistema(pasta, clientes), on="cpf", how="outer").fillna(
        {"saldo_sistema": 0, "saldo": 0, "saldo_sem_vencimento": 0, "vencidos": 0, "resgatado_sem_saldo": 0,
         "acumulado_12m": 0, "categoria": "Bronze"})
    saldos["valor_reais"] = (saldos.saldo * VALOR_PONTO).round(2)
    risco_ouro = ganho[[c for c in ganho.index if cats.get(c) == "Ouro"]].sum() * 0.5
    reg.add(57, "Bônus Ouro (+50%) nunca aplicado: não creditado", "todas", "pontos", None, int(risco_ouro),
            "Quem era Ouro em cada compra não está registrado; valor fica como risco")
    reg.add(40, "Valor do ponto", "config.valor_ponto", "valor_ponto", VALOR_PONTO_SISTEMA, VALOR_PONTO,
            "Regra do programa desde mar/2025 (LEIA ISSO) e o que as lojas aplicam nos resgates em papel")
    impacto.update({54: -int(saldos.vencidos.sum()), 56: int(saldos.resgatado_sem_saldo.sum()), 57: int(risco_ouro)})
    return dict(extrato=extrato, saldos=saldos, ponte=ponte(extrato, saldos, mov, risco_ouro), categorias=cats,
                impacto=impacto, eventos=eventos), reg


def motivo_invalida(v):
    if v.teste:
        return "de teste"
    if not pd.isna(v.duplicada_de):
        return "duplicada"
    if v.tipo == "Troca":
        return "de troca"
    if v.status not in STATUS_VALIDO:
        return {"Cancelada": "cancelada", "Estornada": "estornada", "Chargeback": "com chargeback",
                "Pagamento não confirmado": "não paga"}.get(v.status, "não paga")
    return None


def resgates_em_papel(pasta, clientes, mov, reg):
    """Decisão 55: resgates anotados só na aba Resgates depois de 01/01/2025 (antes disso já estão no saldo
    congelado). Cliente identificado pelo CPF ou por nome completo único no cadastro."""
    r = pd.read_excel(os.path.join(pasta, ARQUIVO), sheet_name="Resgates", dtype=str)
    por_nome = clientes.assign(k=clientes.nome.map(chave)).drop_duplicates("k", keep=False).set_index("k").cpf
    d_banco = mov[mov.tipo == "D"]
    linhas = []
    for i, x in r.iterrows():
        d = data_resgate(x["data"])
        if d is None or d < INICIO_BANCO:
            continue
        origem = f"Resgates!{i + 2}"
        pts = int(dinheiro(x["pontos usados"]) or 0)
        dig = digitos(x["cliente"])
        cpf = dig.zfill(11) if len(dig) >= 9 else por_nome.get(chave(x["cliente"]))
        if cpf is None:
            reg.add(55, "Resgate em papel sem como identificar o cliente", origem, "cliente", x["cliente"], None,
                    f"{pts} pontos em {d:%d/%m/%Y} ficam fora (só primeiro nome)")
            continue
        repetido = d_banco[(d_banco.cpf == cpf) & (d_banco.qtd_corrigida == pts) &
                           ((d_banco.data - d).abs().dt.days <= 31)]
        if len(repetido):
            reg.add(55, "Resgate em papel que também está no sistema", origem, "cliente", x["cliente"], cpf)
            continue
        linhas.append(dict(cpf=cpf, data=d, tipo="RESGATE_PAPEL", qtd=-pts, origem=origem, compra_id=None,
                           obs=f"{x['loja']} / {x['atendente']}"))
        reg.add(55, "Resgate feito só no papel: descontado do saldo", origem, "pontos", None, -pts,
                f"{x['cliente']} em {d:%d/%m/%Y}")
    return linhas


def batel_sem_credito(batel, vendas, mov, reg):
    """Decisão 58: vendas da Batel feitas sem sistema, com CPF, que a loja não lançou no Vulto+ (decisão 20)."""
    creditadas = set(mov.loc[mov.tipo == "C", "compra_id"].dropna())
    relancadas = vendas[vendas.linha_batel.notna()].set_index("linha_batel")
    linhas = []
    for m in batel[(batel.lancado_vulto == "Pendente")].itertuples():
        if m.linha_planilha in relancadas.index:
            v = relancadas.loc[m.linha_planilha]
            if not pd.isna(v.compra_id) and v.compra_id in creditadas:
                continue
        cpf = digitos(m.cpf).zfill(11)
        dh = datetime.combine(pd.Timestamp(m.data).date(), datetime.min.time())
        linhas.append(dict(cpf=cpf, data=dh, tipo="CREDITO_BATEL", qtd=int(round(float(m.valor))),
                           origem=f"Batel!{m.linha_planilha}", compra_id=None, obs="venda sem sistema (mar/26)"))
        reg.add(58, "Pontos da venda sem sistema creditados", f"Batel!{m.linha_planilha}", "pontos", None,
                int(round(float(m.valor))), f"{m.cpf}, R$ {float(m.valor):.2f}")
    return linhas


def simular(g, validade, eventos=None):
    """Saldo em 01/10/2026 com vencimento: cada crédito vira um lote que vence 'validade' meses depois; resgates,
    estornos e ajustes negativos consomem os lotes mais antigos primeiro (FIFO).

    Se 'eventos' for uma lista, recebe os lançamentos que fecham o extrato com o saldo: VENCIMENTO (negativo, na
    data em que o lote venceu) e SALDO_ZERADO (positivo, quando um resgate passou do saldo)."""
    lotes = deque()                                    # [vence_em, pontos]
    vencidos = sem_saldo = 0
    saldo_bruto = 0

    def vencer(ate):
        nonlocal vencidos
        while lotes and lotes[0][0] <= ate:
            vence_em, pts = lotes.popleft()
            vencidos += pts
            if eventos is not None and pts:
                eventos.append(dict(data=vence_em, tipo="VENCIMENTO", qtd=-pts))

    for r in g.itertuples():
        saldo_bruto += r.qtd
        vencer(r.data)
        if r.qtd > 0:
            base = r.data
            lotes.append([mais_meses(base, validade), r.qtd])
            continue
        falta = -r.qtd
        while falta and lotes:
            usa = min(falta, lotes[0][1])
            lotes[0][1] -= usa
            falta -= usa
            if lotes[0][1] == 0:
                lotes.popleft()
        sem_saldo += falta
        if eventos is not None and falta:
            eventos.append(dict(data=r.data, tipo="SALDO_ZERADO", qtd=falta))
    vencer(DATA_CORTE)
    return dict(saldo=int(sum(p for _, p in lotes)), saldo_sem_vencimento=int(max(saldo_bruto, 0)),
                vencidos=int(vencidos), resgatado_sem_saldo=int(sem_saldo))


def saldo_do_sistema(pasta, clientes):
    """O que a tela do sistema da Vulto mostraria para cada CPF (calcula_saldo em main.py): pontos da planilha
    (primeira linha com o CPF, lida pelo parse antigo) + C + AJ − D + E, sem vencer nada."""
    bruto = pd.read_excel(os.path.join(pasta, ARQUIVO), sheet_name="Clientes", dtype=object)
    bruto["cpf_limpo"] = bruto.cpf.map(lambda x: str(x).replace(".", "").replace("-", ""))
    planilha = bruto.drop_duplicates("cpf_limpo").set_index("cpf_limpo").Pontos.map(parse_pontos_sistema)
    with sqlite3.connect(os.path.join(pasta, "vulto.db")) as c:
        m = pd.read_sql("select cpf, tipo, qtd from pontos_mov", c)
    m["s"] = m.qtd.where(m.tipo != "D", -m.qtd)
    banco = m.groupby("cpf").s.sum()
    todos = set(planilha.index) | set(banco.index)
    out = pd.DataFrame({"cpf_sistema": sorted(todos)})
    out["saldo_sistema"] = [int(planilha.get(c, 0)) + int(banco.get(c, 0)) for c in out.cpf_sistema]
    out["cpf"] = out.cpf_sistema.map(lambda x: digitos(x).zfill(11))
    return out.groupby("cpf", as_index=False).saldo_sistema.sum()


def ponte(extrato, saldos, mov, risco_ouro):
    """Do saldo que o sistema mostra até o passivo, uma decisão por vez."""
    soma = lambda tipos: int(extrato[extrato.tipo.isin(tipos)].qtd.sum())
    sistema = int(saldos.saldo_sistema.clip(lower=0).sum())
    passos = [dict(passo="Saldo que o sistema da Vulto mostra (soma dos clientes com saldo positivo)", decisao=None,
                   pontos=sistema, reais_0_05=None, reais_0_10=round(sistema * VALOR_PONTO_SISTEMA, 2))]
    sem_venc = int(saldos.saldo_sem_vencimento.sum())
    passos.append(dict(passo="Corrigido sem vencimento (centavos, compras canceladas, dobro online, resgates em "
                             "papel, Batel)", decisao="51 a 55, 58", pontos=sem_venc,
                       reais_0_05=round(sem_venc * VALOR_PONTO, 2), reais_0_10=round(sem_venc * VALOR_PONTO_SISTEMA, 2)))
    vencidos = int(saldos.vencidos.sum())
    passos.append(dict(passo="Menos pontos vencidos (12 meses; Ouro 18), gastando os mais antigos primeiro",
                       decisao=54, pontos=-vencidos, reais_0_05=None, reais_0_10=None))
    final = int(saldos.saldo.sum())
    passos.append(dict(passo="Passivo de pontos em 01/10/2026", decisao=40, pontos=final,
                       reais_0_05=round(final * VALOR_PONTO, 2), reais_0_10=round(final * VALOR_PONTO_SISTEMA, 2)))
    passos.append(dict(passo="Risco não incluído: bônus Ouro (+50%) prometido e nunca creditado", decisao=57,
                       pontos=int(risco_ouro), reais_0_05=round(risco_ouro * VALOR_PONTO, 2),
                       reais_0_10=round(risco_ouro * VALOR_PONTO_SISTEMA, 2)))
    detalhe = (extrato.groupby("tipo").qtd.agg(["count", "sum"]).reset_index()
               .rename(columns={"count": "movimentos", "sum": "pontos"}))
    return dict(ponte=pd.DataFrame(passos), por_tipo=detalhe)
