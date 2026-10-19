"""Clientes do Vulto+ corrigidos e saldo de pontos de cada um em 01/10/2026.

Entradas (só leitura):
- clientes_fidelidade_FINAL_v3 (1).xlsx (abas Clientes e Resgates) e vulto.db (pontos_mov, compras, vendas)
- clientes atualizado.xlsx (limpeza do time: nome, e-mail e categoria de cada CPF)
- vulto_padronizado.db, tabela vendas_padronizadas (status de cada venda; rode padroniza_vendas.py antes)

Saída: migracao/saida/"clientes atualizado v2.xlsx", com as abas Clientes, Extrato, Resumo,
Regras Vulto+, Decisões e Pendências.

Regras (decididas pelo time em 15/10/2026):
- CPF é texto com 11 dígitos: o zero da frente que o Excel cortou volta antes de validar.
- Cadastros com o mesmo CPF viram um só e os pontos da planilha são somados (sempre).
- aceita_msg: se algum cadastro diz Sim, é Sim; se não disse nada ou disse Não, é Não.
- Validade: só valem os pontos ganhos nos últimos 12 meses (de 01/10/2025 a 01/10/2026). Os resgates
  gastam primeiro os pontos mais antigos.
- Categoria pelo saldo válido em 01/10/2026: 0 sem categoria, 1 a 1.500 Bronze, 1.501 a 5.000 Prata,
  5.001 a 10.000 Ouro, acima de 10.000 Vip.

Uso: python migracao/clientes_pontos.py
"""

import os
import re
import sqlite3
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import caminhos
import registro

PLANILHA = os.path.join(caminhos.DADOS, "clientes_fidelidade_FINAL_v3 (1).xlsx")
LEGADO = caminhos.VULTO_DB
DO_TIME = os.path.join(caminhos.ENTRADA, "clientes atualizado.xlsx")
PADRONIZADO = caminhos.PADRONIZADO
SAIDA = os.path.join(caminhos.SAIDA, "clientes atualizado v2.xlsx")

BRASILIA = ZoneInfo("America/Sao_Paulo")
CORTE = datetime(2026, 10, 1, tzinfo=BRASILIA)
INICIO_VALIDADE = datetime(2025, 10, 1, tzinfo=BRASILIA)
CONGELAMENTO = datetime(2025, 1, 1, tzinfo=BRASILIA)  # a coluna Pontos da planilha foi congelada aqui
VALOR_PONTO = 0.05  # 100 pontos = R$ 5 (regra desde mar/2025)
FAIXAS = ((10001, "Vip"), (5001, "Ouro"), (1501, "Prata"), (1, "Bronze"), (0, "Sem categoria"))

LOJAS = {
    "site": "Site", "ecommerce": "Site", "app": "App", "pinheiros": "Pinheiros", "eldorado": "Eldorado",
    "eldorado shop": "Eldorado", "savassi": "Savassi", "bh": "Savassi", "batel": "Batel", "cwb": "Batel",
}
STATUS_SEM_PONTO = {"Cancelada", "Estornada", "Chargeback"}
MESES = {m: i for i, m in enumerate(["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"], 1)}
PARTICULAS = {"de", "da", "do", "das", "dos", "e"}


# ---------------------------------------------------------------- leitura e normalização


def texto(valor) -> str:
    """Célula como texto ('' quando vazia)."""
    return "" if valor is None or (isinstance(valor, float) and valor != valor) else str(valor).strip()


def so_digitos(valor) -> str:
    if valor is None or (isinstance(valor, float) and valor != valor):
        return ""
    if isinstance(valor, float):
        valor = int(valor)
    return re.sub(r"\D", "", str(valor))


def cpf_11(valor) -> str:
    """CPF com 11 dígitos: devolve os zeros da frente que o Excel cortou."""
    digitos = so_digitos(valor)
    return digitos.zfill(11) if 9 <= len(digitos) <= 11 else digitos


def cpf_valido(cpf: str) -> bool:
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False
    for n in (9, 10):
        if sum(int(cpf[i]) * (n + 1 - i) for i in range(n)) * 10 % 11 % 10 != int(cpf[n]):
            return False
    return True


def cpf_formatado(cpf: str) -> str:
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}" if len(cpf) == 11 else cpf


def pontos(valor) -> int:
    """'2.203', '1389 pts', '2,7k', 2182 -> inteiro."""
    if valor is None or (isinstance(valor, float) and valor != valor):
        return 0
    if isinstance(valor, (int, float)):
        return int(valor)
    texto = str(valor).lower().replace("pts", "").replace(" ", "")
    if not texto:
        return 0
    if texto.endswith("k"):
        return round(float(texto[:-1].replace(",", ".")) * 1000)
    return int(texto.replace(".", ""))


def data_cadastro(valor) -> date | None:
    """Data da planilha: datetime, '01/06/2021', '2021-06-01', 44351 (série do Excel) ou 'ago/25'."""
    if valor is None or (isinstance(valor, float) and valor != valor):
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=int(valor))).date()
    texto = str(valor).strip()
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            pass
    m = re.fullmatch(r"([a-z]{3})/(\d{2})", texto.lower())
    if m and m.group(1) in MESES:
        return date(2000 + int(m.group(2)), MESES[m.group(1)], 1)
    return None


def momento_legado(texto: str) -> datetime:
    """Data do pontos_mov: 'Z' no fim = UTC; sem fuso = Brasília."""
    texto = texto.strip()
    for formato in ("%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M"):
        try:
            momento = datetime.strptime(texto, formato)
        except ValueError:
            continue
        if formato.endswith("Z"):
            return momento.replace(tzinfo=timezone.utc).astimezone(BRASILIA)
        return momento.replace(tzinfo=BRASILIA)
    raise ValueError(texto)


def sem_acento(texto: str) -> str:
    return " ".join(unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode().lower().split())


def nome_bonito(texto) -> str:
    palavras = str(texto).split()
    return " ".join(p.lower() if p.lower() in PARTICULAS and i else p[:1].upper() + p[1:].lower() for i, p in enumerate(palavras))


def aceita(valor) -> str:
    return "Sim" if str(valor).strip().lower() in ("1", "sim", "s", "x") else "Não"


def telefone(valor) -> str:
    digitos = so_digitos(valor)
    if digitos.startswith("55") and len(digitos) in (12, 13):
        digitos = digitos[2:]
    return digitos


def categoria_escrita(valores) -> str:
    """Categoria escrita à mão na planilha (bronze, brz, Silver, ouro ★, vip?...), no padrão do arquivo do time."""
    for valor in valores:
        nome = sem_acento(texto(valor))
        for chave, padrao in (("vip", "Vip"), ("ouro", "Ouro"), ("gold", "Ouro"), ("prata", "Prata"), ("silver", "Prata"),
                              ("bronze", "Bronze"), ("brz", "Bronze")):
            if chave in nome:
                return padrao
    return ""


def categoria(saldo: int) -> str:
    return next(nome for minimo, nome in FAIXAS if saldo >= minimo)


# ---------------------------------------------------------------- clientes


def montar_clientes():
    bruto = pd.read_excel(PLANILHA, sheet_name="Clientes", dtype=object)
    time = pd.read_excel(DO_TIME, dtype=object)
    time["cpf"] = time["CPF"].map(cpf_11)
    do_time = time.set_index("cpf")

    bruto["cpf"] = bruto["cpf"].map(cpf_11)
    bruto["linha"] = bruto.index + 2
    # CPF com dígito verificador inválido fica (marcado), decisão do time em 15/10/2026; só sai quem não tem CPF
    invalidos = bruto[bruto["cpf"] == ""]
    com_cpf = bruto[bruto["cpf"] != ""].copy()
    com_cpf["data"] = com_cpf["Data Cadastro"].map(data_cadastro)

    clientes = {}
    for cpf, grupo in com_cpf.groupby("cpf", sort=False):
        grupo = grupo.sort_values("data", na_position="last")
        primeiro, ultimo = grupo.iloc[0], grupo.iloc[-1]
        no_time = cpf in do_time.index
        loja = texto(primeiro["loja_cadastro"]).lower()
        clientes[cpf] = {
            "cpf": cpf,
            "nome": texto(do_time.at[cpf, "Nome"]) if no_time else nome_bonito(texto(primeiro["Nome"])),
            "email": texto(do_time.at[cpf, "E-mail"] if no_time else primeiro["E-mail"]).lower(),
            "telefone": telefone(next((t for t in reversed(grupo["Tel"].tolist()) if so_digitos(t)), "")),
            "cadastro": grupo["data"].dropna().min() if grupo["data"].notna().any() else None,
            "loja": LOJAS.get(loja, loja.title()),
            "aceita": "Sim" if (grupo["aceita_msg"].map(aceita) == "Sim").any() else "Não",
            "fundidos": len(grupo),
            "pontos_planilha": int(grupo["Pontos"].map(pontos).sum()),
            "categoria_planilha": texto(do_time.at[cpf, "Categoria"]) if no_time else categoria_escrita(grupo["categoria"]),
            "origem": ("Planilha" if no_time else "Planilha (estava fora do arquivo do time: CPF com zero na frente)")
            + ("" if cpf_valido(cpf) else " · CPF inválido"),
            "linhas": ", ".join(str(l) for l in grupo["linha"]),
            "obs": " · ".join(sorted({str(o) for o in grupo["Obs"].dropna()})),
        }
    return clientes, invalidos, len(set(com_cpf["cpf"]) - set(do_time.index))


# ---------------------------------------------------------------- extrato de pontos


def montar_extrato(clientes):
    legado = sqlite3.connect(LEGADO)
    movimentos = pd.read_sql("select * from pontos_mov order by id", legado)
    compras = pd.read_sql("select id, valor from compras", legado).set_index("id")
    vendas = pd.read_sql("select id, compra_id, total, canal from vendas", legado)
    padrao = pd.read_sql(
        "select id_venda, status, tipo, duplicada_de from vendas_padronizadas", sqlite3.connect(PADRONIZADO)
    ).set_index("id_venda")
    venda_da_compra = vendas.dropna(subset=["compra_id"]).set_index(vendas.dropna(subset=["compra_id"])["compra_id"].astype(int))
    estornadas = set(movimentos.loc[movimentos["tipo"] == "E", "compra_id"].dropna().astype(int))

    def compra_em_centavos(compra_id) -> bool:
        if compra_id not in venda_da_compra.index:
            return False
        v = venda_da_compra.loc[compra_id]
        return str(v["canal"]).strip().lower() in ("app", "online") and re.fullmatch(r"-?\d+", str(v["total"]).strip()) is not None

    extrato = []
    for cpf, c in clientes.items():
        if c["pontos_planilha"]:
            extrato.append(dict(cpf=cpf, quando=CONGELAMENTO, tipo="Saldo da planilha", original=c["pontos_planilha"],
                                considerado=c["pontos_planilha"], situacao="Mantido",
                                motivo="Coluna Pontos congelada em 01/01/2025 (soma dos cadastros fundidos)",
                                origem=f"clientes.xlsx:Clientes linhas {c['linhas']}", ref=""))

    bug_corrigido = set()
    for mv in movimentos.itertuples():
        cpf = cpf_11(mv.cpf)
        quando = momento_legado(mv.criado_em)
        sinal = -1 if mv.tipo == "D" else 1
        original = sinal * int(mv.qtd)
        tipo = {"C": "Crédito", "D": "Resgate", "E": "Estorno", "AJ": "Ajuste"}[mv.tipo]
        compra_id = int(mv.compra_id) if pd.notna(mv.compra_id) else None
        considerado, situacao, motivo = original, "Mantido", ""
        venda = venda_da_compra.loc[compra_id] if compra_id in venda_da_compra.index else None
        estado = padrao.loc[int(venda["id"])] if venda is not None else None

        if mv.tipo == "AJ" and "bug app" in str(mv.obs or ""):
            considerado, situacao = 0, "Descartado"
            motivo = "Correção manual aproximada do bug do app: substituída pela correção de cada crédito (/100)"
        elif mv.tipo == "C" and estado is not None and estado["tipo"] == "Teste":
            considerado, situacao, motivo = 0, "Descartado", "Crédito de venda de teste"
        elif mv.tipo == "C" and estado is not None and estado["tipo"] == "Troca":
            considerado, situacao = 0, "Descartado"
            motivo = "Troca: os pontos já foram ganhos na compra original (decisão 16)"
        elif mv.tipo == "C" and estado is not None and estado["duplicada_de"] not in ("", None):
            considerado, situacao = 0, "Descartado"
            motivo = f"Pedido lançado duas vezes: o crédito vale só na venda {estado['duplicada_de']}"
        elif mv.tipo == "C" and estado is not None and estado["status"] == "Aguardando pagamento":
            considerado, situacao, motivo = 0, "Descartado", "Venda nunca paga: ponto só depois do pagamento aprovado"
        elif mv.tipo == "C" and estado is not None and estado["status"] in STATUS_SEM_PONTO and compra_id not in estornadas:
            considerado, situacao = 0, "Descartado"
            motivo = f"Venda {estado['status'].lower()} sem o estorno manual: quem cancela perde os pontos da compra"
        elif mv.tipo in ("C", "E") and compra_id is not None and compra_em_centavos(compra_id):
            reais = float(str(compras.at[compra_id, "valor"]).replace(",", ".")) / 100
            if abs(original) >= reais * 50:  # o app creditou sobre o valor em centavos (×100)
                considerado, situacao = round(original / 100), "Corrigido"
                motivo = "App gravou o valor em centavos e creditou ×100: dividido por 100"
                bug_corrigido.add(cpf)
        extrato.append(dict(cpf=cpf, quando=quando, tipo=tipo, original=original, considerado=considerado, situacao=situacao,
                            motivo=motivo or (str(mv.obs) if mv.obs else ""), origem=f"vulto.db:pontos_mov id {mv.id}",
                            ref=f"compra {compra_id}" + (f" · venda {int(venda['id'])}" if venda is not None else "") if compra_id else ""))
    return extrato, movimentos, bug_corrigido


def resgates_da_planilha(clientes, extrato):
    """Resgate anotado na aba Resgates. Antes de 2025 já está no saldo congelado. De 2025 em diante: se o
    mesmo resgate (CPF e pontos, até 45 dias de diferença) está no sistema, não conta de novo; senão é debitado.
    Só com o nome: procura um cliente com esse nome; sem um único resultado, fica pendente."""
    rs = pd.read_excel(PLANILHA, sheet_name="Resgates", dtype=object)
    por_nome = defaultdict(set)
    for cpf, c in clientes.items():
        por_nome[sem_acento(c["nome"])].add(cpf)
    resgates_sistema = defaultdict(list)
    for e in extrato:
        if e["tipo"] == "Resgate":
            resgates_sistema[e["cpf"]].append(e)

    novos, pendentes, contagem = [], [], Counter()
    for linha, r in zip(rs.index + 2, rs.itertuples()):
        quando = data_cadastro(r.data)
        qtd = pontos(getattr(r, "_3"))
        if quando is None or qtd <= 0:
            pendentes.append((linha, r.cliente, r.data, qtd, "Data ou pontos ilegíveis"))
            contagem["ilegível"] += 1
            continue
        if quando < date(2025, 1, 1):
            contagem["antes de 2025 (já no saldo congelado)"] += 1
            continue
        cpf = cpf_11(r.cliente) if re.search(r"\d{3}", str(r.cliente)) else ""
        if not cpf:
            achados = por_nome.get(sem_acento(r.cliente), set())
            cpf = next(iter(achados)) if len(achados) == 1 else ""
        if not cpf or cpf not in clientes:
            pendentes.append((linha, r.cliente, r.data, qtd, "Cliente não identificado (só o nome, ou CPF fora do cadastro)"))
            contagem["sem cliente identificado"] += 1
            continue
        momento = datetime(quando.year, quando.month, quando.day, 12, tzinfo=BRASILIA)
        par = next((e for e in resgates_sistema[cpf] if -e["original"] == qtd and abs((e["quando"] - momento).days) <= 45
                    and not e.get("pareado")), None)
        if par is not None:
            par["pareado"] = True
            contagem["já no sistema"] += 1
            continue
        contagem["só na planilha (debitado)"] += 1
        novos.append(dict(cpf=cpf, quando=momento, tipo="Resgate (só na planilha)", original=-qtd, considerado=-qtd,
                          situacao="Mantido", motivo=f"Anotado pela loja {r.loja} e não lançado no sistema",
                          origem=f"clientes.xlsx:Resgates linha {linha}", ref=""))
    return novos, pendentes, contagem


def saldo_em(movimentos_do_cliente):
    """(saldo sem vencimento, saldo válido, vencidos, faltou). Os resgates gastam primeiro os pontos mais antigos;
    no corte, vencem os pontos ganhos antes de 01/10/2025."""
    lotes, faltou = [], 0
    for e in sorted(movimentos_do_cliente, key=lambda e: (e["quando"], e["considerado"] < 0)):
        if e["quando"] >= CORTE or not e["considerado"]:
            continue
        if e["considerado"] > 0:
            lotes.append([e["quando"], e["considerado"]])
            continue
        gastar = -e["considerado"]
        for lote in lotes:
            usa = min(gastar, lote[1])
            lote[1] -= usa
            gastar -= usa
            if not gastar:
                break
        faltou += gastar
    total = sum(l[1] for l in lotes)
    valido = sum(l[1] for l in lotes if l[0] >= INICIO_VALIDADE)
    return total, valido, total - valido, faltou


# ---------------------------------------------------------------- planilha de saída


def escrever(clientes, extrato, invalidos, pendentes, resumo, decisoes):
    wb = Workbook()
    titulo = Font(bold=True, size=13)
    negrito = Font(bold=True)
    cabecalho = PatternFill("solid", fgColor="1F3A5F")

    def aba(nome, colunas, linhas, larguras, texto=()):
        ws = wb.create_sheet(nome)
        ws.append(colunas)
        for celula in ws[1]:
            celula.font = Font(bold=True, color="FFFFFF")
            celula.fill = cabecalho
            celula.alignment = Alignment(vertical="center", wrap_text=True)
        for linha in linhas:
            ws.append(linha)
        for i, largura in enumerate(larguras, 1):
            ws.column_dimensions[get_column_letter(i)].width = largura
        for i in texto:  # colunas guardadas como texto (CPF e telefone não perdem o zero)
            for celula in ws.iter_cols(min_col=i, max_col=i, min_row=2):
                for c in celula:
                    c.number_format = "@"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        return ws

    wb.remove(wb.active)
    aba(
        "Clientes",
        ["CPF", "Nome", "E-mail", "Telefone", "Data de cadastro", "Loja de cadastro", "Aceita mensagem",
         "Cadastros fundidos", "Pontos na planilha (01/01/2025)", "Saldo 01/10/2026 sem vencimento",
         "Pontos vencidos", "Saldo válido 01/10/2026", "Valor (R$)", "Categoria", "Categoria na planilha", "Origem", "Obs"],
        [[cpf_formatado(c["cpf"]), c["nome"], c["email"], c["telefone"], c["cadastro"], c["loja"], c["aceita"], c["fundidos"],
          c["pontos_planilha"], c["total"], c["vencidos"], c["valido"], round(c["valido"] * VALOR_PONTO, 2), c["categoria"],
          c["categoria_planilha"], c["origem"], c["obs"]]
         for c in sorted(clientes.values(), key=lambda c: (not c["fundidos"], sem_acento(c["nome"]), c["cpf"]))],
        [16, 30, 32, 14, 14, 14, 10, 10, 14, 14, 12, 14, 12, 14, 14, 34, 40],
        texto=(1, 4),
    )
    for linha in wb["Clientes"].iter_rows(min_row=2, min_col=5, max_col=5):
        linha[0].number_format = "DD/MM/YYYY"
    for linha in wb["Clientes"].iter_rows(min_row=2, min_col=13, max_col=13):
        linha[0].number_format = '"R$" #,##0.00'

    nomes = {cpf: c["nome"] for cpf, c in clientes.items()}
    aba(
        "Extrato",
        ["CPF", "Nome", "Data", "Tipo", "Pontos (original)", "Pontos (considerado)", "Situação", "Motivo", "Origem", "Compra / venda"],
        [[cpf_formatado(e["cpf"]), nomes.get(e["cpf"], ""), e["quando"].replace(tzinfo=None), e["tipo"], e["original"],
          e["considerado"], e["situacao"], e["motivo"], e["origem"], e["ref"]]
         for e in sorted(extrato, key=lambda e: (e["cpf"], e["quando"]))],
        [16, 28, 17, 22, 12, 12, 12, 60, 34, 22],
        texto=(1,),
    )
    for linha in wb["Extrato"].iter_rows(min_row=2, min_col=3, max_col=3):
        linha[0].number_format = "DD/MM/YYYY HH:MM"

    ws = wb.create_sheet("Resumo", 0)
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.append(["Passivo de pontos Vulto+ em 01/10/2026"])
    ws["A1"].font = titulo
    for linha in resumo:
        ws.append(linha)
        if len(linha) == 1 and linha[0]:
            ws.cell(ws.max_row, 1).font = negrito

    ws = wb.create_sheet("Regras Vulto+")
    ws.column_dimensions["A"].width = 120
    for linha in REGRAS:
        ws.append([linha])
    ws["A1"].font = titulo

    aba("Decisões", registro.COLUNAS, decisoes, registro.LARGURAS)
    for ws in (wb["Decisões"],):
        for linha in ws.iter_rows(min_row=2):
            for c in linha:
                c.alignment = Alignment(wrap_text=True, vertical="top")

    aba(
        "Pendências",
        ["Origem", "Linha", "Cliente / CPF", "Data", "Pontos", "Por que ficou de fora"],
        [["clientes.xlsx:Clientes", int(r.linha), f"{r.Nome} · {r.cpf or 'sem CPF'}", str(r["Data Cadastro"] or ""), pontos(r.Pontos),
          "Sem CPF: não dá para identificar o cliente (pontos só com CPF)"] for _, r in invalidos.iterrows()]
        + [["clientes.xlsx:Resgates", linha, str(cliente), str(data), qtd, motivo] for linha, cliente, data, qtd, motivo in pendentes],
        [22, 8, 40, 20, 10, 60],
    )
    wb.save(SAIDA)


def pendencias(clientes, extrato, por_cliente, reais):
    """Impacto das decisões ainda em aberto (40 e 41), sem aplicar: quanto o passivo oficial subiria."""
    legado = sqlite3.connect(LEGADO)
    compras = pd.read_sql("select id, valor, origem from compras", legado).set_index("id")
    vendas = pd.read_sql("select id, compra_id, total, canal from vendas", legado).dropna(subset=["compra_id"])
    centavos = {int(c) for c, t, canal in zip(vendas.compra_id, vendas.total, vendas.canal)
                if str(canal).strip().lower() in ("app", "online") and re.fullmatch(r"-?\d+", str(t).strip())}

    # 40: compra no site/app creditada uma vez só (o código só dobrava se o canal fosse exatamente "site" ou "app")
    dobro = []
    for e in extrato:
        m = re.match(r"compra (\d+)", e["ref"])
        if e["tipo"] != "Crédito" or e["situacao"] == "Descartado" or not m or int(m.group(1)) not in compras.index:
            continue
        c = compras.loc[int(m.group(1))]
        if c.origem not in ("site", "app"):
            continue
        valor = float(str(c.valor).replace(",", ".")) / (100 if int(m.group(1)) in centavos else 1)
        if valor >= 1 and 0.8 <= e["considerado"] / valor <= 1.2:
            dobro.append(dict(e, tipo="Crédito (dobro online)", original=0, situacao="Pendente"))

    # 41: venda manual da Batel com CPF, concluída e não lançada no Vulto+
    import batel
    planilha, _ = batel.ler()
    batel_pts = [dict(cpf=r.cpf, quando=datetime.combine(r.momento.date(), datetime.min.time(), tzinfo=BRASILIA),
                      tipo="Crédito (venda manual Batel)", original=0, considerado=int(round(r.valor)), situacao="Pendente",
                      motivo="", origem=f"planilha Batel Nº {r.numero}", ref="")
                 for r in planilha.itertuples() if r.cpf and r.Status == "Concluída" and r[13] == "Não"]

    def valido_com(extra):
        mais = defaultdict(list)
        for e in extra:
            mais[e["cpf"]].append(e)
        return sum(saldo_em(por_cliente.get(cpf, []) + mais.get(cpf, []))[1] - c["valido"]
                   for cpf, c in clientes.items() if cpf in mais) + sum(
            saldo_em(mais[cpf])[1] for cpf in mais if cpf not in clientes)

    d, b = valido_com(dobro), valido_com(batel_pts)
    return [
        [],
        ["Decisões em aberto (não aplicadas no passivo oficial)", "Pontos a mais no passivo", "R$"],
        [f"40. Dobro online não creditado ({len(dobro)} compras; {sum(e['considerado'] for e in dobro)} pontos no total)", d, reais(d)],
        [f"41. Vendas manuais da Batel sem pontos ({len(batel_pts)} vendas com CPF)", b, reais(b)],
    ]


REGRAS = [
    "REGRAS DO PROGRAMA VULTO+ (atualizado em 15/10/2026 — time de tecnologia, PMI Casa Lorenzi)",
    "",
    "- Cada R$ 1 na loja física = 1 ponto; compras no site e no app valem pontos em dobro.",
    "- 100 pontos = R$ 5,00 de desconto (desde mar/2025; antes eram R$ 10).",
    "- Resgate a partir de 100 pontos, só em múltiplos de 100, e só em compras com pelo menos R$ 500 em produtos (sem o frete).",
    "- O saldo nunca fica negativo.",
    "- Pontos só com CPF (telefone e e-mail não valem).",
    "- Validade: só valem os pontos ganhos nos últimos 12 meses. Os resgates gastam primeiro os pontos mais antigos.",
    "- Ponto só é creditado depois do pagamento aprovado. Cliente que cancela a compra perde os pontos daquela compra.",
    "",
    "CATEGORIAS (pelo saldo válido de pontos)",
    "- 0 pontos: sem categoria",
    "- 1 a 1.500: Bronze",
    "- 1.501 a 5.000: Prata",
    "- 5.001 a 10.000: Ouro",
    "- Acima de 10.000: Vip (categoria nova, criada em 15/10/2026)",
    "",
    "CADASTRO",
    "- CPF é a chave do cliente e é guardado como texto com 11 dígitos (o zero da frente não pode sumir).",
    "- Dois cadastros com o mesmo CPF são a mesma pessoa: viram um só e os pontos são somados.",
    "- Mensagens (LGPD): se o cliente disse Sim em algum cadastro, aceita; se não disse nada ou disse Não, não aceita.",
]


# ---------------------------------------------------------------- principal


def main():
    clientes, invalidos, recuperados = montar_clientes()
    extrato, movimentos, bug_corrigido = montar_extrato(clientes)

    # Quem tem pontos no sistema mas não está na planilha de clientes também é credor
    so_no_sistema = {e["cpf"] for e in extrato} - set(clientes)
    for cpf in so_no_sistema:
        clientes[cpf] = dict(cpf=cpf, nome="(sem cadastro na planilha)", email="", telefone="", cadastro=None, loja="", aceita="Não",
                             fundidos=0, pontos_planilha=0, categoria_planilha="", linhas="",
                             origem="Só no sistema de pontos (vulto.db)" + ("" if cpf_valido(cpf) else " · CPF inválido"), obs="")
    novos, pendentes, contagem_resgates = resgates_da_planilha(clientes, extrato)
    extrato += novos

    por_cliente = defaultdict(list)
    for e in extrato:
        por_cliente[e["cpf"]].append(e)
    faltou_total = 0
    for cpf, c in clientes.items():
        c["total"], c["valido"], c["vencidos"], faltou = saldo_em(por_cliente[cpf])
        c["categoria"] = categoria(c["valido"])
        faltou_total += faltou

    def soma(chave, filtro=lambda c: True):
        return sum(c[chave] for c in clientes.values() if filtro(c))

    def efeito(situacao, trecho):
        return sum(e["considerado"] - e["original"] for e in extrato if e["situacao"] == situacao and trecho in e["motivo"])

    total, valido, vencidos = soma("total"), soma("valido"), soma("vencidos")
    sistema = sum(e["original"] for e in extrato if e["origem"].startswith("vulto.db")) + soma("pontos_planilha")
    por_categoria = Counter()
    for c in clientes.values():
        por_categoria[c["categoria"]] += c["valido"]
    qtd_categoria = Counter(c["categoria"] for c in clientes.values())
    reais = lambda p: round(p * VALOR_PONTO, 2)  # noqa: E731

    pendente_pontos = sum(p[3] for p in pendentes)
    resumo = [
        [],
        ["Número para a diretoria", "Pontos", "R$ (a R$ 0,05)"],
        ["Passivo oficial: saldo válido (só pontos dos últimos 12 meses)", valido, reais(valido)],
        ["Teto (sem aplicar o vencimento)", total, reais(total)],
        ["Pontos vencidos em 01/10/2026 (ganhos antes de 01/10/2025)", vencidos, reais(vencidos)],
        ["Resgates na planilha sem cliente identificado (podem reduzir o passivo)", pendente_pontos, reais(pendente_pontos)],
        [],
        ["Do saldo que o sistema da Vulto mostraria até o nosso número"],
        ["Planilha congelada + todos os movimentos como estão no sistema", sistema, reais(sistema)],
        ["Créditos do app em centavos (×100) divididos por 100", efeito("Corrigido", "centavos"), ""],
        ["Correção manual aproximada do bug do app, descartada", efeito("Descartado", "bug do app"), ""],
        ["Créditos de vendas canceladas/estornadas sem o estorno manual", efeito("Descartado", "sem o estorno"), ""],
        ["Créditos de vendas nunca pagas (aguardando)", efeito("Descartado", "nunca paga"), ""],
        ["Créditos de pedidos lançados duas vezes", efeito("Descartado", "duas vezes"), ""],
        ["Créditos de vendas de teste", efeito("Descartado", "teste"), ""],
        ["Créditos de trocas (a compra original já ganhou os pontos)", efeito("Descartado", "Troca"), ""],
        ["Resgates anotados só na planilha", sum(e["considerado"] for e in novos), ""],
        ["Resgates maiores que o saldo (o sistema deixava; o saldo para em zero)", faltou_total, ""],
        ["= Saldo sem vencimento", total, reais(total)],
        ["Pontos vencidos (validade de 12 meses)", -vencidos, ""],
        ["= Passivo oficial", valido, reais(valido)],
        [],
        ["Por categoria (saldo válido)", "Pontos", "Clientes"],
        *[[nome, por_categoria[nome], qtd_categoria[nome]] for _, nome in FAIXAS],
        [],
        ["Clientes"],
        ["Clientes no cadastro (um por CPF, depois da fusão)", len(clientes) - len(so_no_sistema), ""],
        ["dos quais recuperados (CPF com zero na frente, fora do arquivo do time)", recuperados, ""],
        ["Só no sistema de pontos, sem cadastro na planilha", len(so_no_sistema), ""],
        ["Clientes com CPF inválido (ficam, marcados na coluna Origem)", sum(not cpf_valido(c) for c in clientes), ""],
        ["  saldo válido deles (incluído no passivo oficial)", sum(c["valido"] for cpf, c in clientes.items() if not cpf_valido(cpf)),
         reais(sum(c["valido"] for cpf, c in clientes.items() if not cpf_valido(cpf)))],
        ["Linhas da planilha sem CPF (fora, ver Pendências)", len(invalidos), ""],
        ["Aceitam mensagem (LGPD)", sum(c["aceita"] == "Sim" for c in clientes.values()), ""],
        [],
        ["Aba Resgates da planilha"],
        *[[k, v, ""] for k, v in contagem_resgates.items()],
    ]

    resumo += pendencias(clientes, extrato, por_cliente, reais)
    decisoes = registro.decisoes(18, 26) + registro.decisoes(40, 41)
    escrever(clientes, extrato, invalidos, pendentes, resumo, decisoes)

    print(f"clientes: {len(clientes) - len(so_no_sistema)} no cadastro ({recuperados} recuperados) + {len(so_no_sistema)} só no sistema")
    print(f"passivo oficial: {valido} pontos = R$ {reais(valido)} | sem vencimento: {total} = R$ {reais(total)} | vencidos {vencidos}")
    print(f"sistema como está: {sistema} | categorias: {dict(qtd_categoria)}")
    print("resumo:")
    for linha in resumo[8:21]:
        print("  ", linha)
    print("resgates:", dict(contagem_resgates), "| gravado:", SAIDA)


if __name__ == "__main__":
    main()
