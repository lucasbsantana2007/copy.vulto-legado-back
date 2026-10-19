from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel
from typing import List, Optional
from datetime import datetime
import base64
import csv
import io
import os
import pandas as pd

from banco import query, execute, get_config, conn
from planilha import ler_clientes, busca_cliente, limpa_cpf, cpf_texto, parse_pontos, anota_resgate
from catalogo import ler_catalogo, busca_produto

app = FastAPI(title="Vulto Sistema (Vulto+ / Loja / Estoque)", version="2.3.1")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

CANAIS_ONLINE = ["site", "app"]


class LoginIn(BaseModel):
    cpf: str


class ResgateIn(BaseModel):
    cpf: str
    pontos: int
    loja: str = "Sistema"


class CompraIn(BaseModel):
    cliente_ref: str
    valor: float
    canal: str
    loja_id: int = 1


@app.get("/")
def root():
    return {"status": "ok", "versao": get_config("versao")}


@app.post("/login")
def login(dados: LoginIn):
    # cliente so precisa do cpf pra entrar (pedido da Bia, cliente esquece senha)
    c = busca_cliente(dados.cpf)
    if c is None:
        raise HTTPException(404, "Cliente nao encontrado")
    token = base64.b64encode(limpa_cpf(dados.cpf).encode()).decode()
    return {"token": token, "nome": c["Nome"]}


@app.get("/clientes")
def listar_clientes(q: str = "", page: int = 1):
    df = ler_clientes()
    res = []
    for i, row in df.iterrows():
        if q and q.lower() not in str(row["Nome"]).lower() and q not in str(row["cpf"]):
            continue
        res.append({
            "linha": i + 2,
            "nome": row["Nome"],
            "cpf": row["cpf"],
            "email": row["E-mail"],
            "telefone": row["Tel"],
            "categoria": row["categoria"],
            "pontos_planilha": row["Pontos"],
        })
    inicio = (page - 1) * 50
    return {"total": len(res), "page": page, "clientes": res[inicio:inicio + 50]}


def calcula_saldo(cpf):
    cpf = limpa_cpf(cpf)
    c = busca_cliente(cpf)
    saldo = 0
    if c is not None:
        saldo = parse_pontos(c["Pontos"])
    movs = query("select * from pontos_mov where cpf = ?", (cpf,))
    for m in movs:
        if m["tipo"] == "C" or m["tipo"] == "AJ":
            saldo += m["qtd"]
        elif m["tipo"] == "D":
            saldo -= m["qtd"]
        elif m["tipo"] == "E":
            saldo += m["qtd"]  # estorno ja vem negativo
    # TODO expirar pontos
    return saldo


@app.get("/cliente/{cpf}")
def get_cliente(cpf: str):
    c = busca_cliente(cpf)
    saldo = calcula_saldo(cpf)
    if c is None and saldo == 0:
        raise HTTPException(404, "Cliente nao encontrado")
    valor_ponto = float(get_config("valor_ponto", "0.10"))
    return {
        "nome": c["Nome"] if c else None,
        "email": c["E-mail"] if c else None,
        "categoria": c["categoria"] if c else None,
        "saldo": saldo,
        "valor_em_reais": round(saldo * valor_ponto, 2),
    }


@app.get("/extrato/{cpf}")
def extrato(cpf: str):
    cpf = limpa_cpf(cpf)
    return query("select * from pontos_mov where cpf = ? order by criado_em desc", (cpf,))


@app.post("/resgate")
def resgatar(dados: ResgateIn):
    cpf = limpa_cpf(dados.cpf)
    saldo = calcula_saldo(cpf)
    # if dados.pontos > saldo:
    #     raise HTTPException(400, "saldo insuficiente")
    valor_ponto = float(get_config("valor_ponto", "0.10"))
    valor = round(dados.pontos * valor_ponto, 2)
    agora = datetime.now().isoformat()
    execute("insert into pontos_mov (cpf, tipo, qtd, compra_id, criado_em, obs) values (?,?,?,?,?,?)",
            (cpf, "D", dados.pontos, None, agora, "sistema"))
    try:
        anota_resgate(datetime.now().strftime("%d/%m/%Y"), cpf, dados.pontos, valor, dados.loja)
    except Exception as e:
        print("nao conseguiu salvar na planilha", e)
    return {"ok": True, "desconto": valor, "saldo_novo": saldo - dados.pontos}


@app.post("/compra")
def nova_compra(dados: CompraIn):
    origem = "pdv"
    if dados.canal in CANAIS_ONLINE:
        origem = dados.canal
    compra_id = execute(
        "insert into compras (cliente_ref, valor, canal, loja_id, criado_em, origem) values (?,?,?,?,?,?)",
        (dados.cliente_ref, str(dados.valor), dados.canal, dados.loja_id, datetime.now().isoformat(), origem))
    pontos = dados.valor
    if dados.canal in CANAIS_ONLINE:
        pontos = dados.valor * 2
    pontos = round(pontos)
    # so credita se for cpf
    ref = limpa_cpf(dados.cliente_ref)
    if ref.isdigit():
        execute("insert into pontos_mov (cpf, tipo, qtd, compra_id, criado_em, obs) values (?,?,?,?,?,?)",
                (ref, "C", pontos, compra_id, datetime.now().isoformat(), None))
    return {"ok": True, "compra_id": compra_id, "pontos": pontos}


@app.get("/lojas")
def lojas():
    return query("select * from lojas")


@app.get("/exportar")
def exportar():
    # usado pelo financeiro no excel deles
    # padrão de todo CSV exportado: separador ";", UTF-8 com BOM (o Excel lê os acentos) e CPF como texto
    # 000.000.000-00 (o export antigo saía com o CPF em número e o Excel cortava o zero da frente: decisão 18)
    df = ler_clientes()
    saida = io.StringIO()
    escritor = csv.writer(saida, delimiter=";", lineterminator="\r\n")
    escritor.writerow(["nome", "cpf", "email", "categoria", "pontos"])
    for i, row in df.iterrows():
        escritor.writerow([str(row["Nome"]).strip(), cpf_texto(row["cpf"]), str(row["E-mail"]).strip(),
                           str(row["categoria"]).strip(), str(row["Pontos"]).strip()])
    nome = f"clientes_{datetime.now():%Y-%m-%d}.csv"
    return Response(content=saida.getvalue().encode("utf-8-sig"), media_type="text/csv; charset=utf-8",
                    headers={"Content-Disposition": f"attachment; filename={nome}"})


# @app.get("/admin/recalcular")
# def recalcular():
#     # tentativa de recalcular tudo, travou o servidor em producao. NAO USAR
#     pass


# ------------------------------------------------------------- login das lojas (tela ?admin=1)

class AdminLoginIn(BaseModel):
    codigo: str = ""
    senha: str = ""


@app.post("/admin/login")
def admin_login(dados: AdminLoginIn):
    # pode entrar com o codigo OU com a senha (pedido das lojas, ninguem lembra os dois)
    f = query("select * from funcionarios where (codigo = ? or senha = ?) and ativo = 1",
              (dados.codigo.strip(), dados.senha))
    if len(f) == 0:
        if dados.senha and dados.senha == get_config("senha_master"):
            # senha do TI pra emergencia
            return {"token": base64.b64encode(b"master:0").decode(), "nome": "MASTER", "perfil": "admin", "loja_id": None}
        raise HTTPException(401, "Codigo ou senha invalidos")
    f = f[0]
    token = base64.b64encode(f"{f['codigo']}:{f['loja_id']}".encode()).decode()
    print("login admin", f["codigo"], f["nome"])
    return {"token": token, "nome": f["nome"], "perfil": f["perfil"], "loja_id": f["loja_id"]}


# =====================================================================
#  MODULO LOJA (produtos / estoque / vendas) - v2
#  (antes era um sistema separado, juntamos aqui em 2025)
# =====================================================================

PRAZO_TROCA_DIAS = 7  # CDC
ATENDIMENTOS_CSV = os.path.join(os.path.dirname(__file__), "data", "atendimentos_export_whatsapp_insta.csv")

FORMATOS_DATA = ["%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
                 "%Y-%m-%d"]


def parse_data(s):
    if s is None or s == "":
        return None
    s = str(s).replace("Z", "")  # ja ta em UTC
    for f in FORMATOS_DATA:
        try:
            return datetime.strptime(s, f)
        except:
            pass
    print("data invalida", s)
    return None


def dia_utc(s):
    # tudo em UTC pra bater com o relatorio do app
    d = parse_data(s)
    if d is None:
        return None
    return d.strftime("%Y-%m-%d")


def valor(v):
    try:
        return float(str(v).replace(",", "."))
    except:
        return 0.0


class ItemVendaIn(BaseModel):
    sku: str
    qtd: int = 1
    preco_unit: float
    descricao: Optional[str] = None


class VendaIn(BaseModel):
    canal: str = "loja"
    loja_id: int = 1
    cliente_ref: Optional[str] = None
    vendedor: Optional[str] = None
    itens: List[ItemVendaIn]
    desconto: float = 0
    frete: float = 0
    status: str = "pago"
    pedido_externo: Optional[str] = None
    obs: Optional[str] = None


class AjusteIn(BaseModel):
    sku: str
    loja_id: int
    qtd: int


# ------------------------------------------------------------- produtos

@app.get("/produtos")
def listar_produtos(q: str = "", colecao: str = "", page: int = 1):
    prods = ler_catalogo()
    res = []
    for p in prods:
        if colecao and colecao.lower() not in p["colecao"].lower():
            continue
        if q and q.lower() not in str(p["nome"]).lower() and q.lower() not in p["sku"].lower():
            continue
        res.append(p)
    inicio = (page - 1) * 50
    return {"total": len(res), "page": page, "produtos": res[inicio:inicio + 50]}


@app.get("/produtos/{sku}")
def get_produto(sku: str):
    p = busca_produto(sku)
    if p is None:
        raise HTTPException(404, "Produto nao encontrado")
    estoque = query("select e.loja_id, l.nome as loja, e.qtd, e.atualizado_em from estoque e "
                    "left join lojas l on l.id = e.loja_id where e.sku = ?", (p["sku"],))
    p = dict(p)
    p["estoque"] = estoque
    p["estoque_total"] = sum(e["qtd"] for e in estoque)
    return p


# ------------------------------------------------------------- estoque

@app.get("/estoque")
def listar_estoque(loja: str = "", q: str = "", page: int = 1):
    sql = ("select e.sku, e.loja_id, l.nome as loja, e.qtd, e.atualizado_em from estoque e "
           "left join lojas l on l.id = e.loja_id where 1=1")
    params = []
    if loja:
        if loja.isdigit():
            sql += " and e.loja_id = ?"
            params.append(int(loja))
        else:
            sql += " and l.nome like ?"
            params.append("%" + loja + "%")
    if q:
        sql += " and e.sku like ?"
        params.append("%" + q + "%")
    sql += " order by e.sku, e.loja_id"
    rows = query(sql, params)
    # TODO paginar no sql (ficou lento?)
    total_pecas = sum(r["qtd"] for r in rows)
    inicio = (page - 1) * 50
    return {"total": len(rows), "total_pecas": total_pecas, "page": page, "itens": rows[inicio:inicio + 50]}


@app.post("/estoque/ajuste")
def ajuste_estoque(dados: AjusteIn):
    # usado pelas lojas depois da contagem e pela integracao do CD
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    execute("insert or replace into estoque (sku, loja_id, qtd, atualizado_em) values (?,?,?,?)",
            (dados.sku, dados.loja_id, dados.qtd, agora))
    return {"ok": True, "sku": dados.sku, "loja_id": dados.loja_id, "qtd": dados.qtd}


def baixa_estoque(sku, loja_id, qtd):
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur = conn.cursor()
    cur.execute("update estoque set qtd = qtd - ?, atualizado_em = ? where sku = ? and loja_id = ?",
                (qtd, agora, sku, loja_id))
    if cur.rowcount == 0:
        cur.execute("insert into estoque (sku, loja_id, qtd, atualizado_em) values (?,?,?,?)",
                    (sku, loja_id, -qtd, agora))
    conn.commit()


# ------------------------------------------------------------- vendas

@app.get("/vendas")
def listar_vendas(page: int = 1, loja: str = "", status: str = "", q: str = ""):
    where = " where 1=1"
    params = []
    if loja:
        where += " and v.loja_id = ?"
        params.append(int(loja))
    if status:
        where += " and v.status = ?"
        params.append(status)
    if q:
        where += " and (v.cliente_ref like ? or v.pedido_externo like ? or cast(v.id as text) = ?)"
        params += ["%" + q + "%", "%" + q + "%", q]
    total = query("select count(*) as n from vendas v" + where, params)[0]["n"]
    rows = query("select v.*, l.nome as loja from vendas v left join lojas l on l.id = v.loja_id" + where +
                 " order by v.id desc limit 50 offset ?", params + [(page - 1) * 50])
    return {"total": total, "page": page, "vendas": rows}


@app.get("/vendas/{venda_id}")
def get_venda(venda_id: int):
    v = query("select v.*, l.nome as loja from vendas v left join lojas l on l.id = v.loja_id where v.id = ?",
              (venda_id,))
    if len(v) == 0:
        raise HTTPException(404, "Venda nao encontrada")
    v = v[0]
    v["itens"] = query("select * from itens_venda where venda_id = ?", (venda_id,))
    return v


@app.post("/venda")
def nova_venda(dados: VendaIn):
    bruto = sum(i.qtd * i.preco_unit for i in dados.itens)
    total = bruto - dados.desconto + dados.frete
    if dados.canal in CANAIS_ONLINE:
        criado = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    else:
        criado = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    total_str = str(round(total, 2))
    if dados.canal == "app":
        total_str = str(int(round(total * 100)))  # app trabalha em centavos
    venda_id = execute(
        "insert into vendas (pedido_externo, canal, loja_id, cliente_ref, vendedor, total, desconto, frete, status, "
        "criado_em, obs) values (?,?,?,?,?,?,?,?,?,?,?)",
        (dados.pedido_externo, dados.canal, dados.loja_id, dados.cliente_ref, dados.vendedor, total_str,
         str(dados.desconto), str(dados.frete) if dados.frete else None, dados.status, criado, dados.obs))
    for i in dados.itens:
        execute("insert into itens_venda (venda_id, sku, descricao, qtd, preco_unit) values (?,?,?,?,?)",
                (venda_id, i.sku, i.descricao, i.qtd, str(i.preco_unit)))
        baixa_estoque(i.sku, dados.loja_id, i.qtd)

    # fidelidade (igual /compra)
    pontos = 0
    if dados.cliente_ref:
        compra_id = execute(
            "insert into compras (cliente_ref, valor, canal, loja_id, criado_em, origem) values (?,?,?,?,?,?)",
            (dados.cliente_ref, str(bruto), dados.canal, dados.loja_id, datetime.now().isoformat(),
             "pdv" if dados.canal not in CANAIS_ONLINE else dados.canal))
        pontos = bruto
        if dados.canal in CANAIS_ONLINE:
            pontos = bruto * 2
        pontos = round(pontos)
        ref = limpa_cpf(dados.cliente_ref)
        if ref.isdigit():
            # TODO so creditar depois que o pagamento aprovar
            execute("insert into pontos_mov (cpf, tipo, qtd, compra_id, criado_em, obs) values (?,?,?,?,?,?)",
                    (ref, "C", pontos, compra_id, datetime.now().isoformat(), None))
        execute("update vendas set compra_id = ? where id = ?", (compra_id, venda_id))
    return {"ok": True, "venda_id": venda_id, "total": total_str, "pontos": pontos}


@app.post("/venda/{venda_id}/cancelar")
def cancelar_venda(venda_id: int):
    v = query("select * from vendas where id = ?", (venda_id,))
    if len(v) == 0:
        raise HTTPException(404, "Venda nao encontrada")
    v = v[0]
    criado = parse_data(v["criado_em"])
    if criado is not None and (datetime.now() - criado).days > PRAZO_TROCA_DIAS:
        raise HTTPException(400, f"Venda fora do prazo de {PRAZO_TROCA_DIAS} dias")
    execute("update vendas set status = 'cancelado' where id = ?", (venda_id,))
    # TODO devolver estoque
    # TODO estornar pontos (por enquanto a loja lanca o E na mao)
    return {"ok": True, "venda_id": venda_id, "status": "cancelado"}


# ------------------------------------------------------------- relatorio

@app.get("/relatorio/faturamento")
def relatorio_faturamento(inicio: str = "2025-01-01", fim: str = "2026-12-31", loja: str = ""):
    vendas = query("select id, total, status, criado_em, loja_id, canal from vendas")
    lojas_nome = {l["id"]: l["nome"] for l in query("select * from lojas")}
    por_dia = {}
    por_mes = {}
    por_loja = {}
    total = 0.0
    n = 0
    for v in vendas:
        if loja and str(v["loja_id"]) != str(loja):
            continue
        dia = dia_utc(v["criado_em"])
        if dia is None or dia < inicio or dia > fim:
            continue
        # if v["status"] in ("cancelado", "CANCELADA"):
        #     continue   # TODO confirmar com o financeiro (Paula)
        val = valor(v["total"])
        total += val
        n += 1
        por_dia[dia] = por_dia.get(dia, 0) + val
        por_mes[dia[:7]] = por_mes.get(dia[:7], 0) + val
        nome = lojas_nome.get(v["loja_id"], "?")
        por_loja[nome] = por_loja.get(nome, 0) + val
    return {
        "inicio": inicio,
        "fim": fim,
        "faturamento": round(total, 2),
        "qtd_vendas": n,
        "ticket_medio": round(total / n, 2) if n else 0,
        "por_dia": [{"dia": d, "total": round(t, 2)} for d, t in sorted(por_dia.items())],
        "por_mes": [{"mes": m, "total": round(t, 2)} for m, t in sorted(por_mes.items())],
        "por_loja": [{"loja": l, "total": round(t, 2)} for l, t in sorted(por_loja.items(), key=lambda x: -x[1])],
    }


# ------------------------------------------------------------- atendimento (export da ferramenta)

@app.get("/atendimentos")
def listar_atendimentos(q: str = "", canal: str = "", status: str = "", page: int = 1):
    try:
        df = pd.read_csv(ATENDIMENTOS_CSV, sep=";", encoding="latin-1", skiprows=1, dtype=str)
    except Exception as e:
        print("erro lendo csv atendimento", e)
        return {"total": 0, "page": page, "atendimentos": []}
    df = df.fillna("")
    res = []
    for i, row in df.iterrows():
        if canal and canal.lower() not in row["canal"].lower():
            continue
        if status and status.lower() != row["status"].lower():
            continue
        if q and q.lower() not in (row["mensagem"] + " " + row["contato"] + " " + row["pedido"]).lower():
            continue
        res.append(row.to_dict())
    inicio = (page - 1) * 50
    return {"total": len(res), "page": page, "atendimentos": res[inicio:inicio + 50]}
