"""Padroniza as vendas da Vulto: uma linha por venda, no formato da planilha da Batel feita pelo Lucas.

Lê vulto_padronizado.db (saída de padroniza_valores.py, valores já em reais com vírgula) e grava nele a
tabela vendas_padronizadas, mais o CSV vendas_padronizadas.csv para abrir no Excel. As tabelas originais
não mudam. Cada venda diz se entra no faturamento (Sim / Não / Pendente) e o motivo; "Pendente" é o que
ainda depende de uma decisão do negócio (hoje nenhuma: as decisões 15, 16 e 17 foram tomadas em 19/10/2026).

Também grava as vendas da Batel feitas sem sistema (BATEL-<Nº>, decisão 17, ver batel.py) e deixa os itens com o
SKU do catálogo (coluna sku; o código como veio do sistema fica em sku_original). Rode produtos.py antes.

Uso: python migracao/padroniza_vendas.py
"""

import csv
import os
import re
import sqlite3
from collections import Counter
from datetime import datetime, timezone
from decimal import Decimal
from zoneinfo import ZoneInfo

import batel
import caminhos
import pandas as pd
import produtos

BANCO = caminhos.PADRONIZADO
CSV_SAIDA = os.path.join(caminhos.SAIDA, "vendas_padronizadas.csv")
BRASILIA = ZoneInfo("America/Sao_Paulo")
PERIODO = ("2025-01", "2026-09")

LOJAS = {1: "Pinheiros", 2: "Shopping Eldorado", 3: "Savassi", 4: "Batel", 5: "Site", 6: "App"}
CANAL_POR_LOJA = {1: "Loja física", 2: "Loja física", 3: "Loja física", 4: "Loja física", 5: "Site", 6: "App"}

STATUS = {
    "pago": "Concluída",
    "aprovado": "Concluída",
    "cancelado": "Cancelada",
    "cancelada": "Cancelada",
    "estornado": "Estornada",
    "chargeback": "Chargeback",
    "aguardando": "Aguardando pagamento",
    "": "Sem status",
}

# Vendedor: grafia (minúscula) -> nome do cadastro de funcionários, por loja. "Rafa" na Pinheiros pode ser o
# Rafael ou a Rafaela (os dois têm o crachá 0102): fica ambíguo em vez de chutar.
VENDEDORES = {
    1: {
        "beatriz": "Beatriz Souza", "beatriz s.": "Beatriz Souza", "bia": "Beatriz Souza", "bia.pinheiros": "Beatriz Souza",
        "camila": "Camila Duarte", "camila d.": "Camila Duarte", "cami": "Camila Duarte",
        "rafael": "Rafael Lima", "rafael l.": "Rafael Lima",
        "rafaela": "Rafaela Couto", "rafaela c.": "Rafaela Couto",
        "rafa": "Rafael Lima ou Rafaela Couto", "rafa.pin": "Rafael Lima ou Rafaela Couto",
    },
    2: {
        "jessica": "Jéssica Prado", "jéssica": "Jéssica Prado", "jessica p.": "Jéssica Prado", "jess.eldorado": "Jéssica Prado",
        "thiago": "Thiago Martins", "thiago m": "Thiago Martins", "tiago": "Thiago Martins",
        "larissa": "Larissa Rocha", "larissa r.": "Larissa Rocha", "lari": "Larissa Rocha",
    },
    3: {
        "carlos": "Carlos Mendes", "carlos.savassi": "Carlos Mendes", "carlinhos": "Carlos Mendes",
        "ana paula": "Ana Paula Reis", "anapaula": "Ana Paula Reis", "ana p.": "Ana Paula Reis", "ap": "Ana Paula Reis",
        "bianca": "Bianca Torres", "bianca t.": "Bianca Torres", "bia": "Bianca Torres",
    },
    4: {
        "fernanda": "Fernanda Kowalski", "fernanda k.": "Fernanda Kowalski", "fer": "Fernanda Kowalski", "fer.batel": "Fernanda Kowalski",
        "gustavo": "Gustavo Lenz", "gus": "Gustavo Lenz", "gustavo.batel": "Gustavo Lenz",
        "patricia": "Patrícia Wolff", "patrícia w.": "Patrícia Wolff", "pati": "Patrícia Wolff",
        "gerente": "Gerente Batel", "gerente batel": "Gerente Batel",
    },
}

CONTAS_DE_TESTE = {"qa.app@vulto.com.br", "teste@vulto.com.br", "12345678909"}
PEDIDO_DE_TESTE = re.compile(r"^(S99\d{4}|APP-99\d{5})$")      # séries de QA do site e do app (decisão 11)
SEM_CLIENTE = {"", "consumidor", "nao informado"}

COLUNAS = [
    "id_venda", "pedido_externo", "data", "hora", "mes", "fuso_convertido", "loja_id", "loja", "canal", "canal_original",
    "cliente_tipo", "cliente_cpf", "cpf_valido", "cliente_contato", "cliente_ref_original", "vendedor", "vendedor_original",
    "pecas", "valor_itens", "desconto", "frete", "valor_total", "status", "status_original", "tipo", "venda_original",
    "duplicada_de", "entra_no_faturamento", "motivo", "obs",
]


def num(texto):
    return Decimal(texto.replace(",", ".")) if texto else Decimal(0)


def reais(valor):
    return f"{valor.quantize(Decimal('0.01'))}".replace(".", ",")


def data_local(texto):
    """(momento em Brasília, se veio em UTC). 'Z' no fim = UTC; sem fuso = hora de Brasília."""
    for formato in ("%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M"):
        try:
            momento = datetime.strptime(texto.strip(), formato)
        except ValueError:
            continue
        if formato.endswith("Z"):
            return momento.replace(tzinfo=timezone.utc).astimezone(BRASILIA), True
        return momento.replace(tzinfo=BRASILIA), False
    raise ValueError(f"data em formato desconhecido: {texto}")


def cpf_valido(cpf):
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False
    for n in (9, 10):
        digito = sum(int(cpf[i]) * (n + 1 - i) for i in range(n)) * 10 % 11 % 10
        if digito != int(cpf[n]):
            return False
    return True


def cliente(ref):
    """(tipo, cpf formatado, cpf válido, contato) a partir do cliente_ref (CPF, e-mail ou telefone)."""
    texto = (ref or "").strip()
    if texto.lower() in SEM_CLIENTE:
        return "Outros", "", "", ""
    if "@" in texto:
        return "E-mail", "", "", texto.lower()
    digitos = re.sub(r"\D", "", texto)
    if len(digitos) == 14:
        return "CNPJ", "", "", digitos
    eh_cpf = re.fullmatch(r"\d{3}[.\s]?\d{3}[.\s]?\d{3}[-\s]?\d{2}", texto) is not None
    if eh_cpf:
        formatado = f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
        return "CPF", formatado, "Sim" if cpf_valido(digitos) else "Não", ""
    if digitos.startswith("55") and len(digitos) in (12, 13):
        digitos = digitos[2:]
    return "Telefone", "", "", digitos


def tipo_da_venda(total, obs, ref, externo=""):
    texto = (obs or "").lower()
    if "teste" in texto or (ref or "").strip().lower() in CONTAS_DE_TESTE or PEDIDO_DE_TESTE.match(externo or ""):
        return "Teste"
    if total < 0:
        return "Devolução"
    if texto.startswith("troca"):
        return "Troca"
    return "Venda"


def vendas_da_planilha_batel(planilha):
    """Decisão 17: cada linha da planilha manual da Batel vira uma venda BATEL-<Nº> na data e hora anotadas."""
    linhas = []
    for r in planilha.itertuples():
        cli_tipo, cpf, valido, _ = cliente(r.CPF if isinstance(r.CPF, str) else "")
        status = {"Concluída": "Concluída", "Cancelada": "Cancelada", "Estornada": "Estornada"}[r.Status]
        vend = str(r[11]).strip()
        linhas.append({
            "id_venda": f"BATEL-{r.numero}", "pedido_externo": "", "data": r.momento.strftime("%d/%m/%Y"),
            "hora": r.momento.strftime("%H:%M"), "mes": r.momento.strftime("%Y-%m"), "fuso_convertido": "Não",
            "loja_id": 4, "loja": LOJAS[4], "canal": CANAL_POR_LOJA[4], "canal_original": "planilha manual (PDV fora do ar)",
            "cliente_tipo": cli_tipo, "cliente_cpf": cpf, "cpf_valido": valido, "cliente_contato": "",
            "cliente_ref_original": r.CPF if isinstance(r.CPF, str) else "",
            "vendedor": VENDEDORES[4].get(vend.lower(), vend), "vendedor_original": vend,
            "pecas": int(r.Qtd), "valor_itens": reais(Decimal(str(r.valor))), "desconto": reais(Decimal(0)),
            "frete": reais(Decimal(0)), "valor_total": reais(Decimal(str(r.valor))), "status": status,
            "status_original": r.Status, "tipo": "Venda", "venda_original": "", "duplicada_de": "",
            "entra_no_faturamento": "Sim" if status == "Concluída" else "Não",
            "motivo": "Venda manual da Batel (PDV fora do ar): a planilha é a fonte (decisão 17)" if status == "Concluída"
            else f"Status {status.lower()} (planilha manual da Batel)",
            "obs": " · ".join(x for x in (str(r[14]) if isinstance(r[14], str) else "", f"Peças: {r[7]}") if x),
        })
    return linhas


def itens_padronizados(conn, planilha, itens_batel):
    """Itens com o SKU do catálogo (o código como veio fica em sku_original) e os itens das vendas BATEL-<Nº>."""
    resolver = produtos.resolvedor()
    colunas = [c[1] for c in conn.execute("pragma table_info(itens_venda)")]
    if "sku_original" not in colunas:
        conn.execute("alter table itens_venda add column sku_original")
        conn.execute("update itens_venda set sku_original = sku")
    conn.execute("delete from itens_venda where venda_id like 'BATEL-%'")
    for iid, original in conn.execute("select id, sku_original from itens_venda").fetchall():
        sku = resolver(original) or (produtos.GENERICO if original in ("DIVERSO", "PRODUTO SEM CADASTRO") else original)
        conn.execute("update itens_venda set sku = ? where id = ?", (sku, iid))
    conn.execute("update itens_venda set descricao = 'Peça não identificada no PDV' where sku = ?", (produtos.GENERICO,))
    catalogo = pd.read_sql("select * from produtos", conn)
    novos, nao_achados = batel.itens_com_sku(itens_batel, planilha, catalogo)
    proximo = conn.execute("select max(id) from itens_venda").fetchone()[0] + 1
    conn.executemany("insert into itens_venda (id, venda_id, sku, descricao, qtd, preco_unit, sku_original) values (?,?,?,?,?,?,?)",
                     [(proximo + n, venda, sku, desc, qtd, reais(Decimal(str(preco))), "planilha manual Batel")
                      for n, (venda, sku, desc, qtd, preco) in enumerate(novos)])
    conn.commit()
    sem_sku = conn.execute("select count(*) from itens_venda where sku not in (select sku from produtos)").fetchone()[0]
    print(f"itens: {len(novos)} peças da planilha da Batel ({len(nao_achados)} sem SKU) | itens sem SKU do catálogo: {sem_sku}")


def main():
    conn = sqlite3.connect(BANCO)
    vendas = conn.execute(
        "select id, pedido_externo, canal, loja_id, cliente_ref, vendedor, total, desconto, frete, status, criado_em, obs "
        "from vendas order by id"
    ).fetchall()
    itens = {
        vid: (pecas, valor)
        for vid, pecas, valor in conn.execute(
            "select venda_id, sum(qtd), sum(qtd * cast(replace(preco_unit, ',', '.') as real)) from itens_venda group by venda_id"
        )
    }
    ids = {v[0] for v in vendas}
    # Batel (decisão 17): relançamentos da gerente que casam com uma linha da planilha manual
    planilha_batel, itens_batel = batel.ler()
    relancadas = [(v[0], float(num(v[6])), re.sub(r"\D", "", cliente(v[4])[1])) for v in vendas
                  if v[3] == 4 and "gerente" in (v[5] or "").lower()]
    relancada_de = batel.casar_relancamentos(relancadas, planilha_batel)
    # Pedido do app lançado duas vezes (aguardando -> pago, mesmo valor): fica o último
    ultimo_do_pedido = {}
    for v in vendas:
        if v[1]:
            ultimo_do_pedido[v[1]] = v[0]

    linhas = []
    for vid, externo, canal, loja_id, ref, vendedor, total, desconto, frete, status, criado, obs in vendas:
        momento, convertido = data_local(criado)
        valor_total = num(total)
        tipo = tipo_da_venda(valor_total, obs, ref, externo)
        status_padrao = STATUS[(status or "").strip().lower()]
        cli_tipo, cpf, valido, contato = cliente(ref)
        vend = (vendedor or "").strip()
        vendedor_padrao = VENDEDORES.get(loja_id, {}).get(vend.lower(), vend) if vend else ""
        numero_original = re.search(r"(\d+)", obs or "") if tipo == "Devolução" else None
        original = int(numero_original.group(1)) if numero_original and int(numero_original.group(1)) in ids else ""
        duplicada_de = ultimo_do_pedido[externo] if externo and ultimo_do_pedido[externo] != vid else ""
        relancada = loja_id == 4 and "gerente" in vend.lower()

        mes = momento.strftime("%Y-%m")
        if not PERIODO[0] <= mes <= PERIODO[1]:
            entra, motivo = "Não", "Fora do período (jan/2025 a set/2026)"
        elif tipo == "Teste":
            entra, motivo = "Não", "Venda de teste"
        elif duplicada_de:
            entra, motivo = "Não", f"Pedido lançado duas vezes: vale a venda {duplicada_de}"
        elif status_padrao in ("Cancelada", "Estornada", "Chargeback", "Aguardando pagamento"):
            entra, motivo = "Não", f"Status {status_padrao.lower()}"
        elif relancada and vid in relancada_de:
            entra, motivo = "Não", f"Relançamento da venda manual BATEL-{relancada_de[vid]}: conta pela planilha (decisão 17)"
        elif tipo == "Troca":
            entra, motivo = "Não", "Troca: a compra original já entrou no faturamento (decisão 16)"
        elif relancada:
            entra, motivo = "Sim", "Relançamento sem venda correspondente na planilha manual da Batel: venda real (decisão 17)"
        elif status_padrao == "Sem status":
            entra, motivo = "Sim", "Loja física sem status: venda de balcão, considerada paga (decisão 15)"
        elif tipo == "Devolução":
            entra, motivo = "Sim", "Devolução: entra negativa no mês em que aconteceu"
        else:
            entra, motivo = "Sim", "Venda concluída"

        pecas, valor_itens = itens.get(vid, (0, 0))
        linhas.append({
            "id_venda": vid,
            "pedido_externo": externo or "",
            "data": momento.strftime("%d/%m/%Y"),
            "hora": momento.strftime("%H:%M"),
            "mes": mes,
            "fuso_convertido": "Sim" if convertido else "Não",
            "loja_id": loja_id,
            "loja": LOJAS[loja_id],
            "canal": CANAL_POR_LOJA[loja_id],
            "canal_original": (canal or "").strip(),
            "cliente_tipo": cli_tipo,
            "cliente_cpf": cpf,
            "cpf_valido": valido,
            "cliente_contato": contato,
            "cliente_ref_original": ref or "",
            "vendedor": vendedor_padrao,
            "vendedor_original": vend,
            "pecas": pecas,
            "valor_itens": reais(Decimal(str(valor_itens))),
            "desconto": reais(num(desconto)),
            "frete": reais(num(frete)),
            "valor_total": reais(valor_total),
            "status": status_padrao,
            "status_original": status or "",
            "tipo": tipo,
            "venda_original": original,
            "duplicada_de": duplicada_de,
            "entra_no_faturamento": entra,
            "motivo": motivo,
            "obs": obs or "",
        })

    linhas += vendas_da_planilha_batel(planilha_batel)
    itens_padronizados(conn, planilha_batel, itens_batel)

    conn.execute("drop table if exists vendas_padronizadas")
    conn.execute(f"create table vendas_padronizadas ({', '.join(COLUNAS)}, primary key (id_venda))")
    conn.executemany(
        f"insert into vendas_padronizadas values ({', '.join('?' for _ in COLUNAS)})", [[l[c] for c in COLUNAS] for l in linhas]
    )
    conn.commit()
    conn.close()
    with open(CSV_SAIDA, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLUNAS, delimiter=";")
        w.writeheader()
        w.writerows(linhas)

    def grupo(linha):
        return linha["entra_no_faturamento"], linha["motivo"].split(":")[0]

    print(f"{len(linhas)} vendas padronizadas")
    for (entra, motivo), qtd in sorted(Counter(map(grupo, linhas)).items()):
        soma = sum(num(l["valor_total"]) for l in linhas if grupo(l) == (entra, motivo))
        print(f"  {entra:8} {qtd:6} vendas  R$ {reais(soma):>14}  {motivo}")
    print("vendedor:", dict(Counter(l["vendedor"] for l in linhas if l["loja_id"] <= 4)))
    print("cliente:", dict(Counter(l["cliente_tipo"] for l in linhas)), " cpf inválido:", sum(l["cpf_valido"] == "Não" for l in linhas))
    print("fuso convertido:", sum(l["fuso_convertido"] == "Sim" for l in linhas),
          " mudaram de dia:", sum(l["fuso_convertido"] == "Sim" and l["hora"] >= "21:00" for l in linhas))
    print(f"gravado: tabela vendas_padronizadas em {BANCO}\ncsv: {CSV_SAIDA}")


if __name__ == "__main__":
    main()
