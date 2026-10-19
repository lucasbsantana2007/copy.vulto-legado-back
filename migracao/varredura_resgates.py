"""Varredura dos resgates anotados só com o primeiro nome (roteiro da aba "Varredura (automação)" da planilha de
fidelidade limpa pelo time em 19/10/2026).

Para cada resgate da aba Resgates sem CPF e anotado só com o primeiro nome:
1. procura as vendas da mesma loja, no dia do resgate ou no dia seguinte/anterior (vendas_padronizadas);
2. acha o CPF de cada venda: o CPF da venda, ou o telefone/e-mail batendo com a aba Clientes;
3. fica com os CPFs cujo nome no cadastro começa com o primeiro nome anotado (sem acento, sem maiúsculas);
4. 1 CPF e o desconto da venda igual ao desconto anotado no resgate = confiança alta (preencher o CPF);
   1 CPF com desconto diferente, ou 2 ou mais = sugestão; 0 = continua pendente.
Os atendimentos com o mesmo primeiro nome perto da data servem só para desempatar (nunca preenchem CPF).
O saldo do candidato no dia do resgate vem do extrato corrigido (extrato_pontos), quando a data é de 2025 em diante.

As vendas do vulto.db começam em 01/01/2025: resgates anteriores não têm com o que cruzar.

Entradas (só leitura): migracao/entrada/clientes_fidelidade_FINAL_v3 (1).xlsx e data/vulto.db (rode o carregar.py antes)
Saída: migracao/saida/varredura_resgates.csv (um caso por linha, com candidatos e a decisão)

Uso: python migracao/varredura_resgates.py
"""

import os
import re
import sqlite3
import unicodedata
from datetime import datetime, timedelta

import pandas as pd

import caminhos

PLANILHA = os.path.join(caminhos.ENTRADA, "clientes_fidelidade_FINAL_v3 (1).xlsx")
SAIDA = os.path.join(caminhos.SAIDA, "varredura_resgates.csv")
LOJA_ID = {"Pinheiros": 1, "Eldorado": 2, "Savassi": 3, "Batel": 4, "Site": 5, "App": 6}
INICIO_VENDAS = datetime(2025, 1, 1)
JANELA_VENDA = 1        # dias de diferença aceitos entre a venda e o resgate (lançamento no dia seguinte)
JANELA_ATENDIMENTO = 15  # dias, só para desempate


def chave(texto) -> str:
    return " ".join(unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode().lower().split())


def digitos(valor) -> str:
    return re.sub(r"\D", "", "" if valor is None or valor != valor else str(valor))


def ler():
    clientes = pd.read_excel(PLANILHA, sheet_name="Clientes", dtype=str)
    resgates = pd.read_excel(PLANILHA, sheet_name="Resgates", dtype=object)
    resgates["linha"] = resgates.index + 2
    with sqlite3.connect(caminhos.VULTO_DB) as c:
        vendas = pd.read_sql("select id_venda, data, hora, loja_id, cliente_tipo, cliente_cpf, cliente_contato, desconto, tipo "
                             "from vendas_padronizadas where tipo != 'Teste'", c)
        atend = pd.read_sql("select id, data, nome, cliente_cpf, assunto from atendimentos_padronizados", c)
        extrato = pd.read_sql("select cpf, data, pontos_considerado from extrato_pontos", c)
    vendas["dia"] = pd.to_datetime(vendas.data, format="%d/%m/%Y")
    atend["dia"] = pd.to_datetime(atend.data, format="%d/%m/%Y")
    extrato["quando"] = pd.to_datetime(extrato.data)
    return clientes, resgates, vendas, atend, extrato


def casos(resgates):
    """Resgates sem CPF anotados com uma palavra só (o primeiro nome)."""
    sem_cpf = resgates[resgates["CPF"].fillna("").astype(str).str.strip() == ""]
    anotado = sem_cpf["Cliente (como anotado)"].fillna("").astype(str).str.strip()
    return sem_cpf[anotado.str.fullmatch(r"[^\s\d]+")]


def main():
    clientes, resgates, vendas, atend, extrato = ler()
    nome_do_cpf = dict(zip(clientes.CPF, clientes.Nome))
    por_telefone, por_email = {}, {}
    for cpf, tel, email in zip(clientes.CPF, clientes.Telefone, clientes["E-mail"]):
        d = digitos(tel)
        if len(d) >= 8:
            por_telefone.setdefault(d[-8:], set()).add(cpf)
        if isinstance(email, str) and email.strip():
            por_email.setdefault(email.strip().lower(), set()).add(cpf)

    def cpfs_da_venda(v):
        if v.cliente_cpf:
            return {v.cliente_cpf}, "CPF da venda"
        if v.cliente_tipo == "Telefone" and len(digitos(v.cliente_contato)) >= 8:
            return por_telefone.get(digitos(v.cliente_contato)[-8:], set()), "telefone da venda = telefone do cadastro"
        if v.cliente_tipo == "E-mail":
            return por_email.get(str(v.cliente_contato).strip().lower(), set()), "e-mail da venda = e-mail do cadastro"
        return set(), ""

    def saldo_em(cpf, quando):
        e = extrato[(extrato.cpf == cpf) & (extrato.quando < quando)]
        return int(e.pontos_considerado.sum())

    linhas = []
    for r in casos(resgates).itertuples():
        quando, loja = pd.Timestamp(r.Data), r.Loja
        primeiro = chave(r._10)
        base = dict(linha_resgates=r.linha, data=quando.date(), loja=loja, anotado=r._10, pontos=int(r._4))
        if quando < INICIO_VENDAS:
            linhas.append(base | dict(resultado="Sem fonte", vendas_no_dia=None, candidatos="",
                                      decisao="Continua pendente: o vulto.db só tem vendas a partir de 01/01/2025"))
            continue
        perto = vendas[(vendas.loja_id == LOJA_ID[loja]) & ((vendas.dia - quando).abs() <= timedelta(days=JANELA_VENDA))]
        achados, desconto_bate = {}, set()
        for v in perto.itertuples():
            cpfs, como = cpfs_da_venda(v)
            desconto = float(str(v.desconto or "0").replace(",", "."))
            for cpf in cpfs:
                if cpf in nome_do_cpf and chave(nome_do_cpf[cpf]).split(" ")[0] == primeiro:
                    achados.setdefault(cpf, []).append(f"venda {v.id_venda} de {v.data} ({como}), desconto R$ {desconto:.2f}")
                    if abs(desconto - float(r._5)) < 0.01:
                        desconto_bate.add(cpf)
        candidatos = []
        for cpf, fontes in achados.items():
            at = atend[(atend.cliente_cpf == cpf) & ((atend.dia - quando).abs() <= timedelta(days=JANELA_ATENDIMENTO))]
            extra = f"; atendimento {at.id.iloc[0]} em {at.data.iloc[0]}" if len(at) else ""
            saldo = saldo_em(cpf, quando)
            candidatos.append(f"{nome_do_cpf[cpf]} {cpf} ({'; '.join(fontes)}{extra}; saldo antes do resgate {saldo} pts)")
        if len(achados) == 1 and desconto_bate:
            resultado, decisao = "1 CPF", "Preencher o CPF (confiança alta: venda do dia com o mesmo desconto)"
        elif len(achados) == 1:
            resultado, decisao = "1 CPF, desconto diferente", "Sugestão (confiança média): não preencher o CPF"
        elif achados:
            resultado, decisao = f"{len(achados)} CPFs", "Sugestão (confiança média): não preencher o CPF"
        else:
            resultado, decisao = "0 CPF", "Continua pendente: nenhuma venda da loja no dia com cliente desse nome"
        linhas.append(base | dict(resultado=resultado, vendas_no_dia=len(perto), candidatos=" | ".join(candidatos),
                                  decisao=decisao, cpfs=",".join(achados)))

    out = pd.DataFrame(linhas)
    out.to_csv(SAIDA, index=False, encoding="utf-8-sig", sep=";")
    print(out.drop(columns=["candidatos"]).to_string(index=False))
    print("\ncandidatos:")
    for r in out.itertuples():
        if r.candidatos:
            print(f"  linha {r.linha_resgates}: {r.candidatos}")
    print("\n", out.resultado.value_counts().to_dict(), "| gravado:", SAIDA)


if __name__ == "__main__":
    main()
