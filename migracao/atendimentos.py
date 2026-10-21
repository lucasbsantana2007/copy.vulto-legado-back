"""Atendimentos da Vulto (WhatsApp, Instagram e e-mail) ligados ao cliente e à venda certa (decisões 42 e 43).

1. A fonte é o export limpo pelo time (migracao/entrada/atendimentos_export_whatsapp_insta.xlsx, 19/10/2026): datas,
   canais, atendentes, contato, assunto nas 9 categorias do time, mensagem do cliente separada da resposta, status
   pela última resposta e o número do pedido. O texto como veio da ferramenta (status, assunto, contato e pedido
   originais) vem do export original em data/.
2. Aqui cada atendimento é ligado ao cliente (clientes_padronizados) e à venda (vendas_padronizadas):
   - quem está falando é identificado pelo contato (e-mail ou celular do cadastro); se o contato não identifica
     (ex.: @ do Instagram), vale o cliente da venda citada;
   - a venda só é ligada se existir e, quando os dois lados têm cliente, se for do mesmo cliente e não for posterior
     ao atendimento;
   - sem venda, o atendimento fica no SAC central (loja do site), porque a Lorenzi pede uma loja por atendimento.

Saída: tabela atendimentos no vulto_padronizado.db (o carregar.py grava como atendimentos_padronizados).

Uso: python migracao/atendimentos.py   (depois de padroniza_vendas.py e clientes_pontos.py)
"""

import os
import re
import sqlite3
from datetime import datetime, timedelta

import pandas as pd

import caminhos

CLIENTES = os.path.join(caminhos.SAIDA, "clientes atualizado v2.xlsx")
LIMPO = os.path.join(caminhos.ENTRADA, "atendimentos_export_whatsapp_insta.xlsx")
ORIGINAL = os.path.join(caminhos.DADOS, "atendimentos_export_whatsapp_insta.csv")


def ler():
    """Atendimentos limpos pelo time, com o texto original de cada um ao lado (para conferir o que mudou)."""
    a = pd.read_excel(LIMPO, dtype=str).fillna("")
    a["data_hora"] = pd.to_datetime(a.data_hora).dt.strftime("%d/%m/%Y %H:%M:%S")
    original = pd.read_csv(ORIGINAL, sep=";", encoding="latin-1", skiprows=1, dtype=str).fillna("")
    original = original.drop_duplicates("id").set_index("id")
    for col in ("status", "assunto", "contato", "pedido"):
        a[f"{col}_original"] = a.id.map(original[col])
    return a.rename(columns={"nome_contato": "nome"})


def digitos(texto):
    d = re.sub(r"\D", "", str(texto or ""))
    return d[2:] if d.startswith("55") and len(d) in (12, 13) else d


def main():
    a = ler()
    clientes = pd.read_excel(CLIENTES, sheet_name="Clientes", dtype=str)
    cadastro = set(clientes["CPF"])
    por_email = {e.lower(): c for e, c in zip(clientes["E-mail"], clientes["CPF"]) if isinstance(e, str) and e}
    telefones = clientes.dropna(subset=["Telefone"]).drop_duplicates("Telefone", keep=False)
    por_tel = dict(zip(telefones["Telefone"].map(digitos), telefones["CPF"]))
    conn = sqlite3.connect(caminhos.PADRONIZADO)
    todas = pd.read_sql("select id_venda, pedido_externo, data, loja_id, cliente_cpf, tipo, duplicada_de, motivo "
                        "from vendas_padronizadas", conn)
    relancada = todas.motivo.fillna("").str.extract(r"^Relançamento da venda manual (BATEL-\d+)")[0]
    # só as vendas que ficam no banco corrigido: o pedido lançado duas vezes liga à gravação que vale (decisão 10)
    vendas = todas[(todas.tipo != "Teste") & (todas.duplicada_de.fillna("").astype(str) == "") & relancada.isna()].copy()
    vendas["chave"] = [p if p else str(i) for p, i in zip(vendas.pedido_externo, vendas.id_venda)]
    vendas = vendas.drop_duplicates("chave").set_index("chave")
    # o número de um relançamento da Batel liga à venda da planilha que vale (decisão 17); o CPF anotado no
    # relançamento identifica o cliente quando a planilha não tem
    da_planilha = vendas.set_index(vendas.id_venda.astype(str))
    for numero, batel_id, cpf in zip(todas.id_venda[relancada.notna()].astype(str), relancada.dropna(),
                                     todas.cliente_cpf[relancada.notna()]):
        venda = da_planilha.loc[batel_id].copy()
        venda["cliente_cpf"] = venda.cliente_cpf or cpf
        vendas.loc[numero] = venda

    linhas = []
    for r in a.itertuples():
        quando = datetime.strptime(r.data_hora, "%d/%m/%Y %H:%M:%S")
        contato = str(r.contato or "")
        pelo_contato = por_email.get(contato.lower()) if "@" in contato and not contato.startswith("@") \
            else por_tel.get(digitos(contato))
        tem_pedido = isinstance(r.pedido, str) and r.pedido != ""
        venda = vendas.loc[r.pedido] if tem_pedido and r.pedido in vendas.index else None
        cpf_da_venda = venda.cliente_cpf if venda is not None and venda.cliente_cpf else None
        pelo_pedido = cpf_da_venda if cpf_da_venda in cadastro else None   # só cliente que existe no cadastro
        cpf = pelo_contato or pelo_pedido
        if venda is None:
            confere = "Pedido não encontrado" if tem_pedido else ""
        elif datetime.strptime(venda.data, "%d/%m/%Y") > quando + timedelta(days=1):
            confere = "Não: venda depois do atendimento"
        elif pelo_contato and pelo_pedido and pelo_contato != pelo_pedido:
            confere = "Não: venda de outro cliente"
        elif pelo_contato and pelo_pedido:
            confere = "Sim"
        else:
            confere = "Não verificável (falta o cliente num dos lados)"
        ligada = venda is not None and not confere.startswith("Não:")
        linhas.append(dict(
            id=r.id, data=quando.strftime("%d/%m/%Y"), hora=quando.strftime("%H:%M"), canal=r.canal, contato=contato,
            nome=r.nome, cliente_cpf=cpf or "",
            cliente_identificado_por=("Contato do cadastro" if pelo_contato else "Cliente da venda" if pelo_pedido
                                      else f"CPF da venda sem cadastro no Vulto+ ({cpf_da_venda})" if cpf_da_venda else ""),
            assunto=r.assunto, mensagem_cliente=r.mensagem_cliente, resposta_atendente=r.resposta_atendente,
            status=r.status, atendente=r.atendente, pedido=r.pedido if tem_pedido else "",
            id_venda=str(venda.id_venda) if ligada else "", venda_confere=confere,
            loja_id=int(venda.loja_id) if ligada else 5,
            status_original=r.status_original, assunto_original=r.assunto_original, contato_original=r.contato_original,
            pedido_original=r.pedido_original))
    df = pd.DataFrame(linhas)
    df.to_sql("atendimentos", conn, if_exists="replace", index=False)
    conn.close()
    print(f"{len(df)} atendimentos | com cliente: {(df.cliente_cpf != '').sum()} | ligados a uma venda: {(df.id_venda != '').sum()}")
    print("venda confere:", df.venda_confere.value_counts().to_dict())
    return df


if __name__ == "__main__":
    main()
