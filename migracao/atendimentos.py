"""Atendimentos da Vulto (WhatsApp, Instagram e e-mail) ligados ao cliente e à venda certa (decisões 42 e 43).

1. A padronização do export é a dos scripts do Arthur (migracao/padronizacao/atendimento.py): datas, canais,
   atendentes, contato, assunto em 9 categorias, mensagem do cliente separada da resposta, status pela última
   resposta e o número do pedido (prefixo limpo, recuperado do texto, inversão de dígitos só em pedido do site/app).
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
import sys
from datetime import datetime, timedelta

import pandas as pd

import caminhos

sys.path.append(os.path.join(caminhos.MIGRACAO, "padronizacao"))  # por último: os scripts do time têm prioridade
import atendimento as padronizacao_arthur  # noqa: E402

CLIENTES = os.path.join(caminhos.SAIDA, "clientes atualizado v2.xlsx")


def digitos(texto):
    d = re.sub(r"\D", "", str(texto or ""))
    return d[2:] if d.startswith("55") and len(d) in (12, 13) else d


def main():
    dados, reg = padronizacao_arthur.processar(caminhos.DADOS)
    a = dados["atendimentos"]
    clientes = pd.read_excel(CLIENTES, sheet_name="Clientes", dtype=str)
    cadastro = set(clientes["CPF"])
    por_email = {e.lower(): c for e, c in zip(clientes["E-mail"], clientes["CPF"]) if isinstance(e, str) and e}
    telefones = clientes.dropna(subset=["Telefone"]).drop_duplicates("Telefone", keep=False)
    por_tel = dict(zip(telefones["Telefone"].map(digitos), telefones["CPF"]))
    conn = sqlite3.connect(caminhos.PADRONIZADO)
    vendas = pd.read_sql("select id_venda, pedido_externo, data, loja_id, cliente_cpf from vendas_padronizadas", conn)
    vendas["chave"] = [p if p else str(i) for p, i in zip(vendas.pedido_externo, vendas.id_venda)]
    vendas = vendas.drop_duplicates("chave").set_index("chave")

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
