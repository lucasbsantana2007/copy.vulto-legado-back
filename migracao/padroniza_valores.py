"""Padroniza os valores em dinheiro das vendas da Vulto.

Lê o banco legado (sem alterar) e grava uma cópia em que todo valor fica no
formato brasileiro com duas casas: "1234,50" (vírgula decimal, sem separador de
milhar). Vendas do app gravadas em centavos são convertidas para reais, e o
frete que o site/app embutiu no total vai para a coluna frete.

Saídas, em migracao/saida:
- vulto_padronizado.db: cópia do banco com os valores padronizados
- log_padronizacao_valores.csv: cada valor alterado (tabela, id, coluna, antes, depois, regra)

Uso: python migracao/padroniza_valores.py  (ou passe outro vulto.db como argumento)
"""
import csv
import os
import re
import shutil
import sqlite3
import sys

import caminhos
from decimal import Decimal, ROUND_HALF_UP

DESTINO = caminhos.PADRONIZADO
LOG = os.path.join(caminhos.SAIDA, "log_padronizacao_valores.csv")

CANAIS_EM_CENTAVOS = ("app", "online")
FAIXAS_FRETE = {Decimal("9.90"), Decimal("14.90"), Decimal("19.90")}
INTEIRO = re.compile(r"-?\d+")


def para_decimal(texto):
    """'559,70', '559.70' ou '559' -> Decimal. Vazio ou NULL -> None (não inventa valor)."""
    if texto is None or str(texto).strip() == "":
        return None
    return Decimal(str(texto).strip().replace(",", "."))


def formata(valor):
    """Decimal -> '559,70'."""
    return str(valor.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)).replace(".", ",")


def main(origem):
    shutil.copyfile(origem, DESTINO)
    conn = sqlite3.connect(DESTINO)
    log = []

    # vendas do app/online com total inteiro estão em centavos (total, desconto e itens)
    em_centavos = {
        vid for vid, canal, total in conn.execute("select id, canal, total from vendas")
        if canal and canal.strip().lower() in CANAIS_EM_CENTAVOS and INTEIRO.fullmatch(total.strip())
    }

    def padroniza(tabela, coluna, chave_venda):
        linhas = conn.execute(f"select id, {coluna}, {chave_venda} from {tabela}").fetchall()
        for rid, antes, venda_id in linhas:
            valor = para_decimal(antes)
            if valor is None:
                if antes is not None:  # "" vira NULL: os dois querem dizer "não tem"
                    conn.execute(f"update {tabela} set {coluna} = null where id = ?", (rid,))
                    log.append((tabela, rid, coluna, antes, "", "vazio -> NULL"))
                continue
            regra = "separador decimal -> virgula"
            if venda_id in em_centavos:
                valor = valor / 100
                regra = "centavos do app -> reais"
            depois = formata(valor)
            if depois != antes:
                conn.execute(f"update {tabela} set {coluna} = ? where id = ?", (depois, rid))
                log.append((tabela, rid, coluna, antes, depois, regra))

    for coluna in ("total", "desconto", "frete"):
        padroniza("vendas", coluna, "id")
    padroniza("itens_venda", "preco_unit", "venda_id")
    conn.commit()

    def num(s):
        return para_decimal(s) or Decimal(0)
    itens = {}
    for vid, qtd, preco in conn.execute("select venda_id, qtd, preco_unit from itens_venda"):
        itens[vid] = itens.get(vid, Decimal(0)) + qtd * num(preco)

    # frete é receita (decisão 4): no site/app o frete às vezes está dentro do total e não na
    # coluna frete. A diferença é sempre uma faixa de frete; separa para o número mostrar o frete.
    for vid, total, desc, frete in conn.execute(
            "select id, total, desconto, frete from vendas where loja_id in (5, 6)").fetchall():
        embutido = num(total) - (itens.get(vid, Decimal(0)) - num(desc) + num(frete))
        if num(frete) == 0 and embutido in FAIXAS_FRETE:
            conn.execute("update vendas set frete = ? where id = ?", (formata(embutido), vid))
            log.append(("vendas", vid, "frete", frete or "", formata(embutido), "frete embutido no total -> coluna frete"))
    conn.commit()

    # conferência: total = itens - desconto + frete, agora tudo em reais
    divergentes = [
        vid for vid, total, desc, frete in conn.execute("select id, total, desconto, frete from vendas")
        if abs(num(total) - (itens.get(vid, Decimal(0)) - num(desc) + num(frete))) > Decimal("0.05")
    ]
    conn.close()

    with open(LOG, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["tabela", "id", "coluna", "antes", "depois", "regra"])
        w.writerows(log)

    print(f"vendas em centavos convertidas: {len(em_centavos)}")
    for regra in ("separador decimal -> virgula", "centavos do app -> reais", "vazio -> NULL",
                  "frete embutido no total -> coluna frete"):
        print(f"{regra}: {sum(1 for l in log if l[5] == regra)} valores")
    print(f"vendas com total != itens - desconto + frete: {len(divergentes)}")
    print(f"gravado: {DESTINO}\nlog: {LOG}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else caminhos.VULTO_DB)
