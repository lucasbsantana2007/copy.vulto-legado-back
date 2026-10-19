"""Recarrega no data/vulto.db tudo o que o time padronizou. Um comando só, que qualquer pessoa pode rodar:

    python migracao/carregar.py

1. padroniza_valores.py  valores em reais com vírgula (o app gravava centavos), frete separado
2. produtos.py           catálogo: um produto por SKU (e o tradutor de qualquer código para o SKU)
3. padroniza_vendas.py   uma linha por venda, com o motivo de entrar ou não no faturamento (e as vendas da Batel
                         feitas sem sistema); itens com o SKU do catálogo
4. clientes_pontos.py    clientes do Vulto+ corrigidos e saldo de pontos em 01/10/2026
5. estoque_01_10.py      contagens avançadas até 01/10/2026 (o número oficial)
6. estoque_historico.py  histórico dia a dia desde 01/01/2025, terminando no número oficial
7. atendimentos.py       atendimentos (WhatsApp, Instagram, e-mail) ligados ao cliente e à venda
8. validar.py            confere as tabelas antes de gravar no vulto.db (se algo estiver errado, para aqui)

Depois grava no vulto.db as tabelas abaixo (apagadas e recriadas a cada rodada). As tabelas originais da Vulto
(lojas, vendas, itens_venda, compras, pontos_mov, estoque, entradas_estoque, funcionarios, config) nunca mudam:
são o "sistema da Vulto" com que os nossos números são comparados.
"""

import os
import re
import sqlite3
import sys
import unicodedata

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import caminhos  # noqa: E402
import atendimentos  # noqa: E402
import clientes_pontos  # noqa: E402
import estoque_01_10  # noqa: E402
import estoque_historico  # noqa: E402
import padroniza_valores  # noqa: E402
import padroniza_vendas  # noqa: E402
import produtos  # noqa: E402
import registro  # noqa: E402
import validar  # noqa: E402

ORIGINAIS = {"lojas", "compras", "pontos_mov", "config", "vendas", "itens_venda", "estoque", "funcionarios", "entradas_estoque", "sqlite_sequence"}


def coluna(nome: str) -> str:
    """'Saldo válido 01/10/2026' -> 'saldo_valido_01_10_2026'."""
    texto = unicodedata.normalize("NFKD", str(nome)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"_+", "_", re.sub(r"[^a-z0-9]+", "_", texto)).strip("_")


def aba(arquivo: str, nome: str, cabecalho: int = 0) -> pd.DataFrame:
    df = pd.read_excel(os.path.join(caminhos.SAIDA, arquivo), sheet_name=nome, header=cabecalho, dtype=object)
    df.columns = [coluna(c) for c in df.columns]
    return df


def main():
    os.makedirs(caminhos.SAIDA, exist_ok=True)
    print("1/8 valores");  padroniza_valores.main(caminhos.VULTO_DB)  # noqa: E702
    print("\n2/8 produtos");  produtos.main()  # noqa: E702
    print("\n3/8 vendas");  padroniza_vendas.main()  # noqa: E702
    print("\n4/8 clientes e pontos");  clientes_pontos.main()  # noqa: E702
    print("\n5/8 estoque em 01/10");  estoque_01_10.main()  # noqa: E702
    print("\n6/8 histórico do estoque");  estoque_historico.main()  # noqa: E702
    print("\n7/8 atendimentos");  atendimentos.main()  # noqa: E702

    padronizado = sqlite3.connect(caminhos.PADRONIZADO)
    decisoes = pd.DataFrame(registro.decisoes(1, 10**6), columns=[coluna(c) for c in registro.COLUNAS])
    decisoes = decisoes.rename(columns={decisoes.columns[0]: "numero"})
    tabelas = {
        "vendas_padronizadas": pd.read_sql("select * from vendas_padronizadas", padronizado),
        "itens_venda_padronizados": pd.read_sql("select * from itens_venda", padronizado),
        "clientes_padronizados": aba("clientes atualizado v2.xlsx", "Clientes"),
        "extrato_pontos": aba("clientes atualizado v2.xlsx", "Extrato"),
        "estoque_01_10": aba("estoque_vulto_01-10-2026.xlsx", "Por produto"),
        "estoque_movimentos_set26": aba("estoque_vulto_01-10-2026.xlsx", "Movimentos"),
        "estoque_savassi_por_modelo": aba("estoque_vulto_01-10-2026.xlsx", "Savassi por modelo"),
        "estoque_historico": pd.read_sql("select * from estoque_historico", padronizado),
        "produtos_padronizados": pd.read_sql("select * from produtos", padronizado),
        "atendimentos_padronizados": pd.read_sql("select * from atendimentos", padronizado),
        "registro_decisoes": decisoes,
    }
    padronizado.close()
    print("\n8/8 conferência")
    validar.conferir(tabelas)
    assert not ORIGINAIS & set(tabelas), "nunca sobrescrever uma tabela original da Vulto"

    banco = sqlite3.connect(caminhos.VULTO_DB)
    for nome, df in tabelas.items():
        for c in df.columns:  # datas como texto ISO, para o SQLite
            if pd.api.types.is_datetime64_any_dtype(df[c]) or df[c].map(lambda v: hasattr(v, "isoformat")).any():
                df[c] = df[c].map(lambda v: v.isoformat(sep=" ") if hasattr(v, "isoformat") and not isinstance(v, str) else v)
        df.to_sql(nome, banco, if_exists="replace", index=False)
    banco.execute("vacuum")
    banco.close()

    print(f"\nGravado em {caminhos.VULTO_DB}:")
    for nome, df in tabelas.items():
        print(f"  {nome:28} {len(df):6} linhas")


if __name__ == "__main__":
    main()
