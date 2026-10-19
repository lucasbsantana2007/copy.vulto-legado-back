import os
import pandas as pd
from openpyxl import load_workbook

PLANILHA = os.path.join(os.path.dirname(__file__), "data", "clientes_fidelidade_FINAL_v3 (1).xlsx")

# cache pra nao ler toda hora (as vezes da problema, ai reinicia o servidor)
_cache = None


def ler_clientes(force=False):
    global _cache
    if _cache is not None and not force:
        return _cache
    df = pd.read_excel(PLANILHA, sheet_name="Clientes", dtype=object)
    df = df.fillna("")
    _cache = df
    return df


def limpa_cpf(cpf):
    # tira ponto e traço
    return str(cpf).replace(".", "").replace("-", "")


def cpf_texto(cpf):
    # 000.000.000-00: devolve o zero da frente que o Excel corta e, com os pontos, o Excel não lê como número de novo
    digitos = "".join(c for c in str(cpf) if c.isdigit())
    if 9 <= len(digitos) <= 11:
        digitos = digitos.zfill(11)
        return f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    return digitos


def parse_pontos(p):
    try:
        if p == "" or p is None:
            return 0
        if isinstance(p, (int, float)):
            return int(p)
        p = str(p).replace("pts", "").replace(" ", "")
        p = p.replace(".", "")  # 1.200 -> 1200
        return int(p)
    except:
        print("erro pontos", p)
        return 0


def busca_cliente(cpf):
    df = ler_clientes()
    cpf = limpa_cpf(cpf)
    for i, row in df.iterrows():
        if limpa_cpf(row["cpf"]) == cpf:
            return row.to_dict()
    return None


def anota_resgate(data, cliente, pontos, valor, loja="Sistema", atendente=""):
    # as lojas usam essa aba, entao o sistema escreve aqui tb pra ficar tudo junto
    wb = load_workbook(PLANILHA)
    ws = wb["Resgates"]
    ws.append([data, cliente, pontos, "R$ " + str(valor).replace(".", ","), loja, atendente])
    wb.save(PLANILHA)
