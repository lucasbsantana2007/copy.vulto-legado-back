"""
Clientes e fidelidade (Vulto+) — decisões 1 a 5 do Registro de Decisões.

Entrada : clientes_fidelidade_FINAL_v3 (1).xlsx, aba Clientes
Saída   : clientes_corrigidos.xlsx (Clientes, Alertas por e-mail, Ocorrências, Resumo)
"""
import re
from datetime import datetime, timedelta

import pandas as pd

from comum import Registro, chave, cpf_valido, digitos, dinheiro, nome_proprio, telefone_padrao, vazio

ARQUIVO = "clientes_fidelidade_FINAL_v3 (1).xlsx"

# decisão 4 — faixas de pontos
FAIXAS = [(0, 0, "Nada"), (1, 1500, "Bronze"), (1501, 5000, "Prata"), (5001, 10000, "Ouro"), (10001, None, "VIP")]
CATEGORIA_TEXTO = {
    "bronze": "Bronze", "brz": "Bronze",
    "silver": "Prata", "prata": "Prata", "prata*": "Prata",
    "ouro": "Ouro", "ouro?": "Ouro", "gold": "Ouro", "ouro ★": "Ouro",
    "vip": "VIP", "vip?": "VIP",
}
# decisão 5 — erros de digitação de domínio
DOMINIOS = {"gmial.com": "gmail.com", "hotmail.con": "hotmail.com", "uol.con.br": "uol.com.br"}
SEM_EMAIL = {"", "-", "nao tem", "não tem"}
MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10,
         "nov": 11, "dez": 12}


def categoria_por_pontos(p):
    for ini, fim, nome in FAIXAS:
        if p >= ini and (fim is None or p <= fim):
            return nome


def pontos(v):
    """Só o número: '2.203' -> 2203 | '1389 pts' -> 1389 | '2,7k' -> 2700 | vazio -> None."""
    if vazio(v):
        return None
    s = str(v).strip().lower().replace("pts", "").strip()
    m = re.fullmatch(r"(\d+(?:[.,]\d+)?)\s*k", s)
    if m:
        return int(round(float(m.group(1).replace(",", ".")) * 1000))
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        return int(s.replace(".", ""))
    n = dinheiro(s)
    return None if n is None else int(round(n))


def data_cadastro(v):
    if vazio(v):
        return None
    s = str(v).strip()
    if re.fullmatch(r"\d{5}", s):                                   # número de série do Excel
        return (datetime(1899, 12, 30) + timedelta(days=int(s))).date()
    m = re.fullmatch(r"([a-z]{3})/(\d{2})", s.lower())
    if m and m.group(1) in MESES:                                   # 'ago/21' -> 01/08/2021
        return datetime(2000 + int(m.group(2)), MESES[m.group(1)], 1).date()
    for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def email(v, reg, linha):
    if vazio(v) or chave(v) in SEM_EMAIL:
        if not vazio(v):
            reg.add(5, "E-mail ausente", linha, "e-mail", v, None, "Marcador de ausência ('-', 'nao tem') removido")
        else:
            reg.add(5, "E-mail ausente", linha, "e-mail", None, None, "Cliente sem e-mail")
        return None
    s = str(v)
    e = s.strip().lower()
    if s != s.rstrip():
        reg.add(5, "Espaço sobrando", linha, "e-mail", s, e)
    if s.strip() == s.strip().upper() and s.strip() != e:
        reg.add(5, "E-mail em maiúsculas", linha, "e-mail", s, e)
    for errado, certo in DOMINIOS.items():
        if e.endswith("@" + errado):
            e = e[: -len(errado)] + certo
            reg.add(5, f"Domínio digitado errado ({errado})", linha, "e-mail", s, e)
    return e if re.fullmatch(r"[^@\s]+@[^@\s]+\.[a-z]{2,}(\.[a-z]{2})?", e) else None


def processar(pasta):
    reg = Registro("Clientes")
    bruto = pd.read_excel(f"{pasta}/{ARQUIVO}", sheet_name="Clientes", dtype=str)
    linhas = []
    for i, r in bruto.iterrows():
        linha = i + 2                                               # linha na planilha original
        d = digitos(r["cpf"])
        problemas = []
        # ---------- decisão 1: CPF incompleto
        if len(d) < 11:
            cpf = d.zfill(11)
            reg.add(1, f"CPF com {len(d)} dígitos", linha, "cpf", r["cpf"], cpf,
                    "Zeros à esquerda perdidos; completado para comparar cadastros, mas sinalizado para confirmação")
            problemas.append("CPF incompleto na origem: confirmar com o cliente")
        else:
            cpf = d
        if not cpf_valido(cpf):
            problemas.append("CPF com dígito verificador inválido")
            reg.add(1, "CPF inválido", linha, "cpf", r["cpf"], cpf, "Dígito verificador não confere")
        # ---------- decisão 3: pontos
        p = pontos(r["Pontos"])
        if p is None:
            reg.add(3, "Sem pontos", linha, "Pontos", r["Pontos"], 0, "Cliente sem pontos registrados (vazio)")
        elif str(r["Pontos"]).strip() != str(p):
            reg.add(3, "Pontos padronizados", linha, "Pontos", r["Pontos"], p)
        # ---------- decisão 5: e-mail
        e = email(r["E-mail"], reg, linha)
        if e is None:
            problemas.append("Sem e-mail válido")
        # ---------- decisão 4: categoria escrita na planilha (referência)
        cat_txt = CATEGORIA_TEXTO.get(chave(r["categoria"]).replace(" ", " ")) if not vazio(r["categoria"]) else None
        linhas.append(dict(
            linha_planilha=linha, nome=r["Nome"], cpf=cpf, cpf_original=r["cpf"], email=e, email_original=r["E-mail"],
            telefone=telefone_padrao(r["Tel"]) or (None if vazio(r["Tel"]) else str(r["Tel"]).strip()),
            telefone_original=r["Tel"], data_cadastro=data_cadastro(r["Data Cadastro"]),
            data_cadastro_original=r["Data Cadastro"], pontos=p or 0, pontos_original=r["Pontos"],
            categoria_planilha=cat_txt, categoria_original=r["categoria"], obs=r["Obs"],
            aceita_msg=r["aceita_msg"], loja_cadastro=r["loja_cadastro"], problemas=problemas,
        ))
    df = pd.DataFrame(linhas)

    # ---------- decisão 2: CPFs repetidos
    alertas, manter = [], []
    for cpf, g in df.groupby("cpf", sort=False):
        if len(g) == 1:
            manter.append(g.iloc[0].to_dict() | {"unificado": "Não", "linhas_unificadas": None})
            continue
        primeiros = {chave(n).split(" ")[0] for n in g["nome"]}
        if len(primeiros) > 1:                                      # mesmo CPF, pessoas diferentes: não junta
            for _, r in g.iterrows():
                reg.add(2, "CPF repetido com nomes diferentes", r["linha_planilha"], "cpf", cpf, cpf,
                        "Não unificado; sinalizado para confirmação")
                manter.append(r.to_dict() | {"unificado": "Não", "linhas_unificadas": None,
                                             "problemas": r["problemas"] + ["CPF repetido com outro nome"]})
            continue
        g = g.sort_values(["pontos", "data_cadastro"], ascending=[False, True], na_position="last")
        base = g.iloc[0].to_dict()                                  # e-mail do cadastro com mais pontos
        total = int(g["pontos"].sum())
        datas = [x for x in g["data_cadastro"] if x]
        base.update(pontos=total, data_cadastro=min(datas) if datas else None, unificado="Sim",
                    linhas_unificadas=", ".join(str(x) for x in g["linha_planilha"]),
                    problemas=sorted({p for ps in g["problemas"] for p in ps}))
        if base["email"] is None:                                   # base sem e-mail: usa outro válido
            outros = [x for x in g["email"] if x]
            base["email"] = outros[0] if outros else None
        manter.append(base)
        reg.add(2, "CPF repetido unificado", base["linha_planilha"], "cpf", base["cpf_original"], cpf,
                f"Linhas {base['linhas_unificadas']} viraram 1 cliente; pontos somados ({total}); "
                f"e-mail do cadastro com mais pontos")
        alertas.append(dict(nome=base["nome"], cpf=cpf, email_mantido=base["email"],
                            emails_dos_cadastros=", ".join(sorted({x for x in g["email"] if x})),
                            pontos_por_cadastro=", ".join(str(x) for x in g["pontos"]), pontos_total=total,
                            linhas_unificadas=base["linhas_unificadas"],
                            mensagem="Seus cadastros Vulto+ com o mesmo CPF foram unificados e os pontos somados."))
    out = pd.DataFrame(manter)

    # ---------- decisão 4: categoria pela faixa de pontos (depois de somar os duplicados)
    out["categoria"] = out["pontos"].apply(categoria_por_pontos)
    for _, r in out.iterrows():
        if r["categoria_planilha"] and r["categoria_planilha"] != r["categoria"]:
            reg.add(4, "Categoria diferente da faixa de pontos", r["linha_planilha"], "categoria",
                    r["categoria_original"], r["categoria"], f"{r['pontos']} pontos")
        elif not vazio(r["categoria_original"]) and r["categoria_original"] != r["categoria"]:
            reg.add(4, "Categoria padronizada", r["linha_planilha"], "categoria", r["categoria_original"],
                    r["categoria"])
        elif vazio(r["categoria_original"]):
            reg.add(4, "Categoria em branco preenchida", r["linha_planilha"], "categoria", None, r["categoria"])

    # ---------- 'Dados válidos' (Sim/Não) — se Não, a loja entra em contato com o cliente
    out["dados_validos"] = out["problemas"].apply(lambda p: "Não" if p else "Sim")
    out["motivo"] = out["problemas"].apply(lambda p: "; ".join(p) if p else None)
    colunas = ["linha_planilha", "nome", "cpf", "email", "telefone", "data_cadastro", "pontos", "categoria",
               "dados_validos", "motivo", "unificado", "linhas_unificadas", "categoria_planilha", "obs",
               "aceita_msg", "loja_cadastro", "cpf_original", "email_original", "telefone_original",
               "data_cadastro_original", "pontos_original", "categoria_original"]
    return dict(clientes=out[colunas].sort_values("linha_planilha"), alertas=pd.DataFrame(alertas)), reg


if __name__ == "__main__":
    import sys
    from comum import salvar_excel
    dados, reg = processar(sys.argv[1] if len(sys.argv) > 1 else "../vulto-legado-back/data")
    salvar_excel("clientes_corrigidos.xlsx", {"Clientes": dados["clientes"], "Alertas por e-mail": dados["alertas"],
                                              "Ocorrências": reg.df(), "Resumo": reg.resumo()})
    print(reg.resumo().to_string())
