"""
Clientes e fidelidade (Vulto+) — decisões 1 a 5, 64, 65, 66 e 69 do Registro de Decisões.

Entrada : clientes_fidelidade_FINAL_v3 (1).xlsx, aba Clientes
Saída   : clientes_corrigidos.xlsx (Clientes, Alertas por e-mail, Ocorrências, Resumo)

A coluna Pontos da planilha foi congelada em 01/01/2025 (aba LEIA ISSO). Ela entra como saldo inicial no extrato
do Vulto+ (pontos.py); o saldo e a categoria de 01/10/2026 vêm de lá (função completar).
"""
import re
from datetime import datetime, timedelta

import pandas as pd

from comum import LOJAS, Registro, chave, cpf_valido, digitos, dinheiro, nome_proprio, telefone_padrao, vazio

ARQUIVO = "clientes_fidelidade_FINAL_v3 (1).xlsx"

# decisão 4 (revisada): a categoria segue a regra escrita na aba LEIA ISSO, calculada em pontos.py
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


# decisão 65 — consentimento para mensagens (LGPD: marketing só com 'sim' explícito)
CONSENTIMENTO = {"x": "Sim", "sim": "Sim", "1": "Sim", "s": "Sim", "nao": "Não", "n": "Não", "0": "Não"}
# decisão 69 — loja de cadastro escrita de 10 jeitos ('bh' = Savassi, 'cwb' = Batel)
LOJA_CADASTRO = {"app": 6, "site": 5, "ecommerce": 5, "pinheiros": 1, "eldorado": 2, "eldorado shop": 2,
                 "savassi": 3, "bh": 3, "batel": 4, "cwb": 4}


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

    # ---------- decisão 64: o mesmo e-mail em cadastros de pessoas diferentes
    for e, g in out[out.email.notna()].groupby("email"):
        if len(g) < 2:
            continue
        local = set(re.split(r"[._\-\d]+", e.split("@")[0])) - {""}
        pontuacao = {}
        for i, r in g.iterrows():                                   # primeiro nome no e-mail vale mais
            partes = chave(r["nome"]).split(" ")
            pontuacao[i] = 2 * (partes[0] in local) + len(local & set(partes[1:]))
        melhor = max(pontuacao.values())
        donos = [i for i, p in pontuacao.items() if p == melhor and p > 0]
        dono = donos[0] if len(donos) == 1 else None
        for i, r in g.iterrows():
            if i == dono:
                reg.add(64, "E-mail compartilhado: fica com quem tem o nome no e-mail", r["linha_planilha"], "e-mail",
                        e, e, f"{len(g)} cadastros usavam este e-mail")
                continue
            out.at[i, "email"] = None
            out.at[i, "problemas"] = r["problemas"] + ["E-mail era de outra pessoa (removido)"]
            reg.add(64, "E-mail de outra pessoa removido do cadastro", r["linha_planilha"], "e-mail", e, None,
                    "Mandar mensagem de um cliente para o e-mail de outro é vazamento de dado (LGPD)"
                    if dono is not None else "Não dá para saber de quem é o e-mail: removido de todos")

    # ---------- decisão 65: consentimento; decisão 69: loja de cadastro
    out["aceita_mensagens"] = out.aceita_msg.map(lambda x: CONSENTIMENTO.get(chave(x), "Não informado"))
    for _, r in out.iterrows():
        if vazio(r["aceita_msg"]):
            reg.add(65, "Consentimento em branco: tratado como não", r["linha_planilha"], "aceita_msg", None,
                    "Não informado", "Sem 'sim' explícito não entra em campanha (LGPD)")
        elif str(r["aceita_msg"]) not in ("Sim", "Não"):
            reg.add(65, "Consentimento padronizado", r["linha_planilha"], "aceita_msg", r["aceita_msg"],
                    r["aceita_mensagens"])
    out["loja_cadastro_id"] = out.loja_cadastro.map(lambda x: LOJA_CADASTRO.get(chave(x)))
    for _, r in out.iterrows():
        nome = LOJAS.get(r["loja_cadastro_id"])
        if nome and chave(r["loja_cadastro"]) != chave(nome):
            reg.add(69, "Loja de cadastro padronizada", r["linha_planilha"], "loja_cadastro", r["loja_cadastro"], nome)

    # ---------- 'Dados válidos' (Sim/Não) — se Não, a loja entra em contato com o cliente
    out["dados_validos"] = out["problemas"].apply(lambda p: "Não" if p else "Sim")
    out["motivo"] = out["problemas"].apply(lambda p: "; ".join(p) if p else None)
    colunas = ["linha_planilha", "nome", "cpf", "email", "telefone", "data_cadastro", "pontos",
               "dados_validos", "motivo", "unificado", "linhas_unificadas", "categoria_planilha", "obs",
               "aceita_mensagens", "loja_cadastro_id", "cpf_original", "email_original", "telefone_original",
               "data_cadastro_original", "pontos_original", "categoria_original", "aceita_msg", "loja_cadastro"]
    return dict(clientes=out[colunas].sort_values("linha_planilha"), alertas=pd.DataFrame(alertas)), reg


def completar(clientes, saldos, vendas, reg):
    """Junta o saldo e a categoria de 01/10/2026 (pontos.py) e cria os clientes que só existem no sistema."""
    s = saldos.set_index("cpf")
    out = clientes.rename(columns={"pontos": "pontos_planilha_01jan2025"}).copy()
    out["origem_cadastro"] = "Planilha Vulto+"
    # ---------- decisão 66: CPF com pontos ou compras no sistema, mas sem linha na planilha
    novos = sorted((set(s.index) | set(vendas.cliente_cpf.dropna())) - set(out.cpf))
    loja_compra = vendas.dropna(subset=["cliente_cpf"]).sort_values("data_hora").groupby("cliente_cpf").loja_id.first()
    extra = pd.DataFrame(dict(cpf=novos, origem_cadastro="Só no sistema (sem cadastro na planilha)",
                              pontos_planilha_01jan2025=0, aceita_mensagens="Não informado",
                              loja_cadastro_id=[loja_compra.get(c) for c in novos],
                              dados_validos="Não", motivo="Cadastro só no sistema: sem nome, e-mail e telefone"))
    for c in novos:
        reg.add(66, "Cliente que só existia no sistema", c, "cpf", None, c,
                "Cadastrado depois de 01/01/2025, quando a planilha foi congelada")
    out = pd.concat([out, extra], ignore_index=True)
    out["dados_validos"] = [("Não" if not cpf_valido(c) else d) for c, d in zip(out.cpf, out.dados_validos)]
    # ---------- decisão 4 (revisada): categoria pela regra escrita, saldo em 01/10/2026
    out["categoria"] = out.cpf.map(s.categoria).fillna("Bronze")
    for col in ("acumulado_12m", "saldo_sem_vencimento", "vencidos", "saldo", "valor_reais", "saldo_sistema"):
        out[col] = out.cpf.map(s[col]).fillna(0)
    out = out.rename(columns={"saldo": "saldo_pontos_01out2026", "valor_reais": "saldo_reais_01out2026",
                              "saldo_sistema": "saldo_no_sistema_vulto", "acumulado_12m": "pontos_ganhos_12m"})
    for _, r in out[out.origem_cadastro == "Planilha Vulto+"].iterrows():
        if r["categoria_planilha"] != r["categoria"]:
            reg.add(4, "Categoria pela regra escrita (pontos ganhos em 12 meses)", r["linha_planilha"], "categoria",
                    r["categoria_original"], r["categoria"], f"{int(r['pontos_ganhos_12m'])} pontos ganhos em 12 meses")
    return out


if __name__ == "__main__":
    import sys
    from comum import salvar_excel
    dados, reg = processar(sys.argv[1] if len(sys.argv) > 1 else "../vulto-legado-back/data")
    salvar_excel("clientes_corrigidos.xlsx", {"Clientes": dados["clientes"], "Alertas por e-mail": dados["alertas"],
                                              "Ocorrências": reg.df(), "Resumo": reg.resumo()})
    print(reg.resumo().to_string())
