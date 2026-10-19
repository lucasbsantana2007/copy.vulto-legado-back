"""
Atendimento (export WhatsApp / Instagram / e-mail) — decisões 21 a 27 e 67 do Registro de Decisões.

Entrada : atendimentos_export_whatsapp_insta.csv (latin-1, separador ';', primeira linha 'sep=')
          vulto.db (opcional: só para conferir se o número do pedido existe)
Saída   : atendimento_corrigido.xlsx (Atendimentos, Ocorrências, Resumo)
"""
import os
import re
import sqlite3
from collections import Counter
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import pandas as pd

from comum import Registro, chave, digitos, nome_proprio, telefone_padrao, vazio

ARQUIVO = "atendimentos_export_whatsapp_insta.csv"
FUSO = ZoneInfo("America/Sao_Paulo")
CANAIS = {"whatsapp": "WhatsApp", "wpp": "WhatsApp", "instagram": "Instagram", "insta": "Instagram",
          "instagram dm": "Instagram", "e-mail": "Email", "email": "Email"}
ATENDENTES = {"bot": "Bot Vulto", "bot vulto": "Bot Vulto", "julia": "Júlia", "julia (sac)": "Júlia",
              "renata": "Renata", "renata.sac": "Renata", "marcos sac": "Marcos", "time insta": "Time Insta"}
# decisão 23 — 9 categorias de assunto
ASSUNTOS = {
    "troca": "Troca", "devolucao": "Devolução", "troca/devolucao": None,          # None = decidir pela mensagem
    "pedido": "Entrega e pedido", "entrega": "Entrega e pedido", "atraso entrega": "Entrega e pedido",
    "duvida produto": "Produto e estoque", "estoque": "Produto e estoque",
    "pontos": "Vulto+", "pontos vulto+": "Vulto+", "fidelidade": "Vulto+",
    "pagamento": "Pagamento", "cancelamento": "Cancelamento", "reclamacao": "Reclamação",
    "loja": "Outros", "elogio": "Outros", "spam": "Outros",
}
TAG_ASSUNTO = {"entrega": "Entrega e pedido", "pre-venda": "Produto e estoque", "vulto+": "Vulto+",
               "pagamento": "Pagamento", "loja": "Outros"}
# decisão 26 — status padronizado em 4 valores
STATUS = {"ok": "Resolvido", "resolvido": "Resolvido", "fechado": "Resolvido", "finalizado": "Resolvido",
          "aguardando cliente": "Aguardando cliente", "em andamento": "Em andamento", "pendente": "Aberto",
          "aberto": "Aberto", "": "Aberto"}
STATUS_PELA_RESPOSTA = [("Resolvido!", "Resolvido"), ("Pode me passar", "Aguardando cliente"),
                        ("Já abrimos", "Em andamento"), ("Vou verificar", "Em andamento")]
GENERICOS = {"", "cliente", "?", "nan"}


def data_hora(v, reg, id_):
    s = str(v).strip()
    try:
        dt = datetime.fromisoformat(s)                                # '2026-02-01T09:08:00-03:00'
        dt = dt.astimezone(FUSO) if dt.tzinfo else dt
    except ValueError:
        dt = datetime.strptime(s, "%d/%m/%Y %H:%M")
    novo = dt.strftime("%d/%m/%Y %H:%M:%S")
    if "T" in s:
        reg.add(21, "Data em formato ISO convertida", id_, "data_hora", s, novo)
    return novo


def separar_mensagem(m):
    """Decisão 24: texto do cliente e resposta do atendente em colunas separadas."""
    cliente, resposta = [], []
    for bloco in re.split(r"\n(?=\[(?:Atendente|Bot|Cliente)\]:)", str(m or "")):
        b = bloco.strip()
        if b.startswith("[Atendente]:") or b.startswith("[Bot]:"):
            resposta.append(b.split("]:", 1)[1].strip())
        elif b.startswith("[Cliente]:"):
            cliente.append(b.split("]:", 1)[1].strip())
        elif b:
            cliente.append(b)
    return "\n".join(cliente) or None, "\n".join(resposta) or None


def ultima_resposta_atendente(m):
    r = re.findall(r"\[(Atendente|Bot)\]: (.*)", str(m or ""))
    return r[-1] if r else None


def pedidos_existentes(pasta):
    caminho = os.path.join(pasta, "vulto.db")
    if not os.path.exists(caminho):
        return None
    with sqlite3.connect(caminho) as c:
        pes = {p for (p,) in c.execute("select pedido_externo from vendas where pedido_externo <> ''")}
        pes |= {str(i) for (i,) in c.execute("select id from vendas")}
    return pes


def formato_pedido(p):
    """3 formatos: site S10xxxx, app APP-70xxxxx, loja 4 a 5 dígitos."""
    d = digitos(p)
    if str(p).upper().startswith("S") or (len(d) == 6 and d.startswith("10")):
        return f"S{d}"
    if str(p).upper().startswith("APP") or (len(d) == 7 and d.startswith("70")):
        return f"APP-{d}"
    return d


def pedido(v, msg, existentes, reg, id_):
    """Decisão 25: limpa prefixos, recupera do texto e corrige dígitos invertidos."""
    bruto = None if vazio(v) else str(v).strip()
    if bruto is None:
        texto = str(msg or "")
        m = (re.search(r"(?i)(APP-\d+|\bS\d{6}\b|#\d{4,7})", texto)
             or re.search(r"(?i)pedido\D{0,20}?(\d{4,7})(?!\s*(?:pontos|reais))", texto))
        if not m:
            return None
        bruto = m.group(1)
        recuperado = True
    else:
        recuperado = False
    limpo = formato_pedido(re.sub(r"(?i)^(#|pedido\s*(?:n[ºo°]?)?\s*#?)", "", bruto).strip())
    if existentes is not None and limpo not in existentes:
        invertido = formato_pedido(digitos(limpo)[::-1])
        # decisão 25 (revisada): só vale para pedido do site/app. Número de venda de loja vai de 1 a ~18.600, então
        # quase qualquer número invertido 'existe' e a correção seria um chute.
        if invertido in existentes and invertido[0] in "SA":
            reg.add(25, "Pedido com dígitos invertidos", id_, "pedido", bruto, invertido)
            return invertido
        if invertido in existentes:
            reg.add(25, "Inversão não aplicada (número de loja: não dá para confirmar)", id_, "pedido", bruto, limpo)
    if recuperado:
        reg.add(25, "Pedido recuperado do texto da mensagem", id_, "pedido", None, limpo)
    elif limpo != bruto:
        reg.add(25, "Prefixo do pedido removido / formato padronizado", id_, "pedido", bruto, limpo)
    return limpo


def nome_pelo_contato(contato):
    """Decisão 22: deduz o nome pelo e-mail ou @ (ex.: @nicole.lima -> Nicole Lima)."""
    base = str(contato).lstrip("@").split("@")[0]
    partes = [p for p in re.split(r"[._-]", re.sub(r"\d+", "", base)) if len(p) > 1]
    if not partes or partes[0] in ("user", "contato", "cliente"):
        return None
    return nome_proprio(" ".join(partes))


def processar(pasta):
    reg = Registro("Atendimento")
    bruto = pd.read_csv(os.path.join(pasta, ARQUIVO), sep=";", encoding="latin-1", skiprows=1, dtype=str)
    dup = bruto.duplicated(keep="first")
    for _, r in bruto[dup].iterrows():
        reg.add(21, "Linha duplicada removida", r["id"], "id", r["id"], None, "Mesma linha exportada 2x")
    df = bruto[~dup].copy()
    existentes = pedidos_existentes(pasta)

    linhas = []
    for _, r in df.iterrows():
        id_ = r["id"]
        canal = CANAIS.get(chave(r["canal"]), r["canal"])
        atendente = ATENDENTES.get(chave(r["atendente"]), r["atendente"])
        if canal != r["canal"]:
            reg.add(21, "Canal padronizado", id_, "canal", r["canal"], canal)
        if atendente != r["atendente"]:
            reg.add(21, "Atendente padronizado", id_, "atendente", r["atendente"], atendente)
        # contato
        c = str(r["contato"]).strip()
        if "@" in c and not c.startswith("@"):
            contato = c.lower()
        elif c.startswith("@"):
            contato = c.lower()
        else:
            contato = telefone_padrao(c) or c                         # sem DDD: mantido como estava
            if contato == c and len(digitos(c)) in (8, 9):
                reg.add(22, "Telefone sem DDD mantido", id_, "contato", c, c)
            elif contato != c:
                reg.add(22, "Telefone padronizado", id_, "contato", c, contato)
        # assunto
        msg_cliente, resposta = separar_mensagem(r["mensagem"])
        assunto = classificar_assunto(r["assunto"], r["tags"], msg_cliente)
        if assunto != r["assunto"]:
            reg.add(23, "Assunto reclassificado", id_, "assunto", r["assunto"], assunto)
        if re.search(r"[À-ÿ]", str(r["mensagem"])):
            reg.add(24, "Acentuação recuperada", id_, "mensagem", None, None,
                    "Arquivo lido como latin-1 (no Excel os acentos apareciam quebrados)")
        # status
        base = STATUS.get(chave(r["status"]), "Aberto")
        ult = ultima_resposta_atendente(r["mensagem"])
        status = base
        if ult and ult[0] == "Atendente":
            status = next((s for k, s in STATUS_PELA_RESPOSTA if k in ult[1]), base)
        if status != base:
            reg.add(26, "Status corrigido pela última resposta do atendente", id_, "status", r["status"], status,
                    f"Última resposta: '{ult[1][:45]}'")
            if base == "Resolvido":
                reg.add(27, "Marcado como resolvido, mas conversa em aberto", id_, "status", r["status"], status)
        elif not vazio(r["status"]) and r["status"] != status:
            reg.add(26, "Status padronizado", id_, "status", r["status"], status)
        linhas.append(dict(
            id=int(id_), data_hora=data_hora(r["data_hora"], reg, id_), canal=canal, contato=contato,
            nome=r["nome_contato"], assunto=assunto, mensagem_cliente=msg_cliente, resposta_atendente=resposta,
            status=status, atendente=atendente, pedido=pedido(r["pedido"], r["mensagem"], existentes, reg, id_),
            status_original=r["status"], assunto_original=r["assunto"], contato_original=r["contato"],
            nome_original=r["nome_contato"], pedido_original=r["pedido"],
        ))
    out = pd.DataFrame(linhas)

    # decisão 22: nome unificado por cliente (mesmo contato) e deduzido pelo e-mail/@ quando faltar
    nomes = {}
    for contato, g in out.groupby("contato"):
        bons = [nome_proprio(n) for n in g["nome"] if chave(n) not in GENERICOS]
        nomes[contato] = Counter(bons).most_common(1)[0][0] if bons else None
    final = []
    for _, r in out.iterrows():
        n = nomes[r["contato"]]
        if n is None:
            n = nome_pelo_contato(r["contato"]) if not r["contato"].startswith("+") else None
            if n:
                reg.add(22, "Nome deduzido pelo e-mail ou @", r["id"], "nome", r["nome_original"], n)
            else:
                n = "Não informado"
                reg.add(22, "Contato sem identificação", r["id"], "nome", r["nome_original"], n)
        elif n != r["nome_original"]:
            reg.add(22, "Nome corrigido / unificado por cliente", r["id"], "nome", r["nome_original"], n)
        final.append(n)
    out["nome"] = final
    return dict(atendimentos=out.sort_values("id")), reg


def ligar_clientes(atendimentos, clientes, vendas, reg):
    """Decisão 67: cada atendimento ligado ao cliente e ao pedido certo (pedido da Gerente de Atendimento).

    Quem está falando é identificado pelo contato (e-mail ou celular do cadastro). Se o contato não identifica
    (ex.: @ do Instagram), usa o cliente do pedido citado. O pedido só é ligado se existir e, quando os dois lados
    têm cliente, se for do mesmo cliente e anterior ao atendimento."""
    a = atendimentos.copy()
    por_email = dict(zip(clientes.email.dropna(), clientes.loc[clientes.email.notna(), "cpf"]))
    tel = clientes.dropna(subset=["telefone"]).drop_duplicates("telefone", keep=False)
    por_tel = dict(zip(tel.telefone, tel.cpf))
    v = vendas.assign(chave_pedido=vendas.pedido_externo.fillna(vendas.id.astype(str)))
    v = v.drop_duplicates("chave_pedido", keep="first").set_index("chave_pedido")
    cpfs, como, venda_ids, confere, lojas = [], [], [], [], []
    for r in a.itertuples():
        quando = datetime.strptime(r.data_hora, "%d/%m/%Y %H:%M:%S")
        pelo_contato = por_email.get(r.contato) if "@" in str(r.contato) and not str(r.contato).startswith("@") \
            else por_tel.get(r.contato)
        tem_pedido = not vazio(r.pedido)
        venda = v.loc[r.pedido] if (tem_pedido and r.pedido in v.index) else None
        pelo_pedido = venda.cliente_cpf if venda is not None and not pd.isna(venda.cliente_cpf) else None
        cpf = pelo_contato or pelo_pedido
        como.append("E-mail/celular do cadastro" if pelo_contato else ("Cliente do pedido" if pelo_pedido else None))
        if venda is None:
            ok = "Pedido não encontrado" if tem_pedido else None
        elif venda.data_hora > quando + timedelta(days=1):
            ok = "Não (pedido depois do atendimento)"
        elif pelo_contato and pelo_pedido and pelo_contato != pelo_pedido:
            ok = "Não (pedido de outro cliente)"
        elif pelo_contato and pelo_pedido:
            ok = "Sim"
        else:
            ok = "Não verificável (falta cliente num dos lados)"
        cpfs.append(cpf)
        confere.append(ok)
        venda_ids.append(venda.id if venda is not None and not str(ok).startswith("Não (") else None)
        lojas.append(int(venda.loja_id) if venda is not None else 5)   # sem pedido: SAC central (loja do site)
        reg.add(67, "Atendimento ligado ao cliente" if cpf else "Atendimento sem cliente identificado", r.id,
                "cliente", r.contato, cpf, como[-1] or "contato não está no cadastro e não há pedido")
        if ok and ok.startswith("Não ("):
            reg.add(67, "Pedido citado não confere: não ligado", r.id, "pedido", r.pedido, None, ok)
    a["cliente_cpf"], a["cliente_identificado_por"], a["venda_id"] = cpfs, como, venda_ids
    a["pedido_confere"], a["loja_id"] = confere, lojas
    return a


def classificar_assunto(assunto, tags, mensagem):
    texto = chave(mensagem)
    if not vazio(assunto):
        a = ASSUNTOS.get(chave(assunto), "Outros")
        if a is not None:
            return a
        # 'troca/devolução': a coluna de tags não separa os dois; decide pelo pedido do cliente
        return "Devolução" if re.search(r"devolv|devolu|dinheiro de volta|estorn|reembols", texto) else "Troca"
    # sem assunto: primeiro o texto da mensagem, depois as tags
    for padrao, cat in [(r"devolv|devolu|dinheiro de volta|reembols", "Devolução"), (r"\btroc", "Troca"),
                        (r"pedido|chegou|rastreio|entrega", "Entrega e pedido"), (r"ponto|vulto\+", "Vulto+"),
                        (r"pagamento|pix|boleto|cartao", "Pagamento"), (r"cancel", "Cancelamento")]:
        if re.search(padrao, texto):
            return cat
    for tag, cat in TAG_ASSUNTO.items():
        if tag in chave(tags):
            return cat
    return "Outros"


if __name__ == "__main__":
    import sys
    from comum import salvar_excel
    dados, reg = processar(sys.argv[1] if len(sys.argv) > 1 else "../vulto-legado-back/data")
    salvar_excel("atendimento_corrigido.xlsx", {"Atendimentos": dados["atendimentos"], "Ocorrências": reg.df(),
                                                "Resumo": reg.resumo()})
    print(reg.resumo().to_string())
