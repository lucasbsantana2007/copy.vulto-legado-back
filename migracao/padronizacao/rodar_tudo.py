"""
Roda a padronização completa (planilhas + vulto.db) e gera, na pasta de saída:

  vulto_corrigido.db            o vulto.db consertado: as MESMAS tabelas e colunas, só os dados corrigidos
  planilhas_e_relatorios.db     o que nunca existiu no vulto.db: planilhas corrigidas, histórico de estoque,
                                Registro de Decisões e o log de cada alteração
  numeros_diretoria.xlsx        os três números, cada um com a comparação "sistema da Vulto × nosso número"
  registro_decisoes.xlsx        o Registro de Decisões (as 7 colunas do enunciado) + detalhe por tipo
  clientes_corrigidos.xlsx      clientes com saldo e categoria em 01/10/2026
  catalogo_corrigido.xlsx       catálogo por coleção + cadastro mestre (1 linha por SKU)
  vendas_corrigidas.xlsx        vendas do vulto.db e da planilha da Batel, itens e faturamento
  pontos_vulto_mais.xlsx        extrato corrigido, saldos e passivo
  estoque_historico.xlsx        movimentações (histórico), posição em 01/10/2026 e qualidade por loja
  atendimento_corrigido.xlsx    atendimentos ligados ao cliente e ao pedido
  inventario_corrigido.xlsx     contagem física de set/26 padronizada
  vendas_batel_corrigidas.xlsx  vendas manuais da Batel

Uso:
    python rodar_tudo.py <pasta data da Vulto> [pasta de saída]

Os arquivos originais não são alterados (o relatório antigo precisa continuar reproduzível).
"""
import os
import sqlite3
import sys
from datetime import date, datetime

import pandas as pd

import atendimento
import banco_corrigido
import catalogo
import clientes
import decisoes
import estoque
import inventario
import pontos
import vendas
import vendas_batel
from comum import LOJAS, salvar_excel


def brl(x):
    s = f"{abs(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−" if x < 0 else "") + "R$ " + s


def num(x):
    return f"{x:,.0f}".replace(",", ".")


COMMIT_ORIGINAL = "a792f97"            # commit do copy.vulto-legado-back com o vulto.db original


def ja_corrigido(pasta):
    """O vulto.db do repositório foi trocado pela versão corrigida. Corrigir de novo estragaria os dados (ex.: dividir
    por 100 o que já está em reais), então o script precisa do original."""
    with sqlite3.connect(os.path.join(pasta, "vulto.db")) as c:
        utc = c.execute("select count(*) from vendas where criado_em like '%Z'").fetchone()[0]
        ponto = c.execute("select valor from config where chave = 'valor_ponto'").fetchone()
    return utc == 0 and ponto and ponto[0] == "0.05"


def main(pasta, saida):
    if ja_corrigido(pasta):
        sys.exit(f"""
O vulto.db em {pasta} já é a versão CORRIGIDA. Rode o script com o original, que está no histórico do git:

    mkdir -p entrada
    cp {pasta}/*.xlsx {pasta}/*.csv entrada/
    git -C {os.path.dirname(os.path.abspath(pasta))} show {COMMIT_ORIGINAL}:data/vulto.db > entrada/vulto.db
    python rodar_tudo.py entrada {saida}
""")
    os.makedirs(saida, exist_ok=True)
    caminho = lambda nome: os.path.join(saida, nome)
    registros = []

    def etapa(nome, resultado):
        dados, reg = resultado
        registros.append(reg)
        print(f"{nome} ok")
        return dados, reg

    cli, reg_cli = etapa("clientes", clientes.processar(pasta))
    cat, reg_cat = etapa("catálogo", catalogo.processar(pasta))
    bat, reg_bat = etapa("vendas Batel", vendas_batel.processar(pasta))
    ven, reg_ven = etapa("vendas (vulto.db)", vendas.processar(pasta, cat["mestre"], cli["clientes"], bat["vendas"]))
    cat["mestre"] = incluir_skus_criados(cat["mestre"], ven["skus_criados"])
    pts, reg_pts = etapa("Vulto+", pontos.processar(pasta, cli["clientes"], ven["vendas"], bat["vendas"]))
    clientes_final = clientes.completar(cli["clientes"], pts["saldos"], ven["vendas"], reg_cli)
    ate, reg_ate = etapa("atendimento", atendimento.processar(pasta))
    atendimentos = atendimento.ligar_clientes(ate["atendimentos"], clientes_final, ven["vendas"], reg_ate)
    inv, reg_inv = etapa("inventário", inventario.processar(pasta, cat["mestre"], cat["skus_2024"]))
    est, reg_est = etapa("estoque", estoque.processar(pasta, ven["vendas"], ven["itens"], inv["consolidado"],
                                                      cat["mestre"]))

    # ------------------------------------------------------------------ arquivos por domínio
    salvar_excel(caminho("clientes_corrigidos.xlsx"), {"Clientes": clientes_final, "Alertas por e-mail": cli["alertas"],
                                                      "Ocorrências": reg_cli.df(), "Resumo": reg_cli.resumo()})
    salvar_excel(caminho("catalogo_corrigido.xlsx"), {"Catálogo": cat["catalogo"], "Cadastro mestre": cat["mestre"],
                                                     "Listas": cat["listas"], "Ocorrências": reg_cat.df(),
                                                     "Resumo": reg_cat.resumo()},
                 destaques={"Catálogo": catalogo.destaque_catalogo, "Cadastro mestre": catalogo.destaque_catalogo})
    salvar_excel(caminho("vendas_batel_corrigidas.xlsx"), {"Vendas": bat["vendas"], "Totais por semana": bat["totais"],
                                                          "Ocorrências": reg_bat.df(), "Resumo": reg_bat.resumo()},
                 destaques={"Vendas": vendas_batel.destaque})
    salvar_excel(caminho("vendas_corrigidas.xlsx"), {"Vendas": ven["vendas"], "Itens": ven["itens"],
                                                    "Itens DIVERSO (preços)": ven["diversos"],
                                                    "Faturamento": ven["faturamento"], "Ocorrências": reg_ven.df(),
                                                    "Resumo": reg_ven.resumo()})
    salvar_excel(caminho("pontos_vulto_mais.xlsx"), {"Passivo": pts["ponte"]["ponte"], "Saldos": pts["saldos"],
                                                    "Extrato": pts["extrato"], "Por tipo": pts["ponte"]["por_tipo"],
                                                    "Ocorrências": reg_pts.df(), "Resumo": reg_pts.resumo()})
    salvar_excel(caminho("atendimento_corrigido.xlsx"), {"Atendimentos": atendimentos, "Ocorrências": reg_ate.df(),
                                                        "Resumo": reg_ate.resumo()})
    abas_inv = {"Estoque consolidado": inv["consolidado"], "Resumo por loja": inv["resumo_lojas"]}
    for loja in inv["consolidado"].loja.unique():
        abas_inv[loja.upper()] = inv["consolidado"][inv["consolidado"].loja == loja]
    abas_inv.update({"AVARIA PINHEIROS": inv["avaria"], "Ocorrências": reg_inv.df(), "Resumo": reg_inv.resumo()})
    salvar_excel(caminho("inventario_corrigido.xlsx"), abas_inv,
                 destaques={k: inventario.destaque for k in abas_inv if k.isupper() or k == "Estoque consolidado"})
    salvar_excel(caminho("estoque_historico.xlsx"), {"Resumo por loja": est["resumo"], "Posição 01-10-2026": est["posicao"],
                                                    "Posição 17-09-2026": est["exemplo_17set"],
                                                    "Movimentações": est["movimentacoes"],
                                                    "Ocorrências": reg_est.df(), "Resumo": reg_est.resumo()})

    # ------------------------------------------------------------------ os três números
    numeros = tres_numeros(ven, pts, est)
    salvar_excel(caminho("numeros_diretoria.xlsx"), numeros)

    # ------------------------------------------------------------------ Registro de Decisões
    ocorr = pd.concat([r.df() for r in registros], ignore_index=True)
    detalhe = pd.concat([r.resumo() for r in registros], ignore_index=True)
    registro = decisoes.tabela(ocorr, impactos(ven, pts, est, cli, clientes_final, cat, ocorr))
    salvar_excel(caminho("registro_decisoes.xlsx"), {"Registro de Decisões": registro, "Detalhe por tipo": detalhe})

    # ------------------------------------------------------------------ vulto.db consertado (mesma estrutura)
    mudancas = banco_corrigido.gerar(pasta, caminho("vulto_corrigido.db"), ven, pts, est, clientes_final)
    conferir(caminho("vulto_corrigido.db"), ven, est)

    # ------------------------------------------------------------------ o que não cabe no vulto.db
    db = caminho("planilhas_e_relatorios.db")
    if os.path.exists(db):
        os.remove(db)
    tabelas = {
        # planilhas corrigidas (deixam de ser usadas depois da migração)
        "planilha_clientes": clientes_final, "planilha_clientes_alertas_email": cli["alertas"],
        "planilha_catalogo": cat["catalogo"], "planilha_catalogo_mestre": cat["mestre"],
        "planilha_catalogo_listas": cat["listas"], "planilha_vendas_batel_marco": bat["vendas"],
        "planilha_contagem_set26": inv["consolidado"], "planilha_avaria_pinheiros": inv["avaria"],
        "export_atendimentos": atendimentos,
        # o que o vulto.db não guarda
        "historico_estoque": est["movimentacoes"], "posicao_estoque_01out2026": est["posicao"],
        "relatorio_faturamento": ven["faturamento"], "relatorio_pontos_saldos": pts["saldos"],
        # documentação
        "registro_decisoes": registro, "registro_ocorrencias": ocorr}
    with sqlite3.connect(db) as c:
        for nome, df in tabelas.items():
            para_sql(df).to_sql(nome, c, index=False)
    print(f"\nArquivos em {os.path.abspath(saida)}\n")
    print("vulto_corrigido.db (mesmas tabelas do vulto.db):")
    print(mudancas.to_string(index=False), "\n")
    for nome in ("Faturamento", "Estoque", "Passivo de pontos"):
        print(numeros[nome].to_string(index=False), "\n")


def conferir(db, ven, est):
    """O banco corrigido sozinho tem que dar os mesmos números: faturamento por uma consulta simples e estoque
    pela soma da tabela estoque."""
    with sqlite3.connect(db) as c:
        fat = c.execute("""select round(sum(cast(total as real) - cast(frete as real)), 2) from vendas
                           where status = 'pago' and not (lower(coalesce(obs, '')) like '%troca%' and cast(total as real) >= 0)
                             and criado_em >= '2025-01-01' and criado_em < '2026-10-01'""").fetchone()[0]
        pecas = c.execute("select sum(qtd) from estoque").fetchone()[0]
    esperado_fat = round(ven["ponte"].resultado.iloc[-1], 2)
    esperado_est = int(est["resumo"].set_index("loja").loc["Total"].pecas_01out)
    assert abs(fat - esperado_fat) < 0.01, f"faturamento no banco {fat} ≠ {esperado_fat}"
    assert pecas == esperado_est, f"estoque no banco {pecas} ≠ {esperado_est}"
    print(f"conferência ok: faturamento {fat:,.2f} e estoque {pecas} peças saem direto do vulto_corrigido.db")


def incluir_skus_criados(mestre, skus):
    """Decisão 71: SKU criado no padrão para peça da planilha da Batel (modelo existe, cor/tamanho não)."""
    if not skus:
        return mestre
    from comum import NOME_COR
    linhas = []
    for sku in skus:
        _, pre, numero, cor, tam = sku.split("-")
        base = mestre[mestre.modelo == f"{pre}-{numero}"].iloc[0]
        linhas.append(dict(base.to_dict(), sku=sku, nome_cor=NOME_COR.get(cor), tamanho=tam, codigo_cd=None, ean=None,
                           status="Criado na migração (venda manual da Batel)", colecao=base["colecao"],
                           obs="SKU criado no padrão a partir da planilha da Batel (decisão 71)"))
    return pd.concat([mestre, pd.DataFrame(linhas)], ignore_index=True)


def para_sql(df):
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%Y-%m-%d %H:%M:%S")
        elif out[col].dtype == object:
            out[col] = out[col].map(lambda x: None if x is None or (isinstance(x, float) and x != x)
                                    else x.isoformat(sep=" ") if isinstance(x, datetime)
                                    else x.isoformat() if isinstance(x, date) else str(x))
    return out


def tres_numeros(ven, pts, est):
    v = ven["vendas"]
    fat = ven["faturamento"]
    # 1. faturamento: resumo, ponte, mês x loja e mês x canal, sistema x nosso por mês e por loja
    nosso_mes = v[v.conta_faturamento & v.no_periodo].groupby("mes").valor_mercadorias.sum()
    sis = v[v.no_relatorio]
    sis_mes = sis.groupby(sis.dia_sistema.str[:7]).valor_sistema.sum()
    por_mes = pd.DataFrame({"sistema_vulto": sis_mes, "nosso_numero": nosso_mes}).fillna(0).round(2)
    por_mes["diferenca"] = (por_mes.nosso_numero - por_mes.sistema_vulto).round(2)
    por_mes = por_mes.reset_index(names="mes")
    sis_loja = sis.groupby(sis.loja_id_original.map(LOJAS)).valor_sistema.sum()
    nosso_loja = v[v.conta_faturamento & v.no_periodo].groupby("loja").valor_mercadorias.sum()
    por_loja = pd.DataFrame({"sistema_vulto": sis_loja, "nosso_numero": nosso_loja}).fillna(0).round(2)
    por_loja["diferenca"] = (por_loja.nosso_numero - por_loja.sistema_vulto).round(2)
    por_loja = por_loja.reset_index(names="loja")
    mes_loja = fat.pivot_table(index="mes", columns="loja", values="faturamento", aggfunc="sum", fill_value=0,
                               margins=True, margins_name="Total").round(2).reset_index()
    mes_canal = fat.pivot_table(index="mes", columns="canal", values="faturamento", aggfunc="sum", fill_value=0,
                                margins=True, margins_name="Total").round(2).reset_index()
    ponte = ven["ponte"]
    total_sis, total_nosso = ponte.resultado.iloc[0], ponte.resultado.iloc[-1]
    resumo_fat = pd.DataFrame([
        dict(item="Relatório do sistema da Vulto (jan/25 a set/26)", valor=total_sis),
        dict(item="Nosso faturamento de mercadorias, líquido de devoluções", valor=total_nosso),
        dict(item="  vendas pagas (mercadorias)", valor=round(fat.vendas.sum(), 2)),
        dict(item="  devoluções", valor=round(fat.devolucoes.sum(), 2)),
        dict(item="Frete cobrado dos clientes (fora do faturamento, decisão 46)", valor=round(fat.frete_cobrado_fora.sum(), 2)),
        dict(item="Diferença sistema × nosso", valor=round(total_nosso - total_sis, 2))])

    # 2. estoque
    pos = est["posicao"]
    por_produto = (pos[pos.qtd > 0].pivot_table(index=["sku", "produto"], columns="loja", values="qtd", aggfunc="sum",
                                               fill_value=0, margins=True, margins_name="Total").reset_index())

    # 3. passivo
    s = pts["saldos"]
    por_cat = s.groupby("categoria").agg(clientes_com_saldo=("saldo", lambda x: int((x > 0).sum())),
                                         pontos=("saldo", "sum"), reais=("valor_reais", "sum")).reset_index()
    return {"Faturamento": resumo_fat, "Faturamento - ponte": ponte, "Faturamento mês x loja": mes_loja,
            "Faturamento mês x canal": mes_canal, "Sistema x nosso por mês": por_mes,
            "Sistema x nosso por loja": por_loja, "Faturamento detalhado": fat,
            "Estoque": est["resumo"], "Estoque por produto": por_produto, "Estoque em 17-09 (exemplo)": est["exemplo_17set"],
            "Passivo de pontos": pts["ponte"]["ponte"], "Passivo por categoria": por_cat}


def impactos(ven, pts, est, cli, clientes_final, cat, ocorr):
    """Texto da coluna 'Impacto no número final' do Registro de Decisões."""
    imp = {}
    for r in ven["ponte"].itertuples():
        if not pd.isna(r.decisao) and r.variacao:
            imp[int(r.decisao)] = f"Faturamento {brl(r.variacao)} ({num(r.vendas_afetadas)} vendas)"
    valor = pontos.VALOR_PONTO
    for d, p in pts["impacto"].items():
        if d == 57:
            imp[d] = f"Fora do passivo: risco de {num(p)} pontos ({brl(p * valor)})"
        elif d == 54:
            imp[d] = f"Passivo −{num(abs(p))} pontos ({brl(-abs(p) * valor)})"
        else:
            imp[d] = f"Passivo {'+' if p > 0 else '−'}{num(abs(p))} pontos ({brl(p * valor)}, antes do vencimento)"
    final = pts["ponte"]["ponte"].iloc[3]
    imp[40] = (f"Passivo de {num(final.pontos)} pontos vale {brl(final.reais_0_05)} a R$ 0,05 "
               f"(seria {brl(final.reais_0_10)} a R$ 0,10)")
    imp[4] = f"Muda a validade: {int((clientes_final.categoria == 'Ouro').sum())} clientes Ouro têm 18 meses (decisão 54)"
    v = ven["vendas"]
    app_na_5 = v[(v.loja_id == 6) & (v.loja_id_original == 5) & v.conta_faturamento]
    imp[45] = (f"Não muda o total; {brl(app_na_5.valor_mercadorias.sum())} saem da loja do site e vão para a do app "
               f"({len(app_na_5)} vendas)")
    sfs = v[v.ship_from_store & v.conta_faturamento]
    imp[49] = f"Não muda o faturamento; {num(len(sfs))} pedidos baixam estoque de Pinheiros em vez do CD"
    imp[47] = f"Já dentro do faturamento: devoluções somam {brl(ven['faturamento'].devolucoes.sum())} no período"
    linhas = int(cli["clientes"].linhas_unificadas.fillna("x").map(lambda t: len(str(t).split(","))).sum())
    imp[2] = f"Base de clientes da planilha: {num(linhas)} linhas → {num(cli['clientes'].shape[0])} clientes"
    imp[66] = f"+{num(int((clientes_final.origem_cadastro != 'Planilha Vulto+').sum()))} clientes na base (já tinham pontos ou compras)"
    imp[10] = "Nenhum: preço de venda mantido (a versão anterior subia 73 preços)"
    imp[68] = f"{num(int((cat['mestre'].preco < cat['mestre'].preco_colecao_mais_nova - 0.001).sum()))} peças com preço menor no cadastro"
    r = est["resumo"].set_index("loja").loc["Total"]
    imp[63] = f"Estoque: sistema da Vulto {num(r.pecas_sistema_vulto)} peças → nosso {num(r.pecas_01out)} peças"
    imp[61] = f"Ajuste de inventário de {num(r.ajuste_inventario_pecas)} peças (o que o sistema não registrou)"
    imp[59] = "Tira a contagem dupla/tripla do CD (as lojas do site e do app copiavam o saldo do CD)"
    t = est["transferencias"]
    imp[60] = (f"{num(t['pareadas'])} transferências pareadas; {num(t['so_envio'])} só com envio e "
               f"{num(t['so_recebimento'])} só com recebimento completadas")
    imp[33] = f"Margem de ±{num(r.margem_data_contagem_pecas)} peças nas lojas sem data de contagem"
    em_2024 = cat["mestre"][cat["mestre"].status == catalogo.STATUS_2024]
    imp[70] = f"+{num(len(em_2024))} SKUs no cadastro; vendas e estoque dessas peças passam a apontar para produto existente"
    batel_itens = ven["itens"][ven["itens"].venda_id.isin(v[v.linha_batel.notna()].id)]
    div = ven["itens"][ven["itens"].sku == "DIVERSO"].merge(v[["id", "conta_faturamento"]], left_on="venda_id", right_on="id")
    imp[72] = (f"Faturamento não muda ({brl((div.qtd * div.preco_unit)[div.conta_faturamento].sum())} já incluídos); "
               f"{num(int(div.qtd.sum()))} peças não baixam de nenhuma peça e entram no ajuste de inventário")
    imp[71] = (f"{num(int(batel_itens.qtd.sum()))} peças da Batel com SKU; "
               f"{len(ven['skus_criados'])} SKUs criados no padrão; o estoque da Batel baixa no dia da venda")
    return imp


if __name__ == "__main__":
    aqui = os.path.dirname(os.path.abspath(__file__))
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(aqui, "..", "..", "data"),
         sys.argv[2] if len(sys.argv) > 2 else os.path.join(aqui, "saida"))
