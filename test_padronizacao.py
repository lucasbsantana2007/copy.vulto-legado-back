"""Testes das regras com exemplos reais das planilhas. Rodar: python -m pytest"""
from datetime import datetime

import pandas as pd

import pontos as vulto_mais
from atendimento import formato_pedido, separar_mensagem
from clientes import pontos
from comum import (cpf_mascara, cpf_valido, data_vulto, dinheiro, ean_valido, identificar, ler_codigo_produto,
                   ler_transferencia, montar_codigo_cd, montar_sku, sku_canonico, tamanho_padrao, telefone_padrao)
from estoque import parear
from vendas import canal_da_venda
from vendas_batel import data_marco, Registro


def test_pontos():
    assert pontos("2.203") == 2203
    assert pontos("1389 pts") == 1389
    assert pontos("2,7k") == 2700
    assert pontos("15.000") == 15000
    assert pontos(None) is None


def test_categoria_pela_regra_escrita():
    # aba LEIA ISSO: Bronze até 1.999, Prata 2.000–4.999, Ouro 5.000+ (pontos ganhos em 12 meses)
    assert [vulto_mais.categoria(p) for p in (0, 1999, 2000, 4999, 5000, 15000)] == \
        ["Bronze", "Bronze", "Prata", "Prata", "Ouro", "Ouro"]


def test_cpf():
    assert cpf_valido("00463337549")
    assert cpf_mascara("07058850119") == "070.588.501-19"


def test_tamanhos():
    assert [tamanho_padrao(t) for t in ("36", "38", "40", "42", "44", "46", "PP", "S", "XL", "Unico")] == \
        ["XP", "P", "M", "G", "GG", "XG", "XP", "P", "GG", "U"]


def test_codigos_produto():
    assert ler_codigo_produto("CAM016GP") == ("CAM", "016", "BEG", "P")
    assert ler_codigo_produto("vlt cal 011 lilás 40") == ("CAL", "011", "LIL", "40")
    assert ler_codigo_produto("vlt ves 006 marrom") == ("VES", "006", "MRR", None)     # linha de grade
    assert montar_sku("CAL", "010", "MRR", tamanho_padrao("38")) == "VLT-CAL-010-MRR-P"
    assert montar_codigo_cd("CAL", "010", "MRR", "38") == "CAL010M38"                  # exemplo do PDF


def test_ean():
    assert not ean_valido("7894692000000")        # cortado ao colar
    assert ean_valido("7894903281368")


def test_telefone_e_dinheiro():
    assert telefone_padrao("(11) 98262-6735") == "+55 11 98262-6735"
    assert telefone_padrao("93869-5105") is None   # sem DDD
    assert dinheiro("R$ 509,70") == 509.70


def test_data_batel_dia_mes_trocados():
    reg = Registro("t")
    assert data_marco(datetime(2026, 5, 3), reg, 1).strftime("%d/%m") == "05/03"
    assert data_marco("21/03", reg, 2).strftime("%d/%m/%Y") == "21/03/2026"


def test_atendimento():
    assert formato_pedido("#1094") == "1094"
    assert formato_pedido("7001851") == "APP-7001851"
    c, r = separar_mensagem("Oi\n\n[Atendente]: Resolvido! Qualquer coisa estamos por aqui.\n[Cliente]: obrigadaaa")
    assert c == "Oi\nobrigadaaa" and r.startswith("Resolvido!")


def test_data_em_utc_vira_brasilia():
    assert data_vulto("2026-02-10T15:43:34Z") == datetime(2026, 2, 10, 12, 43, 34)     # Eldorado depois da troca
    assert data_vulto("2026-10-01T02:30:00Z") == datetime(2026, 9, 30, 23, 30)          # volta para setembro
    assert data_vulto("01/03/2025 19:43") == datetime(2025, 3, 1, 19, 43)               # Savassi, horário local
    assert data_vulto("14/04/25") == datetime(2025, 4, 14)


def test_quem_e_o_cliente_da_venda():
    assert identificar("191.281.999-60") == ("cpf", "19128199960")
    assert identificar("41980713796") == ("telefone", "+55 41 98071-3796")             # 11 dígitos, DV não confere
    assert identificar("HENRIQUE.SILVA25@gmail.com") == ("email", "henrique.silva25@gmail.com")
    assert identificar("NAO INFORMADO") == ("anonimo", None)


def test_canal_pelo_numero_do_pedido():
    assert canal_da_venda("APP-7012863") == "App"
    assert canal_da_venda("S100816") == "Site"
    assert canal_da_venda(None) == "Loja física"


def test_transferencia_escrita_no_texto():
    assert ler_transferencia("enviado pinheiros - 1 pçs") == (1, 1)
    assert ler_transferencia("transf p/ savassi 3un") == (3, 3)
    assert ler_transferencia("2 un -> batel") == (4, 2)
    assert ler_transferencia("TRANSF ELDORADO 1") == (2, 1)
    assert ler_transferencia("veio do CD") == (None, None)


def test_sku_canonico_junta_os_dois_codigos():
    assert sku_canonico("CAL013M38") == sku_canonico("VLT-CAL-013-MRR-38") == "VLT-CAL-013-MRR-P"


def test_vencimento_gasta_os_pontos_mais_antigos_primeiro():
    g = pd.DataFrame([dict(data=datetime(2025, 1, 10), qtd=1000), dict(data=datetime(2025, 9, 1), qtd=500),
                      dict(data=datetime(2025, 10, 1), qtd=-300)])
    r = vulto_mais.simular(g, 12)
    # o resgate de 300 sai do lote de janeiro (vence 10/01/2026 com 700); o de setembro vence 01/09/2026
    assert r == dict(saldo=0, saldo_sem_vencimento=1200, vencidos=1200, resgatado_sem_saldo=0)
    r = vulto_mais.simular(g, 18)
    assert r["saldo"] == 500 and r["vencidos"] == 700


def test_resgate_maior_que_o_saldo_nao_deixa_negativo():
    g = pd.DataFrame([dict(data=datetime(2026, 5, 1), qtd=100), dict(data=datetime(2026, 6, 1), qtd=-300)])
    assert vulto_mais.simular(g, 12) == dict(saldo=0, saldo_sem_vencimento=0, vencidos=0, resgatado_sem_saldo=200)


def test_transferencia_pareia_no_mesmo_dia_e_nao_na_semana_seguinte():
    envios = pd.DataFrame([dict(origem_id=7, destino_id=1, sku="X", qtd=1, quando=datetime(2026, 8, 3), ref=1)])
    receb = pd.DataFrame([dict(origem_id=7, destino_id=1, sku="X", qtd=1, quando=datetime(2026, 8, 10), ref=2),
                          dict(origem_id=7, destino_id=1, sku="X", qtd=1, quando=datetime(2026, 8, 3), ref=3)])
    pares, so_envio, so_receb, lag = parear(envios, receb)
    assert [p["ref_receb"] for p in pares] == [3] and [r["ref"] for r in so_receb] == [2] and lag.days == 0
