"""Testes das regras com exemplos reais das planilhas. Rodar: python -m pytest"""
from datetime import datetime

from atendimento import formato_pedido, separar_mensagem
from clientes import categoria_por_pontos, pontos
from comum import (cpf_mascara, cpf_valido, dinheiro, ean_valido, ler_codigo_produto, montar_codigo_cd, montar_sku,
                   tamanho_padrao, telefone_padrao)
from vendas_batel import data_marco, Registro


def test_pontos():
    assert pontos("2.203") == 2203
    assert pontos("1389 pts") == 1389
    assert pontos("2,7k") == 2700
    assert pontos("15.000") == 15000
    assert pontos(None) is None


def test_categoria_por_faixa():
    assert [categoria_por_pontos(p) for p in (0, 1, 1500, 1501, 5000, 5001, 10000, 10001)] == \
        ["Nada", "Bronze", "Bronze", "Prata", "Prata", "Ouro", "Ouro", "VIP"]


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
