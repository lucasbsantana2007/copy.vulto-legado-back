"""
Registro de Decisões (entregável da Entrega 3): uma linha por problema encontrado, com as colunas pedidas no
enunciado. 'Quantos registros afeta' e 'Impacto no número final' são calculados a cada execução de rodar_tudo.py.

autoria: "Arthur" (decisões 1 a 37), "Revisada" (decisão do Arthur alterada nesta versão), "Nova".
"""
import pandas as pd

# numero: (o que encontramos, onde, o que decidimos, por quê, autoria)
DECISOES = {
    1: ("CPF com menos de 11 dígitos", "Planilha de clientes, coluna cpf",
        "Completar com zeros à esquerda e marcar o cadastro para confirmação",
        "O Excel apaga o zero à esquerda; sem completar, o cliente não casa com o banco nem com a Lorenzi", "Arthur"),
    2: ("Mesmo CPF em 2 ou 3 linhas", "Planilha de clientes, coluna cpf",
        "Unificar num cliente só, somar os pontos e ficar com o e-mail do cadastro com mais pontos; avisar o cliente",
        "É a mesma pessoa cadastrada mais de uma vez (a Head de CRM citou isso no kickoff)", "Arthur"),
    3: ("Pontos em vários formatos ('2,7k', '1.200', '999 pts') ou vazios", "Planilha de clientes, coluna Pontos",
        "Converter para número; vazio vira 0. Esse valor é o saldo congelado em 01/01/2025",
        "A aba LEIA ISSO diz que a coluna foi congelada quando o sistema entrou; ela é o saldo inicial do extrato", "Arthur"),
    4: ("Categoria escrita de 10 jeitos e colocada à mão", "Planilha de clientes, coluna categoria",
        "Recalcular pela regra escrita: pontos ganhos nos últimos 12 meses — Bronze até 1.999, Prata 2.000 a 4.999, "
        "Ouro 5.000 ou mais",
        "A regra existe na aba LEIA ISSO. A versão anterior criou faixas próprias (Nada, VIP) sobre o saldo congelado "
        "de 2025. A categoria importa: Ouro tem pontos válidos por 18 meses", "Revisada"),
    5: ("E-mails com erro de digitação, maiúsculas, espaço ou marcador ('-', 'nao tem')",
        "Planilha de clientes, coluna E-mail",
        "Corrigir domínios (gmial→gmail, .con→.com), minúsculas e espaços; inválido ou ausente fica vazio",
        "E-mail errado não entrega mensagem e não identifica o cliente nas vendas e no atendimento", "Arthur"),
    6: ("Custo em branco; dois conceitos de custo (último x médio)", "Catálogo, coluna Custo; entradas_estoque",
        "Estimar custos em branco (mesmo modelo/cor, % da categoria ou 2024). Guardar os dois custos: o último de nota "
        "(Compras, para precificar) e o médio ponderado das notas de entrada (Financeiro). A Lorenzi recebe o médio",
        "REGRAS_OPERACAO: custo do produto = médio ponderado (contabilidade); 'para precificar coleção nova usar o "
        "último custo'. A versão anterior trocava o médio pelo último", "Revisada"),
    7: ("EAN ausente, cortado ou repetido", "Catálogo, coluna EAN",
        "Copiar do mesmo SKU em outra aba ou da planilha de 2024; sem fonte, EAN interno provisório (faixa 200)",
        "Sem EAN a peça não passa no leitor do PDV", "Arthur"),
    8: ("Linhas marcadas em amarelo ('revisar preço')", "Catálogo, aba Verão 26",
        "Concluir a revisão quando o preço é igual ao dos outros tamanhos; senão continua em revisão",
        "Preço igual na grade indica que a revisão já foi feita", "Arthur"),
    9: ("Linhas repetidas na aba Atemporais", "Catálogo, aba Atemporais",
        "Manter a linha original (com custo e menor preço)", "A cópia tinha preço maior e sem custo", "Arthur"),
    10: ("Modelos com margem abaixo de 50%", "Catálogo, Preço x Custo",
         "Manter o preço. A margem baixa vira alerta e o preço sugerido (final ,90) fica para Compras aprovar",
         "A versão anterior aumentava 73 preços com 'aumento aprovado em 15/10/2026', uma aprovação que não "
         "aconteceu. Migrar dados não pode mudar preço de venda", "Revisada"),
    11: ("Fornecedor em branco", "Catálogo, coluna Fornecedor",
         "Preencher pelo mesmo modelo, pelas regras de tipo de peça ou pelo principal da categoria; conferir com as notas",
         "Fornecedor é necessário para recompra e conferência de nota", "Arthur"),
    12: ("Tamanhos fora do padrão e linhas de grade ('P/M/G')", "Catálogo, coluna Tamanho",
         "Padronizar (36→XP, 38→P, … S→P, XL→GG) e expandir a grade em uma linha por tamanho",
         "Regra de tamanhos do grupo; cada tamanho é um SKU com estoque próprio", "Arthur"),
    13: ("Código curto do CD ausente", "Catálogo", "Gerar pelo padrão (CAT + número + letra da cor + tamanho); "
         "provisório quando o CD ainda não usa", "O CD e as vendas usam o código curto", "Arthur"),
    14: ("Datas da Batel com dia e mês trocados ou em texto", "Planilha de vendas manuais da Batel",
         "Todas as vendas são de março/2026: '2026-05-03' era 05/03", "O PDV ficou fora de 02 a 22/03", "Arthur"),
    15: ("Valores da Batel gravados como texto", "Planilha de vendas manuais da Batel", "Converter para número",
         "Texto não soma no total", "Arthur"),
    16: ("CPF da Batel sem zeros à esquerda", "Planilha de vendas manuais da Batel", "Completar e validar",
         "Mesmo motivo da decisão 1", "Arthur"),
    17: ("Vendedora com apelidos diferentes", "Planilha de vendas manuais da Batel", "Unificar o nome",
         "Comissão e ranking por vendedora", "Arthur"),
    18: ("Forma de pagamento escrita de vários jeitos", "Planilha de vendas manuais da Batel", "Padronizar",
         "Conciliar com o extrato da maquininha", "Arthur"),
    19: ("Cancelamento marcado só com texto riscado", "Planilha de vendas manuais da Batel",
         "Riscado vira status Cancelada/Estornada e sai dos totais", "A formatação se perde na migração", "Arthur"),
    20: ("'lancei no vulto+' escrito na observação", "Planilha de vendas manuais da Batel",
         "Virar coluna (Sim / Pendente); as pendentes recebem os pontos (decisão 58)",
         "Saber quais clientes ainda não ganharam pontos", "Arthur"),
    21: ("Datas, canais e atendentes escritos de vários jeitos; linhas duplicadas", "Export de atendimento (CSV)",
         "Padronizar e remover a linha exportada duas vezes", "Relatórios por canal e atendente", "Arthur"),
    22: ("Contato sem padrão e nome do cliente diferente a cada conversa", "Export de atendimento, contato e nome",
         "Celular em +55 DD 9XXXX-XXXX; um nome por contato; deduzir pelo e-mail/@ quando faltar",
         "Encontrar o cliente pelo contato", "Arthur"),
    23: ("Assuntos e tags redundantes", "Export de atendimento", "9 categorias de assunto",
         "Relatórios de atendimento por motivo", "Arthur"),
    24: ("Acentos quebrados", "Export de atendimento (latin-1)", "Ler o arquivo na codificação certa",
         "No Excel os acentos apareciam como '?'", "Arthur"),
    25: ("Número do pedido com prefixo, ausente ou com dígitos invertidos", "Export de atendimento, coluna pedido",
         "Limpar prefixo, recuperar do texto da mensagem e corrigir inversão só quando o resultado é um pedido do "
         "site ou do app",
         "Número de venda de loja vai de 1 a ~18.600: quase todo número invertido 'existe', então a inversão era "
         "um chute (ex.: APP-4904 virava a venda 4094)", "Revisada"),
    26: ("Status de atendimento em vários formatos", "Export de atendimento", "4 status", "Fila de atendimento",
         "Arthur"),
    27: ("Atendimento marcado como resolvido com conversa em aberto", "Export de atendimento",
         "Status pela última resposta do atendente", "O cliente ainda esperava resposta", "Arthur"),
    28: ("Peças com avaria contadas junto", "Contagem de Pinheiros", "Separar a avaria do estoque bom",
         "Avaria não pode ser vendida", "Arthur"),
    29: ("Provador não contado na Batel", "Contagem da Batel", "Considerar zero",
         "Estoque conservador: melhor menor do que vender sem a peça", "Arthur"),
    30: ("Peças marcadas 'recontar' sem recontagem", "Contagem da Batel",
         "Vale a quantidade contada", "A versão anterior estimava pela mediana de outras lojas, que não diz nada "
         "sobre a Batel; se o contado estiver errado, aparece no ajuste de inventário", "Revisada"),
    31: ("Savassi contou só o total por modelo", "Contagem da Savassi",
         "Ratear por cor e tamanho na proporção de Pinheiros (ou Eldorado+Batel)",
         "Sem SKU não dá para vender nem transferir", "Arthur"),
    32: ("Saldo reservado no CD", "Posição do WMS do CD", "Fica no CD (reposição já pedida)",
         "Não somar duas vezes", "Arthur"),
    33: ("Datas de contagem diferentes ou sem data", "Contagem física",
         "Pinheiros 20/09 (após o fechamento); CD 28/09; Eldorado, Savassi e Batel sem data: 28/09",
         "Data mais próxima da posição pedida; a margem de erro está no resumo do estoque", "Arthur"),
    34: ("Lojas sem registro de avaria", "Contagem física", "Avaria considerada zero", "Sem informação", "Arthur"),
    35: ("SKU e código em formatos diferentes", "Contagem física", "Converter tudo para o SKU do catálogo",
         "Uma peça, um código", "Arthur"),
    36: ("Peças fora do catálogo atual", "Contagem física", "Entram no estoque; candidatas a liquidação",
         "A peça existe fisicamente", "Arthur"),
    37: ("Código do CD com dígitos invertidos", "Contagem da Eldorado", "Corrigir quando só um código bate",
         "Erro de digitação", "Arthur"),
    38: ("Vendas do app gravadas em centavos", "vulto.db: vendas.total e itens_venda.preco_unit (loja do app e "
         "pedidos do app na loja do site)",
         "Dividir por 100 quando o preço do item é 20 vezes ou mais o do catálogo",
         "O próprio schema avisa ('app grava centavos, PDV grava reais'). Depois da correção o ticket médio do app "
         "(~R$ 376) fica igual ao do site e das lojas", "Nova"),
    39: ("Status escrito de 10 jeitos; o relatório soma cancelados (a linha que tiraria está comentada)",
         "vulto.db: vendas.status",
         "Só conta venda paga (pago/aprovado; no PDV, status vazio = venda de balcão paga). Cancelada, estornada, "
         "chargeback e 'aguardando' ficam fora",
         "REGRAS_OPERACAO: 'venda cancelada não conta'. 'Aguardando' há meses é pagamento que nunca foi confirmado",
         "Nova"),
    40: ("Valor do ponto: o sistema usa R$ 0,10; a regra mudou para R$ 0,05 em mar/2025", "vulto.db: config."
         "valor_ponto x aba LEIA ISSO",
         "Passivo em R$ 0,05 por ponto; cenário de R$ 0,10 mostrado ao lado",
         "É a regra vigente do programa, e 19 dos 27 resgates em papel de 2025–26 já usam R$ 5 a cada 100 pontos",
         "Nova"),
    41: ("Batel sem PDV: vendas na planilha e parte relançada no sistema depois", "vulto.db (Batel, 23/03 a 02/04) "
         "x planilha manual",
         "Venda relançada que bate com a planilha (mesmo valor e CPF) conta uma vez, na data real; venda só da "
         "planilha entra sem itens",
         "Evita contar duas vezes e não perde as vendas que nunca foram relançadas", "Nova"),
    42: ("Pedido gravado duas vezes ('aguardando' e depois 'pago')", "vulto.db: vendas.pedido_externo",
         "Fica só a gravação com o melhor status", "É o mesmo pedido", "Nova"),
    43: ("Pedidos de teste", "vulto.db: vendas (série S990xxx e APP-990xxxx; obs 'teste')",
         "Fora do faturamento, do estoque e dos pontos", "Não são vendas (inclui um de R$ 99.999)", "Nova"),
    44: ("Datas em dois relógios: UTC (site, app, Eldorado depois de 10/02/2026) e horário local",
         "vulto.db: vendas.criado_em (terminadas em 'Z')",
         "Converter tudo para o horário de Brasília",
         "Venda das 21h às 24h caía no dia (e às vezes no mês) seguinte", "Nova"),
    45: ("Canal escrito de 15 jeitos; venda do app gravada na loja do site", "vulto.db: vendas.canal e loja_id",
         "Canal pelo número do pedido (S = site, APP- = app, sem número = loja física); venda do app vai para a "
         "loja do app", "'online' aparece nos dois canais", "Nova"),
    46: ("Frete: no site fica num campo; no app está dentro do total", "vulto.db: vendas.frete e total",
         "Faturamento = mercadorias (sem frete); o frete aparece em coluna separada",
         "O Financeiro (Paula) diz que frete é repasse; a Lorenzi também calcula receita só pelos itens. Mesmo critério "
         "para as duas redes", "Nova"),
    47: ("Devolução lançada como venda negativa", "vulto.db: vendas com total < 0",
         "Desconta do faturamento no mês da devolução; a peça volta ao estoque", "Regra da operação", "Nova"),
    48: ("Troca lançada como venda nova", "vulto.db: vendas com obs 'troca'",
         "Fora do faturamento; a peça nova sai do estoque",
         "O cliente já pagou na compra original; contar de novo infla a receita", "Nova"),
    49: ("Pedido do site que saiu da loja Pinheiros (ship from store)", "vulto.db: vendas.obs ('sfs', 'saiu da loja')",
         "Faturamento fica no site; a baixa de estoque vai para Pinheiros", "REGRAS_OPERACAO (piloto desde jun/26)",
         "Nova"),
    50: ("Cliente da venda gravado como CPF, e-mail ou telefone", "vulto.db: vendas.cliente_ref",
         "Ligar ao CPF pelo e-mail ou celular do cadastro", "Ver o histórico do cliente na Lorenzi", "Nova"),
    51: ("Crédito de pontos do app calculado em centavos (100x maior); 21 corrigidos à mão em jun/26",
         "vulto.db: pontos_mov tipo C e AJ 'correcao bug app v3.4'",
         "Corrigir o crédito na origem e tirar os ajustes manuais que já compensavam (para não corrigir duas vezes); "
         "CPF de 9 dígitos completado", "Mesmo defeito da decisão 38", "Nova"),
    52: ("Pontos de compra cancelada, estornada, com chargeback, não paga, duplicada, de troca ou de teste",
         "vulto.db: pontos_mov tipo C ligado a vendas",
         "Retirar o crédito (e o estorno manual, quando existe)",
         "LEIA ISSO/REGRAS: 'ponto só depois do pagamento aprovado' e 'quem cancela perde os pontos'", "Nova"),
    53: ("Compra online ganhou pontos simples", "vulto.db: pontos_mov x compras",
         "Creditar o dobro que faltou",
         "A regra é dobro no site/app; o código só dobrava quando o canal era exatamente 'site' ou 'app'. É dívida "
         "com o cliente (a regra foi divulgada)", "Nova"),
    54: ("Pontos nunca venciam", "main.py ('TODO expirar pontos')",
         "Vencer em 12 meses (Ouro 18), usando os pontos mais antigos primeiro nos resgates",
         "Regra escrita do programa; Ouro 18 meses combinado com o Diego", "Nova"),
    55: ("Resgates anotados só na aba Resgates", "Planilha de clientes, aba Resgates",
         "Descontar os de 2025–26 que não estão no sistema; os iguais (mesmo CPF, pontos e mês) não contam de novo",
         "Algumas lojas continuaram anotando no papel (LEIA ISSO)", "Nova"),
    56: ("Resgate maior que o saldo", "vulto.db: pontos_mov tipo D", "Saldo fica zero (o desconto já foi dado)",
         "Regra: o saldo não pode ficar negativo; o sistema não conferia", "Nova"),
    57: ("Bônus Ouro (+50%) nunca aplicado", "Aba LEIA ISSO x sistema",
         "Não creditar; mostrar como risco", "Não há registro de quem era Ouro em cada compra", "Nova"),
    58: ("Vendas da Batel sem sistema não ganharam pontos", "Planilha de vendas manuais da Batel",
         "Creditar 1 ponto por real para quem tinha CPF", "Mesma regra das vendas de loja", "Nova"),
    59: ("As 'lojas' do site e do app têm estoque próprio, cópia do CD", "vulto.db: estoque, lojas 5 e 6",
         "O estoque do site e do app fica no CD; vendas online baixam do CD",
         "REGRAS_OPERACAO: site e app vendem do estoque do CD. Somar as três lojas conta a mesma peça três vezes",
         "Nova"),
    60: ("Transferência anotada pelo CD (só no texto), pela loja ou pelos dois", "vulto.db: entradas_estoque",
         "Parear envio e recebimento da mesma peça no mesmo dia (até 2 dias); se só um lado anotou, criar o outro",
         "'Uma das duas basta' (REGRAS_OPERACAO): o CD nunca baixava a peça enviada", "Nova"),
    61: ("Não existe histórico de estoque", "vulto.db (só saldo atual)",
         "Reconstruir: saldo inicial em 01/01/2025 + entradas, transferências, vendas e devoluções + ajuste de "
         "inventário na data da contagem",
         "Responder 'quanto tinha na loja X há duas semanas' (pedido da Diretora de Operações)", "Nova"),
    62: ("Peça vendida depois da contagem que a contagem não achou", "Contagem x vendas de 21 a 30/09",
         "Saldo fica zero com ajuste", "Estoque não pode ficar negativo", "Nova"),
    63: ("Estoque do sistema com saldos negativos e a mesma peça em dois códigos", "vulto.db: estoque",
         "A posição oficial vem da contagem física levada até 01/10 pelos movimentos; o saldo do sistema só serve "
         "para comparar", "O sistema nunca baixava transferências e duplica as lojas virtuais", "Nova"),
    64: ("O mesmo e-mail em cadastros de pessoas diferentes", "Planilha de clientes, coluna E-mail",
         "O e-mail fica com quem tem o primeiro nome nele; dos outros é removido (se não der para saber, de todos)",
         "Mandar dado de um cliente para o e-mail de outro é vazamento (LGPD); e-mail é único na Lorenzi", "Nova"),
    65: ("Consentimento de mensagens escrito de 8 jeitos ou em branco", "Planilha de clientes, coluna aceita_msg",
         "Sim / Não / Não informado; só 'Sim' recebe campanha",
         "LGPD: marketing exige consentimento explícito (a planilha já alertava)", "Nova"),
    66: ("Clientes que só existem no sistema", "vulto.db: pontos_mov e vendas (CPF fora da planilha)",
         "Criar o cadastro só com o CPF e marcar para completar",
         "A planilha foi congelada em 01/01/2025; quem entrou depois tem pontos, mas não tem linha", "Nova"),
    67: ("Atendimento sem ligação com cliente e pedido", "Export de atendimento",
         "Ligar pelo e-mail ou celular do cadastro, ou pelo cliente do pedido; o pedido só é ligado se for do mesmo "
         "cliente e anterior ao atendimento; sem pedido, o atendimento fica no SAC central (loja do site)",
         "Pedido da Gerente de Atendimento: parar de perguntar o número do pedido", "Nova"),
    68: ("A mesma peça com preços diferentes em coleções diferentes", "Catálogo (todas as abas)",
         "Preço de venda = o menor entre as coleções vigentes; o da coleção mais nova fica guardado",
         "REGRAS_OPERACAO: vale o menor preço (Procon) e, em conflito com a planilha da Marina, vale o arquivo do "
         "Diego", "Nova"),
    69: ("Loja de cadastro escrita de 10 jeitos ('bh', 'cwb', 'eldorado shop')", "Planilha de clientes",
         "Padronizar para as lojas do sistema", "Relatórios por loja", "Nova"),
    70: ("Peças da coleção 2024 que ainda vendem ficaram fora do catálogo", "Catálogo, aba 'PLANILHA ANTIGA NAO USAR' x "
         "vendas e estoque",
         "Incluir no cadastro como 'Coleção 2024 (só estoque)' as que aparecem em vendas, estoque, entradas ou contagem "
         "desde 2025. Preço = último preço vendido; custo e EAN da aba 2024",
         "O sistema da Vulto avisa: 'a antiga tem os SKUs de 2024 que ainda vendem, não tirar'. A aba manda não usar o "
         "PREÇO de 2024, e o preço praticado foi em média 4% maior (decisão do grupo)", "Nova"),
    71: ("Vendas da Batel sem sistema sem itens; relançadas como 'DIVERSO - VENDA MANUAL'", "Planilha de vendas "
         "manuais da Batel, coluna PEÇAS x vulto.db itens_venda",
         "Ler o texto de cada peça (modelo, cor e tamanho) e criar os itens com o SKU do catálogo; se o modelo existir "
         "sem aquela cor/tamanho, criar o SKU no padrão VLT-CAT-NNN-COR-TAM. O preço de cada item segue o do catálogo, "
         "ajustado para fechar com o valor da planilha",
         "Saber quais peças saíram da Batel em março e baixar o estoque certo (decisão do grupo)", "Nova"),
    72: ("Peças vendidas com o código genérico do PDV ('DIVERSO', 'PRODUTO SEM CADASTRO')", "vulto.db: itens_venda",
         "Um código só, 'DIVERSO' (descrição 'Peça não identificada no PDV'). A venda e o valor ficam no faturamento; "
         "não baixa estoque de nenhuma peça (a diferença aparece no ajuste de inventário). Na Lorenzi, vira um "
         "produto genérico só para o histórico",
         "Não há como saber modelo, cor e tamanho: o preço bate com 6 modelos ou mais em 2 de cada 3 itens e só "
         "aponta um modelo em 15 casos, sem cor e tamanho (decisão do grupo)", "Nova"),
}


def tabela(ocorrencias, impactos):
    """ocorrencias: todas as abas 'Ocorrências' juntas. impactos: {decisão: texto}."""
    afetados = ocorrencias.groupby("decisao").origem.nunique()
    linhas = []
    for n, (oque, onde, decisao, porque, autoria) in DECISOES.items():
        linhas.append({"#": n, "O que encontramos": oque, "Onde (arquivo / tabela / campo)": onde,
                       "Quantos registros afeta": int(afetados.get(n, 0)), "O que decidimos fazer": decisao,
                       "Por quê": porque, "Impacto no número final": impactos.get(n, "Não muda os três números"),
                       "Autoria": autoria})
    return pd.DataFrame(linhas)
