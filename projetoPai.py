# -*- coding: utf-8 -*-
"""
Created on Sun Aug 23 14:32:46 2026

@author: Home
"""

from selenium import webdriver
import json
import os
from openpyxl import Workbook


# =========================================================
# FUNÇÕES
# =========================================================

def pegar_dados(navegador, url):
    try:
        navegador.get(url)

        texto = navegador.find_element("tag name", "body").text

        dados = json.loads(texto)

        return dados

    except Exception as erro:
        print("Erro ao buscar os dados:")
        print(erro)
        return None


def pegar_estatisticas(dados, chave, lado):

    if dados is None:
        return None

    if "statistics" not in dados:
        return None

    for i in dados["statistics"]:

        if i["period"] == "ALL":

            for grupo in i["groups"]:

                for estatistica in grupo["statisticsItems"]:

                    if estatistica["key"] == chave:

                        return estatistica.get(lado)


    return None


def pegar_estatisticas_1tempo(dados, chave, lado):

    if dados is None:
        return None

    if "statistics" not in dados:
        return None

    for i in dados["statistics"]:

        if i["period"] == "1ST":

            for grupo in i["groups"]:

                for estatistica in grupo["statisticsItems"]:

                    if estatistica["key"] == chave:

                        return estatistica.get(lado)

    return None


def transoformar_int(valor):

    if valor == None:
        return 0

    else:
        try:
            return int(valor)

        except:
            return 0


# =========================================================
# LISTAS
# =========================================================

resultado_casa = []
resultado_fora = []

handicap_casa = []
handicap_fora = []

defesa_casa = []
defesa_enfrenta_casa = []
defesa_fora = []
defesa_enfrenta_fora = []

finalizacao_casa = []
finalizacao_enfrenta_casa = []
finalizacao_fora = []
finalizacao_enfrenta_fora = []

chutes_nogol_casa = []
chutes_nogol_enfrenta_casa = []
chutes_nogol_fora = []
chutes_nogol_enfrenta_fora = []

faltas_casa = []
faltas_enfreta_casa = []
faltas_fora = []
faltas_enfreta_fora = []

escanteios_casa = []
escanteios_enfrenta_casa = []
escanteios_fora = []
escanteios_enfrenta_fora = []

escanteios1_casa = []
escanteios1_enfrenta_casa = []
escanteios1_fora = []
escanteios1_enfrenta_fora = []

cartoes_casa = []
cartoes_enfrenta_casa = []
cartoes_fora = []
cartoes_enfrenta_fora = []

cartao_vermelho_casa = []
cartao_vermelho_enfrenta_casa = []
cartao_vermelho_fora = []
cartao_vermelho_enfrenta_fora = []

gol_sofrido_casa = []
gol_sofrido_fora = []

gol1_sofrido_casa = []
gol1_sofrido_fora = []

gol_casa = []
gol_fora = []

gol1_casa = []
gol1_fora = []

posse1_casa = []
posse1_enfrenta_casa = []

posse1_fora = []
posse1_enfrenta_fora = []

xg_casa = []
xg_fora = []

xg_enfrenta_casa = []
xg_enfrenta_fora = []


id_jogo_casa = []
id_jogo_visit = []



# =========================================================
# PROGRAMA PRINCIPAL
# =========================================================

while True:

    # =====================================================
    # LIMPA AS LISTAS
    # =====================================================
    
    xg_casa.clear()
    xg_fora.clear()
    
    xg_enfrenta_casa.clear()
    xg_enfrenta_fora.clear()
    
    resultado_casa.clear()
    resultado_fora.clear()

    handicap_casa.clear()
    handicap_fora.clear()

    defesa_casa.clear()
    defesa_enfrenta_casa.clear()
    defesa_fora.clear()
    defesa_enfrenta_fora.clear()

    finalizacao_casa.clear()
    finalizacao_enfrenta_casa.clear()
    finalizacao_fora.clear()
    finalizacao_enfrenta_fora.clear()

    chutes_nogol_casa.clear()
    chutes_nogol_enfrenta_casa.clear()
    chutes_nogol_fora.clear()
    chutes_nogol_enfrenta_fora.clear()

    faltas_casa.clear()
    faltas_enfreta_casa.clear()
    faltas_fora.clear()
    faltas_enfreta_fora.clear()

    escanteios_casa.clear()
    escanteios_enfrenta_casa.clear()
    escanteios_fora.clear()
    escanteios_enfrenta_fora.clear()

    escanteios1_casa.clear()
    escanteios1_enfrenta_casa.clear()
    escanteios1_fora.clear()
    escanteios1_enfrenta_fora.clear()

    cartoes_casa.clear()
    cartoes_enfrenta_casa.clear()
    cartoes_fora.clear()
    cartoes_enfrenta_fora.clear()

    cartao_vermelho_casa.clear()
    cartao_vermelho_enfrenta_casa.clear()
    cartao_vermelho_fora.clear()
    cartao_vermelho_enfrenta_fora.clear()

    gol_sofrido_casa.clear()
    gol_sofrido_fora.clear()

    gol1_sofrido_casa.clear()
    gol1_sofrido_fora.clear()

    gol_casa.clear()
    gol_fora.clear()

    gol1_casa.clear()
    gol1_fora.clear()

    id_jogo_casa.clear()
    id_jogo_visit.clear()
    
    
    posse1_casa.clear()
    posse1_enfrenta_casa.clear()

    posse1_fora.clear()
    posse1_enfrenta_fora.clear()

    # =====================================================
    # PEDE OS IDS DOS JOGOS DO TIME DA CASA
    # =====================================================

    while True:

        id_casa = int(
            input(
                "Digite o Id dos jogos do time da casa no Sofascore: "
            )
        )

        if id_casa == 0:
            break

        else:
            id_jogo_casa.append(id_casa)


    # =====================================================
    # PEDE OS IDS DOS JOGOS DO TIME VISITANTE
    # =====================================================

    while True:

        id_visit = int(
            input(
                "Digite o Id dos jogos do time visitante no Sofascore: "
            )
        )

        if id_visit == 0:
            break

        else:
            id_jogo_visit.append(id_visit)


    # =====================================================
    # ABRE O Edge
    # =====================================================

    navegador = webdriver.Edge()

    os.system("cls")


    # =====================================================
    # JOGOS DO TIME DA CASA
    # =====================================================

    for i in id_jogo_casa:

        print("\nBuscando jogo:", i)

        # ALTERAÇÃO:
        # antes era api.sofascore.com
        # agora é www.sofascore.com

        url = f"https://www.sofascore.com/api/v1/event/{i}"

        dados = pegar_dados(navegador, url)


        # =================================================
        # VERIFICA SE VEIO O EVENT
        # =================================================

        if dados is None or "event" not in dados:

            print("Não foi possível pegar o jogo:", i)
            print("Resposta recebida:", dados)

            continue


        # =================================================
        # GOLS
        # =================================================

        gol_casa.append(
            dados["event"]["homeScore"]["normaltime"]
        )

        gol_sofrido_casa.append(
            dados["event"]["awayScore"]["normaltime"]
        )


        # =================================================
        # RESULTADO
        # =================================================

        if gol_casa[-1] > gol_sofrido_casa[-1]:

            resultado_casa.append("Vitoria")

        elif gol_casa[-1] < gol_sofrido_casa[-1]:

            resultado_casa.append("Derrota")

        else:

            resultado_casa.append("Empate")


        # =================================================
        # HANDICAP
        # =================================================

        if gol_casa[-1] < gol_sofrido_casa[-1]:

            handicap_casa.append(
                gol_casa[-1] - gol_sofrido_casa[-1]
            )

        else:

            handicap_casa.append("-")


        # =================================================
        # GOLS 1º TEMPO
        # =================================================

        gol1_casa.append(
            dados["event"]["homeScore"]["period1"]
        )

        gol1_sofrido_casa.append(
            dados["event"]["awayScore"]["period1"]
        )


        # =================================================
        # ESTATÍSTICAS
        # =================================================

        url = (
            f"https://www.sofascore.com/api/v1/event/"
            f"{i}/statistics"
        )

        dados_estatisticas = pegar_dados(
            navegador,
            url
        )


        if dados_estatisticas is None or "statistics" not in dados_estatisticas:

            print(
                "Não foi possível pegar as estatísticas do jogo:",
                i
            )

            dados_estatisticas = {
                "statistics": []
            }


        # =================================================
        # CARTÕES AMARELOS
        # =================================================

        cartoes = pegar_estatisticas(
            dados_estatisticas,
            "yellowCards",
            "homeValue"
        )
        
        
        cartoes_casa.append(
            transoformar_int(cartoes)
        )

    
        cartoes_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "yellowCards",
            "awayValue"
        )

        cartoes_enfrenta_casa.append(
            transoformar_int(cartoes_enfrenta)
        )

        # =================================================
        # Gols Esperados
        # =================================================
        
        expected_casa = pegar_estatisticas(dados_estatisticas, "expectedGoals", "home")
        xg_casa.append(float(expected_casa or 0))
                
        expected_enfrenta_casa = pegar_estatisticas(dados_estatisticas, "expectedGoals", "away")
        xg_enfrenta_casa.append(float(expected_enfrenta_casa or 0))
        
        
        
        
        # =================================================
        # CARTÕES VERMELHOS
        # =================================================

        cartao_vermelho = pegar_estatisticas(
            dados_estatisticas,
            "redCards",
            "homeValue"
        )

        cartao_vermelho_casa.append(
            transoformar_int(cartao_vermelho)
        )


        cartao_vermelho_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "redCards",
            "awayValue"
        )

        cartao_vermelho_enfrenta_casa.append(
            transoformar_int(cartao_vermelho_enfrenta)
        )


        # =================================================
        # ESCANTEIOS
        # =================================================

        escanteios = pegar_estatisticas(
            dados_estatisticas,
            "cornerKicks",
            "homeValue"
        )

        escanteios_casa.append(
            transoformar_int(escanteios)
        )


        escanteios_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "cornerKicks",
            "awayValue"
        )

        escanteios_enfrenta_casa.append(
            transoformar_int(escanteios_enfrenta)
        )


        # =================================================
        # ESCANTEIOS 1º TEMPO
        # =================================================

        escanteios_1 = pegar_estatisticas_1tempo(
            dados_estatisticas,
            "cornerKicks",
            "homeValue"
        )

        escanteios1_casa.append(
            transoformar_int(escanteios_1)
        )


        escanteios_enfrenta_1 = pegar_estatisticas_1tempo(
            dados_estatisticas,
            "cornerKicks",
            "awayValue"
        )

        escanteios1_enfrenta_casa.append(
            transoformar_int(escanteios_enfrenta_1)
        )
        
                # =================================================
        # POSSE DE BOLA 1º TEMPO
        # =================================================
        
        posse = pegar_estatisticas_1tempo(
            dados_estatisticas,
            "ballPossession",
            "homeValue"
        )
        
        posse1_casa.append(posse)
        
        
        posse_enfrenta = pegar_estatisticas_1tempo(
            dados_estatisticas,
            "ballPossession",
            "awayValue"
        )
        
        posse1_enfrenta_casa.append(posse_enfrenta)

        # =================================================
        # FALTAS
        # =================================================

        faltas = pegar_estatisticas(
            dados_estatisticas,
            "fouls",
            "homeValue"
        )

        faltas_casa.append(
            transoformar_int(faltas)
        )


        faltas_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "fouls",
            "awayValue"
        )

        faltas_enfreta_casa.append(
            transoformar_int(faltas_enfrenta)
        )


        # =================================================
        # DEFESAS DO GOLEIRO
        # =================================================

        defesa = pegar_estatisticas(
            dados_estatisticas,
            "goalkeeperSaves",
            "homeValue"
        )

        defesa_casa.append(
            transoformar_int(defesa)
        )


        defesa_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "goalkeeperSaves",
            "awayValue"
        )

        defesa_enfrenta_casa.append(
            transoformar_int(defesa_enfrenta)
        )


        # =================================================
        # FINALIZAÇÕES
        # =================================================

        finalizacao = pegar_estatisticas(
            dados_estatisticas,
            "totalShotsOnGoal",
            "homeValue"
        )

        finalizacao_casa.append(
            transoformar_int(finalizacao)
        )


        finalizacao_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "totalShotsOnGoal",
            "awayValue"
        )

        finalizacao_enfrenta_casa.append(
            transoformar_int(finalizacao_enfrenta)
        )


        # =================================================
        # FINALIZAÇÕES NO GOL
        # =================================================

        chutes = pegar_estatisticas(
            dados_estatisticas,
            "shotsOnGoal",
            "homeValue"
        )

        chutes_nogol_casa.append(
            transoformar_int(chutes)
        )


        chutes_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "shotsOnGoal",
            "awayValue"
        )

        chutes_nogol_enfrenta_casa.append(
            transoformar_int(chutes_enfrenta)
        )


    # =====================================================
    # JOGOS DO TIME VISITANTE
    # =====================================================

    for i in id_jogo_visit:

        print("\nBuscando jogo:", i)

        # ALTERAÇÃO:
        # antes era api.sofascore.com
        # agora é www.sofascore.com

        url = f"https://www.sofascore.com/api/v1/event/{i}"

        dados = pegar_dados(
            navegador,
            url
        )


        # =================================================
        # VERIFICA SE VEIO O EVENT
        # =================================================

        if dados is None or "event" not in dados:

            print("Não foi possível pegar o jogo:", i)
            print("Resposta recebida:", dados)

            continue


        # =================================================
        # GOLS
        # =================================================

        gol_fora.append(
            dados["event"]["awayScore"]["normaltime"]
        )

        gol_sofrido_fora.append(
            dados["event"]["homeScore"]["normaltime"]
        )


        # =================================================
        # RESULTADO
        # =================================================

        if gol_fora[-1] > gol_sofrido_fora[-1]:

            resultado_fora.append("Vitoria")

        elif gol_fora[-1] < gol_sofrido_fora[-1]:

            resultado_fora.append("Derrota")

        else:

            resultado_fora.append("Empate")


        # =================================================
        # HANDICAP
        # =================================================

        if gol_fora[-1] < gol_sofrido_fora[-1]:

            handicap_fora.append(
                gol_fora[-1] - gol_sofrido_fora[-1]
            )

        else:

            handicap_fora.append("-")


        # =================================================
        # GOLS 1º TEMPO
        # =================================================

        gol1_fora.append(
            dados["event"]["awayScore"]["period1"]
        )

        gol1_sofrido_fora.append(
            dados["event"]["homeScore"]["period1"]
        )


        # =================================================
        # ESTATÍSTICAS
        # =================================================

        url = (
            f"https://www.sofascore.com/api/v1/event/"
            f"{i}/statistics"
        )

        dados_estatisticas = pegar_dados(
            navegador,
            url
        )


        if dados_estatisticas is None or "statistics" not in dados_estatisticas:

            print(
                "Não foi possível pegar as estatísticas do jogo:",
                i
            )

            dados_estatisticas = {
                "statistics": []
            }


        # =================================================
        # CARTÕES AMARELOS
        # =================================================

        cartoes = pegar_estatisticas(
            dados_estatisticas,
            "yellowCards",
            "awayValue"
        )

        cartoes_fora.append(
            transoformar_int(cartoes)
        )


        cartoes_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "yellowCards",
            "homeValue"
        )

        cartoes_enfrenta_fora.append(
            transoformar_int(cartoes_enfrenta)
        )


        # =================================================
        # CARTÕES VERMELHOS
        # =================================================

        cartao_vermelho = pegar_estatisticas(
            dados_estatisticas,
            "redCards",
            "awayValue"
        )

        cartao_vermelho_fora.append(
            transoformar_int(cartao_vermelho)
        )


        cartao_vermelho_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "redCards",
            "homeValue"
        )

        cartao_vermelho_enfrenta_fora.append(
            transoformar_int(cartao_vermelho_enfrenta)
        )


        # =================================================
        # ESCANTEIOS
        # =================================================

        escanteios = pegar_estatisticas(
            dados_estatisticas,
            "cornerKicks",
            "awayValue"
        )

        escanteios_fora.append(
            transoformar_int(escanteios)
        )


        escanteios_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "cornerKicks",
            "homeValue"
        )

        escanteios_enfrenta_fora.append(
            transoformar_int(escanteios_enfrenta)
        )


        # =================================================
        # ESCANTEIOS 1º TEMPO
        # =================================================

        escanteios_1 = pegar_estatisticas_1tempo(
            dados_estatisticas,
            "cornerKicks",
            "awayValue"
        )

        escanteios1_fora.append(
            transoformar_int(escanteios_1)
        )


        escanteios_enfrenta_1 = pegar_estatisticas_1tempo(
            dados_estatisticas,
            "cornerKicks",
            "homeValue"
        )

        escanteios1_enfrenta_fora.append(
            transoformar_int(escanteios_enfrenta_1)
        )
        
        
        # =================================================
        # POSSE DE BOLA 1º TEMPO
        # =================================================
        
        posse = pegar_estatisticas_1tempo(
            dados_estatisticas,
            "ballPossession",
            "awayValue"
        )
        
        posse1_fora.append(posse)
        
        
        posse_enfrenta = pegar_estatisticas_1tempo(
            dados_estatisticas,
            "ballPossession",
            "homeValue"
        )

        posse1_enfrenta_fora.append(posse_enfrenta)

        # =================================================
        # FALTAS
        # =================================================

        faltas = pegar_estatisticas(
            dados_estatisticas,
            "fouls",
            "awayValue"
        )

        faltas_fora.append(
            transoformar_int(faltas)
        )


        faltas_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "fouls",
            "homeValue"
        )

        faltas_enfreta_fora.append(
            transoformar_int(faltas_enfrenta)
        )


        # =================================================
        # DEFESAS DO GOLEIRO
        # =================================================

        defesa = pegar_estatisticas(
            dados_estatisticas,
            "goalkeeperSaves",
            "awayValue"
        )

        defesa_fora.append(
            transoformar_int(defesa)
        )


        defesa_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "goalkeeperSaves",
            "homeValue"
        )

        defesa_enfrenta_fora.append(
            transoformar_int(defesa_enfrenta)
        )


        # =================================================
        # FINALIZAÇÕES
        # =================================================

        finalizacao = pegar_estatisticas(
            dados_estatisticas,
            "totalShotsOnGoal",
            "awayValue"
        )

        finalizacao_fora.append(
            transoformar_int(finalizacao)
        )


        finalizacao_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "totalShotsOnGoal",
            "homeValue"
        )

        finalizacao_enfrenta_fora.append(
            transoformar_int(finalizacao_enfrenta)
        )


        # =================================================
        # FINALIZAÇÕES NO GOL
        # =================================================

        chutes = pegar_estatisticas(
            dados_estatisticas,
            "shotsOnGoal",
            "awayValue"
        )

        chutes_nogol_fora.append(
            transoformar_int(chutes)
        )


        chutes_enfrenta = pegar_estatisticas(
            dados_estatisticas,
            "shotsOnGoal",
            "homeValue"
        )

        chutes_nogol_enfrenta_fora.append(
            transoformar_int(chutes_enfrenta)
        )

        
        # =================================================
        # Gols Esperados
        # =================================================
        
        expected_fora = pegar_estatisticas(dados_estatisticas, "expectedGoals", "away")
        xg_fora.append(float(expected_fora or 0))
                
        expected_enfrenta_fora = pegar_estatisticas(dados_estatisticas, "expectedGoals", "home")
        xg_enfrenta_fora.append(float(expected_enfrenta_fora or 0))
        
        
        
        
    # =====================================================
    # FECHA O NAVEGADOR
    # =====================================================

    navegador.quit()


    # =====================================================
    # VERIFICA SE EXISTEM DADOS
    # =====================================================

    if len(gol_casa) == 0 or len(gol_fora) == 0:

        print("\nNenhum jogo foi carregado corretamente.")

        print(
            "Isso provavelmente significa que o SofaScore "
            "continuou retornando 403."
        )

        print("\nIDs casa:", id_jogo_casa)
        print("IDs visitante:", id_jogo_visit)

        opcao = input(
            "\nDeseja tentar novamente? (s/n): "
        )

        if opcao.lower() == "n":
            break

        else:
            continue


    # =====================================================
    # CALCULA AS MÉDIAS
    # =====================================================

    # GOLS

    m_gol_casa = sum(gol_casa) / len(gol_casa)

    m_gol_fora = sum(gol_fora) / len(gol_fora)


    # GOLS 1º TEMPO

    m_gol1_casa = sum(gol1_casa) / len(gol1_casa)

    m_gol1_fora = sum(gol1_fora) / len(gol1_fora)


    # GOLS SOFRIDOS

    m_gol_sofrido_casa = (
        sum(gol_sofrido_casa) /
        len(gol_sofrido_casa)
    )

    m_gol_sofrido_fora = (
        sum(gol_sofrido_fora) /
        len(gol_sofrido_fora)
    )


    # GOLS SOFRIDOS 1º TEMPO

    m_gol1_sofrido_casa = (
        sum(gol1_sofrido_casa) /
        len(gol1_sofrido_casa)
    )

    m_gol1_sofrido_fora = (
        sum(gol1_sofrido_fora) /
        len(gol1_sofrido_fora)
    )


    # =====================================================
    # CARTÕES
    # =====================================================

    m_cartoes_casa = (
        sum(cartoes_casa) /
        len(cartoes_casa)
    )

    m_cartoes_fora = (
        sum(cartoes_fora) /
        len(cartoes_fora)
    )


    m_cartoes_enfrenta_casa = (
        sum(cartoes_enfrenta_casa) /
        len(cartoes_enfrenta_casa)
    )

    m_cartoes_enfrenta_fora = (
        sum(cartoes_enfrenta_fora) /
        len(cartoes_enfrenta_fora)
    )


    # =====================================================
    # CARTÕES VERMELHOS
    # =====================================================

    m_cartao_vermelho_casa = (
        sum(cartao_vermelho_casa) /
        len(cartao_vermelho_casa)
    )

    m_cartao_vermelho_enfrenta_casa = (
        sum(cartao_vermelho_enfrenta_casa) /
        len(cartao_vermelho_enfrenta_casa)
    )

    m_cartao_vermelho_fora = (
        sum(cartao_vermelho_fora) /
        len(cartao_vermelho_fora)
    )

    m_cartao_vermelho_enfrenta_fora = (
        sum(cartao_vermelho_enfrenta_fora) /
        len(cartao_vermelho_enfrenta_fora)
    )


    # =====================================================
    # ESCANTEIOS
    # =====================================================

    m_escanteios_casa = (
        sum(escanteios_casa) /
        len(escanteios_casa)
    )

    m_escanteios_fora = (
        sum(escanteios_fora) /
        len(escanteios_fora)
    )


    m_escanteios_enfrenta_casa = (
        sum(escanteios_enfrenta_casa) /
        len(escanteios_enfrenta_casa)
    )

    m_escanteios_enfrenta_fora = (
        sum(escanteios_enfrenta_fora) /
        len(escanteios_enfrenta_fora)
    )


    # =====================================================
    # ESCANTEIOS 1º TEMPO
    # =====================================================

    m_escanteios1_casa = (
        sum(escanteios1_casa) /
        len(escanteios1_casa)
    )

    m_escanteios1_fora = (
        sum(escanteios1_fora) /
        len(escanteios1_fora)
    )


    m_escanteios1_enfrenta_casa = (
        sum(escanteios1_enfrenta_casa) /
        len(escanteios1_enfrenta_casa)
    )

    m_escanteios1_enfrenta_fora = (
        sum(escanteios1_enfrenta_fora) /
        len(escanteios1_enfrenta_fora)
    )


    # =====================================================
    # FALTAS
    # =====================================================

    m_faltas_casa = (
        sum(faltas_casa) /
        len(faltas_casa)
    )

    m_faltas_fora = (
        sum(faltas_fora) /
        len(faltas_fora)
    )


    m_faltas_enfrenta_casa = (
        sum(faltas_enfreta_casa) /
        len(faltas_enfreta_casa)
    )

    m_faltas_enfrenta_fora = (
        sum(faltas_enfreta_fora) /
        len(faltas_enfreta_fora)
    )


    # =====================================================
    # DEFESAS DO GOLEIRO
    # =====================================================

    m_defesa_casa = (
        sum(defesa_casa) /
        len(defesa_casa)
    )

    m_defesa_fora = (
        sum(defesa_fora) /
        len(defesa_fora)
    )


    m_defesa_enfrenta_casa = (
        sum(defesa_enfrenta_casa) /
        len(defesa_enfrenta_casa)
    )

    m_defesa_enfrenta_fora = (
        sum(defesa_enfrenta_fora) /
        len(defesa_enfrenta_fora)
    )


    # =====================================================
    # FINALIZAÇÕES
    # =====================================================

    m_finalizacao_casa = (
        sum(finalizacao_casa) /
        len(finalizacao_casa)
    )

    m_finalizacao_fora = (
        sum(finalizacao_fora) /
        len(finalizacao_fora)
    )


    m_finalizacao_enfrenta_casa = (
        sum(finalizacao_enfrenta_casa) /
        len(finalizacao_enfrenta_casa)
    )

    m_finalizacao_enfrenta_fora = (
        sum(finalizacao_enfrenta_fora) /
        len(finalizacao_enfrenta_fora)
    )


    # =====================================================
    # FINALIZAÇÕES NO GOL
    # =====================================================

    m_chutes_nogol_casa = (
        sum(chutes_nogol_casa) /
        len(chutes_nogol_casa)
    )

    m_chutes_nogol_fora = (
        sum(chutes_nogol_fora) /
        len(chutes_nogol_fora)
    )


    m_chutes_nogol_enfrenta_casa = (
        sum(chutes_nogol_enfrenta_casa) /
        len(chutes_nogol_enfrenta_casa)
    )

    m_chutes_nogol_enfrenta_fora = (
        sum(chutes_nogol_enfrenta_fora) /
        len(chutes_nogol_enfrenta_fora)
    )


    # =====================================================
    # PRINTS DAS MÉDIAS
    # =====================================================

    print("\n")
    print("==============================================")
    print("              MÉDIAS TIME DA CASA")
    print("==============================================")

    print(
        "Média de gols feitos:",
        round(m_gol_casa, 2)
    )

    print(
        "Média de gols feitos 1T:",
        round(m_gol1_casa, 2)
    )

    print(
        "Média de gols sofridos:",
        round(m_gol_sofrido_casa, 2)
    )

    print(
        "Média de gols sofridos 1T:",
        round(m_gol1_sofrido_casa, 2)
    )

    print(
        "Média de cartões:",
        round(m_cartoes_casa, 2)
    )

    print(
        "Média de cartões adversário:",
        round(m_cartoes_enfrenta_casa, 2)
    )

    print(
        "Média de cartões vermelhos:",
        round(m_cartao_vermelho_casa, 2)
    )

    print(
        "Média de cartões vermelhos adversário:",
        round(m_cartao_vermelho_enfrenta_casa, 2)
    )

    print(
        "Média de escanteios:",
        round(m_escanteios_casa, 2)
    )

    print(
        "Média de escanteios adversário:",
        round(m_escanteios_enfrenta_casa, 2)
    )

    print(
        "Média de escanteios 1T:",
        round(m_escanteios1_casa, 2)
    )

    print(
        "Média de escanteios adversário 1T:",
        round(m_escanteios1_enfrenta_casa, 2)
    )

    print(
        "Média de faltas:",
        round(m_faltas_casa, 2)
    )

    print(
        "Média de faltas adversário:",
        round(m_faltas_enfrenta_casa, 2)
    )

    print(
        "Média de defesas do goleiro:",
        round(m_defesa_casa, 2)
    )

    print(
        "Média de defesas do goleiro adversário:",
        round(m_defesa_enfrenta_casa, 2)
    )

    print(
        "Média de finalizações:",
        round(m_finalizacao_casa, 2)
    )

    print(
        "Média de finalizações adversário:",
        round(m_finalizacao_enfrenta_casa, 2)
    )

    print(
        "Média de finalizações no gol:",
        round(m_chutes_nogol_casa, 2)
    )

    print(
        "Média de finalizações no gol adversário:",
        round(m_chutes_nogol_enfrenta_casa, 2)
    )


    print("\n")
    print("==============================================")
    print("            MÉDIAS TIME VISITANTE")
    print("==============================================")

    print(
        "Média de gols feitos:",
        round(m_gol_fora, 2)
    )

    print(
        "Média de gols feitos 1T:",
        round(m_gol1_fora, 2)
    )

    print(
        "Média de gols sofridos:",
        round(m_gol_sofrido_fora, 2)
    )

    print(
        "Média de gols sofridos 1T:",
        round(m_gol1_sofrido_fora, 2)
    )

    print(
        "Média de cartões:",
        round(m_cartoes_fora, 2)
    )

    print(
        "Média de cartões adversário:",
        round(m_cartoes_enfrenta_fora, 2)
    )

    print(
        "Média de cartões vermelhos:",
        round(m_cartao_vermelho_fora, 2)
    )

    print(
        "Média de cartões vermelhos adversário:",
        round(m_cartao_vermelho_enfrenta_fora, 2)
    )

    print(
        "Média de escanteios:",
        round(m_escanteios_fora, 2)
    )

    print(
        "Média de escanteios adversário:",
        round(m_escanteios_enfrenta_fora, 2)
    )

    print(
        "Média de escanteios 1T:",
        round(m_escanteios1_fora, 2)
    )

    print(
        "Média de escanteios adversário 1T:",
        round(m_escanteios1_enfrenta_fora, 2)
    )

    print(
        "Média de faltas:",
        round(m_faltas_fora, 2)
    )

    print(
        "Média de faltas adversário:",
        round(m_faltas_enfrenta_fora, 2)
    )

    print(
        "Média de defesas do goleiro:",
        round(m_defesa_fora, 2)
    )

    print(
        "Média de defesas do goleiro adversário:",
        round(m_defesa_enfrenta_fora, 2)
    )

    print(
        "Média de finalizações:",
        round(m_finalizacao_fora, 2)
    )

    print(
        "Média de finalizações adversário:",
        round(m_finalizacao_enfrenta_fora, 2)
    )

    print(
        "Média de finalizações no gol:",
        round(m_chutes_nogol_fora, 2)
    )

    print(
        "Média de finalizações no gol adversário:",
        round(m_chutes_nogol_enfrenta_fora, 2)
    )


    # =====================================================
    # MOSTRA OS IDS
    # =====================================================

    print("\nIDs casa:", id_jogo_casa)
    print("IDs visitante:", id_jogo_visit)


    # =====================================================
    # CRIA O EXCEL
    # =====================================================

    arquivo = Workbook()

    planilha = arquivo.active

    planilha.title = "Estatísticas"


    # =====================================================
    # CABEÇALHO
    # =====================================================

    cabecalho = [

        "Partida",
        "Time",
        "Gols",
        "Gols 1T",
        "Gols Sofridos",
        "Gols Sofridos 1T",
        "Cartões Amarelos",
        "Cartões Amarelos Adversário",
        "Cartões Vermelhos",
        "Cartões Vermelhos Adversário",
        "Escanteios",
        "Escanteios Adversário",
        "Escanteios 1T",
        "Escanteios Adversário 1T",
        "Faltas",
        "Faltas Adversário",
        "Defesas do Goleiro",
        "Defesas do Goleiro Adversário",
        "Finalizações",
        "Finalizações Adversário",
        "Finalizações no gol",
        "Finalizações no gol Adversário",
        "Dupla Chance",
        "Handicap",
        "Posse de Bola 1T",
        "Posse de Bola Adversário 1T",
        "XG",
        "XG Adversário"

    ]


    for coluna in range(len(cabecalho)):

        planilha.cell(
            1,
            coluna + 1,
            cabecalho[coluna]
        )


    # =====================================================
    # PARTIDAS DO TIME DA CASA
    # =====================================================

    linha = 2


    for i in range(len(gol_casa)):

        planilha.cell(
            linha,
            1,
            f"Partida {i + 1}"
        )

        planilha.cell(
            linha,
            2,
            "Casa"
        )


        planilha.cell(
            linha,
            3,
            gol_casa[i]
        )

        planilha.cell(
            linha,
            4,
            gol1_casa[i]
        )


        planilha.cell(
            linha,
            5,
            gol_sofrido_casa[i]
        )

        planilha.cell(
            linha,
            6,
            gol1_sofrido_casa[i]
        )


        planilha.cell(
            linha,
            7,
            cartoes_casa[i]
        )

        planilha.cell(
            linha,
            8,
            cartoes_enfrenta_casa[i]
        )


        planilha.cell(
            linha,
            9,
            cartao_vermelho_casa[i]
        )

        planilha.cell(
            linha,
            10,
            cartao_vermelho_enfrenta_casa[i]
        )


        planilha.cell(
            linha,
            11,
            escanteios_casa[i]
        )

        planilha.cell(
            linha,
            12,
            escanteios_enfrenta_casa[i]
        )


        planilha.cell(
            linha,
            13,
            escanteios1_casa[i]
        )

        planilha.cell(
            linha,
            14,
            escanteios1_enfrenta_casa[i]
        )


        planilha.cell(
            linha,
            15,
            faltas_casa[i]
        )

        planilha.cell(
            linha,
            16,
            faltas_enfreta_casa[i]
        )


        planilha.cell(
            linha,
            17,
            defesa_casa[i]
        )

        planilha.cell(
            linha,
            18,
            defesa_enfrenta_casa[i]
        )


        planilha.cell(
            linha,
            19,
            finalizacao_casa[i]
        )

        planilha.cell(
            linha,
            20,
            finalizacao_enfrenta_casa[i]
        )


        planilha.cell(
            linha,
            21,
            chutes_nogol_casa[i]
        )

        planilha.cell(
            linha,
            22,
            chutes_nogol_enfrenta_casa[i]
        )


        planilha.cell(
            linha,
            23,
            resultado_casa[i]
        )

        planilha.cell(
            linha,
            24,
            handicap_casa[i]
        )
        
        planilha.cell(
            linha,
            25,
            posse1_casa[i]
        )
        
        planilha.cell(
            linha,
            26,
            posse1_enfrenta_casa[i]
        )
        
        planilha.cell(linha, 27, xg_casa[i])
        planilha.cell(linha, 28, xg_enfrenta_casa[i])

        linha += 1


    # =====================================================
    # PARTIDAS DO TIME VISITANTE
    # =====================================================

    for i in range(len(gol_fora)):

        planilha.cell(
            linha,
            1,
            f"Partida {i + 1}"
        )

        planilha.cell(
            linha,
            2,
            "Visitante"
        )


        planilha.cell(
            linha,
            3,
            gol_fora[i]
        )

        planilha.cell(
            linha,
            4,
            gol1_fora[i]
        )


        planilha.cell(
            linha,
            5,
            gol_sofrido_fora[i]
        )

        planilha.cell(
            linha,
            6,
            gol1_sofrido_fora[i]
        )


        planilha.cell(
            linha,
            7,
            cartoes_fora[i]
        )

        planilha.cell(
            linha,
            8,
            cartoes_enfrenta_fora[i]
        )


        planilha.cell(
            linha,
            9,
            cartao_vermelho_fora[i]
        )

        planilha.cell(
            linha,
            10,
            cartao_vermelho_enfrenta_fora[i]
        )


        planilha.cell(
            linha,
            11,
            escanteios_fora[i]
        )

        planilha.cell(
            linha,
            12,
            escanteios_enfrenta_fora[i]
        )


        planilha.cell(
            linha,
            13,
            escanteios1_fora[i]
        )

        planilha.cell(
            linha,
            14,
            escanteios1_enfrenta_fora[i]
        )


        planilha.cell(
            linha,
            15,
            faltas_fora[i]
        )

        planilha.cell(
            linha,
            16,
            faltas_enfreta_fora[i]
        )


        planilha.cell(
            linha,
            17,
            defesa_fora[i]
        )

        planilha.cell(
            linha,
            18,
            defesa_enfrenta_fora[i]
        )


        planilha.cell(
            linha,
            19,
            finalizacao_fora[i]
        )

        planilha.cell(
            linha,
            20,
            finalizacao_enfrenta_fora[i]
        )


        planilha.cell(
            linha,
            21,
            chutes_nogol_fora[i]
        )

        planilha.cell(
            linha,
            22,
            chutes_nogol_enfrenta_fora[i]
        )


        planilha.cell(
            linha,
            23,
            resultado_fora[i]
        )

        planilha.cell(
            linha,
            24,
            handicap_fora[i]
        )
        
        planilha.cell(
            linha,
            25,
            posse1_fora[i]
        )
        
        planilha.cell(
            linha,
            26,
            posse1_enfrenta_fora[i]
        )
        

        planilha.cell(linha, 27, xg_fora[i])
        planilha.cell(linha, 28, xg_enfrenta_fora[i])
                
        linha += 1


    # =====================================================
    # AJUSTAR LARGURA DAS COLUNAS
    # =====================================================

    for coluna in range(1, 29):

        planilha.column_dimensions[
            planilha.cell(
                1,
                coluna
            ).column_letter
        ].width = 25


    # =====================================================
    # SALVAR
    # =====================================================

    arquivo.save(
        "estatisticas.xlsx"
    )


    print(
        "\nArquivo Excel criado com sucesso!"
    )


    # =====================================================
    # NOVA ANÁLISE
    # =====================================================

    opcao = input(
        "Deseja fazer uma nova analise? (s/n): "
    )

    if opcao.lower() == "n":

        print("Programa finalizado!")

        break