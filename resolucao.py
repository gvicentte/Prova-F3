import pulp as pl
from pptx import Presentation
from pptx.util import Inches

# --- PARTE 1: MODELAGEM E SOLUÇÃO (PLI USANDO PuLP) ---

def resolver_problema_logistico():
    """Modela e resolve um exemplo simplificado do problema de custo logístico."""
    
    # 1. Dados Exemplo (Simplificados)
    cidades = ['A', 'B', 'C']  # Destinos de jogo
    modais = ['Aviao', 'Onibus']
    
    # Custo de transporte por pessoa (R$)
    custos_transporte = {
        ('A', 'B', 'Aviao'): 300, ('A', 'B', 'Onibus'): 150,
        ('B', 'C', 'Aviao'): 400, ('B', 'C', 'Onibus'): 200,
        ('C', 'A', 'Aviao'): 600, ('C', 'A', 'Onibus'): 350
    }
    
    # Custo de hospedagem (R$)
    custos_hospedagem = {'A': 2000, 'B': 1500, 'C': 2500}
    
    N = 30  # Número de pessoas na delegação
    B_max = 30000  # Orçamento Máximo (R$)
    
    # 2. Criar o Problema
    prob = pl.LpProblem("Otimizacao_Logistica_Futebol", pl.LpMinimize)
    
    # 3. Variáveis de Decisão
    # X_ij_t: 1 se usar modal t de i para j
    X = pl.LpVariable.dicts("X", custos_transporte.keys(), cat=pl.LpBinary)
    
    # Y_j: 1 se pernoitar na cidade j
    Y = pl.LpVariable.dicts("Y", cidades, cat=pl.LpBinary)

    # 4. Função Objetivo (Minimizar Custo Total)
    custo_transporte_total = pl.lpSum([custos_transporte[i, j, t] * X[i, j, t] * N
                                       for (i, j, t) in custos_transporte])
    
    custo_hospedagem_total = pl.lpSum([custos_hospedagem[j] * Y[j] for j in cidades])
    
    prob += custo_transporte_total + custo_hospedagem_total, "Custo Total"
    
    # 5. Restrições
    
    # R1: Restrição Orçamentária
    prob += custo_transporte_total + custo_hospedagem_total <= B_max, "Orcamento_Maximo"
    
    # R2: Cobertura de Deslocamento (Garantir 3 deslocamentos no ciclo A->B, B->C, C->A)
    # De A para B
    prob += X['A', 'B', 'Aviao'] + X['A', 'B', 'Onibus'] == 1, "Deslocamento_A_B"
    # De B para C
    prob += X['B', 'C', 'Aviao'] + X['B', 'C', 'Onibus'] == 1, "Deslocamento_B_C"
    # De C para A (Retorno ou próximo destino)
    prob += X['C', 'A', 'Aviao'] + X['C', 'A', 'Onibus'] == 1, "Deslocamento_C_A"

    # R3: Exemplo de Regra Logística (Apenas pernoitar em C)
    prob += Y['A'] + Y['B'] == 0, "Pernoite_A_B_Zero"
    prob += Y['C'] == 1, "Pernoite_C_Obrigatoria"
    
    # 6. Resolver
    prob.solve(pl.PULP_CBC_CMD(msg=0)) # Usa o solver padrão (CBC)
    
    solucao = {
        'status': pl.LpStatus[prob.status],
        'custo_minimo': prob.objective.value(),
        'plano': []
    }
    
    for (i, j, t) in custos_transporte.keys():
        if pl.value(X[i, j, t]) == 1:
            solucao['plano'].append(f'De {i} para {j}: {t}')
            
    for j in cidades:
        if pl.value(Y[j]) == 1:
            solucao['plano'].append(f'Hospedagem em {j}: SIM')
    
    return solucao

# --- EXECUÇÃO ---

print("Iniciando a modelagem do problema e geração da apresentação...")

# 1. Resolve o PLI de exemplo
solucao_final = resolver_problema_logistico()
print("\n--- Resultado da Otimização (PuLP) ---")
print(f"Status: {solucao_final['status']}")
print(f"Custo Mínimo Encontrado: R$ {solucao_final['custo_minimo']:.2f}")
print("--------------------------------------\n")
