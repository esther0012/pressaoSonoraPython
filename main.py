import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# Carregar dados de um arquivo Excel
caminho_arquivo_excel = "C:\\Users\\esthe\\Downloads\\Projeto1_A.xlsx"
dados_excel = pd.read_excel(caminho_arquivo_excel)

# Certifique-se de que a coluna de data esteja em formato de data
dados_excel['Período de início'] = pd.to_datetime(dados_excel['Período de início'])

# Ordenar os dados pela coluna de data
dados_excel = dados_excel.sort_values(by='Período de início')

# Aumentar a largura do gráfico
largura_figura = 24  # Largura em polegadas
altura_figura = 6   # Altura em polegadas
plt.figure(figsize=(largura_figura, altura_figura))

# Plotar os pontos de dispersão
plt.scatter(dados_excel['Período de início'], dados_excel['Canal 1'], color='blue', label='Pontos de Dispersão')

# Plotar as linhas retas conectando os pontos
plt.plot(dados_excel['Período de início'], dados_excel['Canal 1'], linestyle='-', color='red', label='Linhas Retas')

# Adicionar legenda ao gráfico
plt.legend()

# Configurar rótulos e título do gráfico
plt.xlabel('Período de início')
plt.ylabel('Canal 1')
plt.title('Net Promoter Score (Canal 1) em Função do Tempo')

# Rotacionar os rótulos do eixo x para facilitar a leitura
plt.xticks(rotation=45)

# Inserir grades no gráfico
plt.grid(True, linestyle='--', alpha=0.7)

# Exibir o gráfico
plt.tight_layout()  # Ajustar layout para evitar sobreposição de elementos
plt.show()

# Converter a coluna "Período de início" para formato de data
dados_excel['Período de início'] = pd.to_datetime(dados_excel['Período de início'], errors='coerce')

# Remover linhas com datas inválidas (NaT após a conversão)
dados_excel = dados_excel.dropna(subset=['Período de início'])

# Remover linhas com valores ausentes na coluna "Canal 1"
dados_sem_nulos = dados_excel.dropna(subset=['Canal 1'])

# Encontrar o índice do maior NPS nas linhas sem nulos
indice_maior_nps = dados_sem_nulos['Canal 1'].idxmax()

# Obter o horário (data) correspondente ao maior NPS nas linhas sem nulos
maior_horario_nps = dados_sem_nulos.loc[indice_maior_nps, 'Período de início']

print("Maior Horário do NPS:", maior_horario_nps)

# Função para calcular a derivada numérica usando a diferença finita central
def calcular_derivada_numerica(valores, h):
    derivadas = []
    for i in range(1, len(valores) - 1):
        derivada = (valores.iloc[i + 1] - valores.iloc[i - 1]) / (2 * h)
        derivadas.append(derivada)
    return derivadas
# Coluna de dados para calcular a derivada
coluna_desejada = 'Canal 1'
# Intervalo para a diferença finita central
h = 1
# Calcular a derivada numérica
derivadas = calcular_derivada_numerica(dados_excel[coluna_desejada], h)
# Redefinir os índices do DataFrame após a ordenação
dados_excel.reset_index(drop=True, inplace=True)
# Criar uma tabela com os resultados
tabela_resultado = pd.DataFrame({
    'Período de início': dados_excel['Período de início'][1:-1],  # Ignorar os primeiros e últimos pontos
    'Valor': dados_excel[coluna_desejada][1:-1],  # Ignorar os primeiros e últimos pontos
    'Derivada Numérica': derivadas
})
# Exibir a tabela de resultados
print(tabela_resultado)

# Criar o gráfico da taxa de variação em função do tempo
# Aumentar a largura do gráfico
largura_figura = 24  # Largura em polegadas
altura_figura = 6   # Altura em polegadas
plt.figure(figsize=(largura_figura, altura_figura))
plt.plot(tabela_resultado['Período de início'], tabela_resultado['Derivada Numérica'], marker='o')

# Configurar rótulos e título do gráfico
plt.xlabel('Período de início')
plt.ylabel('Taxa de Variação (Derivada Numérica)')
plt.title('Taxa de Variação em Função do Tempo')

# Rotacionar os rótulos do eixo x para facilitar a leitura
plt.xticks(rotation=45)

# Inserir grades no gráfico
plt.grid(True, linestyle='--', alpha=0.7)

# Exibir o gráfico
plt.tight_layout()
plt.show()

# Encontrar o índice da maior variação (maior valor na coluna da derivada)
indice_maior_variacao = tabela_resultado['Derivada Numérica'].idxmax()

# Obter o horário correspondente à maior variação
horario_maior_variacao = tabela_resultado.loc[indice_maior_variacao, 'Período de início']

print("Horário da Maior Variação:", horario_maior_variacao)

# Calcular média, mediana e moda
media = tabela_resultado['Valor'].mean()
mediana = tabela_resultado['Valor'].median()
moda = tabela_resultado['Valor'].mode().iloc[0]  # Usando o método mode() do Pandas

# Criar o gráfico
largura_figura = 20  # Largura em polegadas
altura_figura = 6   # Altura em polegadas
plt.figure(figsize=(largura_figura, altura_figura))
plt.plot(tabela_resultado['Período de início'], tabela_resultado['Valor'], marker='o', label='Dados')
plt.axhline(y=media, color='r', linestyle='--', label='Média')
plt.axhline(y=mediana, color='g', linestyle='--', label='Mediana')
plt.axhline(y=moda, color='b', linestyle='--', label='Moda')

# Configurar rótulos e título do gráfico
plt.xlabel('Período de início')
plt.ylabel('Valor')
plt.title('Média, Mediana e Moda em Função do Tempo')
plt.xticks(rotation=45)
plt.legend()

# Inserir grades no gráfico
plt.grid(True, linestyle='--', alpha=0.7)

# Exibir o gráfico
plt.tight_layout()
plt.show()

amplitude_amostral = dados_excel[coluna_desejada].max() - dados_excel[coluna_desejada].min()

print("Amplitude Amostral:", amplitude_amostral)

numero_classes = int(1 + np.log2(len(dados_excel[coluna_desejada])))

print("Número de Classes:", numero_classes)

# Criar o histograma
plt.figure(figsize=(10, 6))
plt.hist(dados_excel[coluna_desejada], bins=numero_classes, edgecolor='k')

# Configurar rótulos e título do gráfico
plt.xlabel('Valor')
plt.ylabel('Frequência')
plt.title('Histograma dos Dados')

# Exibir o gráfico
plt.tight_layout()
plt.show()

numero_classes = 10  # Ajuste conforme necessário

# Calcular a amplitude de classes
maior_valor = dados_excel[coluna_desejada].max()
menor_valor = dados_excel[coluna_desejada].min()
amplitude_classes = (maior_valor - menor_valor) / numero_classes

print("Amplitude de Classes:", amplitude_classes)

# Calcular os limites de classes
limites_classes = [menor_valor + i * amplitude_classes for i in range(numero_classes + 1)]

print("Limites de Classes:", limites_classes)

# Definir limites das classes
limites_classes = [menor_valor + i * amplitude_classes for i in range(numero_classes + 1)]

# Criar os intervalos de classe usando pd.cut
dados_excel['Classes'] = pd.cut(dados_excel[coluna_desejada], bins=limites_classes)

# Calcular as frequências absolutas das classes
frequencias_absolutas = dados_excel['Classes'].value_counts().sort_index()

print("Frequências Absolutas das Classes:")
print(frequencias_absolutas)

# Calcular as frequências relativas das classes
frequencias_relativas = dados_excel['Classes'].value_counts(normalize=True).sort_index()

print("Frequências Relativas das Classes:")
print(frequencias_relativas)

# Calcular os pontos médios das classes
pontos_medios = [(limite_inferior + limite_superior) / 2 for limite_inferior, limite_superior in zip(limites_classes[:-1], limites_classes[1:])]

print("Pontos Médios das Classes:")
print(pontos_medios)

# Calcular o valor mínimo
valor_minimo = dados_excel[coluna_desejada].min()

# Calcular o valor máximo
valor_maximo = dados_excel[coluna_desejada].max()

# Calcular os percentis
percentis = [10, 25, 50, 75, 90]
valores_percentis = dados_excel[coluna_desejada].quantile([p / 100 for p in percentis])

print("Valor Mínimo:", valor_minimo)
print("Valor Máximo:", valor_maximo)
for percentil, valor in zip(percentis, valores_percentis):
    print(f"Percentil {percentil}%:", valor)

