# ============================================================
#  🚢 EXPLORAÇÃO DO TITANIC COM PANDAS
#  Projeto de aprendizado — do básico ao intermediário
# ============================================================
#
# COMO COMEÇAR:
#   1. Instale as dependências:
#        pip install pandas matplotlib seaborn
#
#   2. Baixe o dataset:
#        Kaggle: https://www.kaggle.com/datasets/yasserh/titanic-dataset
#        (arquivo: Titanic-Dataset.csv)
#        OU use o seaborn para carregar automaticamente (ver abaixo)
#
# ============================================================

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# # ────────────────────────────────────────────────────────────
# # 📥 1. CARREGANDO OS DADOS
# # ────────────────────────────────────────────────────────────

# Opção A: carregar do arquivo baixado do Kaggle
# df = pd.read_csv("Titanic-Dataset.csv")

# Opção B: carregar direto pelo seaborn (sem precisar baixar)
df = sns.load_dataset("titanic")
df.to_excel("Titanic-Dataset.xlsx", index=False)  # Salva em Excel para referência futura

print("✅ Dataset carregado com sucesso!")
print(f"   Linhas: {df.shape[0]} | Colunas: {df.shape[1]}\n")

taxa = df['survived'].mean() * 100
print(f"{taxa:.1f}% dos passageiros sobreviveram ({df['survived'].sum()} de {len(df)})\n")

# # 2. ────────────────────────────────────────────────────────────
resultado = df.groupby("sex")["survived"].mean() * 100
contagem = resultado.count()

print(f"Mulheres: {resultado['female']:.1f}% sobreviveram")
print(f"Homens: {resultado['male']:.1f}% sobreviveram")

# # 3. ────────────────────────────────────────────────────────────
media = df['age'].mean()
print(f"A idade média dos passageiros é: {media:.1f} anos")
por_sobrev = df.groupby("survived")['age'].describe().loc[1, 'mean']
print(f"\nIdade média dos sobreviventes: {por_sobrev:.1f} anos")

# # 4. ────────────────────────────────────────────────────────────
resultadoo = df.groupby("pclass")['survived'].agg(
    total='count',
    sobreviventes='sum',
)
resultadoo['taxa_sobrevivencia'] = (resultadoo['sobreviventes'] / resultadoo['total']) * 100
print("\nTaxa de sobrevivência por classe:")
print(resultadoo[['total', 'sobreviventes', 'taxa_sobrevivencia']].round(0))

# # 5. ────────────────────────────────────────────────────────────
