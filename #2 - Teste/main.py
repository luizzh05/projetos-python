import seaborn as sns
import pandas as pd

df = sns.load_dataset("flights")

ano_passageiros = df.groupby("year")["passengers"].sum()
ano_passageiros = ano_passageiros.sort_values(ascending=False).head(1)

ano = ano_passageiros.index[0]
passageiros = ano_passageiros.values[0]

print(f"Ano com mais passageiros: {ano}")
print(f"assageiros: {passageiros}")
print(f"mes passageiros: {mes_passageiros}")
