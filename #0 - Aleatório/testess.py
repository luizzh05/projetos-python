import pandas as pd
import time as time
import matplotlib.pyplot as plt

inicio = time.time()

input_file = "Tabela_Links.csv"
output_xlsx = "teste.xlsx"

df = pd.read_csv(input_file, sep=";", dtype=str, encoding="utf-8")

novo = df[["nom_fantasia", "Nome_Cidade", "Sit_Cliente", "CNPJ_NEO"]].copy()

novo = novo.rename(columns={
    "nom_fantasia": "Nome fantasia",
    "Nome_Cidade": "Nome Cidade",
    "Sit_Cliente": "Sit Cliente",
    "CNPJ_NEO": "CNPJ"
})

novo["CNPJ"] = novo["CNPJ"].str.replace(
    r"(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})",
    r"\1.\2.\3/\4-\5",
    regex=True
)

novo.sort_values('Nome Cidade', inplace=True)

resultado = novo.groupby("Nome Cidade").size().reset_index(name='count')
resultado = resultado.sort_values("Quantidade", ascending=False).head(10)

plt.bar(resultado["Nome Cidade"], resultado["Quantidade"])

plt.title("Top 10 Cidades por Quantidade de Registros")
plt.xlabel("Nome Cidade")
plt.ylabel("Quantidade de Registros")

plt.xticks(rotation=45)

plt.show()

novo.to_excel(output_xlsx, index=False)

fim = time.time()

print(f"Foram gerados {len(novo)} registros no arquivo {output_xlsx}")
print(f"Tempo gasto: {fim - inicio} segundos")