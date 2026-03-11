from Carro import Carro
import streamlit as st
import requests


def buscar_imagem_carro(marca, modelo):
    """Busca imagem do carro na Wikipedia pela marca e modelo."""
    busca = f"{marca} {modelo}"
    url = "https://en.wikipedia.org/w/api.php"
    headers = {"User-Agent": "CarApp/1.0"}
    params = {
        "action": "query",
        "list": "search",
        "srsearch": busca,
        "format": "json",
    }
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=10)
        resp.raise_for_status()
        resultados = resp.json().get("query", {}).get("search", [])
        if not resultados:
            return None

        titulo = resultados[0]["title"]
        params_img = {
            "action": "query",
            "titles": titulo,
            "prop": "pageimages",
            "format": "json",
            "pithumbsize": 600,
        }
        resp_img = requests.get(url, params=params_img, headers=headers, timeout=10)
        resp_img.raise_for_status()
        pages = resp_img.json().get("query", {}).get("pages", {})
        for page in pages.values():
            thumb = page.get("thumbnail", {}).get("source")
            if thumb:
                return thumb
    except requests.RequestException:
        return None
    return None


st.title("Informações do Carro")

marca_input = st.text_input("Digite a marca do carro:")
modelo_input = st.text_input("Digite o modelo do carro:")
ano_input = st.number_input("Digite o ano do carro:", min_value=1900, max_value=2100, step=1)

usar_upload = st.checkbox("Enviar imagem manualmente")
file_input = None
if usar_upload:
    file_input = st.file_uploader("Faça upload de uma imagem do carro", type=["webp", "jpg", "jpeg", "png"])

if st.button("Criar Carro", width="stretch", key="criar_carro"):
    novo_carro = Carro(marca_input, modelo_input, ano_input)
    st.title("Informações do Carro")
    st.write(f"Marca: {novo_carro.marca}")
    st.write(f"Modelo: {novo_carro.modelo}")
    st.write(f"Ano: {novo_carro.ano}")

    if file_input:
        st.image(file_input)
    else:
        with st.spinner("Buscando imagem..."):
            imagem_url = buscar_imagem_carro(marca_input, modelo_input)
        if imagem_url:
            st.image(imagem_url, caption=f"{marca_input} {modelo_input}")
        else:
            st.warning("Imagem não encontrada para este modelo.")