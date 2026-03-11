import streamlit as st

st.title("Meu primeiro app")

form = st.form("meu_form", clear_on_submit=True)

nome = form.text_input("Digite seu nome")
sla = form.form_submit_button("Enviar")
form.write(f"Valor do nome: {nome}")
