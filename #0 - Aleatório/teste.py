import time
import pandas as pd
import xlsxwriter

inicio = time.time()

input_file = "Tabela_Links.csv"
output_xlsx = "teste.xlsx"

# =====================================================================
#  DATA PIPELINE - leitura completa para Power BI
# =====================================================================
COLS = [
	"CNPJ_NEO", "cod_loja", "nom_fantasia", "Nome_loja",
	"estado", "Nome_Cidade", "Associacao",
	"Sit_Cliente", "Sit_contrato", "classificacao",
	"Des_modulo", "sistema", "Possui_dados",
	"Dias em atraso", "Matriz", "qtd_filiais",
]

df = pd.read_csv(
	input_file, sep=";", dtype=str, encoding="utf-8",
	usecols=lambda c: c in set(COLS),
)

# Coluna de farmacia unica
coluna_farmacia = next(
	(c for c in ["cod_loja", "CNPJ_NEO", "nom_fantasia", "Nome_loja"] if c in df.columns), None
)

# Limpeza basica
for col in ["Sit_Cliente", "Nome_Cidade", "estado", "Associacao",
            "Sit_contrato", "classificacao", "Des_modulo", "sistema", "Possui_dados"]:
	if col in df.columns:
		df[col] = df[col].fillna("").str.strip()
df["Sit_Cliente"] = df["Sit_Cliente"].str.upper()
if coluna_farmacia:
	df[coluna_farmacia] = df[coluna_farmacia].fillna("").str.strip()

# Deduplicar: 1 linha por loja
dedup_cols = [c for c in [coluna_farmacia, "Nome_Cidade", "estado"] if c and c in df.columns]
lojas = df.drop_duplicates(subset=dedup_cols) if dedup_cols else df.copy()

if "Dias em atraso" in lojas.columns:
	lojas = lojas.copy()
	lojas["atraso_num"] = pd.to_numeric(lojas["Dias em atraso"], errors="coerce").fillna(0).astype(int)

# --------------- METRICAS GERAIS ---------------
total_registros = len(lojas)
ativas = lojas[lojas["Sit_Cliente"] == "ATIVO"].copy()
total_ativas = len(ativas)
total_inativas = total_registros - total_ativas
taxa_ativa_pct = total_ativas / total_registros if total_registros else 0

cidades_unicas = lojas["Nome_Cidade"].replace("", pd.NA).dropna().nunique()
estados_unicos = lojas["estado"].replace("", pd.NA).dropna().nunique() if "estado" in lojas.columns else 0
assoc_unicas = lojas["Associacao"].replace("", pd.NA).dropna().nunique() if "Associacao" in lojas.columns else 0
media_por_cidade = total_ativas / cidades_unicas if cidades_unicas else 0

# --------------- AGRUPAMENTOS ---------------
def agrupar(frame, col, top_n=None):
	g = (frame[frame[col] != ""].groupby(col, as_index=False)
	     .size().rename(columns={"size": "Qtd"})
	     .sort_values("Qtd", ascending=False).reset_index(drop=True))
	return g.head(top_n).copy() if top_n else g

resumo_cidade = agrupar(ativas, "Nome_Cidade")
top20_cidade = resumo_cidade.head(20).copy()
top20_cidade["Rank"] = range(1, len(top20_cidade) + 1)
top20_cidade["Pct"] = top20_cidade["Qtd"] / total_ativas if total_ativas else 0
top20_cidade["Pct_Acum"] = top20_cidade["Pct"].cumsum()

resumo_estado = agrupar(ativas, "estado") if "estado" in ativas.columns else pd.DataFrame(columns=["estado", "Qtd"])
resumo_status = agrupar(lojas, "Sit_Cliente")
resumo_classif = agrupar(ativas, "classificacao") if "classificacao" in ativas.columns else pd.DataFrame(columns=["classificacao", "Qtd"])
top15_assoc = agrupar(ativas, "Associacao", 15) if "Associacao" in ativas.columns else pd.DataFrame(columns=["Associacao", "Qtd"])
resumo_contrato = agrupar(ativas, "Sit_contrato") if "Sit_contrato" in ativas.columns else pd.DataFrame(columns=["Sit_contrato", "Qtd"])
resumo_modulo = agrupar(ativas, "Des_modulo") if "Des_modulo" in ativas.columns else pd.DataFrame(columns=["Des_modulo", "Qtd"])
resumo_sistema = agrupar(ativas, "sistema") if "sistema" in ativas.columns else pd.DataFrame(columns=["sistema", "Qtd"])
resumo_dados = agrupar(ativas, "Possui_dados") if "Possui_dados" in ativas.columns else pd.DataFrame(columns=["Possui_dados", "Qtd"])

if "atraso_num" in ativas.columns:
	def faixa(d):
		if d <= 0: return "Em dia"
		if d <= 7: return "1-7 dias"
		if d <= 30: return "8-30 dias"
		if d <= 90: return "31-90 dias"
		return "90+ dias"
	at = ativas.copy()
	at["Faixa"] = at["atraso_num"].apply(faixa)
	resumo_atraso = at.groupby("Faixa", as_index=False).size().rename(columns={"size": "Qtd"})
	ordem = {"Em dia": 0, "1-7 dias": 1, "8-30 dias": 2, "31-90 dias": 3, "90+ dias": 4}
	resumo_atraso["_o"] = resumo_atraso["Faixa"].map(ordem)
	resumo_atraso = resumo_atraso.sort_values("_o").drop(columns="_o").reset_index(drop=True)
else:
	resumo_atraso = pd.DataFrame(columns=["Faixa", "Qtd"])

top10_qtd = int(resumo_cidade.head(10)["Qtd"].sum()) if len(resumo_cidade) else 0
top10_pct = top10_qtd / total_ativas if total_ativas else 0

# =====================================================================
#  EXCEL — POWER BI STYLE
# =====================================================================
# Paleta de cores Power BI
C = {
	"navy": "#1F4E78", "blue": "#2F75B5", "light_blue": "#5B9BD5",
	"sky": "#BDD7EE", "bg": "#F2F7FC", "white": "#FFFFFF",
	"red": "#C00000", "orange": "#ED7D31", "green": "#70AD47",
	"gold": "#FFC000", "gray": "#A5A5A5", "dark": "#44546A",
	"teal": "#00B0F0", "purple": "#7030A0",
}
CHART_COLORS = ["#2F75B5", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5",
                "#C00000", "#7030A0", "#00B0F0", "#A5A5A5", "#264478",
                "#9B57A0", "#636363", "#2CA02C", "#FF7F0E", "#1F77B4"]

wb = xlsxwriter.Workbook(output_xlsx, {"strings_to_numbers": False})

# ---- FORMATOS ----
fmt_title = wb.add_format({"bold": True, "font_size": 20, "font_color": C["white"], "bg_color": C["navy"], "align": "center", "valign": "vcenter"})
fmt_subtitle = wb.add_format({"bold": True, "font_size": 11, "font_color": C["navy"], "align": "left", "valign": "vcenter", "bottom": 2, "bottom_color": C["blue"]})
fmt_hdr = wb.add_format({"bold": True, "font_color": C["white"], "bg_color": C["blue"], "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
fmt_cell = wb.add_format({"border": 1, "font_size": 10})
fmt_num = wb.add_format({"border": 1, "align": "center", "num_format": "#,##0", "font_size": 10})
fmt_pct = wb.add_format({"border": 1, "align": "center", "num_format": "0.0%", "font_size": 10})
fmt_stripe = wb.add_format({"border": 1, "bg_color": C["bg"], "font_size": 10})
fmt_stripe_n = wb.add_format({"border": 1, "bg_color": C["bg"], "align": "center", "num_format": "#,##0", "font_size": 10})
fmt_stripe_p = wb.add_format({"border": 1, "bg_color": C["bg"], "align": "center", "num_format": "0.0%", "font_size": 10})

fmt_kpi_lbl = wb.add_format({"bold": True, "font_size": 9, "font_color": C["navy"], "align": "center", "valign": "vcenter", "bg_color": "#EEF4FB", "border": 1, "border_color": "#D9E2F3"})
fmt_kpi_val = wb.add_format({"bold": True, "font_size": 28, "font_color": "#0B3D6B", "align": "center", "valign": "vcenter", "bg_color": "#EEF4FB", "border": 1, "border_color": "#D9E2F3", "num_format": "#,##0"})
fmt_kpi_pct = wb.add_format({"bold": True, "font_size": 28, "font_color": "#0B3D6B", "align": "center", "valign": "vcenter", "bg_color": "#EEF4FB", "border": 1, "border_color": "#D9E2F3", "num_format": "0.0%"})
fmt_card_lbl = wb.add_format({"bold": True, "font_size": 9, "font_color": C["navy"], "align": "left", "valign": "vcenter", "bg_color": "#F7FAFF", "border": 1, "border_color": "#D9E2F3"})
fmt_card_val = wb.add_format({"bold": True, "font_size": 12, "font_color": "#0B3D6B", "align": "left", "valign": "vcenter", "bg_color": "#F7FAFF", "border": 1, "border_color": "#D9E2F3"})
fmt_nav = wb.add_format({"bold": True, "font_size": 9, "font_color": C["white"], "bg_color": C["dark"], "align": "center", "valign": "vcenter"})

# Helper: escreve mini-tabela e retorna nome da aba de dados + posicoes
DATA_SHEET = "_Dados"
ws_data = wb.add_worksheet(DATA_SHEET)
ws_data.hide()  # Hide the data sheet
data_row = [0]  # mutable counter

def write_data_block(names, values, col_name="Cat", val_name="Qtd"):
	"""Writes a block to the hidden data sheet. Returns (start_row, end_row) 1-indexed for chart refs."""
	r = data_row[0]
	ws_data.write(r, 0, col_name, fmt_hdr)
	ws_data.write(r, 1, val_name, fmt_hdr)
	for i, (n, v) in enumerate(zip(names, values)):
		ws_data.write(r + 1 + i, 0, str(n))
		ws_data.write(r + 1 + i, 1, int(v))
	start = r + 1
	end = r + len(names)
	data_row[0] = end + 2
	return start, end

def write_data_block_multi(headers, rows_data):
	"""Writes a multi-column block. Returns (start_row, end_row)."""
	r = data_row[0]
	for c, h in enumerate(headers):
		ws_data.write(r, c, h, fmt_hdr)
	for i, row in enumerate(rows_data):
		for c, val in enumerate(row):
			ws_data.write(r + 1 + i, c, val)
	start = r + 1
	end = r + len(rows_data)
	data_row[0] = end + 2
	return start, end

# Pre-write all data blocks
# Block: Top 20 cidades (multi-col: cidade, qtd, pct, pct_acum)
t20_rows = []
for _, rw in top20_cidade.iterrows():
	t20_rows.append([rw["Nome_Cidade"], int(rw["Qtd"]), float(rw["Pct"]), float(rw["Pct_Acum"])])
t20_s, t20_e = write_data_block_multi(["Cidade", "Qtd", "Pct", "Pct_Acum"], t20_rows)

# Block: Status
st_s, st_e = write_data_block(resumo_status["Sit_Cliente"].tolist(), resumo_status["Qtd"].tolist(), "Status", "Qtd")

# Block: Estado
if len(resumo_estado):
	es_s, es_e = write_data_block(resumo_estado["estado"].tolist(), resumo_estado["Qtd"].tolist(), "Estado", "Qtd")
else:
	es_s = es_e = 0

# Block: Classificacao
if len(resumo_classif):
	cl_s, cl_e = write_data_block(resumo_classif["classificacao"].tolist(), resumo_classif["Qtd"].tolist(), "Classificacao", "Qtd")
else:
	cl_s = cl_e = 0

# Block: Associacao top 15
if len(top15_assoc):
	as_s, as_e = write_data_block(top15_assoc["Associacao"].tolist(), top15_assoc["Qtd"].tolist(), "Associacao", "Qtd")
else:
	as_s = as_e = 0

# Block: Sit_contrato
if len(resumo_contrato):
	co_s, co_e = write_data_block(resumo_contrato["Sit_contrato"].tolist(), resumo_contrato["Qtd"].tolist(), "Contrato", "Qtd")
else:
	co_s = co_e = 0

# Block: Modulo
if len(resumo_modulo):
	mo_s, mo_e = write_data_block(resumo_modulo["Des_modulo"].tolist(), resumo_modulo["Qtd"].tolist(), "Modulo", "Qtd")
else:
	mo_s = mo_e = 0

# Block: Sistema
if len(resumo_sistema):
	si_s, si_e = write_data_block(resumo_sistema["sistema"].tolist(), resumo_sistema["Qtd"].tolist(), "Sistema", "Qtd")
else:
	si_s = si_e = 0

# Block: Possui_dados
if len(resumo_dados):
	pd_s, pd_e = write_data_block(resumo_dados["Possui_dados"].tolist(), resumo_dados["Qtd"].tolist(), "Possui Dados", "Qtd")
else:
	pd_s = pd_e = 0

# Block: Atraso
if len(resumo_atraso):
	at_s, at_e = write_data_block(resumo_atraso["Faixa"].tolist(), resumo_atraso["Qtd"].tolist(), "Faixa Atraso", "Qtd")
else:
	at_s = at_e = 0

# Block: Concentracao (Top 10 vs Demais)
conc_s, conc_e = write_data_block(["Top 10 cidades", "Demais"], [top10_qtd, max(total_ativas - top10_qtd, 0)], "Grupo", "Qtd")

# ---- HELPER: setup dashboard page ----
def setup_page(name, title_text, tab_color=C["navy"]):
	ws = wb.add_worksheet(name)
	ws.hide_gridlines(2)
	ws.set_tab_color(tab_color)
	ws.set_column("A:A", 2)
	ws.set_column("B:C", 16)
	ws.set_column("D:D", 2)
	ws.set_column("E:F", 16)
	ws.set_column("G:G", 2)
	ws.set_column("H:I", 16)
	ws.set_column("J:J", 2)
	ws.set_column("K:L", 16)
	ws.set_column("M:M", 2)
	ws.set_column("N:O", 16)
	ws.set_column("P:P", 2)
	ws.set_column("Q:R", 12)
	# Title bar
	ws.merge_range("B1:R1", title_text, fmt_title)
	ws.set_row(0, 38)
	return ws

def add_kpi(ws, row, col, label, value, fmt_v=fmt_kpi_val, width=2):
	"""Merge width columns for KPI label + value."""
	end_col = col + width - 1
	ws.merge_range(row, col, row, end_col, label, fmt_kpi_lbl)
	ws.merge_range(row + 1, col, row + 1, end_col, value, fmt_v)

def make_chart(chart_type, title, series_list, **kwargs):
	"""Build a chart with standard Power BI styling."""
	ch = wb.add_chart({"type": chart_type, "subtype": kwargs.get("subtype")})
	for s in series_list:
		ch.add_series(s)
	ch.set_title({"name": title, "name_font": {"bold": True, "size": 11, "color": C["navy"]}})
	if chart_type not in ("pie", "doughnut"):
		ch.set_plotarea({"fill": {"color": "#FAFCFF"}, "border": {"color": "#D9E2F3"}})
		x_cfg = {"num_font": {"size": 8, "color": C["dark"]}}
		y_cfg = {"num_font": {"size": 8, "color": C["dark"]}, "major_gridlines": {"visible": True, "line": {"color": "#E2E8F0", "dash_type": "dash"}}}
		if kwargs.get("x_name"):
			x_cfg["name"] = kwargs["x_name"]
			x_cfg["name_font"] = {"bold": True, "size": 9, "color": C["navy"]}
		if kwargs.get("y_name"):
			y_cfg["name"] = kwargs["y_name"]
			y_cfg["name_font"] = {"bold": True, "size": 9, "color": C["navy"]}
		if kwargs.get("reverse_y"):
			y_cfg["reverse"] = True
		ch.set_x_axis(x_cfg)
		ch.set_y_axis(y_cfg)
	else:
		ch.set_plotarea({"fill": {"none": True}})
	ch.set_chartarea({"fill": {"color": C["white"]}, "border": {"none": True}})
	ch.set_legend(kwargs.get("legend", {"position": "bottom"}))
	ch.set_size({"width": kwargs.get("w", 560), "height": kwargs.get("h", 340)})
	return ch

def colored_points(n):
	return [{"fill": {"color": CHART_COLORS[i % len(CHART_COLORS)]}} for i in range(n)]

# ==============================================================
#  PAGE 1 — OVERVIEW DASHBOARD
# ==============================================================
ws1 = setup_page("Overview", "POWER BI  |  Visao Geral das Farmacias")
ws1.set_row(2, 22)
add_kpi(ws1, 3, 1, "Lojas Cadastradas", total_registros)
add_kpi(ws1, 3, 4, "Lojas Ativas", total_ativas)
add_kpi(ws1, 3, 7, "Cidades", cidades_unicas)
add_kpi(ws1, 3, 10, "Estados", estados_unicos)
add_kpi(ws1, 3, 13, "Taxa Ativa", taxa_ativa_pct, fmt_kpi_pct)
ws1.set_row(4, 42)

# Combo chart: Top 20 cidades (colunas + % acumulado)
if len(top20_cidade):
	combo = wb.add_chart({"type": "column"})
	combo.add_series({
		"name": "Qtd farmacias",
		"categories": [DATA_SHEET, t20_s, 0, t20_e, 0],
		"values": [DATA_SHEET, t20_s, 1, t20_e, 1],
		"fill": {"color": C["blue"]}, "border": {"color": C["navy"]},
		"data_labels": {"value": True, "font": {"size": 7, "color": C["navy"]}},
	})
	line = wb.add_chart({"type": "line"})
	line.add_series({
		"name": "% acumulado",
		"categories": [DATA_SHEET, t20_s, 0, t20_e, 0],
		"values": [DATA_SHEET, t20_s, 3, t20_e, 3],
		"y2_axis": True,
		"line": {"color": C["red"], "width": 2},
		"marker": {"type": "circle", "size": 4, "border": {"color": C["red"]}, "fill": {"color": C["white"]}},
	})
	combo.combine(line)
	combo.set_title({"name": "Top 20 Cidades: Volume + Curva Acumulada (Pareto)", "name_font": {"bold": True, "size": 11, "color": C["navy"]}})
	combo.set_x_axis({"num_font": {"size": 7, "color": C["dark"]}})
	combo.set_y_axis({"name": "Quantidade", "name_font": {"bold": True, "size": 9, "color": C["navy"]}, "num_font": {"size": 8, "color": C["dark"]}, "major_gridlines": {"visible": True, "line": {"color": "#E2E8F0", "dash_type": "dash"}}, "min": 0})
	combo.set_y2_axis({"name": "% Acumulado", "name_font": {"bold": True, "size": 9, "color": C["navy"]}, "num_font": {"size": 8, "color": C["dark"]}, "num_format": "0%", "min": 0, "max": 1})
	combo.set_plotarea({"fill": {"color": "#FAFCFF"}, "border": {"color": "#D9E2F3"}})
	combo.set_chartarea({"fill": {"color": C["white"]}, "border": {"none": True}})
	combo.set_legend({"position": "top"})
	combo.set_size({"width": 750, "height": 380})
	ws1.insert_chart("B7", combo)

# Donut: Status (Ativo vs Inativo etc.)
if st_e > st_s or st_s == st_e:
	donut = wb.add_chart({"type": "doughnut"})
	donut.add_series({
		"name": "Status",
		"categories": [DATA_SHEET, st_s, 0, st_e, 0],
		"values": [DATA_SHEET, st_s, 1, st_e, 1],
		"points": colored_points(len(resumo_status)),
		"data_labels": {"percentage": True, "category": True, "separator": "\n", "font": {"size": 8, "bold": True}},
	})
	donut.set_hole_size(60)
	donut.set_title({"name": "Status dos Clientes", "name_font": {"bold": True, "size": 11, "color": C["navy"]}})
	donut.set_legend({"none": True})
	donut.set_chartarea({"fill": {"color": C["white"]}, "border": {"none": True}})
	donut.set_size({"width": 330, "height": 340})
	ws1.insert_chart("L7", donut)

# Donut: Concentracao Top 10
donut2 = wb.add_chart({"type": "doughnut"})
donut2.add_series({
	"name": "Concentracao",
	"categories": [DATA_SHEET, conc_s, 0, conc_e, 0],
	"values": [DATA_SHEET, conc_s, 1, conc_e, 1],
	"points": [{"fill": {"color": C["navy"]}}, {"fill": {"color": C["sky"]}}],
	"data_labels": {"percentage": True, "category": True, "separator": "\n", "font": {"size": 9, "bold": True}},
})
donut2.set_hole_size(62)
donut2.set_title({"name": "Concentracao Top 10 Cidades", "name_font": {"bold": True, "size": 11, "color": C["navy"]}})
donut2.set_legend({"none": True})
donut2.set_chartarea({"fill": {"color": C["white"]}, "border": {"none": True}})
donut2.set_size({"width": 330, "height": 340})
ws1.insert_chart("L27", donut2)

# Mini ranking table
ws1.merge_range("B28:C28", "Top 10 Cidades - Ranking", fmt_subtitle)
ws1.write("B29", "#", fmt_hdr)
ws1.write("C29", "Cidade", fmt_hdr)
ws1.write("D29", "Qtd", fmt_hdr)
ws1.write("E29", "% Total", fmt_hdr)
for i in range(min(10, len(top20_cidade))):
	r = 29 + i
	s = (i % 2 == 1)
	ws1.write(r, 1, i + 1, fmt_stripe_n if s else fmt_num)
	ws1.write(r, 2, top20_cidade.iloc[i]["Nome_Cidade"], fmt_stripe if s else fmt_cell)
	ws1.write(r, 3, int(top20_cidade.iloc[i]["Qtd"]), fmt_stripe_n if s else fmt_num)
	ws1.write(r, 4, float(top20_cidade.iloc[i]["Pct"]), fmt_stripe_p if s else fmt_pct)

# ==============================================================
#  PAGE 2 — GEOGRAPHIC ANALYSIS
# ==============================================================
ws2 = setup_page("Geografico", "POWER BI  |  Analise Geografica", C["blue"])
add_kpi(ws2, 3, 1, "Estados", estados_unicos)
add_kpi(ws2, 3, 4, "Cidades", cidades_unicas)
add_kpi(ws2, 3, 7, "Media/Cidade", round(media_por_cidade, 1), fmt_kpi_val)
add_kpi(ws2, 3, 10, "Top 10 Concentracao", top10_pct, fmt_kpi_pct)
ws2.set_row(4, 42)

# Bar chart horizontal: por estado
if es_e > 0:
	n_es = min(len(resumo_estado), 27)
	bar = make_chart("bar", "Farmacias Ativas por Estado (UF)", [{
		"name": "Ativas",
		"categories": [DATA_SHEET, es_s, 0, es_s + n_es - 1, 0],
		"values": [DATA_SHEET, es_s, 1, es_s + n_es - 1, 1],
		"fill": {"color": C["blue"]}, "border": {"color": C["navy"]},
		"data_labels": {"value": True, "position": "outside_end", "font": {"size": 8, "color": C["navy"], "bold": True}},
	}], x_name="Quantidade", reverse_y=True, legend={"none": True}, w=560, h=500)
	ws2.insert_chart("B7", bar)

# Pie: top 5 estados
if len(resumo_estado) >= 5:
	top5_es_s = es_s
	top5_es_e = es_s + 4
	pie_es = wb.add_chart({"type": "pie"})
	pie_es.add_series({
		"name": "Top 5 Estados",
		"categories": [DATA_SHEET, top5_es_s, 0, top5_es_e, 0],
		"values": [DATA_SHEET, top5_es_s, 1, top5_es_e, 1],
		"points": colored_points(5),
		"data_labels": {"percentage": True, "category": True, "separator": "\n", "font": {"size": 9, "bold": True}, "position": "outside_end"},
	})
	pie_es.set_title({"name": "Top 5 Estados - Participacao", "name_font": {"bold": True, "size": 11, "color": C["navy"]}})
	pie_es.set_legend({"none": True})
	pie_es.set_chartarea({"fill": {"color": C["white"]}, "border": {"none": True}})
	pie_es.set_size({"width": 420, "height": 340})
	ws2.insert_chart("K7", pie_es)

# Mini tabela estados
ws2.merge_range("K25:L25", "Ranking por Estado", fmt_subtitle)
ws2.write("K26", "UF", fmt_hdr)
ws2.write("L26", "Qtd", fmt_hdr)
for i in range(min(27, len(resumo_estado))):
	r = 26 + i
	s = (i % 2 == 1)
	ws2.write(r, 10, resumo_estado.iloc[i]["estado"], fmt_stripe if s else fmt_cell)
	ws2.write(r, 11, int(resumo_estado.iloc[i]["Qtd"]), fmt_stripe_n if s else fmt_num)

# ==============================================================
#  PAGE 3 — BUSINESS ANALYSIS
# ==============================================================
ws3 = setup_page("Negocios", "POWER BI  |  Analise de Negocios", C["green"])
add_kpi(ws3, 3, 1, "Associacoes", assoc_unicas)
add_kpi(ws3, 3, 4, "Classificacoes", len(resumo_classif))
add_kpi(ws3, 3, 7, "Tipos Contrato", len(resumo_contrato))
add_kpi(ws3, 3, 10, "Modulos", len(resumo_modulo))
ws3.set_row(4, 42)

# Pie: Classificacao
if cl_e > 0:
	pie_cl = wb.add_chart({"type": "pie"})
	pie_cl.add_series({
		"name": "Classificacao",
		"categories": [DATA_SHEET, cl_s, 0, cl_e, 0],
		"values": [DATA_SHEET, cl_s, 1, cl_e, 1],
		"points": colored_points(len(resumo_classif)),
		"data_labels": {"percentage": True, "category": True, "separator": "\n", "font": {"size": 9, "bold": True}},
	})
	pie_cl.set_title({"name": "Classificacao das Farmacias", "name_font": {"bold": True, "size": 11, "color": C["navy"]}})
	pie_cl.set_legend({"none": True})
	pie_cl.set_chartarea({"fill": {"color": C["white"]}, "border": {"none": True}})
	pie_cl.set_size({"width": 400, "height": 320})
	ws3.insert_chart("B7", pie_cl)

# Bar: Top 15 Associacoes
if as_e > 0:
	bar_as = make_chart("bar", "Top 15 Associacoes", [{
		"name": "Ativas",
		"categories": [DATA_SHEET, as_s, 0, as_e, 0],
		"values": [DATA_SHEET, as_s, 1, as_e, 1],
		"fill": {"color": C["orange"]}, "border": {"color": "#C05A1C"},
		"data_labels": {"value": True, "position": "outside_end", "font": {"size": 8, "color": C["dark"], "bold": True}},
	}], reverse_y=True, legend={"none": True}, w=580, h=420)
	ws3.insert_chart("H7", bar_as)

# Pie: Sit_contrato
if co_e > 0:
	pie_co = wb.add_chart({"type": "doughnut"})
	pie_co.add_series({
		"name": "Contrato",
		"categories": [DATA_SHEET, co_s, 0, co_e, 0],
		"values": [DATA_SHEET, co_s, 1, co_e, 1],
		"points": colored_points(len(resumo_contrato)),
		"data_labels": {"percentage": True, "category": True, "separator": "\n", "font": {"size": 9, "bold": True}},
	})
	pie_co.set_hole_size(55)
	pie_co.set_title({"name": "Situacao do Contrato", "name_font": {"bold": True, "size": 11, "color": C["navy"]}})
	pie_co.set_legend({"none": True})
	pie_co.set_chartarea({"fill": {"color": C["white"]}, "border": {"none": True}})
	pie_co.set_size({"width": 400, "height": 320})
	ws3.insert_chart("B27", pie_co)

# Column: Modulo
if mo_e > 0:
	n_mo = len(resumo_modulo)
	col_mo = make_chart("column", "Distribuicao por Modulo", [{
		"name": "Ativas",
		"categories": [DATA_SHEET, mo_s, 0, mo_e, 0],
		"values": [DATA_SHEET, mo_s, 1, mo_e, 1],
		"fill": {"color": C["teal"]}, "border": {"color": C["navy"]},
		"data_labels": {"value": True, "font": {"size": 8, "color": C["dark"]}},
	}], legend={"none": True}, w=580, h=340)
	ws3.insert_chart("H27", col_mo)

# ==============================================================
#  PAGE 4 — OPERATIONS & DATA QUALITY
# ==============================================================
ws4 = setup_page("Operacoes", "POWER BI  |  Operacoes e Qualidade de Dados", C["orange"])

# KPIs
em_dia = int(resumo_atraso[resumo_atraso["Faixa"] == "Em dia"]["Qtd"].sum()) if len(resumo_atraso) else 0
pct_em_dia = em_dia / total_ativas if total_ativas else 0
add_kpi(ws4, 3, 1, "Lojas Ativas", total_ativas)
add_kpi(ws4, 3, 4, "Em Dia", em_dia)
add_kpi(ws4, 3, 7, "% Em Dia", pct_em_dia, fmt_kpi_pct)
add_kpi(ws4, 3, 10, "Sistemas", len(resumo_sistema))
ws4.set_row(4, 42)

# Pie: Possui_dados
if pd_e > 0:
	pie_pd = wb.add_chart({"type": "pie"})
	pie_pd.add_series({
		"name": "Possui Dados",
		"categories": [DATA_SHEET, pd_s, 0, pd_e, 0],
		"values": [DATA_SHEET, pd_s, 1, pd_e, 1],
		"points": [{"fill": {"color": C["green"]}}, {"fill": {"color": C["red"]}}] if len(resumo_dados) == 2 else colored_points(len(resumo_dados)),
		"data_labels": {"percentage": True, "category": True, "separator": "\n", "font": {"size": 10, "bold": True}},
	})
	pie_pd.set_title({"name": "Possui Dados?", "name_font": {"bold": True, "size": 11, "color": C["navy"]}})
	pie_pd.set_legend({"none": True})
	pie_pd.set_chartarea({"fill": {"color": C["white"]}, "border": {"none": True}})
	pie_pd.set_size({"width": 380, "height": 310})
	ws4.insert_chart("B7", pie_pd)

# Column: Faixas de atraso
if at_e > 0:
	atraso_colors = [C["green"], C["teal"], C["gold"], C["orange"], C["red"]]
	pts_atraso = [{"fill": {"color": atraso_colors[i]}} for i in range(len(resumo_atraso))]
	col_at = wb.add_chart({"type": "column"})
	col_at.add_series({
		"name": "Farmacias",
		"categories": [DATA_SHEET, at_s, 0, at_e, 0],
		"values": [DATA_SHEET, at_s, 1, at_e, 1],
		"points": pts_atraso,
		"data_labels": {"value": True, "font": {"size": 9, "color": C["dark"], "bold": True}},
	})
	col_at.set_title({"name": "Distribuicao por Faixa de Atraso", "name_font": {"bold": True, "size": 11, "color": C["navy"]}})
	col_at.set_x_axis({"num_font": {"size": 9, "color": C["dark"]}})
	col_at.set_y_axis({"num_font": {"size": 9, "color": C["dark"]}, "major_gridlines": {"visible": True, "line": {"color": "#E2E8F0", "dash_type": "dash"}}})
	col_at.set_plotarea({"fill": {"color": "#FAFCFF"}, "border": {"color": "#D9E2F3"}})
	col_at.set_chartarea({"fill": {"color": C["white"]}, "border": {"none": True}})
	col_at.set_legend({"none": True})
	col_at.set_size({"width": 560, "height": 340})
	ws4.insert_chart("H7", col_at)

# Pie: Sistema
if si_e > 0:
	pie_si = wb.add_chart({"type": "pie"})
	pie_si.add_series({
		"name": "Sistema",
		"categories": [DATA_SHEET, si_s, 0, si_e, 0],
		"values": [DATA_SHEET, si_s, 1, si_e, 1],
		"points": colored_points(len(resumo_sistema)),
		"data_labels": {"percentage": True, "category": True, "separator": "\n", "font": {"size": 9, "bold": True}},
	})
	pie_si.set_title({"name": "Distribuicao por Sistema", "name_font": {"bold": True, "size": 11, "color": C["navy"]}})
	pie_si.set_legend({"none": True})
	pie_si.set_chartarea({"fill": {"color": C["white"]}, "border": {"none": True}})
	pie_si.set_size({"width": 380, "height": 310})
	ws4.insert_chart("B27", pie_si)

# Tabela atraso
ws4.merge_range("K27:L27", "Detalhamento Atraso", fmt_subtitle)
ws4.write("K28", "Faixa", fmt_hdr)
ws4.write("L28", "Qtd", fmt_hdr)
for i in range(len(resumo_atraso)):
	r = 28 + i
	s = (i % 2 == 1)
	ws4.write(r, 10, resumo_atraso.iloc[i]["Faixa"], fmt_stripe if s else fmt_cell)
	ws4.write(r, 11, int(resumo_atraso.iloc[i]["Qtd"]), fmt_stripe_n if s else fmt_num)

# ==============================================================
#  PAGE 5 — DATA TABLE (Resumo completo)
# ==============================================================
ws5 = wb.add_worksheet("Dados Detalhados")
ws5.set_tab_color(C["dark"])
ws5.freeze_panes(2, 0)
ws5.merge_range("A1:F1", "Tabela completa: Farmacias Ativas por Cidade", fmt_title)
ws5.set_row(0, 32)
headers = ["#", "Cidade", "Qtd", "% do Total", "% Acumulado"]
for c, h in enumerate(headers):
	ws5.write(1, c, h, fmt_hdr)

running_sum = 0
for i in range(len(resumo_cidade)):
	r = i + 2
	rw = resumo_cidade.iloc[i]
	qtd = int(rw["Qtd"])
	running_sum += qtd
	pct = qtd / total_ativas if total_ativas else 0
	pct_acum = running_sum / total_ativas if total_ativas else 0
	s = (i % 2 == 1)
	ws5.write(r, 0, i + 1, fmt_stripe_n if s else fmt_num)
	ws5.write(r, 1, rw["Nome_Cidade"], fmt_stripe if s else fmt_cell)
	ws5.write(r, 2, qtd, fmt_stripe_n if s else fmt_num)
	ws5.write(r, 3, pct, fmt_stripe_p if s else fmt_pct)
	ws5.write(r, 4, pct_acum, fmt_stripe_p if s else fmt_pct)

ws5.set_column("A:A", 6)
ws5.set_column("B:B", 34)
ws5.set_column("C:C", 14)
ws5.set_column("D:E", 14)

# Data bar on Qtd column
if len(resumo_cidade):
	ws5.conditional_format(2, 2, len(resumo_cidade) + 1, 2, {"type": "data_bar", "bar_color": C["light_blue"]})

# Activate Overview as the first page
ws1.activate()

wb.close()

fim = time.time()
print(f"Arquivo gerado: {output_xlsx}")
print(f"Lojas: {total_registros} | Ativas: {total_ativas} | Cidades: {cidades_unicas} | Estados: {estados_unicos}")
print(f"Tempo: {fim - inicio:.2f}s")