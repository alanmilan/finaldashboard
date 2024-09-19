import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# Configuração da página do Streamlit
st.set_page_config(page_title="Dashboard de Vendas", page_icon=":bar_chart:", layout="wide")

# Função para ler os dados do Excel
@st.cache_data
def get_data_from_excel():
    df = pd.read_excel(
        io="Base de Dados.xlsx",  # Atualize para o caminho correto no seu ambiente
        engine="openpyxl",
        sheet_name="Sheet1"
    )
    
    # Filtra as colunas necessárias
    df = df[[ 
        "Unidade", "Operador", "Mês", "Leads Recebidos", "Atendimentos no Dia",
        "Ações Realizadas", "Ações Planejadas", "Resgates de Clientes",
        "Pesquisa de Satisfação", "Meta de Vendas", "Vendas Realizadas",
        "Insucessos", "Tempo Médio de Atendimento (min)"
    ]]
    
    return df

df = get_data_from_excel()

# Adicionar o logo da empresa na barra lateral
st.sidebar.image("logo.png", width=250)  # Atualize para o caminho correto no seu ambiente

# Sidebar para filtros
st.sidebar.header("Por favor, selecione os filtros:")
unidade = st.sidebar.multiselect("Selecione a Unidade:", options=df["Unidade"].unique(), default=df["Unidade"].unique())
operador = st.sidebar.multiselect("Selecione o Operador:", options=df["Operador"].unique(), default=df["Operador"].unique())
mes = st.sidebar.multiselect("Selecione o Mês:", options=df["Mês"].unique(), default=df["Mês"].unique())

# Aplicar filtros
df_selection = df.loc[
    (df["Unidade"].isin(unidade)) &
    (df["Operador"].isin(operador)) &
    (df["Mês"].isin(mes))
]

# Verifica se o dataframe está vazio
if df_selection.empty:
    st.warning("Nenhum dado disponível com base nas configurações de filtro atuais!")
    st.stop()

# Título da página
st.title(":bar_chart: Dashboard de Vendas")
st.markdown("##")

# KPIs principais
total_sales = int(df_selection["Vendas Realizadas"].sum())
average_satisfaction = round(df_selection["Pesquisa de Satisfação"].mean(), 1)
average_customer_service = round(df_selection["Atendimentos no Dia"].mean(), 2)

# Exibir os KPIs em colunas
left_column, middle_column, right_column = st.columns(3)

with left_column:
    st.subheader("Total de Vendas")
    st.markdown(f"<h1 style='text-align: center; color: black;'>{total_sales}</h1>", unsafe_allow_html=True)

with middle_column:
    st.subheader("Satisfação Média")
    st.markdown(f"<h1 style='text-align: center; color: black;'>{average_satisfaction}</h1>", unsafe_allow_html=True)

with right_column:
    st.subheader("Média de Atendimento")
    st.markdown(f"<h1 style='text-align: center; color: black;'>{average_customer_service}</h1>", unsafe_allow_html=True)

st.markdown("---")

# Seletor para o tipo de gráfico
chart_type = st.selectbox("Escolha o gráfico a ser exibido:", options=[
    "Vendas por Unidade", "Vendas por Mês", "Leads Recebidos por Mês",
    "Atendimentos por Mês", "Ações Realizadas por Mês", "Ações Planejadas por Mês",
    "Resgate de clientes", "Pesquisa de satisfação", "Insucessos",
    "Tempo Médio de atendimento",
])

# Gerar e exibir o gráfico baseado na escolha
def generate_chart(chart_type):
    fig = go.Figure()
    
    if chart_type == "Vendas por Unidade":
        sales_by_unit = df_selection.groupby("Unidade")["Vendas Realizadas"].sum()
        fig.add_trace(go.Bar(x=sales_by_unit.values, y=sales_by_unit.index, orientation='h', marker_color="#FF6600"))
        fig.update_layout(title="<b>Vendas por Unidade</b>", xaxis_title="Vendas Realizadas", yaxis_title="Unidade", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="black"))
    
    elif chart_type == "Vendas por Mês":
        sales_by_month = df_selection.groupby("Mês")["Vendas Realizadas"].sum()
        fig.add_trace(go.Bar(x=sales_by_month.values, y=sales_by_month.index, orientation='h', marker_color="#FF6600"))
        fig.update_layout(title="<b>Vendas por Mês</b>", xaxis_title="Vendas Realizadas", yaxis_title="Mês", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="black"))
    
    elif chart_type == "Leads Recebidos por Mês":
        leads_by_month = df_selection.groupby("Mês")["Leads Recebidos"].sum()
        fig.add_trace(go.Bar(x=leads_by_month.values, y=leads_by_month.index, orientation='h', marker_color="#FF6600"))
        fig.update_layout(title="<b>Leads Recebidos por Mês</b>", xaxis_title="Leads Recebidos", yaxis_title="Mês", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="black"))

    elif chart_type == "Atendimentos por Mês":
        atendimentos_by_month = df_selection.groupby("Mês")["Atendimentos no Dia"].sum()
        fig.add_trace(go.Bar(x=atendimentos_by_month.values, y=atendimentos_by_month.index, orientation='h', marker_color="#FF6600"))
        fig.update_layout(title="<b>Atendimentos por Mês</b>", xaxis_title="Atendimentos", yaxis_title="Mês", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="black"))
    
    elif chart_type == "Ações Realizadas por Mês":
        acoes_by_month = df_selection.groupby("Mês")["Ações Realizadas"].sum()
        fig.add_trace(go.Bar(x=acoes_by_month.values, y=acoes_by_month.index, orientation='h', marker_color="#FF6600"))
        fig.update_layout(title="<b>Ações Realizadas por Mês</b>", xaxis_title="Ações Realizadas", yaxis_title="Mês", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="black"))

    elif chart_type == "Ações Planejadas por Mês":
        acoes_planejadas_by_month = df_selection.groupby("Mês")["Ações Planejadas"].sum()
        fig.add_trace(go.Bar(x=acoes_planejadas_by_month.values, y=acoes_planejadas_by_month.index, orientation='h', marker_color="#FF6600"))
        fig.update_layout(title="<b>Ações Planejadas por Mês</b>", xaxis_title="Ações Planejadas", yaxis_title="Mês", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="black"))

    elif chart_type == "Resgate de clientes":
        resgates_by_month = df_selection.groupby("Mês")["Resgates de Clientes"].sum()
        fig.add_trace(go.Bar(x=resgates_by_month.values, y=resgates_by_month.index, orientation='h', marker_color="#FF6600"))
        fig.update_layout(title="<b>Resgates de Clientes</b>", xaxis_title="Resgates", yaxis_title="Mês", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="black"))

    elif chart_type == "Pesquisa de satisfação":
        satisfaction_by_month = df_selection.groupby("Mês")["Pesquisa de Satisfação"].mean()
        fig.add_trace(go.Bar(x=satisfaction_by_month.values, y=satisfaction_by_month.index, orientation='h', marker_color="#FF6600"))
        fig.update_layout(title="<b>Pesquisa de Satisfação</b>", xaxis_title="Satisfação Média", yaxis_title="Mês", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="black"))

    elif chart_type == "Insucessos":
        insucessos_by_month = df_selection.groupby("Mês")["Insucessos"].sum()
        fig.add_trace(go.Bar(x=insucessos_by_month.values, y=insucessos_by_month.index, orientation='h', marker_color="#FF6600"))
        fig.update_layout(title="<b>Insucessos</b>", xaxis_title="Total de Insucessos", yaxis_title="Mês", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="black"))

    elif chart_type == "Tempo Médio de atendimento":
        tempo_medio_by_month = df_selection.groupby("Mês")["Tempo Médio de Atendimento (min)"].mean()
        fig.add_trace(go.Bar(x=tempo_medio_by_month.values, y=tempo_medio_by_month.index, orientation='h', marker_color="#FF6600"))
        fig.update_layout(title="<b>Tempo Médio de Atendimento</b>", xaxis_title="Tempo (min)", yaxis_title="Mês", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="black"))

    return fig

# Exibir gráfico
st.plotly_chart(generate_chart(chart_type), use_container_width=True)

# Botão para mostrar/ocultar tabela
if st.button("Mostrar Tabela Excel"):
    st.session_state.show_table = not st.session_state.show_table

if st.session_state.get("show_table", False):
    st.write(df_selection)

# Estilo para esconder a interface padrão do Streamlit
hide_st_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)
