import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# Configuração da página do Streamlit
st.set_page_config(page_title="Dashboard de Vendas", page_icon=":bar_chart:", layout="wide")

# Função para ler os dados do Excel
@st.cache_data
def get_data_from_excel():
    # Usando caminho relativo
    df = pd.read_excel(
        io="Base de Dados.xlsx",  # ou "pasta/Base de Dados.xlsx" se estiver em uma subpasta
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
st.sidebar.image("C:\\Users\\Alan Milan\\Documents\\testemil\\logo.png", width=250)  # Ajuste o caminho e a largura conforme necessário

# Sidebar para filtros
st.sidebar.header("Por favor, selecione os filtros:")
unidade = st.sidebar.multiselect(
    "Selecione a Unidade:",
    options=df["Unidade"].unique(),
    default=df["Unidade"].unique()
)

operador = st.sidebar.multiselect(
    "Selecione o Operador:",
    options=df["Operador"].unique(),
    default=df["Operador"].unique(),
)

mes = st.sidebar.multiselect(
    "Selecione o Mês:",
    options=df["Mês"].unique(),
    default=df["Mês"].unique()
)

# Aplicar filtros usando df.loc[]
df_selection = df.loc[
    (df["Unidade"].isin(unidade)) &
    (df["Operador"].isin(operador)) &
    (df["Mês"].isin(mes))
]

# Verifica se o dataframe está vazio
if df_selection.empty:
    st.warning("Nenhum dado disponível com base nas configurações de filtro atuais!")
    st.stop()

# Inicializa o estado do botão para mostrar tabela
if 'show_table' not in st.session_state:
    st.session_state.show_table = False

# Função para alternar a exibição da tabela
def toggle_table():
    st.session_state.show_table = not st.session_state.show_table

# Título da página
st.title(":bar_chart: Dashboard de Vendas")
st.markdown("##")

# KPIs principais
total_sales = int(df_selection["Vendas Realizadas"].sum())
average_satisfaction = round(df_selection["Pesquisa de Satisfação"].mean(), 1)
average_customer_service = round(df_selection["Atendimentos no Dia"].mean(), 2)

left_column, middle_column, right_column = st.columns(3)

# Gráfico de visão geral para total de vendas
with left_column:
    st.subheader("Visão Geral das Vendas")
    fig_total_sales = go.Figure()
    fig_total_sales.add_trace(go.Indicator(
        mode="number",
        value=total_sales,
        title={"text": "Total de Vendas"},
        number={"prefix": "R$ "},
        domain={"x": [0.1, 0.9], "y": [0.2, 0.8]}
    ))
    fig_total_sales.update_layout(
        height=200,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial, sans-serif", size=12, color="white"),
    )
    st.plotly_chart(fig_total_sales, use_container_width=True)

# Gráfico de medidor para satisfação média
with middle_column:
    st.subheader("Satisfação Média")
    fig_satisfaction = go.Figure()
    fig_satisfaction.add_trace(go.Indicator(
        mode="gauge+number",
        value=average_satisfaction,
        title={"text": "Satisfação Média"},
        gauge={
            "axis": {"range": [0, 10]},
            "bar": {"color": "orange"},
            "steps": [
                {"range": [0, 4], "color": "red"},
                {"range": [4, 7], "color": "yellow"},
                {"range": [7, 10], "color": "green"}
            ]
        },
        domain={"x": [0.1, 0.9], "y": [0.2, 0.8]}
    ))
    fig_satisfaction.update_layout(
        height=200,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial, sans-serif", size=12, color="white"),
    )
    st.plotly_chart(fig_satisfaction, use_container_width=True)

# Gráfico de manômetro para atendimento ao cliente
with right_column:
    st.subheader("Atendimento ao Cliente (Média)")
    fig_customer_service = go.Figure()
    fig_customer_service.add_trace(go.Indicator(
        mode="gauge+number",
        value=average_customer_service,
        title={"text": "Média de Atendimento"},
        gauge={
            "axis": {"range": [0, df_selection["Atendimentos no Dia"].max()]},
            "bar": {"color": "orange"},
            "steps": [
                {"range": [0, 10], "color": "red"},
                {"range": [10, 20], "color": "yellow"},
                {"range": [20, df_selection["Atendimentos no Dia"].max()], "color": "green"}
            ]
        },
        domain={"x": [0.1, 0.9], "y": [0.2, 0.8]}
    ))
    fig_customer_service.update_layout(
        height=200,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial, sans-serif", size=12, color="white"),
    )
    st.plotly_chart(fig_customer_service, use_container_width=True)

st.markdown("---")

# Inicializa a variável 'fig'
fig = go.Figure()

# Seletor para o tipo de gráfico
chart_type = st.selectbox(
    "Escolha o gráfico a ser exibido:",
    options=[
        "Vendas por Unidade",
        "Vendas por Mês",
        "Leads Recebidos por Mês",
        "Atendimentos por Mês",
        "Ações Realizadas por Mês",
        "Ações Planejadas por Mês",
        "Resgate de clientes",
        "Pesquisa de satisfação",
        "Insucessos",
        "Tempo Médio de atendimento",
    ]
)

# Gerar e exibir o gráfico baseado na escolha
if chart_type == "Vendas por Unidade":
    sales_by_unit = df_selection.groupby(by=["Unidade"])[["Vendas Realizadas"]].sum().sort_values(by="Vendas Realizadas")
    fig = go.Figure(go.Bar(
        x=sales_by_unit["Vendas Realizadas"],
        y=sales_by_unit.index,
        orientation="h",
        marker_color="#FF6600"
    ))
    fig.update_layout(
        title="<b>Vendas por Unidade</b>",
        height=300,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Vendas Realizadas",
            title_font_size=14,
            showgrid=False
        ),
        yaxis=dict(
            title="Unidade",
            title_font_size=14,
            showgrid=False
        ),
        font=dict(family="Arial, sans-serif", size=12, color="white")
    )
elif chart_type == "Vendas por Mês":
    sales_by_month = df_selection.groupby(by=["Mês"])[["Vendas Realizadas"]].sum()
    fig = go.Figure(go.Bar(
        x=sales_by_month.index,
        y=sales_by_month["Vendas Realizadas"],
        marker_color="#FF6600"
    ))
    fig.update_layout(
        title="<b>Vendas por Mês</b>",
        height=300,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Mês",
            title_font_size=14,
            showgrid=False
        ),
        yaxis=dict(
            title="Vendas Realizadas",
            title_font_size=14,
            showgrid=False
        ),
        font=dict(family="Arial, sans-serif", size=12, color="white")
    )
elif chart_type == "Leads Recebidos por Mês":
    leads_by_month = df_selection.groupby(by=["Mês"])[["Leads Recebidos"]].sum()
    fig = go.Figure(go.Bar(
        x=leads_by_month.index,
        y=leads_by_month["Leads Recebidos"],
        marker_color="#FF6600"
    ))
    fig.update_layout(
        title="<b>Leads Recebidos por Mês</b>",
        height=300,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Mês",
            title_font_size=14,
            showgrid=False
        ),
        yaxis=dict(
            title="Leads Recebidos",
            title_font_size=14,
            showgrid=False
        ),
        font=dict(family="Arial, sans-serif", size=12, color="white")
    )
elif chart_type == "Atendimentos por Mês":
    atendimentos_by_month = df_selection.groupby(by=["Mês"])[["Atendimentos no Dia"]].sum()
    fig = go.Figure(go.Bar(
        x=atendimentos_by_month.index,
        y=atendimentos_by_month["Atendimentos no Dia"],
        marker_color="#FF6600"
    ))
    fig.update_layout(
        title="<b>Atendimentos por Mês</b>",
        height=300,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Mês",
            title_font_size=14,
            showgrid=False
        ),
        yaxis=dict(
            title="Atendimentos no Dia",
            title_font_size=14,
            showgrid=False
        ),
        font=dict(family="Arial, sans-serif", size=12, color="white")
    )
elif chart_type == "Ações Realizadas por Mês":
    acoes_by_month = df_selection.groupby(by=["Mês"])[["Ações Realizadas"]].sum()
    fig = go.Figure(go.Bar(
        x=acoes_by_month.index,
        y=acoes_by_month["Ações Realizadas"],
        marker_color="#FF6600"
    ))
    fig.update_layout(
        title="<b>Ações Realizadas por Mês</b>",
        height=300,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Mês",
            title_font_size=14,
            showgrid=False
        ),
        yaxis=dict(
            title="Ações Realizadas",
            title_font_size=14,
            showgrid=False
        ),
        font=dict(family="Arial, sans-serif", size=12, color="white")
    )
elif chart_type == "Ações Planejadas por Mês":
    acoes_planejadas_by_month = df_selection.groupby(by=["Mês"])[["Ações Planejadas"]].sum()
    fig = go.Figure(go.Bar(
        x=acoes_planejadas_by_month.index,
        y=acoes_planejadas_by_month["Ações Planejadas"],
        marker_color="#FF6600"
    ))
    fig.update_layout(
        title="<b>Ações Planejadas por Mês</b>",
        height=300,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Mês",
            title_font_size=14,
            showgrid=False
        ),
        yaxis=dict(
            title="Ações Planejadas",
            title_font_size=14,
            showgrid=False
        ),
        font=dict(family="Arial, sans-serif", size=12, color="white")
    )
elif chart_type == "Resgate de clientes":
    resgates_by_month = df_selection.groupby(by=["Mês"])[["Resgates de Clientes"]].sum()
    fig = go.Figure(go.Bar(
        x=resgates_by_month.index,
        y=resgates_by_month["Resgates de Clientes"],
        marker_color="#FF6600"
    ))
    fig.update_layout(
        title="<b>Resgate de Clientes por Mês</b>",
        height=300,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Mês",
            title_font_size=14,
            showgrid=False
        ),
        yaxis=dict(
            title="Resgates de Clientes",
            title_font_size=14,
            showgrid=False
        ),
        font=dict(family="Arial, sans-serif", size=12, color="white")
    )
elif chart_type == "Pesquisa de satisfação":
    satisfaction_by_month = df_selection.groupby(by=["Mês"])[["Pesquisa de Satisfação"]].mean()
    fig = go.Figure(go.Scatter(
        x=satisfaction_by_month.index,
        y=satisfaction_by_month["Pesquisa de Satisfação"],
        mode='lines+markers',
        marker_color="#FF6600"
    ))
    fig.update_layout(
        title="<b>Pesquisa de Satisfação por Mês</b>",
        height=300,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Mês",
            title_font_size=14,
            showgrid=False
        ),
        yaxis=dict(
            title="Satisfação",
            title_font_size=14,
            showgrid=False
        ),
        font=dict(family="Arial, sans-serif", size=12, color="white")
    )
elif chart_type == "Insucessos":
    insucessos_by_month = df_selection.groupby(by=["Mês"])[["Insucessos"]].sum()
    fig = go.Figure(go.Bar(
        x=insucessos_by_month.index,
        y=insucessos_by_month["Insucessos"],
        marker_color="#FF6600"
    ))
    fig.update_layout(
        title="<b>Insucessos por Mês</b>",
        height=300,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Mês",
            title_font_size=14,
            showgrid=False
        ),
        yaxis=dict(
            title="Insucessos",
            title_font_size=14,
            showgrid=False
        ),
        font=dict(family="Arial, sans-serif", size=12, color="white")
    )
elif chart_type == "Tempo Médio de atendimento":
    tempo_medio_by_month = df_selection.groupby(by=["Mês"])[["Tempo Médio de Atendimento (min)"]].mean()
    fig = go.Figure(go.Scatter(
        x=tempo_medio_by_month.index,
        y=tempo_medio_by_month["Tempo Médio de Atendimento (min)"],
        mode='lines+markers',
        marker_color="#FF6600"
    ))
    fig.update_layout(
        title="<b>Tempo Médio de Atendimento por Mês</b>",
        height=300,
        margin=dict(l=20, r=20, t=20, b=20),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            title="Mês",
            title_font_size=14,
            showgrid=False
        ),
        yaxis=dict(
            title="Tempo Médio de Atendimento",
            title_font_size=14,
            showgrid=False
        ),
        font=dict(family="Arial, sans-serif", size=12, color="white")
    )

# Exibir gráfico
st.plotly_chart(fig, use_container_width=True)

# Botão para mostrar/ocultar tabela
if st.button("Mostrar Tabela Excel"):
    toggle_table()

if st.session_state.show_table:
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
