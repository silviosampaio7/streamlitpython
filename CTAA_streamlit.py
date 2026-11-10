from datetime import date, timedelta
from io import BytesIO
import pandas as pd
import streamlit as st
# FUNÇÕES DE CÁLCULO
def numero_por_extenso(numero):
    unidades=["zero","um","dois","três","quatro","cinco","seis","sete","oito","nove"]
    especiais=["dez","onze","doze","treze","quatorze","quinze","dezesseis","dezessete","dezoito","dezenove" ]
    dezenas=["","","vinte","trinta","quarenta","cinquenta","sessenta","setenta","oitenta","noventa" ]
    centenas=["","cento","duzentos","trezentos","quatrocentos","quinhentos","seiscentos","setecentos",
              "oitocentos", "novecentos"    ]
    if not isinstance(numero, int) or not 0 <= numero <= 9999:
        raise ValueError("Informe um número inteiro entre 0 e 9999.")
    if numero < 10: return unidades[numero]
    if numero < 20: return especiais[numero - 10]
    if numero < 100:
        dezena, unidade = divmod(numero, 10)
        return dezenas[dezena] if unidade == 0 else (
            dezenas[dezena] + " e " + unidades[unidade] )
    if numero == 100:  return "cem"
    if numero < 1000:
        centena, resto = divmod(numero, 100)
        return centenas[centena] if resto == 0 else (
            centenas[centena] + " e " + numero_por_extenso(resto) )
    milhar, resto = divmod(numero, 1000)
    prefixo = "mil" if milhar == 1 else unidades[milhar] + " mil"
    if resto == 0:  return prefixo
    if resto < 100 or resto % 100 == 0:
        return prefixo + " e " + numero_por_extenso(resto)
    return prefixo + " " + numero_por_extenso(resto)
def primeira_segunda_marco(ano):
    data_inicio = date(ano, 3, 1)
    while data_inicio.weekday() != 0:
        data_inicio += timedelta(days=1)
    return data_inicio
def ultima_sexta_antes_natal(ano):
    data_fim = date(ano, 12, 24)
    while data_fim.weekday() != 4:
        data_fim -= timedelta(days=1)
    return data_fim
def calcular_ano(ano, faltas):
    inicio = primeira_segunda_marco(ano)
    fim = ultima_sexta_antes_natal(ano)

    # Contagem inclusiva: considera o dia inicial e o dia final.
    total_bruto = (fim - inicio).days + 1
    total_liquido = total_bruto - faltas

    periodo = (
        f"{inicio.strftime('%d/%m/%Y')} "
        f"a {fim.strftime('%d/%m/%Y')}"
    )

    return {
        "Ano": ano,
        "Faltas": faltas,
        "Início letivo": inicio.strftime("%d/%m/%Y"),
        "Fim": fim.strftime("%d/%m/%Y"),
        "Período": periodo,
        "Total dias bruto": total_bruto,
        "Tempo líquido": total_liquido,
    }
# GERAÇÃO DO EXCEL EM MEMÓRIA
def gerar_excel(tabela):
    arquivo = BytesIO()
    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        tabela.to_excel(writer, index=False, sheet_name="Cálculo CTAA")
        planilha = writer.sheets["Cálculo CTAA"]
        planilha.freeze_panes = "A2"
        planilha.auto_filter.ref = planilha.dimensions
        # Ajusta a largura das colunas conforme o conteúdo.
        for coluna in planilha.columns:
            maior = max(
                len(str(celula.value)) if celula.value is not None else 0
                for celula in coluna            )
            letra = coluna[0].column_letter
            planilha.column_dimensions[letra].width = min(maior + 3, 40)
    arquivo.seek(0)
    return arquivo.getvalue()
# ---------------------------------------------------------
# INTERFACE STREAMLIT
# ---------------------------------------------------------
st.set_page_config(
    page_title="Cálculo de CTAA",
    page_icon="📄",
    layout="wide"
)
st.title("📄 Cálculo de Tempo — CTAA")
st.caption("Certidão de Tempo de Aluno-Aprendiz")
st.divider()
st.subheader("1. Dados do cálculo")
col1, col2 = st.columns(2)
with col1:
    primeiro_ano = st.number_input(
        "Informe o primeiro ano",
        min_value=1900,
        max_value=2100,
        value=2026,
        step=1
    )
with col2:
    st.info(
        "Serão calculados automaticamente o ano informado "
        "e os dois anos anteriores."
    )
anos = [int(primeiro_ano) - i for i in range(3)]
st.subheader("2. Informe as faltas")
colunas = st.columns(3)
faltas_por_ano = {}
for coluna, ano in zip(colunas, anos):
    with coluna:
        faltas_por_ano[ano] = st.number_input(
            f"Faltas em {ano}",
            min_value=0,
            max_value=9999,
            value=0,
            step=1,
            key=f"faltas_{ano}"
        )
st.caption(
    "O período é contado em dias corridos, incluindo as datas inicial e final. "
    "As faltas informadas são subtraídas do total bruto.")
if st.button("Calcular CTAA", type="primary", use_container_width=True):
    resultados = []
    for ano in anos:
        faltas = int(faltas_por_ano[ano])
        resultado = calcular_ano(ano, faltas)
        if resultado["Tempo líquido"] < 0:
            st.error(
                f"As faltas informadas para {ano} são maiores que "
                "o total bruto do período."
            )
            st.stop()
        resultados.append(resultado)
    tabela = pd.DataFrame(resultados).sort_values(
        by="Ano"
    ).reset_index(drop=True)
    total_liquido = int(tabela["Tempo líquido"].sum())
    primeiro_inicio = tabela.iloc[0]["Início letivo"]
    ultimo_fim = tabela.iloc[-1]["Fim"]
    periodo_total = f"{primeiro_inicio} a {ultimo_fim}"
    try:
        total_extenso = numero_por_extenso(total_liquido)
    except ValueError as erro:
        st.error(str(erro))
        st.stop()
    st.divider()
    st.subheader("3. Resultado")
    st.dataframe(
        tabela,
        use_container_width=True,
        hide_index=True
    )
    metrica1, metrica2 = st.columns(2)
    with metrica1:
        st.metric("Tempo líquido total", f"{total_liquido:,}".replace(",", ".") + " dias")
    with metrica2:
        st.metric("Período total", periodo_total)
    st.success(
        f"**Total por extenso:** {total_extenso.capitalize()} dias."
    )
    excel = gerar_excel(tabela)
    st.download_button(
        label="📥 Baixar planilha Excel",
        data=excel,
        file_name="calculo_tempo_CTAA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.divider()
    st.info("Informe as faltas de cada ano e clique em **Calcular CTAA**.")