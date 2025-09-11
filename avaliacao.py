import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import unicodedata
import gspread
from gspread_dataframe import set_with_dataframe
import xlsxwriter

# --- CONFIGURAÇÃO INICIAL E CONEXÃO COM GOOGLE SHEETS ---

# Define a configuração da página
st.set_page_config(page_title="Avaliação Qualitativa", layout="wide")

# Conecta ao Google Sheets usando as credenciais armazenadas nos Secrets do Streamlit
try:
    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
    # Substitua 'Dados_Escola' pelo nome exato da sua Planilha Google
    spreadsheet = gc.open("Dados_Escola")
    ws_disciplinas = spreadsheet.worksheet("Disciplinas")
    ws_turmas = spreadsheet.worksheet("Turmas")
    ws_alunos = spreadsheet.worksheet("Alunos")
    ws_notas = spreadsheet.worksheet("Notas")
    st.session_state['db_connection'] = True
except Exception as e:
    st.error("Não foi possível conectar à base de dados (Google Sheets). Verifique as configurações de 'Secrets'.")
    st.error(f"Erro: {e}")
    st.session_state['db_connection'] = False
    st.stop()


# --- FUNÇÕES DE AUXÍLIO E DADOS ---

def remover_acentos(texto):
    """
    Remove acentos de uma string, normalizando-a para a ordenação.
    """
    nfkd_form = unicodedata.normalize('NFKD', str(texto))
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

@st.cache_data(ttl=60)
def load_base():
    """ Carrega as abas Disciplinas, Turmas e Alunos da Planilha Google. """
    try:
        disciplinas_df = pd.DataFrame(ws_disciplinas.get_all_records())
        turmas_df = pd.DataFrame(ws_turmas.get_all_records())
        alunos_df = pd.DataFrame(ws_alunos.get_all_records())
        
        disciplinas = disciplinas_df["Disciplina"].dropna().astype(str).tolist() if "Disciplina" in disciplinas_df else []
        turmas = turmas_df["Turma"].dropna().astype(str).tolist() if "Turma" in turmas_df else []
        
        return disciplinas, turmas, alunos_df
    except Exception as e:
        st.error(f"Erro ao carregar a base de dados: {e}")
        return [], [], pd.DataFrame(columns=["Turma", "Aluno"])


@st.cache_data(ttl=30)
def load_notas():
    """ Carrega todas as notas da Planilha Google. """
    try:
        notas_df = pd.DataFrame(ws_notas.get_all_records())
        if notas_df.empty:
            return pd.DataFrame(columns=["Trimestre", "Disciplina", "Turma", "Aluno", "Nota", "Timestamp"])
        
        # Garante que a coluna 'Nota' está no formato numérico correto para o cálculo.
        if 'Nota' in notas_df.columns:
            # Converte para string para garantir a substituição da vírgula
            notas_df['Nota'] = notas_df['Nota'].astype(str).str.replace(',', '.', regex=False)
            # Converte para numérico, tratando erros para valores inválidos
            notas_df['Nota'] = pd.to_numeric(notas_df['Nota'], errors='coerce')
            
        return notas_df
    except Exception as e:
        st.error(f"Erro ao carregar as notas: {e}")
        return pd.DataFrame(columns=["Trimestre", "Disciplina", "Turma", "Aluno", "Nota", "Timestamp"])


def save_notas(df_final):
    """ Salva o DataFrame completo de notas na Planilha Google. """
    try:
        ws_notas.clear()
        set_with_dataframe(ws_notas, df_final, include_column_header=True)
    except Exception as e:
        st.error(f"Erro ao salvar as notas: {e}")


# --- INTERFACE DO STREAMLIT ---

st.title("📊 Sistema de Avaliação Qualitativa")
st.write("Sistema centralizado para lançamento de notas qualitativas.")

# 1. Carregar os dados base e notas no estado da sessão apenas uma vez
if 'disciplinas' not in st.session_state:
    st.session_state['disciplinas'], st.session_state['turmas'], st.session_state['alunos_df'] = load_base()
if 'notas_df_geral' not in st.session_state:
    st.session_state['notas_df_geral'] = load_notas()

# Se a conexão falhou, o script para no bloco try/except lá em cima.
if not st.session_state['db_connection']:
    st.stop()


# --- BARRA LATERAL COM FILTROS ---
st.sidebar.header("Filtros de Seleção")
trimestres = ["1º Trimestre", "2º Trimestre", "3º Trimestre"]
trimestre = st.sidebar.selectbox("Selecione o trimestre", trimestres)
disciplina = st.sidebar.selectbox("Selecione a disciplina", st.session_state['disciplinas'])
turma = st.sidebar.selectbox("Selecione a turma", st.session_state['turmas'])


# --- LAYOUT PRINCIPAL ---
st.header(f"Lançamento de notas")
st.subheader(f"_{trimestre} — {disciplina} — {turma}_")
st.info("Clique na célula de nota para editar. Use Enter ou as setas do teclado para navegar.")

# 1. Preparar os dados para o editor
alunos_turma_df = st.session_state['alunos_df'][st.session_state['alunos_df']["Turma"] == turma][["Aluno"]].copy()
alunos_turma_df.sort_values(by="Aluno", key=lambda col: col.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8'), inplace=True)

# Filtra as notas apenas para a visualização atual
notas_atuais = st.session_state['notas_df_geral'][
    (st.session_state['notas_df_geral']["Trimestre"] == trimestre) &
    (st.session_state['notas_df_geral']["Disciplina"] == disciplina) &
    (st.session_state['notas_df_geral']["Turma"] == turma)
]

df_para_editar = pd.merge(alunos_turma_df, notas_atuais[["Aluno", "Nota"]], on="Aluno", how="left")
# CORREÇÃO: Converte a coluna 'Nota' para o tipo string antes de passar para o editor.
df_para_editar['Nota'] = df_para_editar['Nota'].astype(str)

# 2. Exibir o editor de dados (st.data_editor)
col_editor, _ = st.columns([2, 1])
with col_editor:
    with st.form(key='data_editor_form'):
        edited_df = st.data_editor(
            df_para_editar,
            column_config={
                "Aluno": st.column_config.TextColumn("Aluno", disabled=True),
                # O tipo é texto para evitar a conversão automática do Streamlit
                "Nota": st.column_config.TextColumn("Nota (0-10)"),
            },
            hide_index=True,
            key=f"editor_{trimestre}_{disciplina}_{turma}",
        )
        submitted = st.form_submit_button("Salvar lançamentos", type="primary")

# --- BOTÕES DE AÇÃO ---
st.markdown("---")
col_a, col_b, col_c, col_d = st.columns(4)

if submitted:
    # A lógica de processamento foi movida para aqui para garantir o formato correto
    notas_inseridas = edited_df.dropna(subset=['Nota'])
    rows_to_save = []
    for _, row in notas_inseridas.iterrows():
        try:
            # Limpa e converte a nota para o formato correto
            nota_limpa = str(row["Nota"]).replace(',', '.')
            nota_final = float(nota_limpa)

            # Adiciona uma validação extra para garantir que a nota está no intervalo
            if 0.0 <= nota_final <= 10.0:
                rows_to_save.append({
                    "Trimestre": trimestre, "Disciplina": disciplina, "Turma": turma,
                    "Aluno": row["Aluno"], "Nota": nota_final, "Timestamp": datetime.now().isoformat()
                })
            else:
                st.warning(f"A nota '{nota_final}' para o aluno '{row['Aluno']}' está fora do intervalo (0-10) e não foi salva.")

        except (ValueError, TypeError):
            st.warning(f"A nota inserida para o aluno '{row['Aluno']}' não é um número válido e não foi salva. Por favor, use um formato como '8.5' ou '8,5'.")

    if rows_to_save:
        new_df = pd.DataFrame(rows_to_save)
        
        # Lógica para atualizar a base de dados geral
        mask = ~(
            (st.session_state['notas_df_geral']["Trimestre"] == trimestre) &
            (st.session_state['notas_df_geral']["Disciplina"] == disciplina) &
            (st.session_state['notas_df_geral']["Turma"] == turma)
        )
        df_mantido = st.session_state['notas_df_geral'][mask]
        df_final = pd.concat([df_mantido, new_df], ignore_index=True)
        
        save_notas(df_final)
        st.session_state['notas_df_geral'] = df_final # Atualiza o estado da sessão com os novos dados
        st.success(f"{len(new_df)} lançamentos salvos/atualizados na base de dados central.")
        st.cache_data.clear() # Limpa o cache para recarregar os dados na próxima ação
        st.rerun()

with col_b:
    if st.button("Relatório (esta turma)", use_container_width=True):
        st.info("Calculando médias para a turma...")
        
        alunos_list = st.session_state['alunos_df'][st.session_state['alunos_df']['Turma'] == turma]['Aluno'].tolist()
        medias = []
        for aluno in sorted(alunos_list, key=remover_acentos):
            df_al = st.session_state['notas_df_geral'][(st.session_state['notas_df_geral']['Aluno'] == aluno) & (st.session_state['notas_df_geral']['Trimestre'] == trimestre)]
            if df_al.empty:
                medias.append({'Trimestre': trimestre, 'Turma': turma, 'Aluno': aluno, 'Média Qualitativa': None, 'Lançamentos': 0})
                continue
            
            latest = df_al.sort_values('Timestamp').groupby('Disciplina', as_index=False).last()
            notas = pd.to_numeric(latest['Nota'], errors='coerce').dropna().tolist()
            if notas:
                media = sum(notas) / len(notas)
                medias.append({'Trimestre': trimestre, 'Turma': turma, 'Aluno': aluno, 'Média Qualitativa': round(media, 1), 'Lançamentos': len(notas)})
            else:
                medias.append({'Trimestre': trimestre, 'Turma': turma, 'Aluno': aluno, 'Média Qualitativa': None, 'Lançamentos': 0})

        df_medias = pd.DataFrame(medias)
        st.subheader(f"Relatório da Turma: {turma}")
        st.dataframe(df_medias)
        
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            df_medias.to_excel(writer, index=False, sheet_name=f'Medias_{turma}')
        
        st.download_button("Baixar Relatório da Turma", data=output_excel.getvalue(), file_name=f"relatorio_{turma.replace(' ', '_')}_{trimestre.replace(' ', '_')}.xlsx")


with col_c:
    if st.button("Relatório (geral do tri)", use_container_width=True):
        st.info("Calculando médias para todas as turmas...")
        
        resultados = []
        for _, row in st.session_state['alunos_df'].iterrows():
            aluno = row['Aluno']
            turma_k = row['Turma']
            
            df_al = st.session_state['notas_df_geral'][(st.session_state['notas_df_geral']['Aluno'] == aluno) & (st.session_state['notas_df_geral']['Trimestre'] == trimestre)]
            if df_al.empty:
                resultados.append({'Trimestre': trimestre, 'Turma': turma_k, 'Aluno': aluno, 'Média Qualitativa': None, 'Lançamentos': 0})
                continue

            latest = df_al.sort_values('Timestamp').groupby('Disciplina', as_index=False).last()
            notas = pd.to_numeric(latest['Nota'], errors='coerce').dropna().tolist()
            if notas:
                media = sum(notas) / len(notas)
                resultados.append({'Trimestre': trimestre, 'Turma': turma_k, 'Aluno': aluno, 'Média Qualitativa': round(media,1), 'Lançamentos': len(notas)})
            else:
                resultados.append({'Trimestre': trimestre, 'Turma': turma_k, 'Aluno': aluno, 'Média Qualitativa': None, 'Lançamentos': 0})
        
        df_result = pd.DataFrame(resultados)
        
        if not df_result.empty:
            df_result['Aluno_sort'] = df_result['Aluno'].apply(remover_acentos)
            df_result = df_result.sort_values(by=["Turma", "Aluno_sort"]).drop(columns=['Aluno_sort'])

        st.subheader(f"Relatório Geral do {trimestre}")
        st.dataframe(df_result)
        
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            df_result.to_excel(writer, index=False, sheet_name=f'Medias_Gerais_{trimestre}')

        st.download_button("Baixar Relatório Geral", data=output_excel.getvalue(), file_name=f"relatorio_geral_{trimestre.replace(' ', '_')}.xlsx")

# --- LÓGICA DE EXCLUSÃO ---
with col_d:
    with st.expander("Excluir Notas"):
        password = st.text_input("Senha para apagar", type="password", key="pwd_delete")
        if password == "qualitativa":
            st.warning(f"Confirma a exclusão das notas de **{disciplina}** ({trimestre}) para a turma **{turma}**?")
            if st.button("Confirmar Exclusão", key="confirm_delete_btn"):
                mask_to_remove = (
                    (st.session_state['notas_df_geral']['Trimestre'] == trimestre) & 
                    (st.session_state['notas_df_geral']['Disciplina'] == disciplina) &
                    (st.session_state['notas_df_geral']['Turma'] == turma)
                )
                df_final = st.session_state['notas_df_geral'][~mask_to_remove]
                save_notas(df_final)
                st.session_state['notas_df_geral'] = df_final # Atualiza o estado da sessão
                st.success("Notas excluídas com sucesso.")
                st.cache_data.clear()
                st.rerun()

# --- INSTRUÇÕES NA BARRA LATERAL ---
st.sidebar.markdown("---")
st.sidebar.info("Como usar:")
st.sidebar.write("- Edite as abas da planilha `Dados_Escola` para gerenciar Disciplinas, Turmas e Alunos.")
st.sidebar.write("- Selecione o trimestre, disciplina e turma nos filtros.")
st.sidebar.write("- Lance as notas na tabela e clique em 'Salvar lançamentos'.")
st.sidebar.write("- Use a seção de relatórios para gerar médias por turma ou de forma geral.")
st.sidebar.write("- Para apagar notas, use a opção 'Excluir Notas'.")
