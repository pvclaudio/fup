import streamlit as st
import pandas as pd
from datetime import datetime, date
import yagmail
from io import BytesIO
from pathlib import Path
import plotly.express as px
import os
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import zipfile
import tempfile
import json
from oauth2client.client import OAuth2Credentials
import httplib2
import traceback
import openai
import json
import httpx
from sentence_transformers import SentenceTransformer, util
from openai import OpenAI
import json
import requests
import tempfile
from difflib import get_close_matches
import re
from datetime import timedelta
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt
from pandas import Timestamp

st.set_page_config(layout = 'wide')

st.write("Hoje:", pd.Timestamp.today())

#st.sidebar.text(f"Diretório atual: {os.getcwd()}")

caminho_csv = "followups.csv"
admin_users = ["cvieira", "amendonca", "mathayde"]
cadastro_users = ["cvieira", "amendonca", "mathayde"]
chat_users = ["cvieira", "amendonca", "mathayde","bromanelli","ysouza"]

hoje = Timestamp.today().normalize()

def enviar_email_gmail(destinatario, assunto, corpo_html):
    try:
        yag = yagmail.SMTP(user=st.secrets["email_user"], password=st.secrets["email_pass"])
        yag.send(to=destinatario, subject=assunto, contents=corpo_html)
        return True
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {e}")
        return False
    
def conectar_drive():
    cred_dict = st.secrets["credentials"]

    credentials = OAuth2Credentials(
        access_token=cred_dict["access_token"],
        client_id=cred_dict["client_id"],
        client_secret=cred_dict["client_secret"],
        refresh_token=cred_dict["refresh_token"],
        token_expiry=datetime.strptime(cred_dict["token_expiry"], "%Y-%m-%dT%H:%M:%SZ"),
        token_uri=cred_dict["token_uri"],
        user_agent="streamlit-app/1.0",
        revoke_uri=cred_dict["revoke_uri"]
    )

    # Atualiza token se expirado
    if credentials.access_token_expired:
        credentials.refresh(httplib2.Http())

    gauth = GoogleAuth()
    gauth.credentials = credentials
    drive = GoogleDrive(gauth)
    return drive

def upload_para_drive():
    try:
        drive = conectar_drive()

        # Verifica ou cria pasta principal 'FUP'
        pastas_fup = drive.ListFile({
            'q': "title = 'FUP' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        }).GetList()

        if pastas_fup:
            pasta_fup = pastas_fup[0]
        else:
            pasta_fup = drive.CreateFile({
                'title': 'FUP',
                'mimeType': 'application/vnd.google-apps.folder'
            })
            pasta_fup.Upload()

        # Verifica ou cria subpasta 'backup' dentro da pasta 'FUP'
        backups = drive.ListFile({
            'q': f"'{pasta_fup['id']}' in parents and title = 'backup' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        }).GetList()

        if backups:
            pasta_backup = backups[0]
        else:
            pasta_backup = drive.CreateFile({
                'title': 'backup',
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [{'id': pasta_fup['id']}]
            })
            pasta_backup.Upload()

        # Cria (ou substitui) o arquivo followups.csv dentro de 'FUP'
        arquivos_existentes = drive.ListFile({
            'q': f"'{pasta_fup['id']}' in parents and title = 'followups.csv' and trashed = false"
        }).GetList()

        if arquivos_existentes:
            arquivo = arquivos_existentes[0]
        else:
            arquivo = drive.CreateFile({
                'title': 'followups.csv',
                'parents': [{'id': pasta_fup['id']}]
            })

        arquivo.SetContentFile(caminho_csv)
        arquivo.Upload()

        # Cria cópia com timestamp em /FUP/backup
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = drive.CreateFile({
            'title': f'followups_backup_{timestamp}.csv',
            'parents': [{'id': pasta_backup['id']}]
        })
        backup_file.SetContentFile(caminho_csv)
        backup_file.Upload()

        st.info("📤 Arquivo 'followups.csv' atualizado na pasta FUP e backup gerado com sucesso.")

    except Exception as e:
        st.warning(f"Erro ao enviar para o Drive: {e}")

def upload_evidencias_para_drive(idx, arquivos, observacao):
    try:
        drive = conectar_drive()

        # Garante que a pasta "FUP" exista
        pastas_fup = drive.ListFile({
            'q': "title = 'FUP' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        }).GetList()

        if pastas_fup:
            pasta_fup = pastas_fup[0]
        else:
            pasta_fup = drive.CreateFile({
                'title': 'FUP',
                'mimeType': 'application/vnd.google-apps.folder'
            })
            pasta_fup.Upload()

        # Verifica ou cria a subpasta "evidencias" dentro de "FUP"
        lista = drive.ListFile({
            'q': f"'{pasta_fup['id']}' in parents and title = 'evidencias' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        }).GetList()

        if lista:
            pasta_principal = lista[0]
        else:
            pasta_principal = drive.CreateFile({
                'title': 'evidencias',
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [{'id': pasta_fup['id']}]
            })
            pasta_principal.Upload()

        # Cria ou encontra a subpasta indice_x
        subpasta_nome = f"indice_{idx}"
        subpastas = drive.ListFile({
            'q': f"'{pasta_principal['id']}' in parents and title='{subpasta_nome}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        }).GetList()

        if subpastas:
            subpasta = subpastas[0]
        else:
            subpasta = drive.CreateFile({
                'title': subpasta_nome,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [{'id': pasta_principal['id']}]
            })
            subpasta.Upload()

        # 📎 Upload dos arquivos
        for arq in arquivos:
            arquivo_drive = drive.CreateFile({
                'title': arq.name,
                'parents': [{'id': subpasta['id']}]
            })
            with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
                tmp_file.write(arq.getvalue())
                tmp_file.flush()
                arquivo_drive.SetContentFile(tmp_file.name)
                arquivo_drive.Upload()

        # 📝 Observação
        if observacao.strip():
            obs_file = drive.CreateFile({
                'title': 'observacao.txt',
                'parents': [{'id': subpasta['id']}]
            })
            obs_file.SetContentString(observacao.strip())
            obs_file.Upload()

        st.success("✅ Evidências enviadas ao Google Drive com sucesso.")
        return True

    except Exception as e:
        st.error(f"Erro ao enviar evidências para o Drive: {e}")
        return False
        
def carregar_followups():
    drive = conectar_drive()

    # Garante que a pasta FUP exista
    pastas_fup = drive.ListFile({
        'q': "title = 'FUP' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    }).GetList()

    if pastas_fup:
        pasta_fup = pastas_fup[0]
    else:
        pasta_fup = drive.CreateFile({
            'title': 'FUP',
            'mimeType': 'application/vnd.google-apps.folder'
        })
        pasta_fup.Upload()

    # Busca o followups.csv dentro da FUP
    arquivos = drive.ListFile({
        'q': f"'{pasta_fup['id']}' in parents and title = 'followups.csv' and trashed=false"
    }).GetList()

    colunas = [
        "Titulo", "Ambiente", "Ano", "Auditoria", "Apontamento", "Risco",
        "Plano de Acao", "Responsavel", "Usuario", "E-mail",
        "Prazo", "Data de Conclusão", "Status", "Avaliação FUP", "Observação"
    ]

    if not arquivos:
        df_vazio = pd.DataFrame(columns=colunas)
        caminho_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv").name
        df_vazio.to_csv(caminho_temp, sep=";", index=False, encoding="utf-8-sig")

        novo_arquivo = drive.CreateFile({
            'title': 'followups.csv',
            'parents': [{'id': pasta_fup['id']}]
        })
        novo_arquivo.SetContentFile(caminho_temp)
        novo_arquivo.Upload()
        return df_vazio

    caminho_temp = tempfile.NamedTemporaryFile(delete=False).name
    arquivos[0].GetContentFile(caminho_temp)

    try:
        df = pd.read_csv(caminho_temp, sep=";", encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(caminho_temp, sep=";", encoding="latin1")

    return df

def aplicar_filtros_df(df, pergunta):
    filtros = {}
    valores_unicos = {}

    for col in df.select_dtypes(include="object").columns:
        valores_unicos[col] = df[col].astype(str).str.lower().str.strip().unique().tolist()

    tokens = re.findall(r"\w+", pergunta.lower())

    for token in tokens:
        for col, valores in valores_unicos.items():
            match = get_close_matches(token, valores, n=1, cutoff=0.8)
            if match:
                filtros[col] = match[0]
                break

    df_filtrado = df.copy()

    for col, valor in filtros.items():
        df_filtrado[col] = df_filtrado[col].astype(str).str.lower().str.strip()
        df_filtrado = df_filtrado[df_filtrado[col].str.contains(valor)]

    return df_filtrado, filtros

# --- Usuários e autenticação simples ---
@st.cache_data
def carregar_usuarios():
    usuarios_config = st.secrets.get("users", {})
    usuarios = {}
    for user, dados in usuarios_config.items():
        try:
            nome, senha = dados.split("|", 1)
            usuarios[user] = {"name": nome, "password": senha}
        except:
            st.warning(f"Erro ao carregar usuário '{user}' nos secrets.")
    return usuarios

users = carregar_usuarios()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""

if not st.session_state.logged_in:
    st.title("🔐 Login")
    username = st.text_input("Usuário")
    password = st.text_input("Senha", type="password")
    if st.button("Entrar"):
        user = users.get(username)
        if user and user["password"] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success(f"Bem-vindo, {user['name']}!")
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos.")
    st.stop()

# --- Layout principal após login ---
st.sidebar.image("PRIO_SEM_POLVO_PRIO_PANTONE_LOGOTIPO_Azul.png")
nome_usuario = users[st.session_state.username]["name"]
st.sidebar.success(f"Logado como: {nome_usuario}")
if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.rerun()

# --- Menu lateral ---
st.sidebar.title("📋 Menu")
menu = st.sidebar.radio("Navegar para:", [
    "Dashboard",
    "Meus Follow-ups",
    "Cadastrar Follow-up",
    "Enviar Evidências",
    "Visualizar Evidências",
    "🔍 Chatbot FUP"
])

# --- Conteúdo das páginas ---

if menu == "Dashboard":
    st.title("📊 Painel de KPIs")

    try:
        # Conecta ao Google Drive
        drive = conectar_drive()
    
        # Procura arquivo chamado 'followups.csv'
        pastas_fup = drive.ListFile({
            'q': "title = 'FUP' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        }).GetList()
        if not pastas_fup:
            st.warning("Pasta 'FUP' não encontrada no Drive.")
            st.stop()
        
        arquivos = drive.ListFile({
            'q': f"'{pastas_fup[0]['id']}' in parents and title = 'followups.csv' and trashed = false"
        }).GetList()

        if not arquivos:
            st.warning("Arquivo followups.csv não encontrado no Drive.")
            st.stop()
    
        arquivo = arquivos[0]
        caminho_temp = tempfile.NamedTemporaryFile(delete=False).name
        arquivo.GetContentFile(caminho_temp)
    
        # Carrega CSV com pandas
        df = pd.read_csv(caminho_temp, sep=";", encoding="utf-8-sig")
        df.columns = df.columns.str.strip()
    
        usuario_logado = st.session_state.username
        nome_usuario = users[usuario_logado]["name"]
    
        if usuario_logado not in admin_users:
            df = df[df["Responsavel"].str.lower() == nome_usuario.lower()]
    
        if df.empty:
            st.info("Nenhum dado disponível para exibir KPIs.")
            st.stop()
    
        df["Prazo"] = pd.to_datetime(df["Prazo"], format="mixed", errors="coerce")
        df["Ano"] = df["Ano"].astype(str)
        df["Status"] = df["Status"].fillna("Não informado")

        # --- Filtros principais ---
        filtro_vencidos = st.radio(
        label='Selecione os status desejados:',
        options=['Todos', 'No Prazo', 'Vencidos'],
        key='filtro_vencidos'
    )
        
        lista_auditorias = sorted(df['Auditoria'].unique().tolist()) + ['Todas']
        filtro_auditoria = st.multiselect('Selecione as auditorias: ', lista_auditorias, default='Todas')
        
        if filtro_vencidos == 'Vencidos':
            df = df[df['Prazo']< hoje]
        elif filtro_vencidos == 'No Prazo':
            df = df[df['Prazo']>= hoje]

        if 'Todas' not in lista_auditorias:
            df = df[df['Auditoria'].isin(filtro_auditoria)]
            
        # --- KPIs principais ---
        total = len(df)
        concluidos = (df["Status"] == "Concluído").sum()
        pendentes = (df["Status"] == "Pendente").sum()
        andamento = (df["Status"] == "Em Andamento").sum()
        taxa_conclusao = round((concluidos / total) * 100, 1) if total > 0 else 0.0
    
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Total Follow-ups", total)
        col2.metric("Concluídos", concluidos)
        col3.metric("Pendentes", pendentes)
        col4.metric("Em Andamento", andamento)
        col5.metric("Conclusão (%)", f"{taxa_conclusao}%")
    
        # --- Gráficos ---
        st.subheader("📌 Distribuição por Status")
        fig_status = px.pie(
            df,
            names="Status",
            title="Distribuição dos Follow-ups por Status",
            hole=0.4
        )
        st.plotly_chart(fig_status, use_container_width=True)
    
        st.subheader("📁 Follow-ups por Auditoria")
        auditoria_counts = df["Auditoria"].value_counts().reset_index()
        auditoria_counts.columns = ["Auditoria", "Quantidade"]

        fig_auditoria = px.bar(
            auditoria_counts,
            x="Auditoria",
            y="Quantidade",
            title="Distribuição de Follow-ups por Auditoria"
        )
        st.plotly_chart(fig_auditoria, use_container_width=True)
    
        st.subheader("📅 Follow-ups por Ano")
        ano_counts = df["Ano"].value_counts().sort_index().reset_index()
        ano_counts.columns = ["Ano", "Quantidade"]
        ano_counts["Ano"] = ano_counts["Ano"].astype(str)
        anos_existentes = ano_counts["Ano"].tolist()
        ano_counts["Ano"] = pd.Categorical(ano_counts["Ano"], categories=anos_existentes, ordered=True)
        
        fig_ano = px.line(
            ano_counts,
            x="Ano",
            y="Quantidade",
            markers=True,
            title="Evolução de Follow-ups por Ano"
        )
        fig_ano.update_xaxes(type='category')
        
        st.plotly_chart(fig_ano, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao acessar dados do Drive: {e}")

elif menu == "Meus Follow-ups":
    st.title("📁 Meus Follow-ups")
    st.info("Esta seção exibirá os follow-ups atribuídos a você.")

    try:
        df = carregar_followups()
        df.columns = df.columns.str.strip()

        usuario_logado = st.session_state.username
        nome_usuario = users[usuario_logado]["name"]

        if usuario_logado not in admin_users:
            df = df[df["Responsavel"].str.lower() == nome_usuario.lower()]
            
        df["Prazo"] = pd.to_datetime(df["Prazo"], format = "mixed", errors="coerce")
        #df["Prazo"] = df["Prazo"].dt.strftime("%d/%m/%Y")
        df["Data de Conclusão"] = pd.to_datetime(df["Data de Conclusão"], format = "mixed", errors="coerce")
        #df["Data de Conclusão"] = df["Data de Conclusão"].dt.strftime("%d/%m/%Y")
        df["Ano"] = df["Ano"].astype(str)
        df["Ambiente"] = df["Ambiente"].str.lower()

        # --- Filtros na sidebar ---
        st.sidebar.subheader("Filtros de Pesquisa")

        if st.sidebar.button("🔄 Limpar Filtros"):
            st.rerun()

        auditorias = ["Todos"] + sorted(df["Auditoria"].dropna().unique().tolist())
        auditoria_selecionada = st.sidebar.selectbox("Auditoria", auditorias)

        status_lista = ["Todos"] + sorted(df["Status"].dropna().unique().tolist())
        status_selecionado = st.sidebar.selectbox("Status", status_lista)

        status_ambiente = ["Todos"] + sorted(df["Ambiente"].dropna().unique().tolist())
        status_ambiente_selecionado = st.sidebar.selectbox("Ambiente", status_ambiente)

        anos = ["Todos"] + sorted(df["Ano"].dropna().unique().tolist())
        ano_selecionado = st.sidebar.selectbox("Ano", anos)

        vencimento = ["Todos", "No Prazo", "Vencido"]
        vencimento_selecionado = st.sidebar.selectbox("Tipo de Vencimento", vencimento)

        prazo_inicial, prazo_final = st.sidebar.date_input(
            "Intervalo de Prazo",
            [df["Prazo"].min().date(), df["Prazo"].max().date()]
        )

        if auditoria_selecionada != "Todos":
            df = df[df["Auditoria"] == auditoria_selecionada]

        if status_selecionado != "Todos":
            df = df[df["Status"] == status_selecionado]

        if status_ambiente_selecionado != "Todos":
            df = df[df["Ambiente"] == status_ambiente_selecionado]
        
        if ano_selecionado != "Todos":
            df = df[df["Ano"] == ano_selecionado]

        if vencimento_selecionado == 'Vencido':
            df = df[df['Prazo']< hoje]
        elif vencimento_selecionado == 'No Prazo':
            df = df[df['Prazo']>= hoje]

        df = df[(df["Prazo"].dt.date >= prazo_inicial) & (df["Prazo"].dt.date <= prazo_final)]
        df = df.sort_values(by="Prazo")

        if not df.empty:
            df["Ambiente"] = df["Ambiente"].str.lower()
            st.dataframe(df, use_container_width=True)
            st.success(f"Total Follow Ups: {len(df)}")

            st.subheader("🛠️ Atualizar / Excluir Follow-up por Índice")

            indices_disponiveis = df.index.tolist()
            indice_selecionado = st.selectbox("Selecione o índice para edição", indices_disponiveis)

            linha = df.loc[indice_selecionado]
            st.markdown(f"""
            🔎 **Título:** {linha['Titulo']}  
            📅 **Prazo:** {linha['Prazo'].strftime('%d/%m/%Y')}  
            👤 **Responsável:** {linha['Responsavel']}  
            📌 **Status:** {linha['Status']}
            """)

            colunas_editaveis = [col for col in df.columns]
            coluna_escolhida = st.selectbox("Selecione a coluna para alterar", colunas_editaveis)

            valor_atual = linha[coluna_escolhida]

            if coluna_escolhida in ["Prazo", "Data de Conclusão"]:
                try:
                    data_inicial = pd.to_datetime(valor_atual).date()
                except:
                    data_inicial = date.today()
                novo_valor = st.date_input(f"Novo valor para '{coluna_escolhida}':", value=data_inicial)
                novo_valor_str = novo_valor.strftime("%Y-%m-%d")
            else:
                novo_valor = st.text_input(f"Valor atual de '{coluna_escolhida}':", value=str(valor_atual))
                novo_valor_str = novo_valor.strip()

            if st.button("💾 Atualizar campo"):
                df_original = carregar_followups()
                df_original.at[indice_selecionado, coluna_escolhida] = novo_valor_str
                df_original.to_csv(caminho_csv, sep=";", index=False, encoding="utf-8-sig")

                try:
                    drive = conectar_drive()
                    arquivos = drive.ListFile({'q': "title = 'followups.csv' and trashed=false"}).GetList()
                    if arquivos:
                        arquivo = arquivos[0]
                        upload_para_drive()
                    st.info("📤 Arquivo 'followups.csv' atualizado no Drive.")
                except Exception as e:
                    st.warning(f"Erro ao enviar para o Drive: {e}")

                st.success(f"'{coluna_escolhida}' atualizado com sucesso.")
                st.rerun()

            if usuario_logado in admin_users:
                if st.button("🗑️ Excluir este follow-up"):
                    df_original = df.drop(index=indice_selecionado)
                    df_original.to_csv(caminho_csv, sep=";", index=False, encoding="utf-8-sig")

                    try:
                        drive = conectar_drive()
                        arquivos = drive.ListFile({'q': "title = 'followups.csv' and trashed=false"}).GetList()
                        if arquivos:
                            arquivo = arquivos[0]
                            upload_para_drive()
                        st.info("📤 Arquivo 'followups.csv' atualizado no Google Drive.")
                    except Exception as e:
                        st.warning(f"Erro ao enviar para o Drive: {e}")

                    st.success("Follow-up excluído com sucesso.")
                    st.rerun()

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='FollowUps')

            st.download_button(
                label="📥 Exportar resultados para Excel",
                data=buffer.getvalue(),
                file_name="followups_filtrados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.info("Nenhum follow-up encontrado com os filtros aplicados.")

    except Exception as e:
        st.error(f"Erro ao acessar dados do Drive: {e}")

elif menu == "Cadastrar Follow-up":
    st.title("📝 Cadastrar Follow-up")
    if st.session_state.username in cadastro_users:
        st.info("Aqui você poderá cadastrar um novo follow-up.")
    
        with st.form("form_followup"):
            titulo = st.text_input("Título")
            ambiente = st.text_input("Ambiente")
            ano = st.selectbox("Ano", list(range(2020, date.today().year + 2)))
            auditoria = st.text_input("Auditoria")
            apontamento = st.text_input("Apontamento")
            risco = st.selectbox("Risco", ["Baixo", "Médio", "Alto"])
            plano = st.text_area("Plano de Ação")
            responsavel = st.text_input("Responsável")
            usuario = st.text_input("Usuário")
            email = st.text_input("E-mail do Responsável")
            prazo = st.date_input("Prazo", min_value=date.today())
            data_conclusao = st.date_input("Data de Conclusão", value=date.today())
            status = st.selectbox("Status", ["Pendente", "Em Andamento", "Concluído"])
            avaliacao = st.selectbox("Avaliação FUP", ["", "Satisfatório", "Insatisfatório"])
            observacao = st.text_area("Observação")
    
            submitted = st.form_submit_button("Salvar Follow-up")
    
        if submitted:
            novo = {
                "Titulo": titulo,
                "Ambiente": ambiente,
                "Ano": ano,
                "Auditoria": auditoria,
                "Apontamento": apontamento,
                "Risco": risco,
                "Plano de Acao": plano,
                "Responsavel": responsavel,
                "Usuario": usuario,
                "E-mail": email,
                "Prazo": prazo.strftime("%Y-%m-%d"),
                "Data de Conclusão": data_conclusao.strftime("%Y-%m-%d"),
                "Status": status,
                "Avaliação FUP": avaliacao,
                "Observação": observacao
            }
    
            try:
                # Conecta ao Drive e busca o followups.csv
                drive = conectar_drive()
                arquivos = drive.ListFile({
                    'q': "title = 'followups.csv' and trashed=false"
                }).GetList()
    
                if arquivos:
                    arquivo = arquivos[0]
                    caminho_temp = tempfile.NamedTemporaryFile(delete=False).name
                    arquivo.GetContentFile(caminho_temp)
                    df = pd.read_csv(caminho_temp, sep=";", encoding="utf-8-sig")
                else:
                    df = pd.DataFrame()
                    arquivo = drive.CreateFile({'title': 'followups.csv'})
    
                # Atualiza e salva
                df = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
                df.to_csv(caminho_csv, sep=";", index=False, encoding="utf-8-sig")
    
                upload_para_drive()
    
                st.success("✅ Follow-up salvo e sincronizado com o Drive!")
    
                corpo = f"""
                <p>Olá <b>{responsavel}</b>,</p>
                <p>Um novo follow-up foi atribuído a você:</p>
                <ul>
                    <li><b>Título:</b> {titulo}</li>
                    <li><b>Auditoria:</b> {auditoria}</li>
                    <li><b>Apontamento:</b> {apontamento}</li>
                    <li><b>Plano de Acao:</b> {plano}</li>
                    <li><b>Prazo:</b> {prazo.strftime('%d/%m/%Y')}</li>
                    <li><b>Status:</b> {status}</li>
                </ul>
                <p>Acesse o aplicativo para incluir evidências e acompanhar o andamento:</p>
                <p><a href='https://fup-auditoria.streamlit.app/' target='_blank'>🔗 fup-auditoria.streamlit.app</a></p>
                <br>
                <p>Atenciosamente,<br>Auditoria Interna</p>
                """
    
                if email:
                    sucesso_envio = enviar_email_gmail(
                        destinatario=email,
                        assunto=f"[Follow-up] Nova Atribuição: {titulo}",
                        corpo_html=corpo
                    )
                    if sucesso_envio:
                        st.success("📧 E-mail de notificação enviado com sucesso!")
    
            except Exception as e:
                st.error(f"Erro ao cadastrar follow-up: {e}")
    else:
        st.warning("Você não possui permissão para cadastrar follow ups!")

elif menu == "Enviar Evidências":
    st.title("📌 Enviar Evidências")
    st.info("Aqui você poderá enviar comprovantes e observações para follow-ups.")

    try:
        # 🔄 Puxa o arquivo mais recente do Drive
        drive = conectar_drive()
        arquivos_drive = drive.ListFile({
            'q': "title = 'followups.csv' and trashed=false"
        }).GetList()

        if not arquivos_drive:
            st.warning("Arquivo followups.csv não encontrado no Google Drive.")
            st.stop()

        arquivo_drive = arquivos_drive[0]
        caminho_temp = tempfile.NamedTemporaryFile(delete=False).name
        arquivo_drive.GetContentFile(caminho_temp)
        df = pd.read_csv(caminho_temp, sep=";", encoding="utf-8-sig")
        df.columns = df.columns.str.strip()

        usuario_logado = st.session_state.username
        nome_usuario = users[usuario_logado]["name"]

        if usuario_logado not in admin_users:
            df = df[df["Responsavel"].str.lower() == nome_usuario.lower()]

        if df.empty:
            st.info("Nenhum follow-up disponível para envio de evidência.")
            st.stop()

        idx = st.selectbox("Selecione o índice do follow-up:", df.index.tolist())
        linha = df.loc[idx]

        st.markdown(f"""
        🔎 **Título:** {linha['Titulo']}  
        🚩 **Apontamento:** {linha['Apontamento']}  
        📅 **Prazo:** {linha['Prazo']}  
        👤 **Responsável:** {linha['Responsavel']}  
        📝 **Plano de Ação:** {linha['Plano de Acao']}
        """)

        arquivos = st.file_uploader(
            "Anexe arquivos de evidência",
            type=["pdf", "png", "jpg", "jpeg", "zip", "doc", "docx", "eml", "msg"],
            accept_multiple_files=True
        )
        observacao = st.text_area("Observações (opcional)")

        submitted = st.button("📨 Enviar Evidência")
        if submitted:
            if not arquivos:
                st.warning("Você precisa anexar pelo menos um arquivo.")
                st.stop()

            # Upload direto para o Drive
            sucesso_upload = upload_evidencias_para_drive(idx, arquivos, observacao)

            # Registro em log (local)
            if sucesso_upload:
                try:
                    log_path = Path("log_evidencias.csv")
                    log_data = {
                        "indice": idx,
                        "titulo": linha["Titulo"],
                        "responsavel": linha["Responsavel"],
                        "arquivos": "; ".join([arq.name for arq in arquivos]),
                        "observacao": observacao,
                        "data_envio": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "enviado_por": nome_usuario,
                    }
                    log_df = pd.DataFrame([log_data])
                    if log_path.exists():
                        log_df.to_csv(log_path, mode="a", header=False, index=False)
                    else:
                        log_df.to_csv(log_path, index=False)
                    st.success("✅ Registro salvo no log local.")
                except Exception as e:
                    st.error(f"Erro ao salvar log local: {e}")

                # Enviar e-mail de notificação
                corpo = f"""
                <p>🕵️ Evidência enviada para o follow-up:</p>
                <ul>
                    <li><b>Índice:</b> {idx}</li>
                    <li><b>Título:</b> {linha['Titulo']}</li>
                    <li><b>Apontamento:</b> {linha['Apontamento']}</li>
                    <li><b>Responsável:</b> {linha['Responsavel']}</li>
                    <li><b>Arquivos:</b> {"; ".join([arq.name for arq in arquivos])}</li>
                    <li><b>Data:</b> {datetime.now().strftime("%d/%m/%Y %H:%M")}</li>
                </ul>
                <p>Evidências armazenadas no Drive (pasta: <b>evidencias/indice_{idx}</b>).</p>
                """

                destinatarios_evidencias = ["cvieira@prio3.com.br","mathayde@prio3.com.br","amendonca@prio3.com.br"]
                
                sucesso_envio = enviar_email_gmail(
                    destinatario=destinatarios_evidencias,
                    assunto=f"[Evidência] Follow-up #{idx} - {linha['Titulo']}",
                    corpo_html=corpo
                )
                if sucesso_envio:
                    st.success("📧 Notificação enviada ao time de auditoria!")

    except Exception as e:
        st.error(f"Erro ao carregar dados do Drive: {e}")

elif menu == "Visualizar Evidências":

    st.title("📂 Visualização de Evidências")

    try:
        drive = conectar_drive()
        df = carregar_followups()
        df.columns = df.columns.str.strip()

        usuario_logado = st.session_state.username
        nome_usuario = users[usuario_logado]["name"].lower()

        # Pasta "evidencias"
        pasta_principal = drive.ListFile({
            'q': "title='evidencias' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        }).GetList()

        if not pasta_principal:
            st.warning("Nenhuma pasta de evidências encontrada.")
            st.stop()

        pasta_id = pasta_principal[0]['id']

        subpastas = drive.ListFile({
            'q': f"'{pasta_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        }).GetList()

        if usuario_logado in admin_users:
            opcoes = {
                p['title'].split('_')[1]: {'id': p['id'], 'obj': p}
                for p in subpastas if p['title'].startswith('indice_') and '_' in p['title']
            }
        else:
            indices_usuario = df[df['Responsavel'].str.lower() == nome_usuario].index.astype(str).tolist()
            opcoes = {
                p['title'].split('_')[1]: {'id': p['id'], 'obj': p}
                for p in subpastas
                if p['title'].startswith('indice_') and '_' in p['title'] and p['title'].split('_')[1] in indices_usuario
            }

        if not opcoes:
            st.warning("Você não possui evidências associadas.")
            st.stop()

        indices_disponiveis = sorted(opcoes.keys(), key=int)
        indice_escolhido = st.selectbox("Selecione o índice do follow-up:", indices_disponiveis)

        if indice_escolhido not in opcoes:
            st.error(f"Índice '{indice_escolhido}' não encontrado.")
            st.stop()

        pasta_selecionada_id = opcoes[indice_escolhido]['id']
        pasta_obj = opcoes[indice_escolhido]['obj']

        st.subheader(f"📁 Evidências para Follow-up #{indice_escolhido}")

        arquivos = drive.ListFile({
            'q': f"'{pasta_selecionada_id}' in parents and trashed=false"
        }).GetList()

        if not arquivos:
            st.info("Nenhum arquivo nesta pasta.")
            st.stop()

        buffer_zip = BytesIO()
        with zipfile.ZipFile(buffer_zip, "w") as zipf:
            for arq in arquivos:
                nome = arq['title']
                if nome.lower() == "observacao.txt":
                    conteudo = arq.GetContentString()
                    st.markdown("**📝 Observação:**")
                    st.info(conteudo)
                    zipf.writestr(nome, conteudo)
                else:
                    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
                        arq.GetContentFile(tmp_file.name)
                        tmp_file.seek(0)
                        zipf.write(tmp_file.name, arcname=nome)
                        link = arq['alternateLink']
                        st.markdown(f"📎 [{nome}]({link})", unsafe_allow_html=True)

        buffer_zip.seek(0)
        st.download_button(
            label="📦 Baixar todos como .zip",
            data=buffer_zip,
            file_name=f"evidencias_indice_{indice_escolhido}.zip",
            mime="application/zip"
        )

        if usuario_logado in admin_users:
            if st.button("🗑️ Excluir todas as evidências deste índice"):
                try:
                    pasta_obj.Delete()
                    st.success(f"Evidências do índice {indice_escolhido} excluídas com sucesso.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao excluir evidências: {e}")

    except Exception as e:
        st.error("Erro ao acessar evidências no Drive.")
        st.code(traceback.format_exc())

elif menu == "🔍 Chatbot FUP":

    st.title("🤖 Chatbot dos Relatórios de Auditoria")
    
    usuario_logado = st.session_state.username
    nome_usuario = users[usuario_logado]["name"]
    
    df = carregar_followups()
    if df.empty:
        st.warning("Nenhum dado disponível.")
        st.stop()
    
    if usuario_logado not in admin_users:
        df = df[df["Responsavel"].str.lower() == nome_usuario.lower()]
    
    st.markdown("### 📝 Digite sua pergunta sobre os follow-ups:")
    pergunta = st.text_input(
        "Ex: Quais são os principais riscos dos meus follow-ups? Ou: Me mostre os pontos críticos no ambiente SAP.",
        key="pergunta_fup"
    )
    
    if 'executar_analise' not in st.session_state:
        st.session_state.executar_analise = False
    if 'executar_consultor' not in st.session_state:
        st.session_state.executar_consultor = False
    
    # 🔘 Botão da análise executiva
    if st.button("📨 Executar Análise"):
        st.session_state.executar_analise = True
        st.session_state.executar_consultor = False
    
    if st.session_state.executar_analise:
        df_filtrado, filtros = aplicar_filtros_df(df, pergunta)
    
        dados_filtrados = df_filtrado.fillna("").astype(str).to_markdown(index=False) if not df_filtrado.empty else "Nenhum follow-up encontrado."
        dados_completo = df.fillna("").astype(str).to_markdown(index=False)
    
        system_prompt = f"""
    Você é um especialista sênior em Auditoria, Riscos, Governança e Controles Internos, com domínio dos frameworks:
    - COSO, COBIT, ISO 27001, NIST CSF, ITIL e PMBOK.
    
    ### 🎯 Sua missão:
    1. Gerar um **SUMÁRIO EXECUTIVO** robusto com:
    - Principais riscos dos follow-ups filtrados.
    - Temas críticos, controles deficientes, prazos críticos.
    - Status (atrasados, pendentes, em andamento).
    - Distribuição por ambiente, ano, risco e auditoria.
    - Referência aos frameworks relevantes para os riscos identificados.
    
    2. Na sequência, liste os follow-ups encontrados:
    - Para cada um, apresente:
      - 📜 Descrição breve.
      - 🔥 Status e Risco.
      - 📌 Ambiente e Auditoria relacionada.
    
    ---
    
    ### 🗂️ Base filtrada:
    {dados_filtrados}
    
    ### 🏛️ Base total:
    {dados_completo}
    
    ---
    
    ⚠️ Seja técnico, objetivo e aderente às melhores práticas profissionais.
    """
    
        payload = {
            "model": "gpt-4o",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": pergunta}
            ],
            "temperature": 0.2
        }
    
        headers = {
            "Authorization": f"Bearer {st.secrets['openai']['api_key']}",
            "Content-Type": "application/json"
        }
    
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=payload,
            verify=False
        )
    
        if response.status_code == 200:
            resposta_analise = response.json()["choices"][0]["message"]["content"]
        else:
            resposta_analise = f"Erro na API: {response.status_code} - {response.text}"
    
        st.subheader("💡 Resultado da Análise Executiva")
        st.markdown(resposta_analise)
    
        if st.button("🚀 Consultor de Planos de Ação"):
            st.session_state.executar_consultor = True
    
    if st.session_state.executar_consultor:
        df_filtrado, filtros = aplicar_filtros_df(df, pergunta)
    
        dados_filtrados = df_filtrado.fillna("").astype(str).to_markdown(index=False) if not df_filtrado.empty else "Nenhum follow-up encontrado."
    
        prompt_consultor = f"""
Você é um consultor sênior, especialista em governança, riscos, compliance, auditoria e gestão de projetos.

Sua missão é ajudar o usuário a **sanar os follow-ups identificados**, propondo **formas práticas e detalhadas de executar cada plano de ação existente na base de dados**.

---

### 🎯 Para cada follow-up listado na base:
1. **Leia atentamente o conteúdo do campo "Plano de Acao"** e interprete qual é a ação que está sendo proposta.

2. Gere um **plano de execução detalhado**, incluindo:
   - 📜 **Descrição prática de como executar o plano de ação.**
   - 🔧 **Ferramentas, metodologias ou sistemas que podem ser utilizados.**
   - ✅ **Critérios de avaliação, checklists ou requisitos que devem ser analisados.**
   - 🚩 **Principais riscos e cuidados que precisam ser tomados durante a execução.**
   - 🧠 **Boas práticas de mercado e referência aos frameworks aplicáveis (COBIT, COSO, ISO 27001, NIST, ITIL, PMBOK).**

---

### 💡 **Exemplo esperado:**
- Se o plano de ação diz: "**Executar due diligence do fornecedor**":
   - Descreva:
     - Como estruturar um processo de due diligence.
     - Quais critérios devem ser avaliados (ex.: integridade, questões financeiras, trabalhistas, ambientais).
     - Quais ferramentas podem ser usadas (ex.: sites públicos, bases de dados, softwares como LexisNexis, Refinitiv, D&B).
     - Principais cuidados, como veracidade das informações e atualização dos dados.
     - Frameworks que apoiam essa prática (ex.: ISO 37001, COSO, Compliance Programs).

---

### 🗂️ Base de follow-ups:
{dados_filtrados}

---

⚠️ Importante:
- O plano deve ser **100% personalizado com base no conteúdo real dos planos de ação da base**.
- Não escreva respostas genéricas.
- Cada follow-up deve gerar uma análise própria, com orientações práticas, específicas e acionáveis.
- Seja extremamente profissional, técnico, detalhado e aderente às melhores práticas internacionais.
"""
    
        payload2 = {
            "model": "gpt-4o",
            "messages": [
                {"role": "system", "content": prompt_consultor},
                {"role": "user", "content": "Como posso estruturar um projeto para resolver meus follow-ups?"}
            ],
            "temperature": 0.2
        }
    
        response2 = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=payload2,
            verify=False
        )
    
        if response2.status_code == 200:
            resposta_consultor = response2.json()["choices"][0]["message"]["content"]
        else:
            resposta_consultor = f"Erro na API: {response2.status_code} - {response2.text}"
    
        st.subheader("🏗️ Consultoria - Plano de Ação")
        st.markdown(resposta_consultor)
    
    # 🔍 Visualizar follow-ups encontrados
    st.markdown("### 📋 Follow-ups encontrados:")
    df_filtrado, filtros = aplicar_filtros_df(df, pergunta)
    if not df_filtrado.empty:
        st.dataframe(df_filtrado, use_container_width=True)
    else:
        st.info("Nenhum follow-up encontrado.")
        
# Função para enviar e-mail mensal com follow-ups vencidos

def enviar_emails_followups_vencidos():
    df = carregar_followups()
    df.columns = df.columns.str.strip()

    df["Prazo"] = pd.to_datetime(df["Prazo"], format="mixed", errors="coerce")
    df["Prazo"] = df["Prazo"].dt.normalize()
    hoje = pd.Timestamp.today().normalize()

    df_vencidos = df[
        (df["Status"].str.lower() != "concluído") &
        (df["Prazo"] < hoje)
    ]

    if df_vencidos.empty:
        st.info("✅ Nenhum follow-up vencido identificado para envio.")
        return

    # Agrupar por responsável
    responsaveis = df_vencidos["E-mail"].dropna().unique().tolist()

    for email in responsaveis:
        df_resp = df_vencidos[df_vencidos["E-mail"] == email]

        if df_resp.empty:
            continue

        corpo = f"""
        <p>Olá,</p>
        <p>Você possui os seguintes follow-ups vencidos:</p>
        <table border='1' cellpadding='4' cellspacing='0'>
            <tr><th>Título</th><th>Auditoria</th><th>Plano de Ação</th><th>Responsável</th><th>Prazo</th><th>Status</th></tr>
        """

        for _, row in df_resp.iterrows():
            corpo += f"<tr><td>{row['Titulo']}</td><td>{row['Auditoria']}</td><td>{row['Plano de Acao']}</td><td>{row['Responsavel']}</td><td>{row['Prazo'].date()}</td><td>{row['Status']}</td></tr>"

        corpo += """
        </table>
        <p>Por favor, atualize os registros no sistema ou entre em contato com a Auditoria Interna.</p>
        <p>Acesse o aplicativo para incluir evidências e acompanhar o andamento:</p>
        <p><a href='https://fup-auditoria.streamlit.app/' target='_blank'>🔗 fup-auditoria.streamlit.app</a></p>
        <br>
        <p>Atenciosamente,<br>Time de Auditoria</p>
        """

        try:
            yag = yagmail.SMTP(user=st.secrets["email_user"], password=st.secrets["email_pass"])
            yag.send(to=email, subject="📌 Follow-ups vencidos - Auditoria Interna", contents=corpo)
            st.success(f"📧 E-mail enviado para: {email}")
        except Exception as e:
            st.warning(f"Erro ao enviar para {email}: {e}")

# 🔁 Botão para envio manual

if st.session_state.username in admin_users:
    if st.sidebar.button("✉️ Enviar lembrete de follow-ups vencidos"):
        enviar_emails_followups_vencidos()
#-------------------------------------------------------------------- e-mail de follow ups a vencer
def enviar_emails_followups_a_vencer():
    df = carregar_followups()
    df.columns = df.columns.str.strip()
    df["Prazo"] = pd.to_datetime(df["Prazo"], errors="coerce")

    hoje = pd.Timestamp.today()
    limite = hoje + timedelta(days=30)

    df_a_vencer = df[
        (df["Status"].str.lower() != "concluído") &
        (df["Prazo"] >= hoje) &
        (df["Prazo"] <= limite)
    ]

    if df_a_vencer.empty:
        st.info("✅ Nenhum follow-up com prazo a vencer em 30 dias.")
        return

    responsaveis = df_a_vencer["E-mail"].dropna().unique().tolist()

    for email in responsaveis:
        df_resp = df_a_vencer[df_a_vencer["E-mail"] == email]
        if df_resp.empty:
            continue

        corpo = f"""
        <p>Olá,</p>
        <p>Você possui os seguintes follow-ups com prazo a vencer em até 30 dias:</p>
        <table border='1' cellpadding='4' cellspacing='0'>
            <tr><th>Título</th><th>Auditoria</th><th>Plano de Ação</th><th>Responsável</th><th>Prazo</th><th>Status</th></tr>
        """

        for _, row in df_resp.iterrows():
            corpo += f"<tr><td>{row['Titulo']}</td><td>{row['Auditoria']}</td><td>{row['Plano de Acao']}</td><td>{row['Responsavel']}</td><td>{row['Prazo'].date()}</td><td>{row['Status']}</td></tr>"

        corpo += """
        </table>
        <p>Por favor, antecipe ações necessárias e atualize o status no sistema.</p>
        <p>Acesse o aplicativo para mais detalhes:</p>
        <p><a href='https://fup-auditoria.streamlit.app/' target='_blank'>🔗 fup-auditoria.streamlit.app</a></p>
        <br>
        <p>Atenciosamente,<br>Time de Auditoria</p>
        """

        sucesso = enviar_email_gmail(destinatario=email, assunto="⏳ Follow-ups próximos do vencimento", corpo_html=corpo)
        if sucesso:
            st.success(f"📧 E-mail enviado para: {email}")

if st.session_state.username in admin_users:
    if st.sidebar.button("📅 Enviar lembrete de follow-ups a vencer"):
        enviar_emails_followups_a_vencer()
