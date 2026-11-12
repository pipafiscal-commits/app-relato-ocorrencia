import streamlit as st
import pandas as pd
from datetime import date, datetime
import hmac
import io
from pathlib import Path
import socket
from PIL import Image
import qrcode
import gspread
import requests
from xml.sax.saxutils import escape
from urllib.request import urlopen
from streamlit.errors import StreamlitSecretNotFoundError
from google.auth.exceptions import GoogleAuthError

# Dependências para geração de PDF
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ==========================================
# CONFIGURAÇÃO DA PÁGINA (MOBILE FIRST)
# ==========================================
st.set_page_config(
    page_title="Relato de Ocorrência - EROCP/10ª RM",
    page_icon="🚛",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Estilização CSS para otimização em Smartphone
st.markdown("""
    <style>
    .main { padding: 0.5rem; }
    .stButton>button { width: 100%; height: 3rem; font-weight: bold; }
    textarea { font-size: 16px !important; }
    </style>
""", unsafe_allow_html=True)

def exigir_autenticacao():
    try:
        credenciais = st.secrets["auth"]
        usuario_configurado = credenciais["username"]
        senha_configurada = credenciais["password"]
    except (StreamlitSecretNotFoundError, KeyError):
        st.error("Configure auth.username e auth.password nas secrets do Streamlit.")
        st.stop()

    if not st.session_state.get("autenticado", False):
        with st.form("formulario_login"):
            st.subheader("Acesso ao App Carro-Pipa")
            usuario = st.text_input("Usuário")
            senha = st.text_input("Senha", type="password")
            enviado = st.form_submit_button("Entrar")

        if enviado:
            usuario_valido = hmac.compare_digest(usuario, str(usuario_configurado))
            senha_valida = hmac.compare_digest(senha, str(senha_configurada))
            if usuario_valido and senha_valida:
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos.")

        st.stop()

    with st.sidebar:
        if st.button("Sair"):
            st.session_state.autenticado = False
            st.rerun()


exigir_autenticacao()

def obter_ip_local():
    conexao = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        conexao.connect(("8.8.8.8", 80))
        return conexao.getsockname()[0]
    except OSError:
        enderecos = socket.gethostbyname_ex(socket.gethostname())[2]
        return next((ip for ip in enderecos if not ip.startswith("127.")), None)
    finally:
        conexao.close()

# URL export CSV da sua Planilha do Google
DADOS_LOCAIS_PATH = Path(__file__).with_name("dados_local.csv")
BASE_PRESTADORES_PATH = Path(__file__).with_name("base_prestadores.csv")
COLUNAS_BASE = ["Placa", "Motorista", "Nr Contrato", "Município", "Lote", "Empresa", "Ocorrência"]
COLUNAS_BASE_OBRIGATORIAS = COLUNAS_BASE[:-1]
COLUNAS_HISTORICO = [
    "id",
    "data",
    "placa",
    "assunto",
    "relator",
    "relato",
    "motorista",
    "contrato",
    "municipio",
    "lote",
    "empresa",
    "status",
]
OCORRENCIAS_PADRAO = [
    "Abastecimento fora do horário permitido",
    "Ausência do módulo rastreador/embarcado",
    "Desvio de rota da carrada",
    "Cobrança indevida de água ao morador",
]

# ==========================================
# CARREGAMENTO DE DADOS (ONLINE / OFFLINE)
# ==========================================
def filtrar_colunas_base(df):
    ausentes = [coluna for coluna in COLUNAS_BASE_OBRIGATORIAS if coluna not in df.columns]
    if ausentes:
        raise ValueError(f"A planilha não contém as colunas necessárias: {', '.join(ausentes)}")
    colunas_disponiveis = [coluna for coluna in COLUNAS_BASE if coluna in df.columns]
    return df.loc[:, colunas_disponiveis].copy()


def usar_google_sheets():
    try:
        backend = st.secrets["storage"]["backend"]
    except (StreamlitSecretNotFoundError, KeyError):
        return False

    if backend not in {"local", "google_sheets"}:
        st.error("storage.backend deve ser 'local' ou 'google_sheets'.")
        st.stop()
    return backend == "google_sheets"


@st.cache_data(ttl=600)
def buscar_dados_online():
    try:
        url_planilha = st.secrets["data"]["sheet_url"]
    except (StreamlitSecretNotFoundError, KeyError) as erro:
        raise ValueError("Configure data.sheet_url nas secrets do Streamlit.") from erro

    with urlopen(url_planilha, timeout=15) as resposta:
        conteudo = resposta.read()
    df = filtrar_colunas_base(pd.read_csv(io.BytesIO(conteudo)))
    if df.empty:
        raise ValueError("A planilha não contém registros.")
    return df


def carregar_dados():
    if not usar_google_sheets():
        if BASE_PRESTADORES_PATH.exists():
            try:
                df = filtrar_colunas_base(
                    pd.read_csv(BASE_PRESTADORES_PATH, dtype=str).fillna("")
                )
                if not df.empty:
                    return df, "Offline (base local)"
            except (OSError, UnicodeDecodeError, ValueError, pd.errors.EmptyDataError, pd.errors.ParserError) as erro:
                st.error(f"Não foi possível carregar a cópia local da planilha: {erro}")
        return pd.DataFrame(columns=COLUNAS_BASE), "Offline (sem cópia local da base)"

    try:
        df = buscar_dados_online()
    except (OSError, ValueError, pd.errors.ParserError) as erro_online:
        if BASE_PRESTADORES_PATH.exists():
            try:
                df = filtrar_colunas_base(
                    pd.read_csv(BASE_PRESTADORES_PATH, dtype=str).fillna("")
                )
                if df.empty:
                    raise pd.errors.EmptyDataError("A cópia local não contém registros.")
                df.to_csv(BASE_PRESTADORES_PATH, index=False, encoding="utf-8-sig")
                return df, f"Offline (cópia local salva; falha ao acessar a planilha: {erro_online})"
            except (OSError, UnicodeDecodeError, ValueError, pd.errors.EmptyDataError, pd.errors.ParserError) as erro_local:
                st.error(f"Não foi possível carregar a cópia local da planilha: {erro_local}")

        return pd.DataFrame(), f"Offline (sem cópia válida da planilha; falha ao acessar a planilha: {erro_online})"

    try:
        df.to_csv(BASE_PRESTADORES_PATH, index=False, encoding="utf-8-sig")
    except OSError as erro:
        st.error(f"Não foi possível salvar a cópia local da planilha: {erro}")
    return df, "Online (Google Sheets; cópia local atualizada)"

if "dados_planilha" not in st.session_state:
    dados_iniciais, status_inicial = carregar_dados()
    st.session_state.dados_planilha = dados_iniciais
    st.session_state.status_base = status_inicial

with st.sidebar:
    st.header("Acesso pelo celular")
    url_publica = st.secrets.get("app", {}).get("public_url")
    ip_local = obter_ip_local() if not url_publica else None
    if url_publica:
        link_celular = url_publica.rstrip("/")
        st.code(link_celular)
        st.image(qrcode.make(link_celular).get_image(), caption="Acesse pelo link público")
    elif ip_local:
        link_celular = f"http://{ip_local}:8501"
        st.code(link_celular)
        st.image(qrcode.make(link_celular).get_image(), caption="Escaneie na mesma rede Wi-Fi")
    else:
        st.warning("Não foi possível detectar o endereço da rede local.")

    st.header("Base de dados")
    st.caption(
        "Para trabalhar sem Google Sheets, defina storage.backend como 'local' "
        "nas secrets. A base e os relatos serão guardados em arquivos CSV neste servidor."
    )
    if usar_google_sheets() and st.button("Atualizar base da planilha"):
        buscar_dados_online.clear()
        try:
            df_atualizado = buscar_dados_online()
            df_atualizado.to_csv(BASE_PRESTADORES_PATH, index=False, encoding="utf-8-sig")
        except (OSError, ValueError, pd.errors.ParserError) as erro:
            st.error(f"Não foi possível atualizar a base: {erro}")
        else:
            st.session_state.dados_planilha = df_atualizado
            st.session_state.status_base = "Online (Google Sheets; cópia local atualizada)"
            st.rerun()

    df_dados = st.session_state.dados_planilha
    status_conexao = st.session_state.status_base
    csv_para_baixar = df_dados.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Baixar base CSV para o celular",
        data=csv_para_baixar,
        file_name="base_prestadores.csv",
        mime="text/csv",
        disabled=df_dados.empty,
        use_container_width=True,
    )
    arquivo_base_local = st.file_uploader(
        "Importar CSV da base para uso offline",
        type=["csv"],
        key="arquivo_base_local",
    )
    if st.button("Salvar CSV como base offline"):
        if arquivo_base_local is None:
            st.error("Selecione primeiro o arquivo base_prestadores.csv.")
        else:
            try:
                df_local = filtrar_colunas_base(
                    pd.read_csv(arquivo_base_local, dtype=str).fillna("")
                )
                if df_local.empty:
                    raise ValueError("O arquivo CSV não contém registros.")
                df_local.to_csv(BASE_PRESTADORES_PATH, index=False, encoding="utf-8-sig")
            except (OSError, UnicodeDecodeError, ValueError, pd.errors.EmptyDataError, pd.errors.ParserError) as erro:
                st.error(f"Não foi possível salvar a base offline: {erro}")
            else:
                st.session_state.dados_planilha = df_local
                st.session_state.status_base = "Offline (base local importada)"
                st.success("Base CSV salva para uso offline.")
                st.rerun()

def carregar_historico():
    if usar_google_sheets():
        try:
            linhas = obter_aba_relatorios().get_all_values()
            historico = [
                dict(zip(COLUNAS_HISTORICO, linha))
                for linha in linhas[1:]
                if linha
            ]
            ids_remotos = {item.get("id") for item in historico}
            historico.extend(
                item
                for item in carregar_historico_local()
                if item.get("id") and item.get("id") not in ids_remotos
            )
            try:
                salvar_historico_local(historico)
            except OSError as erro:
                st.warning(f"Não foi possível atualizar a cópia local dos relatos: {erro}")
            return historico
        except (
            ValueError,
            gspread.exceptions.GSpreadException,
            requests.exceptions.RequestException,
            GoogleAuthError,
        ) as erro:
            st.warning(
                "Não foi possível acessar os relatos no Google Sheets. "
                f"Usando o histórico local: {erro}"
            )

    return carregar_historico_local()


def carregar_historico_local():
    if not DADOS_LOCAIS_PATH.exists():
        return []
    try:
        return pd.read_csv(DADOS_LOCAIS_PATH, dtype=str).fillna("").to_dict("records")
    except (OSError, UnicodeDecodeError, pd.errors.EmptyDataError, pd.errors.ParserError) as e:
        st.error(f"Não foi possível carregar o histórico local: {e}")
        return []


def salvar_historico_local(historico):
    pd.DataFrame(historico, columns=COLUNAS_HISTORICO).to_csv(
        DADOS_LOCAIS_PATH, index=False, encoding="utf-8-sig"
    )


@st.cache_resource
def obter_aba_relatorios():
    try:
        configuracao_storage = st.secrets["storage"]
        service_account = dict(st.secrets["gcp_service_account"])
        spreadsheet_id = configuracao_storage["spreadsheet_id"]
        worksheet_title = configuracao_storage.get("worksheet", "Registro")
        chave_privada = service_account.get("private_key")
        if isinstance(chave_privada, str):
            service_account["private_key"] = (
                chave_privada.replace("\\n", "\n").replace("\r\n", "\n").strip()
            )
        client = gspread.service_account_from_dict(service_account)
        worksheet = client.open_by_key(spreadsheet_id).worksheet(worksheet_title)
    except (StreamlitSecretNotFoundError, KeyError) as erro:
        raise ValueError(
            "Configure storage.spreadsheet_id e gcp_service_account nas secrets."
        ) from erro

    cabecalho = worksheet.row_values(1)
    if not cabecalho:
        worksheet.append_row(COLUNAS_HISTORICO, value_input_option="RAW")
    elif cabecalho != COLUNAS_HISTORICO:
        raise ValueError(
            f"A primeira linha da aba '{worksheet_title}' deve conter os cabeçalhos do app."
        )
    return worksheet


def sincronizar_historico_local():
    worksheet = obter_aba_relatorios()
    historico_local = carregar_historico_local()
    linhas_remotas = worksheet.get_all_values()
    indice_id = COLUNAS_HISTORICO.index("id")
    ids_remotos = {
        linha[indice_id]
        for linha in linhas_remotas[1:]
        if len(linha) > indice_id and linha[indice_id]
    }
    pendentes = [
        item for item in historico_local
        if item.get("id") and item["id"] not in ids_remotos
    ]
    if pendentes:
        worksheet.append_rows(
            [
                [item.get(coluna, "") for coluna in COLUNAS_HISTORICO]
                for item in pendentes
            ],
            value_input_option="RAW",
        )
    return len(pendentes)


# Inicialização do histórico de ocorrências
if "historico_ocorrencias" not in st.session_state:
    st.session_state.historico_ocorrencias = carregar_historico()

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================
st.title("📋 Relato de Ocorrência")
st.caption(f"Status da Base de Dados: **{status_conexao}**")
if status_conexao.startswith("Offline"):
    st.warning(status_conexao)

# Navegação por Abas
aba_novo, aba_historico = st.tabs(["✍️ Novo Relato", "🔍 Histórico / Consultas"])

with aba_novo:
    st.subheader("Dados da Fiscalização")
    
    # Campo 00: Data
    campo_data = st.date_input("Campo 00: Data da Ocorrência", datetime.now())

    # Campo 01: Assunto da Ocorrência
    ocorrencias_planilha = (
        df_dados["Ocorrência"].dropna().astype(str).unique().tolist()
        if "Ocorrência" in df_dados.columns
        else []
    )
    lista_ocorrencias = list(dict.fromkeys(OCORRENCIAS_PADRAO + ocorrencias_planilha))
    campo_assunto = st.selectbox(
        "Campo 01: Assunto da Ocorrência (Selecione ou digite)",
        options=["-- Digitar novo assunto --"] + lista_ocorrencias
    )
    if campo_assunto == "-- Digitar novo assunto --":
        campo_assunto = st.text_input("Digite o assunto personalizado:")

    st.markdown("---")
    st.markdown("### Busca do Prestador (Ênfase na Placa)")
    
    # Busca com Ênfase na Placa
    placas = df_dados["Placa"].dropna().astype(str).unique().tolist() if "Placa" in df_dados.columns else []
    opcoes_placa = [""] + placas
    placa_selecionada = st.selectbox("Campo 02: Placa do Veículo (Busca Principal)", opcoes_placa)
    
    # Autofill dos dados baseados na Placa
    motorista_val, contrato_val, mun_val, lote_val, empresa_val = "", "", "", "", ""
    if placa_selecionada:
        registro = df_dados[df_dados["Placa"].astype(str) == placa_selecionada].iloc[0]

        def valor_texto(campo):
            valor = registro.get(campo, "")
            return "" if pd.isna(valor) else str(valor)

        motorista_val = valor_texto("Motorista")
        contrato_val = valor_texto("Nr Contrato")
        mun_val = valor_texto("Município")
        lote_val = valor_texto("Lote")
        empresa_val = valor_texto("Empresa")

    col1, col2 = st.columns(2)
    with col1:
        sufixo_chave = placa_selecionada or "sem_placa"
        campo_motorista = st.text_input("Campo 03: Motorista", value=motorista_val, key=f"motorista_{sufixo_chave}")
        campo_contrato = st.text_input("Campo 04: Nº Contrato", value=contrato_val, key=f"contrato_{sufixo_chave}")
        campo_municipio = st.text_input("Campo 05: Município", value=mun_val, key=f"municipio_{sufixo_chave}")
    with col2:
        campo_lote = st.text_input("Campo 06: Lote", value=lote_val, key=f"lote_{sufixo_chave}")
        campo_empresa = st.text_input("Campo 07: Empresa", value=empresa_val, key=f"empresa_{sufixo_chave}")

    st.markdown("---")
    # Campo 08: Relato da Ocorrência (Caixa Longa)
    st.markdown("### Campo 08: Relato da Ocorrência")
    st.caption("Use o ditado por voz do teclado do smartphone se preferir. Salve o relato para mantê-lo no arquivo local.")
    
    campo_relato = st.text_area(
        "Descrição detalhada dos fatos:",
        height=200,
        max_chars=4000,
        placeholder="Descreva a ocorrência detalhadamente..."
    )
    st.caption(f"Caracteres digitados: {len(campo_relato)} / 4.000")

    # Campo 09: Evidências
    st.markdown("---")
    st.markdown("### Campo 09: Evidências (Fotos / Anexos)")
    fotos_anexadas = st.file_uploader(
        "Anexar fotos da ocorrência",
        type=["jpg", "png", "jpeg"],
        accept_multiple_files=True
    )
    st.caption("Fotos e assinatura são incluídas no PDF baixado; o CSV local guarda os dados textuais.")
    assinatura_anexada = st.file_uploader(
        "Anexar imagem da assinatura do prestador (opcional)",
        type=["jpg", "jpeg", "png"],
        key="assinatura_prestador",
    )

    # Campos 10 e 11: Identificação do Relator
    st.markdown("---")
    st.markdown("### Identificação do Relator")
    
    postos_grad = ["Cel", "Ten Cel", "Maj", "Cap", "1º Ten", "2º Ten", "Asp Of", "ST", "1º Sgt", "2º Sgt", "3º Sgt", "Cb", "Sd"]
    col_p, col_r = st.columns([1, 2])
    with col_p:
        campo_posto = st.selectbox("Campo 11: Posto/Grad", postos_grad)
    with col_r:
        campo_relator = st.text_input("Campo 10: Relator da Ocorrência (Nome Completo)")

    # ==========================================
    # ROTINA DE GERAÇÃO DO PDF (PADRÃO A4)
    # ==========================================
    def gerar_pdf():
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        elements = []
        styles = getSampleStyleSheet()

        # Estilos customizados
        style_header = ParagraphStyle('HeaderStyle', parent=styles['Normal'], alignment=1, fontSize=9, leading=11, fontName='Helvetica-Bold')
        style_title = ParagraphStyle('TitleStyle', parent=styles['Normal'], alignment=1, fontSize=11, leading=14, fontName='Helvetica-Bold')
        style_body = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=9, leading=12, fontName='Helvetica')
        style_bold = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontSize=9, leading=12, fontName='Helvetica-Bold')

        # Cabeçalho Oficial
        header_text = """
        MINISTÉRIO DA DEFESA<br/>
        EXÉRCITO BRASILEIRO<br/>
        COMANDO DA 10ª REGIÃO MILITAR<br/>
        (REGIÃO MARTIM SOARES MORENO)<br/>
        ESCRITÓRIO REGIONAL DA OPERAÇÃO CARRO-PIPA DA 10ª REGIÃO MILITAR
        """
        elements.append(Paragraph(header_text, style_header))
        elements.append(Spacer(1, 15))

        # Título
        elements.append(Paragraph("RELATO DE OCORRÊNCIA NA FISCALIZAÇÃO OPERACIONAL", style_title))
        elements.append(Spacer(1, 15))

        # Tabela de Dados Gerais
        dados_tabela = [
            [Paragraph("<b>Data:</b> " + escape(campo_data.strftime("%d/%m/%Y")), style_body), Paragraph("<b>Assunto:</b> " + escape(campo_assunto), style_body)],
            [Paragraph("<b>Placa:</b> " + escape(placa_selecionada), style_body), Paragraph("<b>Motorista:</b> " + escape(campo_motorista), style_body)],
            [Paragraph("<b>Contrato:</b> " + escape(campo_contrato), style_body), Paragraph("<b>Empresa:</b> " + escape(campo_empresa), style_body)],
            [Paragraph("<b>Município:</b> " + escape(campo_municipio), style_body), Paragraph("<b>Lote:</b> " + escape(campo_lote), style_body)]
        ]
        t = Table(dados_tabela, colWidths=[260, 260])
        t.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 15))

        # Bloco do Relato
        elements.append(Paragraph("<b>RELATO DA OCORRÊNCIA:</b>", style_bold))
        elements.append(Spacer(1, 5))
        elements.append(Paragraph(escape(campo_relato).replace("\n", "<br/>"), style_body))
        elements.append(Spacer(1, 15))

        # Bloco de Anexos / Fotos
        if fotos_anexadas:
            elements.append(Paragraph("<b>EVIDÊNCIAS FOTOGRÁFICAS:</b>", style_bold))
            elements.append(Spacer(1, 5))
            for foto in fotos_anexadas:
                foto.seek(0)
                img = Image.open(foto)
                img.thumbnail((400, 250))
                img_buffer = io.BytesIO()
                img.save(img_buffer, format="PNG")
                img_buffer.seek(0)
                elements.append(RLImage(img_buffer))
                elements.append(Spacer(1, 10))

        # Bloco de Assinaturas
        elements.append(Spacer(1, 20))
        
        # Preparar imagem da assinatura, se anexada.
        img_assinatura = Paragraph("_______________________________<br/>Assinatura do Prestador", style_body)
        if assinatura_anexada is not None:
            assinatura_anexada.seek(0)
            img_pad = Image.open(assinatura_anexada)
            img_pad.thumbnail((400, 160))
            pad_buffer = io.BytesIO()
            img_pad.convert("RGB").save(pad_buffer, format="PNG")
            pad_buffer.seek(0)
            img_assinatura = RLImage(pad_buffer, width=150, height=60)

        texto_govbr = Paragraph(
            f"_______________________________<br/><b>{escape(campo_posto)} {escape(campo_relator)}</b><br/>Espaço reservado p/ Assinatura GOV.BR",
            style_body,
        )

        t_assinaturas = Table([[img_assinatura, texto_govbr]], colWidths=[260, 260])
        t_assinaturas.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
        ]))
        elements.append(t_assinaturas)

        doc.build(elements)
        buffer.seek(0)
        return buffer

    # Botões de Ação
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("💾 Salvar Ocorrência"):
            nova_oc = {
                "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                "data": campo_data.strftime("%Y-%m-%d"),
                "placa": placa_selecionada,
                "assunto": campo_assunto,
                "relator": f"{campo_posto} {campo_relator}",
                "relato": campo_relato,
                "motorista": campo_motorista,
                "contrato": campo_contrato,
                "municipio": campo_municipio,
                "lote": campo_lote,
                "empresa": campo_empresa,
                "status": "Rascunho",
            }
            try:
                if usar_google_sheets():
                    try:
                        obter_aba_relatorios().append_row(
                            [nova_oc.get(coluna, "") for coluna in COLUNAS_HISTORICO],
                            value_input_option="RAW",
                        )
                    except (
                        OSError,
                        ValueError,
                        gspread.exceptions.GSpreadException,
                        requests.exceptions.RequestException,
                        GoogleAuthError,
                    ) as erro_online:
                        historico_local = carregar_historico_local()
                        historico_local.append(nova_oc)
                        salvar_historico_local(historico_local)
                        st.session_state.historico_ocorrencias.append(nova_oc)
                        st.warning(
                            "Google Sheets indisponível. Ocorrência salva neste dispositivo "
                            f"para uso offline: {erro_online}"
                        )
                    else:
                        historico_local = carregar_historico_local()
                        if not any(item.get("id") == nova_oc["id"] for item in historico_local):
                            historico_local.append(nova_oc)
                            try:
                                salvar_historico_local(historico_local)
                            except OSError as erro_local:
                                st.warning(f"Relato enviado ao Google Sheets, mas não foi possível salvar a cópia local: {erro_local}")
                        st.session_state.historico_ocorrencias.append(nova_oc)
                        st.success("Ocorrência salva no Google Sheets.")
                else:
                    historico_atualizado = carregar_historico_local()
                    historico_atualizado.append(nova_oc)
                    salvar_historico_local(historico_atualizado)
                    st.session_state.historico_ocorrencias.append(nova_oc)
                    st.success("Ocorrência salva localmente.")
            except (
                OSError,
                ValueError,
                gspread.exceptions.GSpreadException,
                requests.exceptions.RequestException,
                GoogleAuthError,
            ) as e:
                st.error(f"Não foi possível salvar a ocorrência: {e}")

    with col_btn2:
        pdf_bytes = gerar_pdf()
        st.download_button(
            label="🖨️ Imprimir em PDF (A4)",
            data=pdf_bytes,
            file_name=f"Relato_{placa_selecionada}_{campo_data.strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )

# Aba de Histórico e Consultas
with aba_historico:
    st.subheader("Consultar Ocorrências Salvas")
    if usar_google_sheets() and st.button("Sincronizar relatos salvos offline"):
        try:
            quantidade = sincronizar_historico_local()
        except (
            OSError,
            ValueError,
            gspread.exceptions.GSpreadException,
            requests.exceptions.RequestException,
            GoogleAuthError,
        ) as erro:
            st.error(f"Não foi possível sincronizar os relatos locais: {erro}")
        else:
            st.success(f"Sincronização concluída: {quantidade} relato(s) enviado(s).")
            st.session_state.pop("historico_ocorrencias", None)
            st.rerun()

    data_filtro = st.date_input("Filtrar por data:", date.today())
    
    if st.session_state.historico_ocorrencias:
        df_hist = pd.DataFrame(st.session_state.historico_ocorrencias)
        if "data" in df_hist.columns:
            df_hist = df_hist[df_hist["data"].astype(str).str.startswith(data_filtro.isoformat())]
        st.dataframe(df_hist)
        if df_hist.empty:
            st.info("Nenhuma ocorrência encontrada para esta data.")
    else:
        st.info("Nenhuma ocorrência salva neste arquivo local.")
