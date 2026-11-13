import streamlit as st
import pandas as pd
from datetime import date, datetime
import hmac
import io
import re
from pathlib import Path
from PIL import Image, ImageOps
import gspread
import numpy as np
from xml.sax.saxutils import escape
from streamlit.errors import StreamlitSecretNotFoundError
from streamlit_drawable_canvas import st_canvas

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

# Caminhos locais e constantes
DADOS_LOCAIS_PATH = Path(__file__).with_name("dados_local.csv")
BASE_PRESTADORES_PATH = Path(__file__).with_name("base_prestadores.csv")
OCORRENCIAS_LOCAIS_PATH = Path(__file__).with_name("ocorrencias_local.csv")

DEFAULT_SPREADSHEET_ID = "1PRy0XlQliIHXvM9Ou2HwmHwCrgq3iPpEWPY4XwxJ5qI"
COLUNAS_BASE = ["Placa", "Motorista", "Nr Contrato", "Município", "Lote", "Empresa"]
COLUNAS_BASE_OBRIGATORIAS = COLUNAS_BASE[:-1]
COLUNAS_HISTORICO = [
    "id",
    "numero_relato",
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

# ==========================================
# AUTENTICAÇÃO GOOGLE SHEETS (RECONSTRUÇÃO PEM)
# ==========================================
def reconstruir_pem_seguro(chave_bruta: str) -> str:
    """Extrai o payload Base64 puro e remonta o arquivo PEM estritamente a cada 64 caracteres."""
    texto = chave_bruta.replace("\\n", "\n").replace("\r", "")
    # Remove cabeçalhos e qualquer espaço ou quebra de linha
    corpo = re.sub(r"-----BEGIN [^-]+-----", "", texto)
    corpo = re.sub(r"-----END [^-]+-----", "", corpo)
    corpo = re.sub(r"\s+", "", corpo)
    # Quebra estritamente em blocos de 64 caracteres conforme padrão RFC 7468
    linhas = [corpo[i:i+64] for i in range(0, len(corpo), 64)]
    return "-----BEGIN PRIVATE KEY-----\n" + "\n".join(linhas) + "\n-----END PRIVATE KEY-----\n"

@st.cache_resource
def obter_cliente_gsheets():
    try:
        service_account = dict(st.secrets["gcp_service_account"])
        chave_privada = service_account.get("private_key", "")
        if isinstance(chave_privada, str) and chave_privada.strip():
            service_account["private_key"] = reconstruir_pem_seguro(chave_privada)
        return gspread.service_account_from_dict(service_account)
    except (StreamlitSecretNotFoundError, KeyError) as erro:
        raise ValueError("Configure gcp_service_account nas secrets.") from erro

def obter_spreadsheet_id():
    return st.secrets.get("storage", {}).get("spreadsheet_id", DEFAULT_SPREADSHEET_ID)

@st.cache_resource
def obter_aba_relatorios():
    spreadsheet_id = obter_spreadsheet_id()
    worksheet_title = st.secrets.get("storage", {}).get("worksheet", "Registro")
    client = obter_cliente_gsheets()
    planilha = client.open_by_key(spreadsheet_id)
    try:
        worksheet = planilha.worksheet(worksheet_title)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = planilha.add_worksheet(title=worksheet_title, rows=2000, cols=15)

    cabecalho = worksheet.row_values(1)
    if not cabecalho:
        worksheet.append_row(COLUNAS_HISTORICO, value_input_option="RAW")
    elif cabecalho != COLUNAS_HISTORICO:
        # Atualiza cabeçalho se estiver no formato antigo sem 'numero_relato'
        if "numero_relato" not in cabecalho and len(cabecalho) == len(COLUNAS_HISTORICO) - 1:
            worksheet.insert_cols([["numero_relato"]], col=2)
        else:
            raise ValueError(
                f"A primeira linha da aba '{worksheet_title}' deve conter: {', '.join(COLUNAS_HISTORICO)}"
            )
    return worksheet

# ==========================================
# NUMERAÇÃO AUTOMÁTICA SEQUENCIAL (Nº xxxx/AAAA)
# ==========================================
def calcular_proximo_numero(worksheet, ano_referencia: int) -> str:
    """Lê a coluna numero_relato diretamente da planilha para evitar colisão entre usuários."""
    coluna_numeros = worksheet.col_values(2)[1:]  # Coluna B = numero_relato
    maior_seq = 0
    sufixo_ano = f"/{ano_referencia}"
    for val in coluna_numeros:
        val_str = str(val).strip()
        if val_str.endswith(sufixo_ano):
            parte_num = val_str.replace("Nº", "").replace(sufixo_ano, "").strip()
            if parte_num.isdigit():
                maior_seq = max(maior_seq, int(parte_num))
    return f"Nº {maior_seq + 1:04d}/{ano_referencia}"

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
    if st.session_state.get("forcar_offline", False):
        return False
    backend = st.secrets.get("storage", {}).get("backend", "google_sheets")
    return backend == "google_sheets"

@st.cache_data(ttl=300)
def buscar_dados_online():
    spreadsheet_id = obter_spreadsheet_id()
    client = obter_cliente_gsheets()
    planilha = client.open_by_key(spreadsheet_id)
    
    # 1. Carrega a base de prestadores
    aba_base = st.secrets.get("storage", {}).get("base_worksheet", "Base")
    try:
        ws_base = planilha.worksheet(aba_base)
    except gspread.exceptions.WorksheetNotFound:
        abas = [w for w in planilha.worksheets() if w.title not in ["Registro", "Ocorrência"]]
        ws_base = abas[0] if abas else planilha.worksheets()[0]
    
    registros = ws_base.get_all_records()
    df_base = filtrar_colunas_base(pd.DataFrame(registros).astype(str)) if registros else pd.DataFrame(columns=COLUNAS_BASE)

    # 2. Carrega todas as ocorrências da Coluna A da aba "Ocorrência"
    lista_ocorrencias = []
    try:
        ws_ocorrencia = planilha.worksheet("Ocorrência")
        coluna_a = ws_ocorrencia.col_values(1)
        if coluna_a:
            lista_ocorrencias = [
                x.strip() for x in coluna_a 
                if x and x.strip().lower() not in ("ocorrência", "ocorrencia")
            ]
    except gspread.exceptions.WorksheetNotFound:
        pass

    return df_base, lista_ocorrencias

def carregar_dados():
    if not usar_google_sheets():
        df_base = pd.DataFrame(columns=COLUNAS_BASE)
        if BASE_PRESTADORES_PATH.exists():
            try:
                df_base = filtrar_colunas_base(pd.read_csv(BASE_PRESTADORES_PATH, dtype=str).fillna(""))
            except Exception:
                pass
        ocorrencias = []
        if OCORRENCIAS_LOCAIS_PATH.exists():
            try:
                df_oc = pd.read_csv(OCORRENCIAS_LOCAIS_PATH, dtype=str)
                ocorrencias = df_oc.iloc[:, 0].dropna().astype(str).tolist()
            except Exception:
                pass
        return df_base, ocorrencias, "Offline (base local carregada)"

    try:
        df_base, ocorrencias = buscar_dados_online()
    except Exception as erro_online:
        df_base = pd.DataFrame(columns=COLUNAS_BASE)
        if BASE_PRESTADORES_PATH.exists():
            try:
                df_base = filtrar_colunas_base(pd.read_csv(BASE_PRESTADORES_PATH, dtype=str).fillna(""))
            except Exception:
                pass
        ocorrencias = []
        if OCORRENCIAS_LOCAIS_PATH.exists():
            try:
                df_oc = pd.read_csv(OCORRENCIAS_LOCAIS_PATH, dtype=str)
                ocorrencias = df_oc.iloc[:, 0].dropna().astype(str).tolist()
            except Exception:
                pass
        return df_base, ocorrencias, f"Offline (usando cópia local; falha online: {erro_online})"

    try:
        df_base.to_csv(BASE_PRESTADORES_PATH, index=False, encoding="utf-8-sig")
        pd.DataFrame({"Ocorrência": ocorrencias}).to_csv(OCORRENCIAS_LOCAIS_PATH, index=False, encoding="utf-8-sig")
    except Exception:
        pass
    return df_base, ocorrencias, "Online (Google Sheets; cópia local atualizada)"

if "dados_planilha" not in st.session_state:
    df_b, oc_l, status_inicial = carregar_dados()
    st.session_state.dados_planilha = df_b
    st.session_state.lista_ocorrencias = oc_l
    st.session_state.status_base = status_inicial

with st.sidebar:
    st.header("Modo de Operação")
    modo_offline_manual = st.toggle(
        "Trabalhar Offline (Local)",
        value=st.session_state.get("forcar_offline", False),
        help="Ative em áreas sem internet para salvar localmente e sincronizar depois."
    )
    if modo_offline_manual != st.session_state.get("forcar_offline", False):
        st.session_state.forcar_offline = modo_offline_manual
        df_b, oc_l, status_inicial = carregar_dados()
        st.session_state.dados_planilha = df_b
        st.session_state.lista_ocorrencias = oc_l
        st.session_state.status_base = status_inicial
        st.rerun()

    st.header("Base de dados")
    if usar_google_sheets() and st.button("Atualizar base da planilha"):
        buscar_dados_online.clear()
        try:
            df_atualizado, oc_atualizado = buscar_dados_online()
            df_atualizado.to_csv(BASE_PRESTADORES_PATH, index=False, encoding="utf-8-sig")
            pd.DataFrame({"Ocorrência": oc_atualizado}).to_csv(OCORRENCIAS_LOCAIS_PATH, index=False, encoding="utf-8-sig")
        except Exception as erro:
            st.error(f"Não foi possível atualizar a base: {erro}")
        else:
            st.session_state.dados_planilha = df_atualizado
            st.session_state.lista_ocorrencias = oc_atualizado
            st.session_state.status_base = "Online (Google Sheets; cópia local atualizada)"
            st.rerun()

    df_dados = st.session_state.dados_planilha
    status_conexao = st.session_state.status_base
    csv_para_baixar = df_dados.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Baixar base CSV para backup",
        data=csv_para_baixar,
        file_name="base_prestadores.csv",
        mime="text/csv",
        disabled=df_dados.empty,
        use_container_width=True,
    )

def carregar_historico_local():
    if not DADOS_LOCAIS_PATH.exists():
        return []
    try:
        return pd.read_csv(DADOS_LOCAIS_PATH, dtype=str).fillna("").to_dict("records")
    except Exception as e:
        st.error(f"Não foi possível carregar o histórico local: {e}")
        return []

def salvar_historico_local(historico):
    pd.DataFrame(historico, columns=COLUNAS_HISTORICO).to_csv(
        DADOS_LOCAIS_PATH, index=False, encoding="utf-8-sig"
    )

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
            except OSError:
                pass
            return historico
        except Exception as erro:
            st.warning(
                "Não foi possível acessar os relatos no Google Sheets. "
                f"Usando o histórico local: {erro}"
            )

    return carregar_historico_local()

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
        ano_atual = datetime.now().year
        linhas_para_enviar = []
        for item in pendentes:
            # Se foi gerado offline, atribui agora o número oficial sequencial
            num_atual = str(item.get("numero_relato", ""))
            if not num_atual or num_atual.startswith("OFF-"):
                num_oficial = calcular_proximo_numero(worksheet, ano_atual)
                item["numero_relato"] = num_oficial
            linhas_para_enviar.append([item.get(coluna, "") for coluna in COLUNAS_HISTORICO])
            # Envia linha por linha para garantir incrementação segura
            worksheet.append_row(linhas_para_enviar[-1], value_input_option="RAW")
        salvar_historico_local(historico_local)
    return len(pendentes)

if "historico_ocorrencias" not in st.session_state:
    st.session_state.historico_ocorrencias = carregar_historico()

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================
st.title("📋 Relato de Ocorrência")
st.caption(f"Status da Base de Dados: **{status_conexao}**")
if status_conexao.startswith("Offline"):
    st.info(status_conexao)

aba_novo, aba_historico = st.tabs(["✍️ Novo Relato", "🔍 Histórico / Consultas"])

with aba_novo:
    st.subheader("Dados da Fiscalização")
    
    # Exibe o número oficial atribuído caso o relato atual já tenha sido salvo
    num_exibicao = st.session_state.get("ultimo_numero_gerado", "Gerado automaticamente ao salvar")
    st.info(f"**Identificação do Documento:** {num_exibicao}")

    campo_data = st.date_input("Campo 00: Data da Ocorrência", datetime.now())

    # Campo 01: Assunto da Ocorrência (Puxa direto da coluna A da aba 'Ocorrência')
    lista_ocorrencias = st.session_state.get("lista_ocorrencias", [])
    if not lista_ocorrencias:
        lista_ocorrencias = ["Nenhuma ocorrência encontrada na aba 'Ocorrência'"]

    campo_assunto = st.selectbox(
        "Campo 01: Assunto da Ocorrência",
        options=lista_ocorrencias
    )

    st.markdown("---")
    st.markdown("### Busca do Prestador (Ênfase na Placa)")
    
    placas = df_dados["Placa"].dropna().astype(str).unique().tolist() if "Placa" in df_dados.columns else []
    if placas:
        opcoes_placa = ["", "-- Digitar placa manualmente --"] + placas
        placa_escolha = st.selectbox("Campo 02: Placa do Veículo (Busca Principal)", opcoes_placa)
        if placa_escolha == "-- Digitar placa manualmente --":
            placa_selecionada = st.text_input("Digite a Placa do Veículo (Ex: ABC1D23):").upper()
        else:
            placa_selecionada = placa_escolha
    else:
        placa_selecionada = st.text_input("Campo 02: Placa do Veículo (Ex: ABC1D23):").upper()
    
    motorista_val, contrato_val, mun_val, lote_val, empresa_val = "", "", "", "", ""
    if placa_selecionada and not df_dados.empty and "Placa" in df_dados.columns:
        filtro = df_dados[df_dados["Placa"].astype(str) == placa_selecionada]
        if not filtro.empty:
            registro = filtro.iloc[0]

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
    st.markdown("### Campo 08: Relato da Ocorrência")
    st.caption("Use o ditado por voz do teclado do smartphone se preferir.")
    
    campo_relato = st.text_area(
        "Descrição detalhada dos fatos:",
        height=200,
        max_chars=4000,
        placeholder="Descreva a ocorrência detalhadamente..."
    )
    st.caption(f"Caracteres digitados: {len(campo_relato)} / 4.000")

    st.markdown("---")
    st.markdown("### Campo 09: Evidências (Fotos / Anexos)")
    fotos_anexadas = st.file_uploader(
        "Anexar fotos da ocorrência",
        type=["jpg", "png", "jpeg"],
        accept_multiple_files=True
    )

    st.markdown("---")
    st.markdown("### Assinatura do Prestador (Toque e Assine na Tela)")
    st.caption("Utilize o dedo ou uma caneta stylus para assinar no quadro abaixo:")
    
    canvas_result = st_canvas(
        fill_color="rgba(255, 255, 255, 0)",
        stroke_width=3,
        stroke_color="#000000",
        background_color="#FFFFFF",
        height=150,
        width=340,
        drawing_mode="freedraw",
        key="canvas_assinatura",
    )

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
    def gerar_pdf(numero_documento: str):
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        elements = []
        styles = getSampleStyleSheet()

        style_header = ParagraphStyle('HeaderStyle', parent=styles['Normal'], alignment=1, fontSize=9, leading=11, fontName='Helvetica-Bold')
        style_title = ParagraphStyle('TitleStyle', parent=styles['Normal'], alignment=1, fontSize=11, leading=14, fontName='Helvetica-Bold')
        style_body = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=9, leading=12, fontName='Helvetica')
        style_bold = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontSize=9, leading=12, fontName='Helvetica-Bold')

        header_text = """
        MINISTÉRIO DA DEFESA<br/>
        EXÉRCITO BRASILEIRO<br/>
        COMANDO DA 10ª REGIÃO MILITAR<br/>
        (REGIÃO MARTIM SOARES MORENO)<br/>
        ESCRITÓRIO REGIONAL DA OPERAÇÃO CARRO-PIPA DA 10ª REGIÃO MILITAR
        """
        elements.append(Paragraph(header_text, style_header))
        elements.append(Spacer(1, 15))

        elements.append(Paragraph(f"RELATO DE OCORRÊNCIA NA FISCALIZAÇÃO OPERACIONAL — {escape(numero_documento)}", style_title))
        elements.append(Spacer(1, 15))

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

        elements.append(Paragraph("<b>RELATO DA OCORRÊNCIA:</b>", style_bold))
        elements.append(Spacer(1, 5))
        elements.append(Paragraph(escape(campo_relato).replace("\n", "<br/>"), style_body))
        elements.append(Spacer(1, 15))

        if fotos_anexadas:
            elements.append(Paragraph("<b>EVIDÊNCIAS FOTOGRÁFICAS:</b>", style_bold))
            elements.append(Spacer(1, 5))
            for foto in fotos_anexadas:
                foto.seek(0)
                img = Image.open(foto)
                img = ImageOps.exif_transpose(img)
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                img.thumbnail((400, 250))
                img_buffer = io.BytesIO()
                img.save(img_buffer, format="JPEG", quality=85)
                img_buffer.seek(0)
                elements.append(RLImage(img_buffer, width=img.width, height=img.height))
                elements.append(Spacer(1, 10))

        elements.append(Spacer(1, 20))
        
        img_assinatura = Paragraph("_______________________________<br/>Assinatura do Prestador", style_body)
        if canvas_result.image_data is not None:
            arr = canvas_result.image_data
            if arr.sum() > 0:
                img_assinatura_pil = Image.fromarray(arr.astype('uint8'), mode="RGBA")
                pad_buffer = io.BytesIO()
                img_assinatura_pil.convert("RGB").save(pad_buffer, format="PNG")
                pad_buffer.seek(0)
                img_assinatura = RLImage(pad_buffer, width=160, height=70)

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

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("💾 Salvar Ocorrência"):
            ano_oc = campo_data.year
            numero_gerado = f"OFF-{datetime.now().strftime('%H%M%S')}/{ano_oc}"
            
            nova_oc = {
                "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                "numero_relato": numero_gerado,
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
                "status": "Registrado",
            }
            try:
                if usar_google_sheets():
                    try:
                        ws = obter_aba_relatorios()
                        numero_gerado = calcular_proximo_numero(ws, ano_oc)
                        nova_oc["numero_relato"] = numero_gerado
                        ws.append_row(
                            [nova_oc.get(coluna, "") for coluna in COLUNAS_HISTORICO],
                            value_input_option="RAW",
                        )
                    except Exception as erro_online:
                        historico_local = carregar_historico_local()
                        historico_local.append(nova_oc)
                        salvar_historico_local(historico_local)
                        st.session_state.historico_ocorrencias.append(nova_oc)
                        st.session_state.ultimo_numero_gerado = numero_gerado
                        st.warning(
                            f"Google Sheets indisponível. Salvo localmente com protocolo provisório {numero_gerado}: {erro_online}"
                        )
                    else:
                        historico_local = carregar_historico_local()
                        if not any(item.get("id") == nova_oc["id"] for item in historico_local):
                            historico_local.append(nova_oc)
                            try:
                                salvar_historico_local(historico_local)
                            except OSError:
                                pass
                        st.session_state.historico_ocorrencias.append(nova_oc)
                        st.session_state.ultimo_numero_gerado = numero_gerado
                        st.success(f"Ocorrência salva com sucesso! Protocolo: **{numero_gerado}**")
                else:
                    historico_atualizado = carregar_historico_local()
                    historico_atualizado.append(nova_oc)
                    salvar_historico_local(historico_atualizado)
                    st.session_state.historico_ocorrencias.append(nova_oc)
                    st.session_state.ultimo_numero_gerado = numero_gerado
                    st.success(f"Ocorrência salva localmente (Offline). Protocolo: **{numero_gerado}**")
            except Exception as e:
                st.error(f"Não foi possível salvar a ocorrência: {e}")

    with col_btn2:
        numero_para_pdf = st.session_state.get("ultimo_numero_gerado", f"Nº 0000/{campo_data.year}")
        pdf_bytes = gerar_pdf(numero_para_pdf)
        nome_arquivo_num = numero_para_pdf.replace("Nº ", "").replace("/", "-")
        st.download_button(
            label="🖨️ Imprimir em PDF (A4)",
            data=pdf_bytes,
            file_name=f"Relato_{nome_arquivo_num}_{placa_selecionada}.pdf",
            mime="application/pdf"
        )

with aba_historico:
    st.subheader("Consultar Ocorrências Salvas")
    if usar_google_sheets() and st.button("Sincronizar relatos salvos offline"):
        try:
            quantidade = sincronizar_historico_local()
        except Exception as erro:
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
