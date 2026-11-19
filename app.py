import streamlit as st
import pandas as pd
from datetime import date, datetime
import hmac
import io
import json
import re
from PIL import Image, ImageOps
import gspread
from xml.sax.saxutils import escape
from streamlit.errors import StreamlitSecretNotFoundError

# Canvas opcional para assinatura na tela
try:
    from streamlit_drawable_canvas import st_canvas
    HAS_CANVAS = True
except ImportError:
    st_canvas = None
    HAS_CANVAS = False

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

# Constantes da Planilha (100% em memória, sem CSVs no GitHub)
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
    texto = chave_bruta.replace("\\n", "\n").replace("\r", "")
    corpo = re.sub(r"-----BEGIN [^-]+-----", "", texto)
    corpo = re.sub(r"-----END [^-]+-----", "", corpo)
    corpo = re.sub(r"\s+", "", corpo)
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
        if "numero_relato" not in cabecalho and len(cabecalho) == len(COLUNAS_HISTORICO) - 1:
            worksheet.insert_cols([["numero_relato"]], col=2)
    return worksheet

def calcular_proximo_numero(worksheet, ano_referencia: int) -> str:
    coluna_numeros = worksheet.col_values(2)[1:]
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
# CARREGAMENTO DE DADOS EM MEMÓRIA
# ==========================================
def filtrar_colunas_base(df):
    ausentes = [coluna for coluna in COLUNAS_BASE_OBRIGATORIAS if coluna not in df.columns]
    if ausentes:
        raise ValueError(f"A planilha não contém as colunas necessárias: {', '.join(ausentes)}")
    colunas_disponiveis = [coluna for coluna in COLUNAS_BASE if coluna in df.columns]
    return df.loc[:, colunas_disponiveis].copy()

@st.cache_data(ttl=300)
def buscar_dados_online():
    spreadsheet_id = obter_spreadsheet_id()
    client = obter_cliente_gsheets()
    planilha = client.open_by_key(spreadsheet_id)

    # 1. Aba Dados (Prestadores)
    try:
        ws_base = planilha.worksheet("Dados")
    except gspread.exceptions.WorksheetNotFound:
        ws_base = planilha.worksheets()[0]

    registros = ws_base.get_all_records()
    df_base = filtrar_colunas_base(pd.DataFrame(registros).astype(str)) if registros else pd.DataFrame(columns=COLUNAS_BASE)

    # 2. Aba Ocorrência (Coluna A)
    lista_ocorrencias = []
    try:
        ws_ocorrencia = planilha.worksheet("Ocorrência")
        coluna_a = ws_ocorrencia.col_values(1)
        if coluna_a:
            lista_ocorrencias = [
                str(x).strip() for x in coluna_a
                if x and str(x).strip() and str(x).strip().lower() not in ("ocorrência", "ocorrencia")
            ]
    except Exception:
        pass

    return df_base, lista_ocorrencias

def carregar_dados_memoria():
    if st.session_state.get("forcar_offline", False):
        return (
            st.session_state.get("dados_planilha", pd.DataFrame(columns=COLUNAS_BASE)),
            st.session_state.get("lista_ocorrencias", []),
            "Modo Offline Manual (dados mantidos na memória da sessão)"
        )
    try:
        df_base, ocorrencias = buscar_dados_online()
        return df_base, ocorrencias, "Online (Google Sheets ativo)"
    except Exception as erro_online:
        return (
            st.session_state.get("dados_planilha", pd.DataFrame(columns=COLUNAS_BASE)),
            st.session_state.get("lista_ocorrencias", []),
            f"Sem conexão com Google Sheets ({erro_online})"
        )

def carregar_historico_memoria():
    if st.session_state.get("forcar_offline", False):
        return st.session_state.get("historico_ocorrencias", [])
    try:
        linhas = obter_aba_relatorios().get_all_values()
        historico_remoto = [
            dict(zip(COLUNAS_HISTORICO, linha))
            for linha in linhas[1:]
            if linha
        ]
        ids_remotos = {item.get("id") for item in historico_remoto}
        pendentes_sessao = [
            item for item in st.session_state.get("historico_ocorrencias", [])
            if item.get("id") and item.get("id") not in ids_remotos
        ]
        return historico_remoto + pendentes_sessao
    except Exception as erro:
        st.warning(f"Não foi possível carregar o histórico do Google Sheets: {erro}")
        return st.session_state.get("historico_ocorrencias", [])

def sincronizar_pendentes_sessao():
    worksheet = obter_aba_relatorios()
    historico_sessao = st.session_state.get("historico_ocorrencias", [])
    linhas_remotas = worksheet.get_all_values()
    indice_id = COLUNAS_HISTORICO.index("id")
    ids_remotos = {
        linha[indice_id]
        for linha in linhas_remotas[1:]
        if len(linha) > indice_id and linha[indice_id]
    }
    pendentes = [
        item for item in historico_sessao
        if item.get("id") and item["id"] not in ids_remotos
    ]
    if pendentes:
        ano_atual = datetime.now().year
        for item in pendentes:
            num_atual = str(item.get("numero_relato", ""))
            if not num_atual or num_atual.startswith("OFF-"):
                item["numero_relato"] = calcular_proximo_numero(worksheet, ano_atual)
            linha_formatada = [item.get(coluna, "") for coluna in COLUNAS_HISTORICO]
            worksheet.append_row(linha_formatada, value_input_option="RAW")
    return len(pendentes)

# ==========================================
# GERADOR DE APP OFFLINE AUTÔNOMO (.HTML)
# ==========================================
def gerar_app_offline_html(df_base: pd.DataFrame, ocorrencias: list[str]) -> str:
    """Gera um arquivo HTML autônomo com os dados embutidos para uso 100% offline no celular."""
    base_registros = df_base.fillna("").to_dict(orient="records") if not df_base.empty else []
    base_json = json.dumps(base_registros, ensure_ascii=False)
    ocorrencias_json = json.dumps(ocorrencias, ensure_ascii=False)
    ano_atual = datetime.now().year

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Relato de Ocorrência Offline - EROCP/10ª RM</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{
      font-family: Arial, Helvetica, sans-serif;
      margin: 0;
      padding: 14px;
      background: #f4f6f8;
      color: #111;
    }}
    .container {{
      max-width: 780px;
      margin: 0 auto;
      background: #fff;
      padding: 18px;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }}
    .cabecalho-militar {{
      text-align: center;
      font-weight: bold;
      font-size: 12px;
      line-height: 1.35;
      margin-bottom: 12px;
    }}
    h1 {{
      text-align: center;
      font-size: 15px;
      margin: 8px 0 16px;
      text-transform: uppercase;
    }}
    .badge-offline {{
      background: #fff3cd;
      color: #664d03;
      border: 1px solid #ffecb5;
      padding: 8px 12px;
      border-radius: 6px;
      font-size: 13px;
      margin-bottom: 14px;
      text-align: center;
      font-weight: bold;
    }}
    .grid-2 {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
    }}
    @media (max-width: 600px) {{
      .grid-2 {{ grid-template-columns: 1fr; }}
    }}
    .campo {{ margin-bottom: 12px; }}
    label {{
      display: block;
      font-weight: bold;
      font-size: 13px;
      margin-bottom: 4px;
    }}
    input, select, textarea {{
      width: 100%;
      padding: 9px;
      font-size: 15px;
      border: 1px solid #bbb;
      border-radius: 5px;
      background: #fff;
    }}
    textarea {{ min-height: 140px; resize: vertical; }}
    canvas {{
      width: 100%;
      height: 150px;
      border: 1.5px solid #333;
      border-radius: 6px;
      background: #fff;
      touch-action: none;
    }}
    .acoes {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 16px;
    }}
    button {{
      flex: 1 1 180px;
      padding: 12px 14px;
      font-size: 14px;
      font-weight: bold;
      border: none;
      border-radius: 6px;
      cursor: pointer;
    }}
    .btn-imprimir {{ background: #198754; color: #fff; }}
    .btn-salvar {{ background: #0d6efd; color: #fff; }}
    .btn-limpar {{ background: #6c757d; color: #fff; }}
    .btn-csv {{ background: #212529; color: #fff; }}
    .fotos-preview {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 8px;
    }}
    .fotos-preview img {{
      max-width: 220px;
      max-height: 180px;
      border: 1px solid #ccc;
      border-radius: 4px;
    }}
    .assinaturas-print {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 20px;
      margin-top: 28px;
      text-align: center;
      font-size: 12px;
    }}
    .assinatura-img {{
      max-height: 65px;
      display: block;
      margin: 0 auto 4px;
    }}
    .historico-box {{
      margin-top: 24px;
      padding-top: 14px;
      border-top: 2px dashed #ccc;
    }}
    @media print {{
      body {{ background: #fff; padding: 0; }}
      .container {{ box-shadow: none; max-width: 100%; padding: 10px; }}
      .no-print, .badge-offline, .historico-box {{ display: none !important; }}
      input, select, textarea {{
        border: none;
        padding: 2px 0;
        font-size: 12px;
        appearance: none;
        -webkit-appearance: none;
      }}
    }}
  </style>
</head>
<body>
  <div class="container">
    <div class="badge-offline">
      MODO OFFLINE ATIVO — Funciona sem internet no celular (Preencher, Assinar e Imprimir PDF)
    </div>

    <div class="cabecalho-militar">
      MINISTÉRIO DA DEFESA<br/>
      EXÉRCITO BRASILEIRO<br/>
      COMANDO DA 10ª REGIÃO MILITAR<br/>
      (REGIÃO MARTIM SOARES MORENO)<br/>
      ESCRITÓRIO REGIONAL DA OPERAÇÃO CARRO-PIPA DA 10ª REGIÃO MILITAR
    </div>

    <h1>RELATO DE OCORRÊNCIA NA FISCALIZAÇÃO OPERACIONAL — <span id="num-relato-display">Nº 0001/{ano_atual}</span></h1>

    <div class="grid-2">
      <div class="campo">
        <label for="data_oc">Data da Ocorrência</label>
        <input type="date" id="data_oc" />
      </div>
      <div class="campo">
        <label for="num_oc">Número do Relato</label>
        <input type="text" id="num_oc" value="Nº 0001/{ano_atual}" oninput="document.getElementById('num-relato-display').textContent = this.value" />
      </div>
    </div>

    <div class="campo">
      <label for="assunto_oc">Assunto da Ocorrência</label>
      <select id="assunto_oc"></select>
    </div>

    <div class="grid-2">
      <div class="campo">
        <label for="placa_select">Placa (Base Salva)</label>
        <select id="placa_select" onchange="preencherPorPlaca()">
          <option value="">-- Selecione ou digite abaixo --</option>
        </select>
      </div>
      <div class="campo">
        <label for="placa_manual">Placa do Veículo</label>
        <input type="text" id="placa_manual" placeholder="Ex: ABC1D23" oninput="this.value = this.value.toUpperCase()" />
      </div>
    </div>

    <div class="grid-2">
      <div class="campo">
        <label for="motorista">Motorista</label>
        <input type="text" id="motorista" />
      </div>
      <div class="campo">
        <label for="contrato">Nº Contrato</label>
        <input type="text" id="contrato" />
      </div>
      <div class="campo">
        <label for="municipio">Município</label>
        <input type="text" id="municipio" />
      </div>
      <div class="campo">
        <label for="lote">Lote</label>
        <input type="text" id="lote" />
      </div>
      <div class="campo">
        <label for="empresa">Empresa</label>
        <input type="text" id="empresa" />
      </div>
      <div class="campo">
        <label for="relator">Relator (Posto/Grad e Nome)</label>
        <input type="text" id="relator" placeholder="Ex: 1º Sgt Silva" />
      </div>
    </div>

    <div class="campo">
      <label for="relato">Relato da Ocorrência</label>
      <textarea id="relato" placeholder="Descreva os fatos constatados em campo..."></textarea>
    </div>

    <div class="campo no-print">
      <label for="fotos_input">Evidências Fotográficas (Opcional para impressão)</label>
      <input type="file" id="fotos_input" accept="image/*" multiple onchange="renderizarFotos(this.files)" />
    </div>
    <div id="fotos_preview" class="fotos-preview"></div>

    <div class="campo">
      <label>Assinatura do Prestador (Assine na tela)</label>
      <canvas id="assinatura_canvas" width="600" height="150"></canvas>
      <div class="no-print" style="margin-top:6px;">
        <button type="button" class="btn-limpar" style="flex:none; padding:6px 12px;" onclick="limparAssinatura()">Limpar Assinatura</button>
      </div>
    </div>

    <div class="assinaturas-print">
      <div>
        <img id="assinatura_print_img" class="assinatura-img" style="display:none;" />
        <div>_______________________________<br/>Assinatura do Prestador</div>
      </div>
      <div>
        <div style="height:65px;"></div>
        <div>_______________________________<br/>Fiscal / Relator (GOV.BR)</div>
      </div>
    </div>

    <div class="acoes no-print">
      <button type="button" class="btn-imprimir" onclick="imprimirRelato()">🖨️ Imprimir / Salvar PDF</button>
      <button type="button" class="btn-salvar" onclick="salvarLocal()">💾 Salvar no Celular</button>
      <button type="button" class="btn-csv" onclick="exportarCSV()">⬇️ Exportar CSV</button>
      <button type="button" class="btn-limpar" onclick="novoRelato()">🆕 Novo Relato</button>
    </div>

    <div class="historico-box">
      <label>Relatos Salvos Neste Celular (<span id="qtd_salvos">0</span>)</label>
      <div id="lista_salvos" style="font-size:13px;"></div>
    </div>
  </div>

  <script>
    const BASE_PRESTADORES = {base_json};
    const LISTA_OCORRENCIAS = {ocorrencias_json};
    const STORAGE_KEY = "relatos_carro_pipa_offline_v1";
    const ANO_REF = {ano_atual};

    const hoje = new Date().toISOString().split("T")[0];
    document.getElementById("data_oc").value = hoje;

    const selOc = document.getElementById("assunto_oc");
    const opcoesOc = LISTA_OCORRENCIAS.length ? LISTA_OCORRENCIAS : [
      "Abastecimento fora do horário regulamentar estabelecido",
      "Ausência ou desligamento do módulo rastreador/embarcado no veículo",
      "Cobrança indevida de água ao morador",
      "Desvio de rota da carrada"
    ];
    opcoesOc.forEach(item => {{
      const opt = document.createElement("option");
      opt.value = item;
      opt.textContent = item;
      selOc.appendChild(opt);
    }});

    const selPlaca = document.getElementById("placa_select");
    BASE_PRESTADORES.forEach(reg => {{
      if (reg.Placa) {{
        const opt = document.createElement("option");
        opt.value = reg.Placa;
        opt.textContent = reg.Placa;
        selPlaca.appendChild(opt);
      }}
    }});

    function preencherPorPlaca() {{
      const placa = selPlaca.value;
      if (!placa) return;
      document.getElementById("placa_manual").value = placa;
      const reg = BASE_PRESTADORES.find(r => String(r.Placa) === String(placa));
      if (reg) {{
        document.getElementById("motorista").value = reg["Motorista"] || "";
        document.getElementById("contrato").value = reg["Nr Contrato"] || "";
        document.getElementById("municipio").value = reg["Município"] || "";
        document.getElementById("lote").value = reg["Lote"] || "";
        document.getElementById("empresa").value = reg["Empresa"] || "";
      }}
    }}

    const canvas = document.getElementById("assinatura_canvas");
    const ctx = canvas.getContext("2d");
    ctx.strokeStyle = "#000000";
    ctx.lineWidth = 3;
    ctx.lineCap = "round";
    let desenhando = false;

    function getPos(e) {{
      const rect = canvas.getBoundingClientRect();
      const clientX = e.touches ? e.touches[0].clientX : e.clientX;
      const clientY = e.touches ? e.touches[0].clientY : e.clientY;
      const scaleX = canvas.width / rect.width;
      const scaleY = canvas.height / rect.height;
      return {{
        x: (clientX - rect.left) * scaleX,
        y: (clientY - rect.top) * scaleY
      }};
    }}

    function iniciarDesenho(e) {{
      e.preventDefault();
      desenhando = true;
      const pos = getPos(e);
      ctx.beginPath();
      ctx.moveTo(pos.x, pos.y);
    }}

    function moverDesenho(e) {{
      if (!desenhando) return;
      e.preventDefault();
      const pos = getPos(e);
      ctx.lineTo(pos.x, pos.y);
      ctx.stroke();
    }}

    function pararDesenho() {{
      desenhando = false;
    }}

    canvas.addEventListener("mousedown", iniciarDesenho);
    canvas.addEventListener("mousemove", moverDesenho);
    window.addEventListener("mouseup", pararDesenho);
    canvas.addEventListener("touchstart", iniciarDesenho, {{ passive: false }});
    canvas.addEventListener("touchmove", moverDesenho, {{ passive: false }});
    window.addEventListener("touchend", pararDesenho);

    function limparAssinatura() {{
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      document.getElementById("assinatura_print_img").style.display = "none";
    }}

    function imprimirRelato() {{
      const imgPrint = document.getElementById("assinatura_print_img");
      imgPrint.src = canvas.toDataURL("image/png");
      imgPrint.style.display = "block";
      window.print();
    }}

    function renderizarFotos(files) {{
      const preview = document.getElementById("fotos_preview");
      preview.innerHTML = "";
      Array.from(files).forEach(file => {{
        const reader = new FileReader();
        reader.onload = ev => {{
          const img = document.createElement("img");
          img.src = ev.target.result;
          preview.appendChild(img);
        }};
        reader.readAsDataURL(file);
      }});
    }}

    function obterSalvos() {{
      try {{
        return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
      }} catch {{
        return [];
      }}
    }}

    function atualizarNumeroSugerido() {{
      const salvos = obterSalvos();
      const prox = String(salvos.length + 1).padStart(4, "0");
      const numStr = `Nº ${{prox}}/${{ANO_REF}}`;
      document.getElementById("num_oc").value = numStr;
      document.getElementById("num-relato-display").textContent = numStr;
    }}

    function salvarLocal() {{
      const registro = {{
        numero_relato: document.getElementById("num_oc").value,
        data: document.getElementById("data_oc").value,
        assunto: document.getElementById("assunto_oc").value,
        placa: document.getElementById("placa_manual").value || document.getElementById("placa_select").value,
        motorista: document.getElementById("motorista").value,
        contrato: document.getElementById("contrato").value,
        municipio: document.getElementById("municipio").value,
        lote: document.getElementById("lote").value,
        empresa: document.getElementById("empresa").value,
        relator: document.getElementById("relator").value,
        relato: document.getElementById("relato").value
      }};
      const salvos = obterSalvos();
      salvos.push(registro);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(salvos));
      renderizarListaSalvos();
      alert("Relato salvo na memória do celular!");
    }}

    function renderizarListaSalvos() {{
      const salvos = obterSalvos();
      document.getElementById("qtd_salvos").textContent = salvos.length;
      const div = document.getElementById("lista_salvos");
      if (!salvos.length) {{
        div.innerHTML = "<p>Nenhum relato salvo localmente.</p>";
        return;
      }}
      div.innerHTML = salvos.map((s, i) =>
        `<div style="padding:6px 0; border-bottom:1px solid #eee;">
          <strong>${{i + 1}}. ${{s.numero_relato}}</strong> — ${{s.data}} — Placa: <strong>${{s.placa || "-"}}</strong> — ${{s.assunto}}
        </div>`
      ).join("");
    }}

    function exportarCSV() {{
      const salvos = obterSalvos();
      if (!salvos.length) {{
        alert("Não há relatos salvos para exportar.");
        return;
      }}
      const cols = ["numero_relato","data","placa","assunto","relator","relato","motorista","contrato","municipio","lote","empresa"];
      const linhas = [cols.join(";")];
      salvos.forEach(item => {{
        const row = cols.map(c => `"${{String(item[c] || "").replace(/"/g, '""')}}"`).join(";");
        linhas.push(row);
      }});
      const blob = new Blob(["\\ufeff" + linhas.join("\\n")], {{ type: "text/csv;charset=utf-8;" }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `relatos_offline_${{hoje}}.csv`;
      a.click();
      URL.revokeObjectURL(url);
    }}

    function novoRelato() {{
      document.getElementById("placa_select").value = "";
      document.getElementById("placa_manual").value = "";
      document.getElementById("motorista").value = "";
      document.getElementById("contrato").value = "";
      document.getElementById("municipio").value = "";
      document.getElementById("lote").value = "";
      document.getElementById("empresa").value = "";
      document.getElementById("relato").value = "";
      document.getElementById("fotos_preview").innerHTML = "";
      document.getElementById("fotos_input").value = "";
      limparAssinatura();
      atualizarNumeroSugerido();
    }}

    atualizarNumeroSugerido();
    renderizarListaSalvos();
  </script>
</body>
</html>
"""

# ==========================================
# CARREGAMENTO INICIAL EM MEMÓRIA
# ==========================================
if "dados_planilha" not in st.session_state:
    df_b, oc_l, status_inicial = carregar_dados_memoria()
    st.session_state.dados_planilha = df_b
    st.session_state.lista_ocorrencias = oc_l
    st.session_state.status_base = status_inicial

if "historico_ocorrencias" not in st.session_state:
    st.session_state.historico_ocorrencias = carregar_historico_memoria()

# ==========================================
# BARRA LATERAL
# ==========================================
with st.sidebar:
    st.header("App Offline para Celular")
    df_dados_sidebar = st.session_state.dados_planilha
    oc_sidebar = st.session_state.get("lista_ocorrencias", [])
    html_offline = gerar_app_offline_html(df_dados_sidebar, oc_sidebar)
    st.download_button(
        label="📲 Baixar App Offline (.html)",
        data=html_offline.encode("utf-8"),
        file_name="relato_ocorrencia_offline.html",
        mime="text/html",
        use_container_width=True,
        help="Baixe no celular e abra no navegador mesmo em Modo Avião para preencher e imprimir em PDF."
    )

    st.header("Modo de Operação")
    modo_offline_manual = st.toggle(
        "Trabalhar Offline (Sessão)",
        value=st.session_state.get("forcar_offline", False),
        help="Mantém os dados em memória durante a sessão e sincroniza quando voltar à rede."
    )
    if modo_offline_manual != st.session_state.get("forcar_offline", False):
        st.session_state.forcar_offline = modo_offline_manual
        df_b, oc_l, status_inicial = carregar_dados_memoria()
        st.session_state.dados_planilha = df_b
        st.session_state.lista_ocorrencias = oc_l
        st.session_state.status_base = status_inicial
        st.rerun()

    st.header("Base de dados")
    if not st.session_state.get("forcar_offline", False) and st.button("Atualizar dados da planilha"):
        buscar_dados_online.clear()
        df_atualizado, oc_atualizado, status_atualizado = carregar_dados_memoria()
        st.session_state.dados_planilha = df_atualizado
        st.session_state.lista_ocorrencias = oc_atualizado
        st.session_state.status_base = status_atualizado
        st.rerun()

df_dados = st.session_state.dados_planilha
status_conexao = st.session_state.status_base

# ==========================================
# INTERFACE PRINCIPAL
# ==========================================
st.title("📋 Relato de Ocorrência")
st.caption(f"Status da Base de Dados: **{status_conexao}**")
if not status_conexao.startswith("Online"):
    st.info(status_conexao)

aba_novo, aba_historico = st.tabs(["✍️ Novo Relato", "🔍 Histórico / Consultas"])

with aba_novo:
    st.subheader("Dados da Fiscalização")

    num_exibicao = st.session_state.get("ultimo_numero_gerado", "Gerado automaticamente ao salvar")
    st.info(f"**Identificação do Documento:** {num_exibicao}")

    campo_data = st.date_input("Campo 00: Data da Ocorrência", datetime.now())

    lista_ocorrencias = st.session_state.get("lista_ocorrencias", [])
    if not lista_ocorrencias:
        lista_ocorrencias = [
            "Abastecimento fora do horário regulamentar estabelecido",
            "Abastecimento realizado com volume inferior à cota diária",
            "Ausência do motorista ou prestador responsável"
        ]

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
    st.markdown("### Assinatura do Prestador")
    canvas_result = None
    assinatura_anexada = None
    if HAS_CANVAS and st_canvas is not None:
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
    else:
        st.caption("Anexe a imagem da assinatura do prestador (ou baixe o App Offline .html para assinar direto na tela):")
        assinatura_anexada = st.file_uploader(
            "Imagem da assinatura do prestador",
            type=["jpg", "jpeg", "png"],
            key="assinatura_fallback"
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
        try:
            if canvas_result is not None and getattr(canvas_result, "image_data", None) is not None:
                arr = canvas_result.image_data
                if arr is not None and arr.sum() > 0:
                    img_assinatura_pil = Image.fromarray(arr.astype('uint8'), mode="RGBA")
                    pad_buffer = io.BytesIO()
                    img_assinatura_pil.convert("RGB").save(pad_buffer, format="PNG")
                    pad_buffer.seek(0)
                    img_assinatura = RLImage(pad_buffer, width=160, height=70)
            elif assinatura_anexada is not None:
                assinatura_anexada.seek(0)
                img_ass = ImageOps.exif_transpose(Image.open(assinatura_anexada))
                img_ass.thumbnail((320, 140))
                pad_buffer = io.BytesIO()
                img_ass.convert("RGB").save(pad_buffer, format="PNG")
                pad_buffer.seek(0)
                img_assinatura = RLImage(pad_buffer, width=160, height=70)
        except Exception:
            pass

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
            if not st.session_state.get("forcar_offline", False):
                try:
                    ws = obter_aba_relatorios()
                    numero_gerado = calcular_proximo_numero(ws, ano_oc)
                    nova_oc["numero_relato"] = numero_gerado
                    ws.append_row(
                        [nova_oc.get(coluna, "") for coluna in COLUNAS_HISTORICO],
                        value_input_option="RAW",
                    )
                    st.session_state.historico_ocorrencias.append(nova_oc)
                    st.session_state.ultimo_numero_gerado = numero_gerado
                    st.success(f"Ocorrência salva no Google Sheets! Protocolo: **{numero_gerado}**")
                except Exception as erro_online:
                    st.session_state.historico_ocorrencias.append(nova_oc)
                    st.session_state.ultimo_numero_gerado = numero_gerado
                    st.warning(
                        f"Google Sheets indisponível. Mantido na memória com protocolo {numero_gerado}: {erro_online}"
                    )
            else:
                st.session_state.historico_ocorrencias.append(nova_oc)
                st.session_state.ultimo_numero_gerado = numero_gerado
                st.success(f"Ocorrência mantida na sessão offline. Protocolo: **{numero_gerado}**")

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
    if not st.session_state.get("forcar_offline", False) and st.button("Sincronizar pendentes da sessão"):
        try:
            quantidade = sincronizar_pendentes_sessao()
        except Exception as erro:
            st.error(f"Não foi possível sincronizar: {erro}")
        else:
            st.success(f"Sincronização concluída: {quantidade} relato(s) enviado(s).")
            st.session_state.historico_ocorrencias = carregar_historico_memoria()
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
        st.info("Nenhuma ocorrência registrada.")
