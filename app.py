from datetime import date

import pandas as pd
import streamlit as st

from tc6m import (
    FORMULAS_DPP,
    PatientData,
    build_default_during_table,
    build_default_pre_table,
    build_default_recovery_table,
    build_effort_figure,
    build_excel_bytes,
    build_oscillation_figure,
    build_patient_dataframe,
    build_pdf_bytes,
    build_safe_filename,
    build_summary_dataframe,
    build_curve_findings,
    calcular_dpp_ben_saad,
    calcular_dpp_enright,
    calcular_dpp_iwama,
    calcular_dpp_por_formula,
    calcular_fc_maxima,
    calcular_fc_submaxima,
    calculate_tc6m_professional,
    combine_timeseries,
    normalize_timeseries,
)


CONTRAINDICACOES_ABSOLUTAS = [
    "Angina instável no mês anterior",
    "Infarto do miocárdio no mês anterior",
    "Arritmias não controladas",
    "Estenose aórtica",
    "Endocardite ativa",
    "Miocardite ou pericardite aguda",
    "Tromboembolismo pulmonar",
    "Trombose de membros inferiores",
    "Suspeita de aneurisma dissecante",
    "Doenças agudas que possam influenciar no teste",
    "Distúrbio mental que limite a colaboração",
]

CONTRAINDICACOES_RELATIVAS = [
    "Frequência cardíaca em repouso > 120 bpm ou bradicardia",
    "Pressão arterial sistólica > 180 mmHg",
    "Pressão arterial diastólica > 100 mmHg",
    "Bloqueio atrioventricular de 3º grau",
    "Cardiomiopatia hipertrófica",
    "Gestação avançada ou complicada",
    "Anormalidade de eletrólitos",
    "Disfunção ortopédica que limite a caminhada",
]


st.set_page_config(
    page_title="Teste de Caminhada de 6 Minutos",
    layout="wide",
)

st.markdown(
    """
    <style>
        :root {
            --primary-color: #0B4238;
            --secondary-color: #2F5D55;
            --bg-color: #F0F4F3;
            --card-bg: #FFFFFF;
            --text-dark: #1A202C;
            --accent-color: #007AFF;
            --muted-text: #718096;
            --border-color: #D7E1DE;
            --success-color: #28A745;
            --warning-color: #F6AD55;
            --danger-color: #E53E3E;
            --border-radius: 12px;
            --shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        }

        .stApp {
            background: var(--bg-color);
        }

        html, body {
            color: var(--text-dark);
            font-family: Inter, Roboto, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }

        h1, h2, h3 {
            color: var(--primary-color) !important;
            letter-spacing: -0.02em;
        }

        h3 {
            background: var(--primary-color);
            color: #FFFFFF !important;
            border-radius: var(--border-radius);
            padding: 10px 14px;
            margin-bottom: 18px;
        }

        .section-label {
            color: var(--primary-color) !important;
            font-weight: 800;
            font-size: 1rem;
            margin: 14px 0 8px;
        }

        .hero {
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--secondary-color) 100%);
            border-radius: 18px;
            padding: 28px 32px;
            margin-bottom: 22px;
            box-shadow: var(--shadow);
        }

        .hero h1, .hero p {
            color: #ffffff !important;
            margin: 0;
        }

        .hero p {
            opacity: 0.92;
            margin-top: 8px;
            font-size: 1.02rem;
        }

        .clinical-card {
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-left: 5px solid var(--primary-color);
            border-radius: var(--border-radius);
            padding: 22px 24px;
            margin-bottom: 18px;
            box-shadow: var(--shadow);
        }

        [data-testid="stWidgetLabel"] p,
        label,
        .stCheckbox p,
        .stRadio p {
            color: var(--primary-color) !important;
            font-weight: 700 !important;
        }

        input, textarea,
        div[data-baseweb="input"] input,
        div[data-baseweb="textarea"] textarea {
            color: var(--text-dark) !important;
            -webkit-text-fill-color: var(--text-dark) !important;
            background-color: #FFFFFF !important;
        }

        div[data-baseweb="select"] span,
        div[data-baseweb="select"] div {
            color: #FFFFFF !important;
        }

        div[data-baseweb="select"] > div {
            background: var(--primary-color) !important;
            border-color: var(--primary-color) !important;
            color: #FFFFFF !important;
        }

        div[data-baseweb="select"] svg {
            color: #FFFFFF !important;
            fill: #FFFFFF !important;
        }

        div[data-baseweb="popover"] {
            background: var(--primary-color) !important;
            border-radius: var(--border-radius) !important;
            overflow: hidden !important;
        }

        div[data-baseweb="popover"] ul,
        ul[role="listbox"] {
            background: var(--primary-color) !important;
        }

        div[data-baseweb="popover"] li,
        li[role="option"],
        li[role="option"] span,
        li[role="option"] div {
            color: #FFFFFF !important;
            background: var(--primary-color) !important;
            font-weight: 700 !important;
        }

        div[data-baseweb="popover"] li:hover,
        li[role="option"]:hover,
        li[aria-selected="true"] {
            background: var(--accent-color) !important;
            color: #FFFFFF !important;
        }

        .interpretation-box {
            background: #F7FCF9;
            border: 1px solid #CFE2DA;
            border-left: 5px solid var(--primary-color);
            border-radius: var(--border-radius);
            padding: 16px 18px;
            color: var(--primary-color) !important;
            font-weight: 700;
            line-height: 1.55;
            margin: 14px 0;
        }

        .findings-box {
            background: #FFFFFF;
            border: 1px solid #CFE2DA;
            border-left: 5px solid var(--primary-color);
            border-radius: var(--border-radius);
            padding: 14px 18px;
            color: var(--text-dark) !important;
            margin-bottom: 14px;
        }

        .findings-box li {
            color: var(--text-dark) !important;
            margin-bottom: 6px;
        }

        .danger-panel, .warning-panel {
            border-radius: var(--border-radius);
            padding: 14px 16px;
            min-height: 260px;
            box-shadow: var(--shadow);
        }

        .danger-panel {
            background: #FFF5F5;
            border: 1px solid #FEB2B2;
            border-left: 5px solid #7B1E1E;
        }

        .warning-panel {
            background: #FFF8E6;
            border: 1px solid #F6AD55;
            border-left: 5px solid #B7791F;
        }

        .danger-panel h4 {
            color: #7B1E1E !important;
            margin-top: 0;
        }

        .warning-panel h4 {
            color: #8A5A00 !important;
            margin-top: 0;
        }

        .danger-panel li, .warning-panel li {
            color: var(--text-dark) !important;
            margin-bottom: 4px;
        }

        .status-box {
            border-radius: var(--border-radius);
            padding: 14px 16px;
            font-weight: 700;
            margin-top: 12px;
        }

        .status-ok {
            background: #E6F4EA;
            border-left: 5px solid #14532D;
            color: #14532D !important;
        }

        .status-warning-box {
            background: #FFF4D6;
            border-left: 5px solid #B7791F;
            color: #5F3B00 !important;
        }

        .status-danger-box {
            background: #FDE2E2;
            border-left: 5px solid #7B1E1E;
            color: #7B1E1E !important;
        }

        .result-card {
            background: var(--card-bg);
            border-radius: var(--border-radius);
            padding: 18px 18px;
            box-shadow: var(--shadow);
            border-left: 5px solid var(--primary-color);
            min-height: 112px;
        }

        .result-card .label {
            color: var(--muted-text) !important;
            font-size: 0.84rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            margin-bottom: 6px;
        }

        .result-card .value {
            color: var(--primary-color) !important;
            font-size: 1.28rem;
            font-weight: 800;
            line-height: 1.2;
            white-space: normal;
            overflow-wrap: anywhere;
        }

        .result-card.success {
            border-left-color: var(--success-color);
        }

        .result-card.warning {
            border-left-color: var(--warning-color);
        }

        .result-card.danger {
            border-left-color: var(--danger-color);
        }

        .soft-note {
            background: #FFFFFF;
            border-left: 5px solid var(--accent-color);
            border-radius: var(--border-radius);
            padding: 13px 15px;
            margin: 12px 0 2px;
            color: var(--text-dark);
        }

        .test-panel {
            background: #FFFFFF;
            border: 1px solid var(--border-color);
            border-radius: var(--border-radius);
            padding: 18px 20px;
            margin-bottom: 18px;
            box-shadow: var(--shadow);
        }

        [data-testid="stMetric"] {
            background: #FFFFFF;
            border: 1px solid var(--border-color);
            border-radius: var(--border-radius);
            padding: 16px;
            box-shadow: var(--shadow);
        }

        [data-testid="stMetricLabel"] p {
            color: var(--muted-text) !important;
            font-weight: 600;
        }

        [data-testid="stMetricValue"] {
            color: var(--primary-color) !important;
            font-weight: 800;
        }

        .stButton button, .stDownloadButton button {
            background: var(--primary-color) !important;
            color: #ffffff !important;
            border: 1px solid var(--primary-color) !important;
            border-radius: var(--border-radius) !important;
            font-weight: 700 !important;
            box-shadow: var(--shadow);
        }

        .stButton button:hover, .stDownloadButton button:hover {
            background: var(--accent-color) !important;
            color: #ffffff !important;
            border: 1px solid var(--accent-color) !important;
        }

        div[data-testid="stDataFrame"] {
            border-radius: var(--border-radius);
            overflow: hidden;
            border: 1px solid var(--border-color);
        }

        button[data-baseweb="tab"] {
            background: var(--primary-color) !important;
            border: 1px solid var(--primary-color) !important;
            border-radius: 999px !important;
            padding: 8px 16px !important;
            margin-right: 8px !important;
        }

        button[data-baseweb="tab"] p {
            color: #FFFFFF !important;
            font-weight: 700 !important;
        }

        button[data-baseweb="tab"][aria-selected="true"] {
            background: var(--accent-color) !important;
            border-color: var(--accent-color) !important;
        }

        button[data-baseweb="tab"][aria-selected="true"] p {
            color: #ffffff !important;
        }

        div[data-testid="stExpander"] details summary {
            background: var(--primary-color) !important;
            border: 1px solid var(--primary-color) !important;
            border-radius: var(--border-radius) !important;
            padding: 10px 14px !important;
            color: #FFFFFF !important;
        }

        div[data-testid="stExpander"] details summary p,
        div[data-testid="stExpander"] details summary span,
        div[data-testid="stExpander"] details summary svg {
            color: #FFFFFF !important;
            fill: #FFFFFF !important;
            font-weight: 700 !important;
        }

        div[data-testid="stExpander"] details[open] summary {
            background: var(--accent-color) !important;
            border-color: var(--accent-color) !important;
        }

        .status-normal [data-testid="stMetricValue"] {
            color: var(--success-color) !important;
        }

        .status-warning [data-testid="stMetricValue"] {
            color: var(--warning-color) !important;
        }

        .status-danger [data-testid="stMetricValue"] {
            color: var(--danger-color) !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


def limpar_resultado() -> None:
    """Remove o resumo antigo quando algum dado do paciente ou teste muda."""

    st.session_state.resultado_tc6m = None
    st.session_state.paciente_tc6m = None
    st.session_state.serie_tc6m = None


def responder_triagem(status: str) -> None:
    """Registra a resposta da triagem e mantém a trava invisível até haver escolha."""

    st.session_state.triagem_status = status
    limpar_resultado()


def converter_pa_rapida(valor: object) -> tuple[int, int]:
    """Converte PA digitada como 12x8, 12/8, 120x80 ou 120/80 em PAS/PAD."""

    texto = str(valor or "").lower().strip()
    if not texto:
        return 0, 0

    apenas_digitos = "".join(caractere for caractere in texto if caractere.isdigit())
    if texto.isdigit() and len(apenas_digitos) == 3:
        return int(apenas_digitos[:2]) * 10, int(apenas_digitos[2]) * 10
    if texto.isdigit() and len(apenas_digitos) == 5:
        return int(apenas_digitos[:3]), int(apenas_digitos[3:])
    if texto.isdigit() and len(apenas_digitos) == 6:
        return int(apenas_digitos[:3]), int(apenas_digitos[3:])

    for separador in [" por ", "x", "/", "\\", "-", " "]:
        texto = texto.replace(separador, " ")

    partes = [parte for parte in texto.split() if parte.replace(".", "", 1).isdigit()]
    if len(partes) < 2:
        return 0, 0

    pas = float(partes[0])
    pad = float(partes[1])

    if pas < 30:
        pas *= 10
    if pad < 30:
        pad *= 10

    return int(round(pas)), int(round(pad))


def formatar_pa(pas: int | float, pad: int | float) -> str:
    """Mostra a PA no formato rápido usado no atendimento."""

    if float(pas) <= 0 or float(pad) <= 0:
        return ""
    return f"{int(round(float(pas)))}/{int(round(float(pad)))}"


def preparar_editor_com_pa(df):
    """Prepara PA em campo único para evitar erro de edição da diastólica."""

    tabela = normalize_timeseries(df)
    tabela["PA"] = tabela.apply(lambda row: formatar_pa(row["PAS"], row["PAD"]), axis=1)
    return tabela[["Tempo", "FC", "SpO2", "FR", "PA", "Borg Respiratório", "Borg MMII"]]


def restaurar_pas_pad(df):
    """Transforma a PA digitada em PAS/PAD para o motor de cálculo."""

    linhas = []
    for _, row in df.iterrows():
        pas, pad = converter_pa_rapida(row.get("PA", ""))
        linhas.append(
            {
                "Tempo": row["Tempo"],
                "FC": row["FC"],
                "SpO2": row["SpO2"],
                "FR": row["FR"],
                "PAS": pas,
                "PAD": pad,
                "Borg Respiratório": row["Borg Respiratório"],
                "Borg MMII": row["Borg MMII"],
            }
        )
    return normalize_timeseries(pd.DataFrame(linhas))


def estilizar_tabela(df):
    """Aplica uma aparência mais clínica e clara às tabelas de resultado."""

    return (
        df.style.set_properties(
            **{
                "background-color": "#ffffff",
                "color": "#07130f",
                "border-color": "#d8e6df",
                "font-size": "0.95rem",
            }
        )
        .set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        ("background-color", "#dcefe7"),
                        ("color", "#102c25"),
                        ("font-weight", "700"),
                        ("border", "1px solid #c9ded4"),
                    ],
                },
                {
                    "selector": "td",
                    "props": [("border", "1px solid #e2ece7"), ("padding", "8px")],
                },
            ]
        )
        .hide(axis="index")
    )


def montar_lista_html(itens: list[str]) -> str:
    """Monta uma lista HTML simples para painéis visuais."""

    return "".join(f"<li>{item}</li>" for item in itens)


def classe_desempenho(percentual: float) -> str:
    """Define cor visual do desempenho previsto."""

    if percentual >= 80:
        return "success"
    if percentual >= 60:
        return "warning"
    return "danger"


def card_resultado(label: str, valor: str, status: str = "") -> None:
    """Renderiza card de resultado sem cortar texto com reticências."""

    classe = f"result-card {status}".strip()
    st.markdown(
        f"""
        <div class="{classe}">
            <div class="label">{label}</div>
            <div class="value">{valor}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def iniciar_estado() -> None:
    """Cria valores iniciais da interface e das tabelas clínicas."""

    defaults = {
        "nome": "",
        "prontuario": "",
        "avaliador": "",
        "diagnostico": "",
        "data_avaliacao": date.today(),
        "sexo_label": "Masculino",
        "idade": 60,
        "peso": 70.0,
        "altura_cm": 170.0,
        "formula_principal": FORMULAS_DPP[0],
        "perfil_teste": "Homem adulto",
        "triagem_status": "Selecione",
        "contra_abs": False,
        "contra_rel": False,
        "observacao_triagem": "",
        "distancia": 420.0,
        "interrompeu_label": "Não",
        "motivo_interrupcao": "",
        "distancia_interrupcao": 0.0,
        "resultado_tc6m": None,
        "paciente_tc6m": None,
        "serie_tc6m": None,
    }

    for key, value in defaults.items():
        st.session_state.setdefault(key, value)

    st.session_state.setdefault("pre_df", build_default_pre_table())
    st.session_state.setdefault("during_df", build_default_during_table())
    st.session_state.setdefault("recovery_df", build_default_recovery_table())


def preencher_ambiente_teste(perfil: str) -> None:
    """Preenche dados fictícios por perfil para testar fluxo, gráficos, PDF e Excel."""

    perfis = {
        "Homem adulto": {
            "nome": "Paciente Teste - Homem",
            "prontuario": "TC6M-H-001",
            "sexo_label": "Masculino",
            "idade": 62,
            "peso": 74.0,
            "altura_cm": 171.0,
            "formula": FORMULAS_DPP[0],
            "diagnostico": "Avaliação funcional cardiorrespiratória",
            "distancia": 438.0,
            "pre": [78, 97, 18, 122, 78, 0.0, 0.0],
            "fc": [94, 106, 116, 124, 132, 139],
            "spo2": [96, 95, 95, 94, 93, 92],
            "borg_r": [1, 2, 3, 4, 5, 6],
            "borg_m": [1, 2, 3, 4, 5, 5],
            "rec": [[120, 94, 24, 146, 82, 4, 3], [98, 96, 20, 132, 80, 2, 2], [84, 97, 18, 124, 78, 1, 1]],
        },
        "Mulher adulta": {
            "nome": "Paciente Teste - Mulher",
            "prontuario": "TC6M-M-001",
            "sexo_label": "Feminino",
            "idade": 55,
            "peso": 66.0,
            "altura_cm": 160.0,
            "formula": FORMULAS_DPP[0],
            "diagnostico": "Avaliação funcional cardiorrespiratória",
            "distancia": 462.0,
            "pre": [74, 98, 17, 118, 74, 0.0, 0.0],
            "fc": [88, 98, 108, 116, 122, 128],
            "spo2": [98, 97, 97, 96, 96, 95],
            "borg_r": [1, 1, 2, 3, 4, 5],
            "borg_m": [1, 2, 2, 3, 4, 4],
            "rec": [[110, 96, 22, 136, 78, 3, 3], [90, 97, 19, 124, 76, 2, 1], [78, 98, 17, 118, 74, 1, 0]],
        },
        "Criança/adolescente": {
            "nome": "Paciente Teste - Adolescente",
            "prontuario": "TC6M-C-001",
            "sexo_label": "Masculino",
            "idade": 12,
            "peso": 42.0,
            "altura_cm": 150.0,
            "formula": FORMULAS_DPP[2],
            "diagnostico": "Avaliação funcional pediátrica",
            "distancia": 610.0,
            "pre": [82, 99, 20, 108, 68, 0.0, 0.0],
            "fc": [102, 118, 132, 144, 152, 158],
            "spo2": [99, 98, 98, 97, 97, 97],
            "borg_r": [1, 2, 3, 4, 5, 6],
            "borg_m": [1, 2, 3, 4, 4, 5],
            "rec": [[132, 98, 26, 118, 72, 4, 3], [104, 99, 22, 110, 70, 2, 2], [88, 99, 20, 106, 68, 1, 1]],
        },
        "Paciente com DPOC": {
            "nome": "Paciente Teste - DPOC",
            "prontuario": "TC6M-DPOC-001",
            "sexo_label": "Masculino",
            "idade": 68,
            "peso": 69.0,
            "altura_cm": 168.0,
            "formula": FORMULAS_DPP[0],
            "diagnostico": "DPOC - ambiente fictício de teste",
            "distancia": 285.0,
            "pre": [86, 94, 21, 132, 82, 1.0, 1.0],
            "fc": [98, 108, 116, 124, 130, 134],
            "spo2": [93, 91, 89, 88, 87, 86],
            "borg_r": [2, 3, 4, 5, 7, 8],
            "borg_m": [1, 2, 3, 4, 5, 5],
            "rec": [[122, 88, 28, 150, 86, 7, 4], [104, 91, 24, 140, 84, 5, 3], [92, 93, 22, 134, 82, 3, 2]],
        },
    }

    dados = perfis[perfil]
    st.session_state.nome = dados["nome"]
    st.session_state.prontuario = dados["prontuario"]
    st.session_state.avaliador = "Equipe Cardiorrespiratória"
    st.session_state.diagnostico = dados["diagnostico"]
    st.session_state.data_avaliacao = date.today()
    st.session_state.sexo_label = dados["sexo_label"]
    st.session_state.idade = dados["idade"]
    st.session_state.peso = dados["peso"]
    st.session_state.altura_cm = dados["altura_cm"]
    st.session_state.formula_principal = dados["formula"]
    st.session_state.triagem_status = "Sem contraindicações"
    st.session_state.contra_abs = False
    st.session_state.contra_rel = False
    st.session_state.observacao_triagem = f"Ambiente de teste fictício: {perfil}."
    st.session_state.distancia = dados["distancia"]
    st.session_state.interrompeu_label = "Não"
    st.session_state.motivo_interrupcao = ""
    st.session_state.distancia_interrupcao = 0.0

    pre = dados["pre"]
    st.session_state.pre_df = normalize_timeseries(
        build_default_pre_table().assign(
            FC=[pre[0]], SpO2=[pre[1]], FR=[pre[2]], PAS=[pre[3]], PAD=[pre[4]],
            **{"Borg Respiratório": [pre[5]], "Borg MMII": [pre[6]]},
        )
    )

    st.session_state.during_df = build_default_during_table().assign(
        FC=dados["fc"],
        SpO2=dados["spo2"],
        **{"Borg Respiratório": dados["borg_r"], "Borg MMII": dados["borg_m"]},
    )

    rec = dados["rec"]
    st.session_state.recovery_df = normalize_timeseries(
        build_default_recovery_table().assign(
            FC=[linha[0] for linha in rec],
            SpO2=[linha[1] for linha in rec],
            FR=[linha[2] for linha in rec],
            PAS=[linha[3] for linha in rec],
            PAD=[linha[4] for linha in rec],
            **{
                "Borg Respiratório": [linha[5] for linha in rec],
                "Borg MMII": [linha[6] for linha in rec],
            },
        )
    )

    limpar_resultado()


iniciar_estado()

with st.sidebar:
    st.header("Menu")
    if st.button("Limpar resumo final", use_container_width=True):
        limpar_resultado()
        st.rerun()

    st.divider()
    st.write("Rode o resumo somente depois de preencher avaliação prévia, execução, recuperação e resultado final.")


st.markdown(
    """
    <div class="hero">
        <h1>Teste de Caminhada de 6 Minutos (TC6M)</h1>
        <p>Registro clínico organizado por etapas: paciente, triagem, predição, execução, recuperação e resultado final.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="test-panel">', unsafe_allow_html=True)
teste_col1, teste_col2 = st.columns([2, 1])
with teste_col1:
    st.selectbox(
        "Ambiente de teste para preencher automaticamente",
        ["Homem adulto", "Mulher adulta", "Criança/adolescente", "Paciente com DPOC"],
        key="perfil_teste",
        help="Escolha um perfil fictício para testar o sistema rapidamente.",
    )
with teste_col2:
    st.write("")
    st.write("")
    if st.button("Preencher ambiente de teste", use_container_width=True):
        preencher_ambiente_teste(st.session_state.perfil_teste)
        st.rerun()
st.markdown("</div>", unsafe_allow_html=True)


st.markdown('<div class="clinical-card">', unsafe_allow_html=True)
st.subheader("Etapa 1 - Identificação e antropometria")

col1, col2, col3 = st.columns(3)
with col1:
    st.text_input("Nome do paciente", key="nome")
    st.text_input("Prontuário/ID", key="prontuario")
    st.date_input("Data da avaliação", key="data_avaliacao")

with col2:
    st.selectbox("Sexo biológico", ["Masculino", "Feminino"], key="sexo_label")
    st.number_input("Idade (anos)", min_value=1, max_value=120, step=1, key="idade")
    st.number_input("Peso (kg)", min_value=1.0, max_value=350.0, step=0.1, key="peso")

with col3:
    st.number_input("Altura (cm)", min_value=50.0, max_value=240.0, step=0.1, key="altura_cm")
    st.text_input("Avaliador", key="avaliador")
    st.text_input("Diagnóstico/condição clínica", key="diagnostico")

st.markdown("</div>", unsafe_allow_html=True)


st.markdown('<div class="clinical-card">', unsafe_allow_html=True)
st.subheader("Etapa 2 - Triagem de segurança")

st.markdown(
    '<div class="soft-note">Antes de iniciar o teste, responda obrigatoriamente a triagem: sem contraindicações, contraindicação relativa ou contraindicação absoluta.</div>',
    unsafe_allow_html=True,
)

triagem1, triagem2 = st.columns(2)
with triagem1:
    with st.expander("Abrir contraindicações absolutas"):
        st.markdown(
            f"""
            <div class="danger-panel">
                <h4>⚠ Contraindicações absolutas</h4>
                <ul>{montar_lista_html(CONTRAINDICACOES_ABSOLUTAS)}</ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

with triagem2:
    with st.expander("Abrir contraindicações relativas"):
        st.markdown(
            f"""
            <div class="warning-panel">
                <h4>⚠ Contraindicações relativas</h4>
                <ul>{montar_lista_html(CONTRAINDICACOES_RELATIVAS)}</ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.markdown('<div class="section-label">Resultado da triagem de segurança</div>', unsafe_allow_html=True)
triagem_botoes = st.columns(3)
with triagem_botoes[0]:
    if st.button("Sem contraindicações", use_container_width=True):
        responder_triagem("Sem contraindicações")
        st.rerun()
with triagem_botoes[1]:
    if st.button("Contraindicação relativa", use_container_width=True):
        responder_triagem("Contraindicação relativa")
        st.rerun()
with triagem_botoes[2]:
    if st.button("Contraindicação absoluta", use_container_width=True):
        responder_triagem("Contraindicação absoluta")
        st.rerun()

st.session_state.contra_abs = st.session_state.triagem_status == "Contraindicação absoluta"
st.session_state.contra_rel = st.session_state.triagem_status == "Contraindicação relativa"

st.text_area(
    "Observações da triagem",
    key="observacao_triagem",
    placeholder="Ex.: sintomas, medicamentos, queixas, limitações ortopédicas ou observações relevantes.",
)

triagem_respondida = st.session_state.triagem_status != "Selecione"
triagem_liberada = st.session_state.triagem_status in ["Sem contraindicações", "Contraindicação relativa"]

if st.session_state.triagem_status == "Selecione":
    st.markdown(
        '<div class="status-box status-warning-box">⚠ Trava ativa: responda a triagem de segurança para liberar a avaliação prévia.</div>',
        unsafe_allow_html=True,
    )
elif st.session_state.contra_abs:
    st.markdown(
        '<div class="status-box status-danger-box">⚠ Trava de segurança ativada: há contraindicação absoluta. O teste não deve ser iniciado.</div>',
        unsafe_allow_html=True,
    )
elif st.session_state.contra_rel:
    st.markdown(
        '<div class="status-box status-warning-box">⚠ Contraindicação relativa registrada. Prossiga apenas com critério clínico e supervisão adequada.</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<div class="status-box status-ok">✓ Triagem sem contraindicação marcada.</div>',
        unsafe_allow_html=True,
    )

st.markdown("</div>", unsafe_allow_html=True)

if not triagem_respondida:
    st.stop()

if not triagem_liberada:
    st.stop()


st.markdown('<div class="clinical-card">', unsafe_allow_html=True)
st.subheader("Etapa 3 - Avaliação prévia e cálculos preditos")

sexo = "M" if st.session_state.sexo_label == "Masculino" else "F"
fc_max = calcular_fc_maxima(int(st.session_state.idade))
fc_submax = calcular_fc_submaxima(int(st.session_state.idade))

st.selectbox(
    "Escolha a fórmula principal para calcular a distância predita",
    FORMULAS_DPP,
    key="formula_principal",
)

dpp_principal, lin_principal = calcular_dpp_por_formula(
    st.session_state.formula_principal,
    sexo,
    int(st.session_state.idade),
    float(st.session_state.peso),
    float(st.session_state.altura_cm),
)
dpp_enright, lin_enright = calcular_dpp_enright(sexo, int(st.session_state.idade), float(st.session_state.peso), float(st.session_state.altura_cm))
dpp_iwama = calcular_dpp_iwama(sexo, int(st.session_state.idade))
dpp_ben_saad = calcular_dpp_ben_saad(int(st.session_state.idade), float(st.session_state.peso), float(st.session_state.altura_cm))

metricas = st.columns(5)
metricas[0].metric("FC máxima estimada", f"{fc_max} bpm")
metricas[1].metric("FC submáxima 85%", f"{fc_submax} bpm")
metricas[2].metric("DPP principal", f"{dpp_principal:.2f} m")
metricas[3].metric("LIN principal", f"{lin_principal:.2f} m" if lin_principal is not None else "Não definido")
metricas[4].metric("Fórmula", st.session_state.formula_principal.split(" - ")[0])

with st.expander("Ver todas as fórmulas preditas"):
    st.dataframe(
        {
            "Fórmula": [
                "Enright e Sherrill (1998)",
                "Iwama et al. (2009)",
                "Ben Saad et al. (2009)",
            ],
            "Resultado": [
                f"{dpp_enright:.2f} m | LIN {lin_enright:.2f} m",
                f"{dpp_iwama:.2f} m",
                f"{dpp_ben_saad:.2f} m",
            ],
        },
        use_container_width=True,
        hide_index=True,
    )

st.markdown(
    '<div class="soft-note">Preencha os sinais de repouso antes de iniciar o teste. Na PA, digite em campo único: 120/80, 120x80, 12x8, 128 ou 12080.</div>',
    unsafe_allow_html=True,
)

pre_before = st.session_state.pre_df.copy()
pre_df = st.data_editor(
    preparar_editor_com_pa(st.session_state.pre_df),
    key="editor_pre_pa_v4",
    hide_index=True,
    use_container_width=True,
    num_rows="fixed",
    column_config={
        "Tempo": st.column_config.TextColumn("Momento", disabled=True),
        "FC": st.column_config.NumberColumn("FC", min_value=0, max_value=260, step=1),
        "SpO2": st.column_config.NumberColumn("SpO2", min_value=0, max_value=100, step=1),
        "FR": st.column_config.NumberColumn("FR", min_value=0, max_value=80, step=1),
        "PA": st.column_config.TextColumn("PA", help="Ex.: 120/80, 120x80, 12x8, 128 ou 12080."),
        "Borg Respiratório": st.column_config.NumberColumn("Borg respiratório", min_value=0.0, max_value=10.0, step=0.5),
        "Borg MMII": st.column_config.NumberColumn("Borg MMII", min_value=0.0, max_value=10.0, step=0.5),
    },
)
st.session_state.pre_df = restaurar_pas_pad(pre_df)
if not st.session_state.pre_df.equals(pre_before):
    limpar_resultado()

st.markdown("</div>", unsafe_allow_html=True)


st.markdown('<div class="clinical-card">', unsafe_allow_html=True)
st.subheader("Etapa 4 - Execução do teste e recuperação")

st.markdown(
    '<div class="soft-note">Durante a caminhada, registre apenas FC, SpO2 e Borg. PA e FR ficam para antes e recuperação.</div>',
    unsafe_allow_html=True,
)

during_before = st.session_state.during_df.copy()
during_df = st.data_editor(
    st.session_state.during_df,
    key="editor_during_v2",
    hide_index=True,
    use_container_width=True,
    num_rows="fixed",
    column_config={
        "Tempo": st.column_config.TextColumn("Minuto", disabled=True),
        "FC": st.column_config.NumberColumn("FC", min_value=0, max_value=260, step=1),
        "SpO2": st.column_config.NumberColumn("SpO2", min_value=0, max_value=100, step=1),
        "Borg Respiratório": st.column_config.NumberColumn("Borg respiratório", min_value=0.0, max_value=10.0, step=0.5),
        "Borg MMII": st.column_config.NumberColumn("Borg MMII", min_value=0.0, max_value=10.0, step=0.5),
    },
)
st.session_state.during_df = during_df
if not st.session_state.during_df.equals(during_before):
    limpar_resultado()

st.markdown(
    '<h3>Recuperação</h3>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="soft-note">Após o teste, registre novamente os sinais completos em 1, 3 e 6 minutos. Na PA, digite em campo único: 120/80, 120x80, 12x8, 128 ou 12080.</div>',
    unsafe_allow_html=True,
)

recovery_before = st.session_state.recovery_df.copy()
recovery_df = st.data_editor(
    preparar_editor_com_pa(st.session_state.recovery_df),
    key="editor_recovery_pa_v4",
    hide_index=True,
    use_container_width=True,
    num_rows="fixed",
    column_config={
        "Tempo": st.column_config.TextColumn("Momento", disabled=True),
        "FC": st.column_config.NumberColumn("FC", min_value=0, max_value=260, step=1),
        "SpO2": st.column_config.NumberColumn("SpO2", min_value=0, max_value=100, step=1),
        "FR": st.column_config.NumberColumn("FR", min_value=0, max_value=80, step=1),
        "PA": st.column_config.TextColumn("PA", help="Ex.: 120/80, 120x80, 12x8, 128 ou 12080."),
        "Borg Respiratório": st.column_config.NumberColumn("Borg respiratório", min_value=0.0, max_value=10.0, step=0.5),
        "Borg MMII": st.column_config.NumberColumn("Borg MMII", min_value=0.0, max_value=10.0, step=0.5),
    },
)
st.session_state.recovery_df = restaurar_pas_pad(recovery_df)
if not st.session_state.recovery_df.equals(recovery_before):
    limpar_resultado()

st.markdown("</div>", unsafe_allow_html=True)


st.markdown('<div class="clinical-card">', unsafe_allow_html=True)
st.subheader("Etapa 5 - Resultado final do teste")

resultado_col1, resultado_col2, resultado_col3 = st.columns(3)
with resultado_col1:
    st.number_input("Distância percorrida ao final do TC6M (m)", min_value=0.0, max_value=2000.0, step=1.0, key="distancia")
with resultado_col2:
    st.radio("O paciente interrompeu o teste?", ["Não", "Sim"], horizontal=True, key="interrompeu_label")
with resultado_col3:
    if st.session_state.interrompeu_label == "Sim":
        st.number_input("Distância no momento da interrupção (m)", min_value=0.0, max_value=2000.0, step=1.0, key="distancia_interrupcao")

if st.session_state.interrompeu_label == "Sim":
    st.text_area("Motivo da interrupção", key="motivo_interrupcao")

serie_completa = combine_timeseries(
    st.session_state.pre_df,
    st.session_state.during_df,
    st.session_state.recovery_df,
)

gerar = st.button(
    "Gerar resumo final do TC6M",
    type="primary",
    disabled=st.session_state.contra_abs,
    use_container_width=True,
)

if gerar:
    try:
        paciente = PatientData(
            nome=st.session_state.nome,
            prontuario=st.session_state.prontuario,
            data_avaliacao=st.session_state.data_avaliacao,
            avaliador=st.session_state.avaliador,
            diagnostico=st.session_state.diagnostico,
            sexo=sexo,
            idade=int(st.session_state.idade),
            peso=float(st.session_state.peso),
            altura_cm=float(st.session_state.altura_cm),
            distancia=float(st.session_state.distancia),
            formula_principal=st.session_state.formula_principal,
            interrompeu=st.session_state.interrompeu_label == "Sim",
            motivo_interrupcao=st.session_state.motivo_interrupcao,
            distancia_interrupcao=float(st.session_state.distancia_interrupcao),
            contraindicacao_absoluta=bool(st.session_state.contra_abs),
            contraindicacao_relativa=bool(st.session_state.contra_rel),
            observacao_triagem=st.session_state.observacao_triagem,
        )

        resultado = calculate_tc6m_professional(paciente, serie_completa)
        st.session_state.paciente_tc6m = paciente
        st.session_state.resultado_tc6m = resultado
        st.session_state.serie_tc6m = serie_completa
        st.success("Resumo final gerado com sucesso.")
    except ValueError as erro:
        st.error(str(erro))

st.markdown("</div>", unsafe_allow_html=True)


if st.session_state.resultado_tc6m and st.session_state.paciente_tc6m is not None:
    paciente = st.session_state.paciente_tc6m
    resultado = st.session_state.resultado_tc6m
    serie = st.session_state.serie_tc6m

    st.markdown('<div class="clinical-card">', unsafe_allow_html=True)
    st.subheader("Resumo final do TC6M")

    cards = st.columns(5)
    with cards[0]:
        card_resultado("Distância percorrida", f"{paciente.distancia:.2f} m")
    with cards[1]:
        card_resultado("DPP escolhida", f"{resultado.dpp_principal:.2f} m")
    with cards[2]:
        card_resultado("Desempenho previsto", f"{resultado.percentual_atingido:.2f}%", classe_desempenho(resultado.percentual_atingido))
    with cards[3]:
        card_resultado("Qualificador funcional", resultado.qualificador_funcional, classe_desempenho(resultado.percentual_atingido))
    with cards[4]:
        card_resultado("Classificação", resultado.classificacao_risco)

    aba1, aba2, aba3, aba4, aba5 = st.tabs(
        ["Interpretação", "Resultados completos", "Sinais vitais", "Gráficos", "Exportar"]
    )

    with aba1:
        resumo_interpretacao = st.columns(3)
        with resumo_interpretacao[0]:
            card_resultado("Qualificador funcional", resultado.qualificador_funcional, classe_desempenho(resultado.percentual_atingido))
        with resumo_interpretacao[1]:
            card_resultado("Classificação", resultado.classificacao_risco)
        with resumo_interpretacao[2]:
            card_resultado("Risco", resultado.risco, classe_desempenho(resultado.percentual_atingido))
        st.markdown(
            f'<div class="interpretation-box">{resultado.interpretacao}</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(estilizar_tabela(build_patient_dataframe(paciente)), use_container_width=True)

    with aba2:
        st.dataframe(estilizar_tabela(build_summary_dataframe(paciente, resultado)), use_container_width=True)

    with aba3:
        st.dataframe(estilizar_tabela(serie), use_container_width=True)

    with aba4:
        with st.expander("Abrir achados dos gráficos", expanded=False):
            st.markdown(
                f'<div class="findings-box"><ul>{montar_lista_html(build_curve_findings(serie))}</ul></div>',
                unsafe_allow_html=True,
            )
        st.pyplot(build_oscillation_figure(serie), use_container_width=True)
        st.pyplot(build_effort_figure(serie), use_container_width=True)

    with aba5:
        excel_bytes = build_excel_bytes(paciente, resultado, serie)
        pdf_bytes = build_pdf_bytes(paciente, resultado, serie)

        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "Baixar Excel estruturado",
                data=excel_bytes,
                file_name=build_safe_filename(paciente.nome, "xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "Baixar PDF clínico",
                data=pdf_bytes,
                file_name=build_safe_filename(paciente.nome, "pdf"),
                mime="application/pdf",
                use_container_width=True,
            )

    st.markdown("</div>", unsafe_allow_html=True)
