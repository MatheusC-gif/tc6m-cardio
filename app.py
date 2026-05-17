from datetime import date
import base64
from html import escape
from pathlib import Path
from unicodedata import normalize as normalize_unicode

import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

from tc6m import (
    FORMULAS_DPP,
    PatientData,
    build_curve_findings,
    build_default_during_table,
    build_default_pre_table,
    build_default_recovery_table,
    build_dp_recovery_figure,
    build_effort_figure,
    build_excel_bytes,
    build_integrated_recovery_analysis,
    build_oscillation_figure,
    build_pdf_bytes,
    build_report_payload,
    build_safe_filename,
    calcular_dpp_ben_saad,
    calcular_dpp_enright,
    calcular_dpp_iwama,
    calcular_dpp_por_formula,
    calcular_fc_maxima,
    calcular_fc_submaxima,
    calculate_tc6m_professional,
    combine_timeseries,
    format_analysis_value,
    normalize_timeseries,
)


st.set_page_config(
    page_title="Protocolo do TC6M",
    layout="wide",
    initial_sidebar_state="expanded",
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

TEST_PROFILES = ["Homem adulto", "Mulher adulta", "Criança/adolescente", "Paciente com DPOC"]

ICON_USER = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="M20 21a8 8 0 0 0-16 0"/>
  <circle cx="12" cy="7" r="4"/>
</svg>
"""

ICON_SHIELD_CHECK = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10Z"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="m9 12 2 2 4-4"/>
</svg>
"""

ICON_CALCULATOR = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <rect x="4" y="2" width="16" height="20" rx="2"/>
  <path stroke-linecap="round" d="M8 6h8"/>
  <path stroke-linecap="round" d="M8 10h.01M12 10h.01M16 10h.01M8 14h.01M12 14h.01M16 14h.01M8 18h.01M12 18h.01M16 18h.01"/>
</svg>
"""

ICON_FOOTPRINTS = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <circle cx="11" cy="3.8" r="1.9"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M8.4 7.2h3.2c.9 0 1.6.5 2 1.3l1 2.4 2.9 1.1"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M8.5 7.4 7.3 13c-.2.9-1.1 1.4-1.9 1.2-.8-.2-1.3-1-1.1-1.8l1.2-4.7c.2-.8.8-1.4 1.6-1.7"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M12.2 8.3 11 13.2 8.4 19.5"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M11 13.2 14 16l2.6 3.4"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M2.8 20h.01M5.6 20h.01M18.4 20h.01M21.2 20h.01"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M8.8 20h6.4"/>
</svg>
"""

ICON_REFRESH = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="M21 12a9 9 0 0 1-15.5 6.3"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M3 12a9 9 0 0 1 15.5-6.3"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M18 3v4h-4"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M6 21v-4h4"/>
</svg>
"""

ICON_GAUGE = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="M3.5 20.5h17"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M6.2 20.2v-4.1"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M10 20.2v-7.1"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M13.8 20.2v-5.7"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M17.6 20.2V9.4"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M4.8 15.6 8.4 12l4.5-.1 5.5-5.5"/>
  <circle cx="4.8" cy="15.6" r="1.3"/>
  <circle cx="8.4" cy="12" r="1.3"/>
  <circle cx="12.9" cy="11.9" r="1.3"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M15.2 6.4h3.2v3.2"/>
</svg>
"""

ICON_DROPLET = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="M12 2.5S5 10 5 15a7 7 0 0 0 14 0c0-5-7-12.5-7-12.5Z"/>
</svg>
"""

ICON_HEART = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="M20.8 4.6a5.5 5.5 0 0 0-7.8 0L12 5.6l-1-1a5.5 5.5 0 0 0-7.8 7.8l1 1L12 21l7.8-7.6 1-1a5.5 5.5 0 0 0 0-7.8Z"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M3 12h4l2-4 4 8 2-4h6"/>
</svg>
"""

ICON_LUNGS = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="M12 3v8"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M12 11c-2.3-2.4-4.6-3.6-6.4-2.4-1.6 1-2.4 3.5-2.4 7.4v2.2c0 2.2 2.4 3.5 4.3 2.4 2.2-1.3 3.4-4.6 4.5-9.6Z"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M12 11c2.3-2.4 4.6-3.6 6.4-2.4 1.6 1 2.4 3.5 2.4 7.4v2.2c0 2.2-2.4 3.5-4.3 2.4-2.2-1.3-3.4-4.6-4.5-9.6Z"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M9.5 12.5c-1.4 1.2-2 3.2-2 5M14.5 12.5c1.4 1.2 2 3.2 2 5"/>
</svg>
"""

ICON_CHECK = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <circle cx="12" cy="12" r="9"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="m8.5 12.5 2.5 2.5 4.5-5"/>
</svg>
"""

ICON_TREND_UP = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="m3 17 6-6 4 4 8-8"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M14 7h7v7"/>
</svg>
"""

ICON_TREND_DOWN = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="m3 7 6 6 4-4 8 8"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M14 17h7v-7"/>
</svg>
"""

ICON_USERS = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/>
  <circle cx="9" cy="7" r="4"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M22 21v-2a4 4 0 0 0-3-3.9"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M16 3.1a4 4 0 0 1 0 7.8"/>
</svg>
"""

ICON_TARGET = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <circle cx="12" cy="12" r="9"/>
  <circle cx="12" cy="12" r="5"/>
  <circle cx="12" cy="12" r="1"/>
</svg>
"""

ICON_CLIPBOARD = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <rect x="8" y="2" width="8" height="4" rx="1"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/>
  <path stroke-linecap="round" d="M8 12h8M8 16h5"/>
</svg>
"""

ICON_INFO = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <circle cx="12" cy="12" r="9"/>
  <path stroke-linecap="round" d="M12 11v5"/>
  <path stroke-linecap="round" d="M12 8h.01"/>
</svg>
"""

ICON_MENU = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" d="M12 6h.01M12 12h.01M12 18h.01"/>
</svg>
"""

ICON_STAR = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="m12 3 2.6 5.3 5.9.9-4.3 4.1 1 5.8L12 16.4l-5.2 2.7 1-5.8-4.3-4.1 5.9-.9L12 3Z"/>
</svg>
"""

ICON_ACTIVITY = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <circle cx="12" cy="12" r="9"/>
  <path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12h4l2.1-4.5 3.4 9 2-4.5h3.5"/>
</svg>
"""

ICON_SHIELD = """
<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
  <path stroke-linecap="round" stroke-linejoin="round" d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10Z"/>
</svg>
"""

ASSET_ICON_DIR = Path(__file__).resolve().parent / "assets" / "icons"

ICON_MAP = {
    "achados_automaticos": ASSET_ICON_DIR / "achados_automaticos.png",
    "caminhada_tc6m": ASSET_ICON_DIR / "caminhada_tc6m.png",
    "classificacao_risco": ASSET_ICON_DIR / "classificacao_risco.png",
    "distancia_percorrida": ASSET_ICON_DIR / "distancia_percorrida.png",
    "duplo_produto_recuperacao": ASSET_ICON_DIR / "duplo_produto_recuperacao.png",
    "fator_limitante": ASSET_ICON_DIR / "fator_limitante.png",
    "graficos_achados": ASSET_ICON_DIR / "graficos_achados.png",
    "metricas_hemodinamicas": ASSET_ICON_DIR / "metricas_hemodinamicas.png",
    "interpretacao_integrada": ASSET_ICON_DIR / "interpretacao_integrada.png",
    "nota_metodologica": ASSET_ICON_DIR / "nota_metodologica.png",
    "pontos_atencao": ASSET_ICON_DIR / "pontos_atencao.png",
    "predicoes_comparativas": ASSET_ICON_DIR / "predicoes_comparativas.png",
    "resumo_clinico": ASSET_ICON_DIR / "resumo_clinico.png",
    "spo2": ASSET_ICON_DIR / "spo2.png",
}


def load_asset_icon(icon_key: str, fallback: str) -> str:
    """Carrega PNG oficial do sistema para uso inline no Streamlit."""

    icon_path = ICON_MAP.get(icon_key)
    if icon_path and icon_path.exists():
        encoded = base64.b64encode(icon_path.read_bytes()).decode("ascii")
        alt = escape(icon_key.replace("_", " "))
        return f'<img class="official-icon" src="data:image/png;base64,{encoded}" alt="{alt}">'
    return fallback


ICON_USER = load_asset_icon("caminhada_tc6m", ICON_USER)
ICON_SHIELD_CHECK = load_asset_icon("classificacao_risco", ICON_SHIELD_CHECK)
ICON_CALCULATOR = load_asset_icon("predicoes_comparativas", ICON_CALCULATOR)
ICON_FOOTPRINTS = load_asset_icon("caminhada_tc6m", ICON_FOOTPRINTS)
ICON_REFRESH = load_asset_icon("duplo_produto_recuperacao", ICON_REFRESH)
ICON_GAUGE = load_asset_icon("graficos_achados", ICON_GAUGE)
ICON_DROPLET = load_asset_icon("spo2", ICON_DROPLET)
ICON_HEART = load_asset_icon("metricas_hemodinamicas", ICON_HEART)
ICON_LUNGS = load_asset_icon("metricas_hemodinamicas", ICON_LUNGS)
ICON_CHECK = load_asset_icon("classificacao_risco", ICON_CHECK)
ICON_TREND_UP = load_asset_icon("graficos_achados", ICON_TREND_UP)
ICON_TREND_DOWN = load_asset_icon("graficos_achados", ICON_TREND_DOWN)
ICON_USERS = load_asset_icon("fator_limitante", ICON_USERS)
ICON_TARGET = load_asset_icon("interpretacao_integrada", ICON_TARGET)
ICON_CLIPBOARD = load_asset_icon("resumo_clinico", ICON_CLIPBOARD)
ICON_INFO = load_asset_icon("nota_metodologica", ICON_INFO)
ICON_MENU = load_asset_icon("pontos_atencao", ICON_MENU)
ICON_STAR = load_asset_icon("classificacao_risco", ICON_STAR)
ICON_ACTIVITY = load_asset_icon("classificacao_risco", ICON_ACTIVITY)
ICON_SHIELD = load_asset_icon("classificacao_risco", ICON_SHIELD)


def clean_text(value: object) -> str:
    return str(value or "")


def br_number(value: float, decimals: int = 2) -> str:
    return f"{float(value):.{decimals}f}".replace(".", ",")


def remove_accents(text: str) -> str:
    return normalize_unicode("NFKD", text).encode("ascii", "ignore").decode("ascii")


def patient_initials(name: str) -> str:
    clean_name = remove_accents(name).upper()
    parts = ["".join(char for char in part if char.isalnum()) for part in clean_name.split()]
    parts = [part for part in parts if part]
    if not parts:
        return "PAC"
    if len(parts) == 1:
        return parts[0][:3].ljust(3, "X")
    return "".join(part[0] for part in parts[:3])


def generate_evaluation_id(name: str, evaluation_date: date, age: int, test_number: int) -> str:
    date_code = evaluation_date.strftime("%Y%m%d")
    return f"TC6M-{date_code}-{patient_initials(name)}-{max(int(age), 0)}-{max(int(test_number), 1):02d}"


def borg_resp_col(df: pd.DataFrame) -> str:
    return next((column for column in df.columns if "Borg" in column and "MMII" not in column), "Borg Respiratório")


def borg_mmii_col(df: pd.DataFrame) -> str:
    return next((column for column in df.columns if "Borg" in column and "MMII" in column), "Borg MMII")


def clear_result() -> None:
    st.session_state.resultado_tc6m = None
    st.session_state.paciente_tc6m = None
    st.session_state.serie_tc6m = None


def reset_patient_progress() -> None:
    clear_result()
    st.session_state.patient_saved = False
    st.session_state.prediction_saved = False
    st.session_state.execution_saved = False
    st.session_state.recovery_saved = False


def update_evaluation_id(force: bool = False) -> None:
    if force:
        st.session_state.id_auto_ativo = True
    if not st.session_state.get("id_auto_ativo", True):
        reset_patient_progress()
        return
    if not st.session_state.get("nome", "").strip():
        reset_patient_progress()
        return
    st.session_state.prontuario = generate_evaluation_id(
        st.session_state.nome,
        st.session_state.data_avaliacao,
        int(st.session_state.idade),
        int(st.session_state.numero_teste),
    )
    reset_patient_progress()


def force_evaluation_id() -> None:
    update_evaluation_id(force=True)


def apply_pending_identification_actions() -> None:
    should_force_id = st.session_state.get("pending_force_id", False)
    should_save_patient = st.session_state.get("pending_patient_save", False)

    if should_force_id:
        st.session_state.pending_force_id = False
        st.session_state.id_auto_ativo = True
        if st.session_state.get("nome", "").strip():
            st.session_state.prontuario = generate_evaluation_id(
                st.session_state.nome,
                st.session_state.data_avaliacao,
                int(st.session_state.idade),
                int(st.session_state.numero_teste),
            )

    if should_save_patient:
        st.session_state.pending_patient_save = False
        clear_result()
        st.session_state.patient_saved = True
        st.session_state.prediction_saved = False
        st.session_state.execution_saved = False
        st.session_state.recovery_saved = False


def mark_manual_id() -> None:
    if st.session_state.get("prontuario", "").strip():
        st.session_state.id_auto_ativo = False
    reset_patient_progress()


def set_triage(status: str) -> None:
    st.session_state.triagem_status = status
    st.session_state.contra_abs = status == "Contraindicação absoluta"
    st.session_state.contra_rel = status == "Contraindicação relativa"
    st.session_state.prediction_saved = False
    st.session_state.execution_saved = False
    st.session_state.recovery_saved = False
    clear_result()


def parse_quick_bp(value: object) -> tuple[int, int]:
    text = str(value or "").lower().strip()
    digits = "".join(char for char in text if char.isdigit())
    if not digits:
        return 0, 0
    if any(separator in text for separator in ["/", "x", "por", "-", " "]):
        normalized = text.replace("por", " ").replace("/", " ").replace("x", " ").replace("-", " ")
        parts = [part for part in normalized.split() if part.isdigit()]
        if len(parts) >= 2:
            pas = int(parts[0]) * 10 if int(parts[0]) < 30 else int(parts[0])
            pad = int(parts[1]) * 10 if int(parts[1]) < 30 else int(parts[1])
            return pas, pad
    if len(digits) == 3:
        return int(digits[:2]) * 10, int(digits[2]) * 10
    if len(digits) == 4:
        return int(digits[:2]) * 10, int(digits[2:]) * 10
    if len(digits) == 5:
        return int(digits[:3]), int(digits[3:])
    if len(digits) >= 6:
        return int(digits[:3]), int(digits[3:5])
    return 0, 0


def format_bp(pas: int | float, pad: int | float) -> str:
    if float(pas) <= 0 or float(pad) <= 0:
        return ""
    return f"{int(round(float(pas)))}/{int(round(float(pad)))}"


def prepare_editor_with_bp(df: pd.DataFrame) -> pd.DataFrame:
    table = normalize_timeseries(df)
    resp_col = borg_resp_col(table)
    mmii_col = borg_mmii_col(table)
    table["PA"] = table.apply(lambda row: format_bp(row["PAS"], row["PAD"]), axis=1)
    return table[["Tempo", "FC", "SpO2", "FR", "PA", resp_col, mmii_col]]


def restore_pas_pad(df: pd.DataFrame) -> pd.DataFrame:
    resp_col = borg_resp_col(df)
    mmii_col = borg_mmii_col(df)
    rows = []
    for _, row in df.iterrows():
        pas, pad = parse_quick_bp(row.get("PA", ""))
        rows.append(
            {
                "Tempo": row["Tempo"],
                "FC": row["FC"],
                "SpO2": row["SpO2"],
                "FR": row["FR"],
                "PAS": pas,
                "PAD": pad,
                resp_col: row[resp_col],
                mmii_col: row[mmii_col],
            }
        )
    return normalize_timeseries(pd.DataFrame(rows))


def display_vitals_table(df: pd.DataFrame, include_full: bool = True) -> pd.DataFrame:
    table = normalize_timeseries(df)
    resp_col = borg_resp_col(table)
    mmii_col = borg_mmii_col(table)
    output = pd.DataFrame(
        {
            "Momento" if include_full else "Minuto": table["Tempo"],
            "FC (bpm)": table["FC"],
            "SpO2 (%)": table["SpO2"],
        }
    )
    if include_full:
        output["FR (ipm)"] = table["FR"]
        output["PA (mmHg)"] = [format_bp(pas, pad) for pas, pad in zip(table["PAS"], table["PAD"])]
    output["Borg dispneia"] = table[resp_col].map(lambda item: f"{float(item):.1f}")
    output["Borg MMII"] = table[mmii_col].map(lambda item: f"{float(item):.1f}")
    return output


def dataframe_to_table(df: pd.DataFrame) -> str:
    header = "".join(f"<th>{escape(str(column))}</th>" for column in df.columns)
    rows = []
    for _, row in df.iterrows():
        cells = "".join(f"<td>{escape(clean_text(value))}</td>" for value in row)
        rows.append(f"<tr>{cells}</tr>")
    return f'<table class="compact-table"><thead><tr>{header}</tr></thead><tbody>{"".join(rows)}</tbody></table>'


def init_state() -> None:
    defaults = {
        "nome": "",
        "prontuario": "",
        "numero_teste": 1,
        "id_auto_ativo": True,
        "avaliador": "",
        "diagnostico": "",
        "data_avaliacao": date.today(),
        "sexo_label": "Masculino",
        "idade": 60,
        "peso": 70.0,
        "altura_cm": 170.0,
        "comprimento_membro_inferior_m": 0.0,
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
        "nav_section": "avaliacao",
        "sidebar_compact": False,
        "patient_saved": False,
        "prediction_saved": False,
        "execution_saved": False,
        "recovery_saved": False,
        "pending_force_id": False,
        "pending_patient_save": False,
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)
    st.session_state.setdefault("pre_df", build_default_pre_table())
    st.session_state.setdefault("during_df", build_default_during_table())
    st.session_state.setdefault("recovery_df", build_default_recovery_table())


def fill_demo_profile(profile: str) -> None:
    profiles = {
        "Homem adulto": {
            "nome": "Homem",
            "prontuario": "TC6M-H-001",
            "sexo_label": "Masculino",
            "idade": 62,
            "peso": 74.0,
            "altura_cm": 171.0,
            "comprimento_membro_inferior_m": 0.90,
            "formula": FORMULAS_DPP[0],
            "diagnostico": "Avaliação funcional cardiorrespiratória",
            "distancia": 438.0,
            "pre": [78, 97, 18, 122, 78, 0.0, 0.0],
            "fc": [94, 106, 116, 124, 132, 139],
            "spo2": [96, 95, 95, 94, 93, 92],
            "borg_r": [1, 2, 3, 4, 5, 6],
            "borg_m": [1, 2, 3, 4, 5, 5],
            "rec": [[120, 94, 24, 146, 82, 4, 4], [98, 96, 20, 132, 80, 2, 2], [84, 97, 18, 124, 78, 1, 1]],
        },
        "Mulher adulta": {
            "nome": "Mulher",
            "prontuario": "TC6M-M-001",
            "sexo_label": "Feminino",
            "idade": 55,
            "peso": 66.0,
            "altura_cm": 160.0,
            "comprimento_membro_inferior_m": 0.84,
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
            "nome": "Adolescente",
            "prontuario": "TC6M-C-001",
            "sexo_label": "Masculino",
            "idade": 12,
            "peso": 42.0,
            "altura_cm": 150.0,
            "comprimento_membro_inferior_m": 0.78,
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
            "nome": "Paciente DPOC",
            "prontuario": "TC6M-DPOC-001",
            "sexo_label": "Masculino",
            "idade": 68,
            "peso": 69.0,
            "altura_cm": 168.0,
            "comprimento_membro_inferior_m": 0.88,
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
    data = profiles[profile]
    st.session_state.nome = data["nome"]
    st.session_state.prontuario = data["prontuario"]
    st.session_state.numero_teste = 1
    st.session_state.id_auto_ativo = True
    st.session_state.avaliador = "Equipe Cardiorrespiratória"
    st.session_state.diagnostico = data["diagnostico"]
    st.session_state.data_avaliacao = date.today()
    st.session_state.sexo_label = data["sexo_label"]
    st.session_state.idade = data["idade"]
    st.session_state.peso = data["peso"]
    st.session_state.altura_cm = data["altura_cm"]
    st.session_state.comprimento_membro_inferior_m = data["comprimento_membro_inferior_m"]
    st.session_state.formula_principal = data["formula"]
    st.session_state.triagem_status = "Sem contraindicações"
    st.session_state.contra_abs = False
    st.session_state.contra_rel = False
    st.session_state.observacao_triagem = "Paciente assintomático. Sem queixas ou limitações relatadas. Apto para realização do exame."
    st.session_state.distancia = data["distancia"]
    st.session_state.interrompeu_label = "Não"
    st.session_state.motivo_interrupcao = ""
    st.session_state.distancia_interrupcao = 0.0

    pre = data["pre"]
    pre_base = build_default_pre_table()
    st.session_state.pre_df = normalize_timeseries(
        pre_base.assign(
            FC=[pre[0]],
            SpO2=[pre[1]],
            FR=[pre[2]],
            PAS=[pre[3]],
            PAD=[pre[4]],
            **{borg_resp_col(pre_base): [pre[5]], borg_mmii_col(pre_base): [pre[6]]},
        )
    )
    during_base = build_default_during_table()
    st.session_state.during_df = during_base.assign(
        FC=data["fc"],
        SpO2=data["spo2"],
        **{borg_resp_col(during_base): data["borg_r"], borg_mmii_col(during_base): data["borg_m"]},
    )
    recovery_base = build_default_recovery_table()
    st.session_state.recovery_df = normalize_timeseries(
        recovery_base.assign(
            FC=[row[0] for row in data["rec"]],
            SpO2=[row[1] for row in data["rec"]],
            FR=[row[2] for row in data["rec"]],
            PAS=[row[3] for row in data["rec"]],
            PAD=[row[4] for row in data["rec"]],
            **{
                borg_resp_col(recovery_base): [row[5] for row in data["rec"]],
                borg_mmii_col(recovery_base): [row[6] for row in data["rec"]],
            },
        )
    )
    st.session_state.patient_saved = True
    st.session_state.prediction_saved = True
    st.session_state.execution_saved = True
    st.session_state.recovery_saved = True
    clear_result()


def build_patient() -> PatientData:
    sex = "M" if st.session_state.sexo_label == "Masculino" else "F"
    return PatientData(
        nome=st.session_state.nome,
        prontuario=st.session_state.prontuario,
        data_avaliacao=st.session_state.data_avaliacao,
        avaliador=st.session_state.avaliador,
        diagnostico=st.session_state.diagnostico,
        sexo=sex,
        idade=int(st.session_state.idade),
        peso=float(st.session_state.peso),
        altura_cm=float(st.session_state.altura_cm),
        comprimento_membro_inferior_m=float(st.session_state.comprimento_membro_inferior_m) or None,
        distancia=float(st.session_state.distancia),
        formula_principal=st.session_state.formula_principal,
        interrompeu=st.session_state.interrompeu_label == "Sim",
        motivo_interrupcao=st.session_state.motivo_interrupcao,
        distancia_interrupcao=float(st.session_state.distancia_interrupcao),
        contraindicacao_absoluta=bool(st.session_state.contra_abs),
        contraindicacao_relativa=bool(st.session_state.contra_rel),
        observacao_triagem=st.session_state.observacao_triagem,
    )


def generate_result() -> None:
    series = combine_timeseries(st.session_state.pre_df, st.session_state.during_df, st.session_state.recovery_df)
    patient = build_patient()
    result = calculate_tc6m_professional(patient, series)
    st.session_state.paciente_tc6m = patient
    st.session_state.resultado_tc6m = result
    st.session_state.serie_tc6m = series


def inject_css() -> None:
    st.markdown(
        """
        <style>
            :root {
                --bg-page: #F3F7F5;
                --card-bg: #FFFFFF;
                --primary: #064C3B;
                --primary-dark: #02382D;
                --primary-soft: #E6F3EE;
                --text-main: #10201C;
                --text-muted: #66736F;
                --border: #DDE7E3;
                --success: #0E7A4F;
                --warning: #D98B18;
                --danger: #D94A4A;
                --blue-soft: #E9F2FF;
                --shadow: 0 8px 24px rgba(0,0,0,0.06);
            }

            .stApp {
                background: radial-gradient(circle at top left, rgba(6,76,59,0.07), transparent 30%), var(--bg-page);
                color: var(--text-main);
                font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            }

            [data-testid="stAppViewContainer"] > .main .block-container {
                max-width: 1160px;
                margin-left: auto;
                margin-right: auto;
                padding: 2rem 2rem 3rem;
            }

            [data-testid="stToolbar"],
            [data-testid="stDecoration"],
            [data-testid="stStatusWidget"] {
                display: none !important;
            }

            header[data-testid="stHeader"] {
                background: transparent !important;
                box-shadow: none !important;
                height: 3rem !important;
            }

            [data-testid="stSidebarCollapsedControl"] {
                align-items: center !important;
                background: #064C3B !important;
                border: 1px solid rgba(255,255,255,.35) !important;
                border-radius: 999px !important;
                box-shadow: 0 10px 24px rgba(2,56,45,.22) !important;
                display: flex !important;
                height: 38px !important;
                justify-content: center !important;
                left: 16px !important;
                opacity: 1 !important;
                position: fixed !important;
                top: 14px !important;
                visibility: visible !important;
                width: 38px !important;
                z-index: 999999 !important;
            }

            [data-testid="stSidebarCollapsedControl"] button {
                background: transparent !important;
                border: 0 !important;
                color: #FFFFFF !important;
                min-height: 0 !important;
                padding: 0 !important;
            }

            [data-testid="stSidebarCollapsedControl"] svg {
                color: #FFFFFF !important;
                fill: none !important;
                stroke: #FFFFFF !important;
                stroke-width: 2.8 !important;
            }

            [data-testid="stSidebar"] {
                background:
                    radial-gradient(circle at 18% 14%, rgba(34, 197, 94, .25), transparent 28%),
                    linear-gradient(180deg, #064C3B 0%, #02382D 100%) !important;
                border-right: 0 !important;
            }

            [data-testid="stSidebar"] * {
                color: #FFFFFF;
            }

            [data-testid="stSidebar"] button {
                background: rgba(255,255,255,.08) !important;
                border: 1px solid rgba(255,255,255,.12) !important;
                border-radius: 12px !important;
                color: #FFFFFF !important;
                font-weight: 850 !important;
                min-height: 44px;
            }

            [data-testid="stSidebar"] button:hover {
                background: rgba(255,255,255,.16) !important;
                border-color: rgba(255,255,255,.22) !important;
            }

            .protocol-header {
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                gap: 20px;
                margin: 4px 0 24px;
            }

            .protocol-title {
                color: #07172D !important;
                font-size: 2.25rem;
                font-weight: 900;
                line-height: 1.05;
                margin: 0;
            }

            .protocol-subtitle {
                color: var(--text-muted) !important;
                font-size: .98rem;
                font-weight: 600;
                margin-top: 10px;
            }

            .protocol-actions {
                display: flex;
                gap: 12px;
                align-items: center;
                color: #1F2E43 !important;
                font-weight: 800;
            }

            .date-pill, .export-pill {
                background: rgba(255,255,255,.9);
                border: 1px solid #C8D5D1;
                border-radius: 10px;
                padding: 10px 15px;
                box-shadow: 0 8px 20px rgba(6,76,59,.05);
            }

            .stepper {
                display: grid;
                grid-template-columns: repeat(6, minmax(0, 1fr));
                margin: 0 34px 32px;
            }

            .step-item {
                position: relative;
                text-align: center;
            }

            .step-item::before {
                content: "";
                position: absolute;
                top: 22px;
                left: 0;
                right: 0;
                height: 2px;
                background: #CCD7D3;
                z-index: 0;
            }

            .step-item:first-child::before { left: 50%; }
            .step-item:last-child::before { right: 50%; }

            .step-number {
                position: relative;
                z-index: 1;
                width: 46px;
                height: 46px;
                margin: 0 auto 10px;
                display: grid;
                place-items: center;
                border-radius: 50%;
                background: #FFFFFF;
                border: 1px solid #CEDAD6;
                color: #203248 !important;
                font-weight: 900;
                box-shadow: 0 8px 20px rgba(6,76,59,.08);
            }

            .step-item.active .step-number {
                background: var(--primary);
                border-color: var(--primary);
                color: #FFFFFF !important;
            }

            .step-label {
                color: #24384E !important;
                font-size: .86rem;
                font-weight: 850;
            }

            .clinical-card {
                background: var(--card-bg);
                border: 1px solid var(--border);
                border-radius: 18px;
                box-shadow: var(--shadow);
                margin-bottom: 16px;
                padding: 24px;
            }

            .app-sidebar {
                background:
                    radial-gradient(circle at 18% 14%, rgba(34, 197, 94, .25), transparent 28%),
                    linear-gradient(180deg, #064C3B 0%, #02382D 100%);
                bottom: 0;
                box-shadow: 14px 0 34px rgba(2, 56, 45, .16);
                color: #FFFFFF;
                display: flex;
                flex-direction: column;
                left: 0;
                padding: 28px 16px;
                position: fixed;
                top: 0;
                width: 220px;
                z-index: 999;
            }

            .sidebar-brand {
                align-items: center;
                display: flex;
                gap: 12px;
                margin-bottom: 34px;
            }

            .sidebar-brand-icon {
                color: #8AD879;
                height: 42px;
                width: 42px;
            }

            .sidebar-brand-title {
                font-size: 1.55rem;
                font-weight: 950;
                line-height: 1;
            }

            .sidebar-brand-sub {
                color: rgba(255,255,255,.78);
                font-size: .78rem;
                font-weight: 650;
                margin-top: 4px;
            }

            .sidebar-menu {
                display: grid;
                gap: 10px;
            }

            .sidebar-item {
                align-items: center;
                border-radius: 12px;
                color: rgba(255,255,255,.9);
                display: flex;
                font-weight: 850;
                gap: 11px;
                padding: 13px 12px;
                text-decoration: none !important;
            }

            .sidebar-item.active {
                background: rgba(255,255,255,.13);
                box-shadow: inset 4px 0 0 #35C979;
            }

            .sidebar-item svg {
                height: 20px;
                stroke-width: 2.8 !important;
                width: 20px;
            }

            .official-icon {
                display: inline-block;
                height: 22px;
                object-fit: contain;
                vertical-align: middle;
                width: 22px;
            }

            .stage-icon .official-icon,
            .class-icon .official-icon,
            .quick-icon .official-icon,
            .monitor-icon .official-icon,
            .finding-icon .official-icon,
            .finding-row-icon .official-icon,
            .analytics-title-icon .official-icon,
            .graph-metric-icon .official-icon,
            .chart-note-icon .official-icon,
            .result-tile-label .official-icon {
                height: 100%;
                max-height: none;
                object-fit: contain;
                width: 100%;
            }

            .stage-icon:has(.official-icon),
            .class-icon:has(.official-icon),
            .quick-icon:has(.official-icon) {
                background: transparent !important;
                box-shadow: none !important;
            }

            .stage-icon .official-icon { height: 40px; width: 40px; }
            .class-icon .official-icon { height: 56px; width: 56px; }
            .quick-icon .official-icon { height: 50px; width: 50px; }
            .monitor-icon .official-icon { height: 42px; width: 42px; }
            .finding-icon .official-icon { height: 56px; width: 56px; }
            .finding-row-icon .official-icon { height: 48px; width: 48px; }
            .graph-metric-icon .official-icon { height: 54px; width: 54px; }
            .analytics-title-icon .official-icon { height: 28px; width: 28px; }
            .chart-note-icon .official-icon { height: 24px; width: 24px; }
            .result-tile-label .official-icon { height: 24px; width: 24px; }

            .sidebar-new {
                display: block;
                background: linear-gradient(135deg, #24B96F, #0D7A4F);
                border-radius: 12px;
                box-shadow: 0 14px 30px rgba(0,0,0,.15);
                color: #FFFFFF;
                font-weight: 900;
                margin-top: 28px;
                padding: 14px 13px;
                text-align: center;
                text-decoration: none !important;
            }

            .sidebar-bottom {
                display: grid;
                gap: 12px;
                margin-top: auto;
            }

            .sidebar-toggle {
                align-items: center;
                background: rgba(255,255,255,.12);
                border: 1px solid rgba(255,255,255,.14);
                border-radius: 999px;
                color: #FFFFFF !important;
                display: flex;
                font-size: 1.25rem;
                font-weight: 950;
                height: 30px;
                justify-content: center;
                position: absolute;
                right: -14px;
                text-decoration: none !important;
                top: 22px;
                width: 30px;
            }

            .dashboard-grid-2 {
                display: grid;
                grid-template-columns: .74fr 1.26fr;
                gap: 16px;
                margin-bottom: 16px;
            }

            .dashboard-grid-even {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 16px;
                margin-bottom: 16px;
            }

            .dashboard-card {
                background: #FFFFFF;
                border: 1px solid var(--border);
                border-radius: 16px;
                box-shadow: var(--shadow);
                padding: 20px;
            }

            .dashboard-title {
                align-items: center;
                color: var(--primary);
                display: flex;
                gap: 10px;
                font-size: 1.04rem;
                font-weight: 950;
                margin-bottom: 16px;
            }

            .dashboard-step {
                align-items: center;
                background: var(--primary);
                border-radius: 8px;
                color: #FFFFFF;
                display: inline-flex;
                font-size: .88rem;
                font-weight: 950;
                height: 28px;
                justify-content: center;
                min-width: 28px;
                padding: 0 8px;
            }

            .dashboard-title-muted {
                color: #64748B;
                font-size: .9rem;
                font-weight: 750;
            }

            .patient-card {
                align-items: stretch;
                display: grid;
                grid-template-columns: 245px 1fr;
                gap: 18px;
            }

            .patient-primary {
                align-items: center;
                border-right: 1px solid var(--border);
                display: grid;
                gap: 12px;
                grid-template-columns: 64px 1fr;
                padding-right: 18px;
            }

            .patient-avatar {
                align-items: center;
                background: var(--primary);
                border-radius: 999px;
                color: #FFFFFF;
                display: flex;
                height: 58px;
                justify-content: center;
                width: 58px;
            }

            .patient-avatar svg {
                height: 34px;
                stroke-width: 2.9 !important;
                width: 34px;
            }

            .patient-grid {
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 0;
            }

            .patient-info {
                border-left: 1px solid var(--border);
                min-height: 68px;
                padding: 10px 16px;
            }

            .patient-label, .mini-label {
                color: var(--text-muted);
                font-size: .76rem;
                font-weight: 850;
                margin-bottom: 7px;
            }

            .patient-value, .mini-value {
                color: #07172D;
                font-size: .98rem;
                font-weight: 950;
                line-height: 1.2;
            }

            .triage-stack {
                display: grid;
                gap: 10px;
            }

            .triage-stack .triage-pill {
                min-height: 42px;
            }

            .prediction-tiles {
                display: grid;
                grid-template-columns: repeat(3, minmax(0, 1fr));
                gap: 12px;
            }

            .prediction-tile {
                border: 1px solid var(--border);
                border-radius: 12px;
                min-height: 98px;
                padding: 16px;
            }

            .prediction-more {
                align-items: center;
                border: 1px solid var(--border);
                border-radius: 12px;
                color: var(--primary);
                display: flex;
                font-weight: 900;
                justify-content: space-between;
                margin-top: 12px;
                padding: 13px 16px;
            }

            .result-band {
                background:
                    radial-gradient(circle at 12% 18%, rgba(54, 211, 153, .22), transparent 30%),
                    radial-gradient(circle at 86% 12%, rgba(36, 185, 111, .16), transparent 28%),
                    linear-gradient(135deg, #075B46 0%, #034434 48%, #012A23 100%);
                border-radius: 16px;
                box-shadow: 0 18px 40px rgba(2, 56, 45, .18);
                color: #FFFFFF;
                margin-bottom: 16px;
                padding: 18px;
            }

            .result-band-title {
                align-items: center;
                display: flex;
                gap: 10px;
                font-size: 1.28rem;
                font-weight: 950;
                margin-bottom: 16px;
            }

            .result-band .dashboard-step {
                background: #1DB879;
            }

            .result-band-grid {
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 14px;
            }

            .result-tile {
                background: linear-gradient(180deg, rgba(255,255,255,.12), rgba(255,255,255,.055));
                border: 1px solid rgba(255,255,255,.18);
                border-radius: 14px;
                min-height: 175px;
                padding: 18px;
                box-shadow: inset 0 1px 0 rgba(255,255,255,.08), 0 14px 28px rgba(0,0,0,.08);
            }

            .result-tile-label {
                align-items: center;
                color: rgba(255,255,255,.86);
                display: flex;
                font-size: .8rem;
                font-weight: 850;
                gap: 8px;
                margin-bottom: 14px;
            }

            .result-tile-label svg {
                color: currentColor;
                height: 22px;
                stroke-width: 2.9 !important;
                width: 22px;
            }

            .result-tile:nth-child(1) .result-tile-label { color: #7EE37E !important; }
            .result-tile:nth-child(2) .result-tile-label { color: #BEEA2E !important; }
            .result-tile:nth-child(3) .result-tile-label { color: #D5EA2E !important; }
            .result-tile:nth-child(4) .result-tile-label { color: #FFB22D !important; }

            .result-main-number {
                color: #FFFFFF;
                font-size: 2.9rem;
                font-weight: 950;
                line-height: 1;
            }

            .result-main-number span {
                font-size: 1.25rem;
            }

            .result-accent {
                color: #FFB22D;
                font-size: 1.8rem;
                font-weight: 950;
                margin-top: 14px;
            }

            .result-progress-mini {
                background: rgba(255,255,255,.18);
                border-radius: 999px;
                height: 8px;
                margin: 14px 0 8px;
                overflow: hidden;
            }

            .result-progress-mini span {
                background: linear-gradient(90deg, #FFB22D, #FF8A00);
                border-radius: inherit;
                display: block;
                height: 100%;
                width: var(--progress);
            }

            .risk-scale-mini {
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 6px;
                margin-top: 18px;
                position: relative;
            }

            .risk-scale-mini span {
                border-radius: 999px;
                height: 8px;
            }

            .risk-scale-mini span:nth-child(1) { background: #CDEB37; }
            .risk-scale-mini span:nth-child(2) { background: #FFB22D; }
            .risk-scale-mini span:nth-child(3) { background: #FF6C44; }
            .risk-scale-mini span:nth-child(4) { background: #FF4158; }

            .risk-scale-mini::after {
                background: #FFB22D;
                border: 2px solid rgba(255,255,255,.85);
                border-radius: 999px;
                bottom: -9px;
                box-shadow: 0 4px 10px rgba(0,0,0,.18);
                content: "";
                height: 10px;
                left: var(--risk-marker, 37.5%);
                position: absolute;
                transform: translateX(-50%);
                width: 10px;
            }

            .monitor-row {
                display: grid;
                grid-template-columns: repeat(5, minmax(0, 1fr));
                gap: 14px;
            }

            .monitor-card {
                align-items: center;
                border: 1px solid var(--border);
                border-radius: 12px;
                display: flex;
                gap: 12px;
                min-height: 84px;
                padding: 14px;
            }

            .monitor-icon {
                color: #64748B;
                flex: 0 0 auto;
            }

            .monitor-icon svg {
                height: 28px;
                stroke-width: 2.65 !important;
                width: 28px;
            }

            .chart-card.compact-chart {
                margin-bottom: 0;
                padding: 18px;
            }

            .attention-list {
                display: grid;
                gap: 10px;
                margin-top: 14px;
            }

            .attention-row {
                align-items: center;
                border-bottom: 1px solid var(--border);
                display: flex;
                gap: 10px;
                justify-content: space-between;
                padding: 8px 0;
            }

            .attention-row:last-child {
                border-bottom: none;
            }

            .attention-badge {
                border-radius: 8px;
                font-size: .74rem;
                font-weight: 900;
                padding: 6px 9px;
            }

            .attention-badge.ok { background: #E7F8EC; color: #0E7A4F; }
            .attention-badge.warning { background: #FFF1DB; color: #B65314; }
            .attention-badge.danger { background: #FFE1E1; color: #B42318; }

            .summary-quote {
                border-left: 5px solid #22A86F;
                background: linear-gradient(90deg, rgba(34,168,111,.09), rgba(34,168,111,0));
                border-radius: 12px;
                padding: 14px 16px;
            }

            .protocol-footnote {
                align-items: center;
                background: #EAF4EF;
                border: 1px solid #CFE3DA;
                border-radius: 14px;
                color: #24483D;
                display: flex;
                font-size: .9rem;
                font-weight: 750;
                gap: 12px;
                margin: 18px 0;
                padding: 16px 18px;
            }

            .stage-header {
                display: flex;
                align-items: center;
                gap: 12px;
                margin-bottom: 22px;
            }

            .stage-number {
                width: 34px;
                height: 34px;
                border-radius: 50%;
                background: var(--primary);
                color: #FFFFFF !important;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 15px;
                font-weight: 800;
                flex: 0 0 auto;
            }

            .stage-icon {
                width: 40px;
                height: 40px;
                border-radius: 999px;
                background: var(--primary-soft);
                color: var(--primary) !important;
                display: flex;
                align-items: center;
                justify-content: center;
                flex: 0 0 auto;
                box-shadow: inset 0 0 0 1px rgba(6, 76, 59, .12);
            }

            .stage-icon svg {
                width: 24px;
                height: 24px;
                stroke-width: 2.85 !important;
                color: currentColor;
            }

            .stage-title {
                color: var(--text-main) !important;
                font-size: 18px;
                font-weight: 800;
                line-height: 1.25;
                margin: 0;
            }

            .info-grid {
                border: 1px solid var(--border);
                border-radius: 14px;
                display: grid;
                grid-template-columns: repeat(5, minmax(0, 1fr));
                overflow: hidden;
            }

            .info-field {
                border-right: 1px solid var(--border);
                border-bottom: 1px solid var(--border);
                min-height: 76px;
                padding: 16px 18px;
            }

            .info-field:nth-child(5), .info-field:nth-child(9) {
                border-right: none;
            }

            .info-field.wide {
                grid-column: span 2;
            }

            .info-label, .table-label, .metric-label, .class-label {
                color: var(--text-muted) !important;
                font-size: .78rem;
                font-weight: 800;
                margin-bottom: 8px;
            }

            .info-value {
                color: #0C172A !important;
                font-size: 1rem;
                font-weight: 900;
                line-height: 1.3;
            }

            .triage-layout {
                display: grid;
                grid-template-columns: 1fr;
                gap: 18px;
            }

            .triage-pills {
                display: grid;
                grid-template-columns: repeat(3, minmax(0, 1fr));
                gap: 10px;
                margin: 12px 0 6px;
            }

            .triage-pill {
                align-items: center;
                background: #FFFFFF;
                border: 1px solid var(--border);
                border-radius: 10px;
                color: #24384E !important;
                display: flex;
                font-size: .83rem;
                font-weight: 850;
                gap: 9px;
                min-height: 44px;
                padding: 10px 14px;
                box-shadow: 0 8px 18px rgba(6,76,59,.04);
            }

            .triage-pill .dot {
                width: 15px;
                height: 15px;
                border-radius: 50%;
                border: 2px solid #B6C4C0;
                flex: 0 0 auto;
                align-items: center;
                display: inline-flex;
                font-size: .68rem;
                font-weight: 950;
                justify-content: center;
                line-height: 1;
            }

            .triage-pill.active-ok {
                background: linear-gradient(135deg, #12A365, #08724F);
                border-color: #0E7A4F;
                color: #FFFFFF !important;
                box-shadow: 0 10px 24px rgba(14,122,79,.22);
            }

            .triage-pill.active-warning {
                background: #FFF8E8;
                border-color: #E9B762;
                color: #7A4B07 !important;
            }

            .triage-pill.active-danger {
                background: #FFF0F0;
                border-color: #EEA3A3;
                color: #8D2525 !important;
            }

            .triage-pill.active-ok .dot {
                background: #FFFFFF;
                border-color: #FFFFFF;
                box-shadow: inset 0 0 0 4px #12A365;
                color: #12A365;
            }

            .triage-pill.active-warning .dot {
                background: #FFFFFF;
                border-color: #F97316;
                box-shadow: inset 0 0 0 4px #FFF8E8;
                color: #F97316;
            }

            .triage-pill.active-danger .dot {
                background: #FFFFFF;
                border-color: #FF4B55;
                box-shadow: inset 0 0 0 4px #FFF0F0;
                color: #FF4B55;
            }

            .triage-note, .select-look {
                border: 1px solid var(--border);
                border-radius: 10px;
                color: #24384E !important;
                font-weight: 750;
                line-height: 1.45;
                padding: 15px 16px;
                background: #FFFFFF;
            }

            .contra-details {
                display: grid;
                gap: 10px;
                margin-top: 12px;
            }

            .contra-details details {
                background: #FFFFFF;
                border: 1px solid var(--border);
                border-radius: 10px;
                color: #24384E;
                font-size: .82rem;
                font-weight: 750;
                padding: 10px 12px;
            }

            .contra-details summary {
                color: var(--primary);
                cursor: pointer;
                font-weight: 900;
            }

            .contra-details details:first-child {
                background: #FFF0F0;
                border-color: #EEA3A3;
            }

            .contra-details details:first-child summary {
                color: #8D2525;
            }

            .contra-details details:nth-child(2) {
                background: #FFF8E8;
                border-color: #E9B762;
            }

            .contra-details details:nth-child(2) summary {
                color: #7A4B07;
            }

            .st-key-triage_ok_btn button {
                background: linear-gradient(135deg, #12A365, #08724F) !important;
                border: 1px solid #0E7A4F !important;
                color: #FFFFFF !important;
                font-weight: 900 !important;
                min-height: 46px;
            }

            .st-key-triage_relative_btn button {
                background: #FFF8E8 !important;
                border: 1px solid #E9B762 !important;
                color: #7A4B07 !important;
                font-weight: 900 !important;
                min-height: 46px;
            }

            .st-key-triage_absolute_btn button {
                background: #FFF0F0 !important;
                border: 1px solid #EEA3A3 !important;
                color: #8D2525 !important;
                font-weight: 900 !important;
                min-height: 46px;
            }

            .contra-details ul {
                margin: 10px 0 0 18px;
                padding: 0;
            }

            .contra-details li {
                margin-bottom: 6px;
            }

            .select-look {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 18px;
            }

            .soft-note {
                color: #405066 !important;
                font-weight: 650;
                margin: 0 0 14px;
            }

            .metric-grid {
                display: grid;
                grid-template-columns: repeat(5, minmax(0, 1fr));
                gap: 16px;
                margin: 18px 0;
            }

            .metric-card {
                border: 1px solid #DDE7E3;
                border-radius: 13px;
                padding: 18px 18px;
                min-height: 104px;
                background: linear-gradient(180deg, #FFFFFF 0%, #FBFDFB 100%);
                box-shadow: 0 10px 22px rgba(2,56,45,.045);
            }

            .metric-value {
                color: #07172D !important;
                font-size: 1.55rem;
                font-weight: 900;
                line-height: 1.1;
            }

            .metric-value.small {
                font-size: 1.02rem;
            }

            .compact-table {
                width: 100%;
                border-collapse: separate;
                border-spacing: 0;
                border: 1px solid var(--border);
                border-radius: 10px;
                overflow: hidden;
                color: #24384E;
                font-size: .88rem;
                text-align: center;
            }

            .compact-table th {
                background: #F7FAF9;
                color: #405066;
                font-weight: 900;
                padding: 10px 12px;
                border-right: 1px solid var(--border);
                border-bottom: 1px solid var(--border);
            }

            .compact-table td {
                background: #FFFFFF;
                border-right: 1px solid var(--border);
                border-bottom: 1px solid var(--border);
                font-weight: 800;
                padding: 11px 12px;
            }

            .compact-table th:last-child, .compact-table td:last-child {
                border-right: none;
            }

            .compact-table tr:last-child td {
                border-bottom: none;
            }

            .result-grid-main {
                display: grid;
                grid-template-columns: .9fr 1.9fr;
                gap: 18px;
                margin-bottom: 18px;
            }

            .distance-card, .prediction-card, .class-card, .insight-card {
                border: 1px solid var(--border);
                border-radius: 14px;
                padding: 22px;
                background: #FFFFFF;
            }

            .distance-number {
                color: var(--primary) !important;
                font-size: 4.1rem;
                font-weight: 950;
                line-height: 1;
                text-align: center;
            }

            .distance-caption {
                color: var(--text-muted) !important;
                font-weight: 800;
                text-align: center;
                margin-top: 8px;
            }

            .progress-wrap {
                margin-top: 26px;
            }

            .progress-track {
                height: 16px;
                border-radius: 999px;
                background: #D5D8D7;
                position: relative;
            }

            .progress-fill {
                height: 16px;
                border-radius: 999px;
                background: linear-gradient(90deg, #064C3B 0%, #0B7A60 100%);
                width: var(--progress);
            }

            .progress-marker {
                position: absolute;
                left: var(--progress);
                top: -10px;
                width: 3px;
                height: 34px;
                background: var(--warning);
                border-radius: 999px;
            }

            .progress-marker-label {
                position: absolute;
                left: var(--progress);
                top: -32px;
                transform: translateX(-50%);
                color: var(--warning) !important;
                font-size: .86rem;
                font-weight: 900;
                white-space: nowrap;
            }

            .progress-footer {
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                color: #25364A !important;
                font-size: .86rem;
                font-weight: 850;
                margin-top: 17px;
            }

            .class-grid {
                display: grid;
                grid-template-columns: repeat(3, minmax(0, 1fr));
                gap: 18px;
                margin-bottom: 18px;
            }

            .class-card {
                display: flex;
                gap: 14px;
                min-height: 112px;
            }

            .class-icon {
                border-radius: 50%;
                display: flex;
                align-items: center;
                flex: 0 0 auto;
                height: 56px;
                justify-content: center;
                width: 56px;
            }

            .class-icon svg {
                width: 28px;
                height: 28px;
                stroke-width: 2.8 !important;
            }

            .icon-green { background: #DDF8E5; color: #0E7A4F !important; }
            .icon-blue { background: #E9F2FF; color: #2877D9 !important; }
            .icon-orange { background: #FFF1DB; color: #D98B18 !important; }

            .class-value {
                color: var(--primary) !important;
                font-size: 1.28rem;
                font-weight: 950;
                line-height: 1.2;
            }

            .class-sub {
                color: var(--primary) !important;
                font-style: italic;
                font-weight: 800;
                margin-top: 5px;
            }

            .quick-grid {
                grid-template-columns: repeat(4, minmax(0, 1fr)) !important;
            }

            .quick-card {
                align-items: center;
                background: #FFFFFF;
                border: 1px solid var(--border);
                border-radius: 16px;
                display: flex;
                gap: 14px;
                min-height: 92px;
                padding: 18px 20px;
            }

            .quick-icon {
                width: 50px;
                height: 50px;
                border-radius: 999px;
                background: var(--primary-soft);
                color: var(--primary);
                display: flex;
                align-items: center;
                justify-content: center;
                flex: 0 0 auto;
            }

            .quick-icon svg {
                width: 26px;
                height: 26px;
                stroke-width: 2.8 !important;
            }

            .quick-label {
                color: var(--text-muted);
                font-size: 12px;
                font-weight: 700;
                margin-bottom: 6px;
            }

            .quick-value {
                color: var(--text-main);
                font-size: 24px;
                font-weight: 900;
                line-height: 1.1;
            }

            .insight-grid {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 18px;
            }

            .insight-title {
                color: var(--primary) !important;
                font-size: .92rem;
                font-weight: 850;
                margin-bottom: 14px;
            }

            .insight-main {
                color: var(--primary) !important;
                font-size: 1.25rem;
                font-weight: 950;
                margin-bottom: 14px;
            }

            .insight-text, .insight-card li {
                color: #24384E !important;
                font-size: .92rem;
                font-weight: 650;
                line-height: 1.55;
            }

            .insight-card {
                background:
                    radial-gradient(circle at 6% 12%, rgba(14,122,79,.07), transparent 26%),
                    #FFFFFF;
                box-shadow: 0 10px 24px rgba(2,56,45,.055);
            }

            .analytics-card {
                background: #FFFFFF;
                border: 1px solid var(--border);
                border-radius: 18px;
                box-shadow: var(--shadow);
                margin-top: 18px;
                padding: 22px;
            }

            .analytics-grid {
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 16px;
                margin-bottom: 18px;
            }

            .finding-metric {
                background: #FFFFFF;
                border: 1px solid var(--border);
                border-radius: 16px;
                min-height: 150px;
                padding: 22px 18px;
                text-align: center;
            }

            .finding-icon {
                width: 58px;
                height: 58px;
                border-radius: 999px;
                display: inline-flex;
                align-items: center;
                justify-content: center;
                margin-bottom: 14px;
            }

            .finding-icon svg {
                width: 32px;
                height: 32px;
                stroke-width: 2.8 !important;
            }

            .tone-green { background: #DDF8E5; color: #0E7A4F; }
            .tone-blue { background: #E0F0FF; color: #0B68C8; }
            .tone-orange { background: #FFF0D9; color: #B65314; }
            .tone-purple { background: #EFE6FF; color: #6F3CC3; }
            .tone-red { background: #FFE1E1; color: #D62828; }

            .finding-label {
                color: var(--text-main);
                font-size: .98rem;
                font-weight: 850;
                margin-bottom: 10px;
            }

            .finding-value {
                color: var(--primary);
                font-size: 2.05rem;
                font-weight: 950;
                line-height: 1.05;
            }

            .finding-unit {
                color: var(--text-muted);
                font-size: 1.05rem;
                font-weight: 750;
                margin-top: 2px;
            }

            .finding-list {
                border: 1px solid var(--border);
                border-radius: 16px;
                overflow: hidden;
            }

            .finding-row {
                align-items: center;
                border-bottom: 1px solid var(--border);
                display: grid;
                grid-template-columns: 72px 1fr;
                gap: 18px;
                padding: 18px 20px;
            }

            .finding-row:last-child {
                border-bottom: none;
            }

            .finding-row-icon {
                width: 52px;
                height: 52px;
                border-radius: 999px;
                display: flex;
                align-items: center;
                justify-content: center;
            }

            .finding-row-icon svg {
                width: 29px;
                height: 29px;
                stroke-width: 2.8 !important;
            }

            .finding-row-text {
                color: #24384E;
                font-size: 1rem;
                font-weight: 600;
                line-height: 1.55;
            }

            .finding-row-text strong {
                color: var(--text-main);
                font-size: 1.08rem;
                font-weight: 900;
            }

            .analytics-title {
                align-items: center;
                color: var(--text-main);
                display: flex;
                gap: 12px;
                font-size: 1.15rem;
                font-weight: 900;
                margin-bottom: 18px;
            }

            .analytics-title-icon {
                width: 40px;
                height: 40px;
                border-radius: 10px;
                color: var(--primary);
                display: flex;
                align-items: center;
                justify-content: center;
            }

            .analytics-title-icon svg {
                width: 27px;
                height: 27px;
                stroke-width: 2.8 !important;
            }

            .reading-grid {
                display: grid;
                grid-template-columns: repeat(3, minmax(0, 1fr));
                gap: 14px;
            }

            .reading-pill {
                align-items: center;
                border: 1px solid var(--border);
                border-radius: 16px;
                display: flex;
                gap: 14px;
                min-height: 78px;
                padding: 14px 16px;
            }

            .reading-text {
                color: var(--primary);
                font-weight: 900;
                line-height: 1.35;
            }

            .graph-metrics {
                display: grid;
                grid-template-columns: 1fr 1fr 1.25fr;
                gap: 16px;
                margin-bottom: 18px;
            }

            .graph-metric-card {
                align-items: center;
                background: #FFFFFF;
                border: 1px solid var(--border);
                border-radius: 16px;
                display: flex;
                gap: 18px;
                min-height: 112px;
                padding: 18px 22px;
            }

            .graph-metric-icon {
                width: 64px;
                height: 64px;
                border-radius: 999px;
                display: flex;
                align-items: center;
                justify-content: center;
                flex: 0 0 auto;
            }

            .graph-metric-icon svg {
                width: 36px;
                height: 36px;
                stroke-width: 2.8 !important;
            }

            .graph-label {
                color: var(--text-main);
                font-size: .98rem;
                font-weight: 850;
                margin-bottom: 4px;
            }

            .graph-value {
                color: var(--primary);
                font-size: 2.2rem;
                font-weight: 950;
                line-height: 1.1;
            }

            .graph-value.blue {
                color: #0B68C8;
            }

            .chart-card {
                background: #FFFFFF;
                border: 1px solid var(--border);
                border-radius: 16px;
                margin-top: 14px;
                padding: 18px 20px 8px;
            }

            .chart-card-header {
                align-items: center;
                display: flex;
                justify-content: space-between;
                margin-bottom: 4px;
            }

            .chart-title {
                color: var(--text-main);
                font-size: 1.05rem;
                font-weight: 900;
            }

            .chart-actions {
                color: #64748B;
                display: flex;
                gap: 12px;
            }

            .chart-actions svg {
                width: 21px;
                height: 21px;
                stroke-width: 2.75 !important;
            }

            .chart-note {
                align-items: center;
                background: #E9F8F0;
                border: 1px solid #BEE7CC;
                border-radius: 12px;
                color: #174D38;
                display: flex;
                gap: 12px;
                font-size: .95rem;
                font-weight: 750;
                margin-top: 14px;
                padding: 13px 16px;
            }

            .chart-note-icon {
                background: var(--primary);
                border-radius: 999px;
                color: #FFFFFF;
                display: flex;
                align-items: center;
                justify-content: center;
                width: 36px;
                height: 36px;
                flex: 0 0 auto;
            }

            .chart-note-icon svg {
                width: 22px;
                height: 22px;
                stroke-width: 2.8 !important;
            }

            button[data-baseweb="tab"] {
                background: #FFFFFF !important;
                border: 1px solid var(--border) !important;
                border-radius: 999px !important;
                padding: 8px 16px !important;
                margin-right: 8px !important;
            }

            button[data-baseweb="tab"][aria-selected="true"] {
                background: var(--primary) !important;
                border-color: var(--primary) !important;
            }

            button[data-baseweb="tab"][aria-selected="true"] p {
                color: #FFFFFF !important;
            }

            @media (max-width: 820px) {
                [data-testid="stAppViewContainer"] > .main .block-container {
                    padding: 1.25rem 1rem 2rem;
                }
                .protocol-header, .triage-layout, .result-grid-main, .insight-grid, .analytics-grid, .reading-grid, .graph-metrics {
                    grid-template-columns: 1fr;
                    display: grid;
                }
                .protocol-title {
                    font-size: 2rem;
                }
                [data-testid="stAppViewContainer"] > .main .block-container {
                    margin-left: 0;
                    max-width: 100%;
                    padding: 1.25rem 1rem 2rem;
                }
                .app-sidebar {
                    display: none;
                }
                .stage-header {
                    align-items: center;
                    gap: 9px;
                    display: grid;
                    grid-template-columns: 30px 36px minmax(0, 1fr);
                }
                .stage-number {
                    width: 30px;
                    height: 30px;
                    font-size: 14px;
                }
                .stage-icon {
                    width: 36px;
                    height: 36px;
                }
                .stage-title {
                    font-size: .98rem;
                    line-height: 1.25;
                    word-break: normal;
                    overflow-wrap: normal;
                    hyphens: none;
                }
                .stepper, .info-grid, .metric-grid, .class-grid, .triage-pills, .quick-grid {
                    grid-template-columns: 1fr !important;
                    margin-left: 0;
                    margin-right: 0;
                }
                .step-item::before {
                    display: none;
                }
                .info-field, .info-field:nth-child(5), .info-field:nth-child(9) {
                    border-right: none;
                }
                .info-field.wide {
                    grid-column: span 1;
                }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header_with_stepper_unused() -> None:
    current_date = date.today().strftime("%d/%m/%Y")
    steps = [
        ("1", "Identificação", True),
        ("2", "Triagem", False),
        ("3", "Predição", False),
        ("4", "Execução", False),
        ("5", "Recuperação", False),
        ("6", "Resultado", False),
    ]
    steps_html = "".join(
        f'<div class="step-item {"active" if active else ""}">'
        f'<div class="step-number">{number}</div>'
        f'<div class="step-label">{label}</div>'
        "</div>"
        for number, label, active in steps
    )
    st.markdown(
        f'<div class="protocol-header">'
        f'<div><h1 class="protocol-title">Protocolo do TC6M</h1>'
        f'<div class="protocol-subtitle">Fluxo clínico organizado por etapas</div></div>'
        f'<div class="protocol-actions"><div class="date-pill">{current_date}</div>'
        f'<div class="export-pill">Exportar relatório</div></div></div>'
        f'<div class="stepper">{steps_html}</div>',
        unsafe_allow_html=True,
    )


def render_app_sidebar() -> None:
    section = st.session_state.get("nav_section", "avaliacao")

    def go_to(target: str) -> None:
        st.session_state.nav_section = target

    with st.sidebar:
        st.markdown(
            """
            <div style="display:flex;align-items:center;gap:12px;margin:8px 0 28px;">
                <div style="font-size:2rem;line-height:1;">♡</div>
                <div>
                    <div style="font-size:1.55rem;font-weight:950;line-height:1;">TC6M</div>
                    <div style="font-size:.78rem;font-weight:700;opacity:.8;margin-top:4px;">Avaliação Funcional</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        nav_items = [
            ("avaliacao", "Avaliação"),
            ("execucao", "Execução"),
            ("resultados", "Resultados"),
            ("graficos", "Gráficos"),
            ("relatorio", "Relatório"),
        ]

        for key, label in nav_items:
            button_label = f"● {label}" if section == key else label
            if st.button(button_label, key=f"nav_{key}", use_container_width=True):
                go_to(key)
                st.rerun()

        st.divider()

        if st.button("+ Nova avaliação", key="nav_new", use_container_width=True):
            clear_result()
            st.session_state.nav_section = "avaliacao"
            st.rerun()

        st.markdown("<div style='height:24px;'></div>", unsafe_allow_html=True)
        st.caption("Equipe Cardiorrespiratória")


def render_app_sidebar() -> None:
    with st.sidebar:
        st.markdown(
            """
            <div style="margin:8px 0 28px;">
                <div style="font-size:1.75rem;font-weight:950;line-height:1;">TC6M</div>
                <div style="font-size:.82rem;font-weight:700;opacity:.82;margin-top:5px;">
                    Avaliação Funcional
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        section_labels = {
            "Avaliação": "avaliacao",
            "Execução": "execucao",
            "Resultados": "resultados",
            "Gráficos": "graficos",
            "Relatório": "relatorio",
        }
        labels = list(section_labels.keys())
        reverse_labels = {value: key for key, value in section_labels.items()}
        current_label = reverse_labels.get(st.session_state.get("nav_section", "avaliacao"), "Avaliação")
        selected_label = st.radio(
            "Navegação",
            labels,
            index=labels.index(current_label),
            key="sidebar_nav_radio",
        )
        st.session_state.nav_section = section_labels[selected_label]

        st.divider()

        if st.button("+ Nova avaliação", key="nav_new", use_container_width=True):
            clear_result()
            st.session_state.nav_section = "avaliacao"
            st.rerun()

        st.caption("Use a seta da lateral para minimizar/expandir.")
        st.caption("Equipe Cardiorrespiratória")


def render_header() -> None:
    current_date = date.today().strftime("%d/%m/%Y")
    st.markdown(
        f'<div class="protocol-header">'
        f'<div><h1 class="protocol-title">TC6M — Protocolo Integrado</h1>'
        f'<div class="protocol-subtitle">Fluxo organizado por etapas do teste de caminhada de 6 minutos</div></div>'
        f'<div class="protocol-actions"><div class="date-pill">{current_date}</div></div></div>',
        unsafe_allow_html=True,
    )


def sync_navigation_from_query() -> None:
    return


def stage_header(number: int, icon_svg: str, title: str) -> str:
    return (
        f'<div class="stage-header">'
        f'<div class="stage-number">{number}</div>'
        f'<div class="stage-icon">{icon_svg}</div>'
        f'<div class="stage-title">Etapa {number} · {escape(title)}</div>'
        f"</div>"
    )


def render_stage_card(number: int, icon_svg: str, title: str, body_html: str) -> None:
    st.markdown(
        f'<section class="clinical-card">{stage_header(number, icon_svg, title)}{body_html}</section>',
        unsafe_allow_html=True,
    )


def info_grid(items: list[tuple[str, str, bool]]) -> str:
    fields = "".join(
        f'<div class="info-field {"wide" if wide else ""}">'
        f'<div class="info-label">{escape(label)}</div>'
        f'<div class="info-value">{escape(clean_text(value))}</div></div>'
        for label, value, wide in items
    )
    return f'<div class="info-grid">{fields}</div>'


def metric_grid(items: list[tuple[str, str, bool]]) -> str:
    cards = "".join(
        f'<div class="metric-card"><div class="metric-label">{escape(label)}</div>'
        f'<div class="metric-value {"small" if small else ""}">{escape(clean_text(value))}</div></div>'
        for label, value, small in items
    )
    return f'<div class="metric-grid">{cards}</div>'


def triage_html(status: str, observation: str) -> str:
    absolute_items = "".join(f"<li>{escape(item)}</li>" for item in CONTRAINDICACOES_ABSOLUTAS)
    relative_items = "".join(f"<li>{escape(item)}</li>" for item in CONTRAINDICACOES_RELATIVAS)
    contraindications = (
        '<div class="contra-details">'
        f"<details><summary>Contraindicações absolutas</summary><ul>{absolute_items}</ul></details>"
        f"<details><summary>Contraindicações relativas</summary><ul>{relative_items}</ul></details>"
        "</div>"
    )
    return (
        '<div class="triage-layout">'
        f'<div>{contraindications}</div>'
        '<div><div class="table-label">Observações da triagem (opcional)</div>'
        f'<div class="triage-note">{escape(clean_text(observation))}</div></div></div>'
    )


def result_html(patient: PatientData, result, series: pd.DataFrame) -> str:
    payload = build_report_payload(patient, result, series)
    percent = min(max(float(result.percentual_atingido), 0), 100)
    lin_label = f"{br_number(result.lin_principal)} m" if result.lin_principal is not None else "Não definido"
    pico = payload["phase"]["pico"]
    spo2_values = pd.to_numeric(normalize_timeseries(series)["SpO2"], errors="coerce")
    spo2_values = spo2_values[spo2_values > 0]
    min_spo2 = int(spo2_values.min()) if not spo2_values.empty else 0
    risk = clean_text(result.risco)
    risk_lower = risk.lower()
    if "muito elevado" in risk_lower or "elevadíssimo" in risk_lower:
        risk_subtitle = "Alto risco"
    elif "elevado" in risk_lower:
        risk_subtitle = "Risco elevado"
    elif "moderado" in risk_lower:
        risk_subtitle = "Moderado risco"
    else:
        risk_subtitle = "Baixo risco"

    classification_text = clean_text(result.classificacao_risco).lower()
    if "4" in classification_text:
        risk_marker = 12.5
    elif "3" in classification_text:
        risk_marker = 37.5
    elif "2" in classification_text:
        risk_marker = 62.5
    else:
        risk_marker = 87.5
    summary = clean_text(payload["clinical_summary"]).replace("Resumo clínico: ", "")
    factor = clean_text(result.fator_limitante)

    return f"""
    <div class="result-grid-main">
        <div class="distance-card">
            <div class="table-label" style="text-align:center;">Distância percorrida no TC6M</div>
            <div class="distance-number">{patient.distancia:.0f} <span style="font-size:2rem;">m</span></div>
            <div class="distance-caption">Distância percorrida</div>
        </div>
        <div class="prediction-card">
            <div class="table-label">Desempenho em relação ao predito ({escape(clean_text(result.formula_principal).split(" - ")[0])})</div>
            <div class="progress-wrap" style="--progress:{percent:.1f}%;">
                <div class="progress-track">
                    <div class="progress-fill"></div>
                    <div class="progress-marker"></div>
                    <div class="progress-marker-label">{br_number(result.percentual_atingido)}% do previsto</div>
                </div>
            </div>
            <div class="progress-footer">
                <div>LIN<br><strong>{lin_label}</strong></div>
                <div style="text-align:center;">Percentual do previsto<br><strong>{br_number(result.percentual_atingido)}%</strong></div>
                <div style="text-align:right;">Predito<br><strong>{br_number(result.dpp_principal)} m</strong></div>
            </div>
        </div>
    </div>
    <div class="class-grid">
        <div class="class-card"><div class="class-icon icon-green">{ICON_ACTIVITY}</div><div><div class="class-label">Qualificador funcional</div><div class="class-value">{escape(clean_text(result.qualificador_funcional))}</div></div></div>
        <div class="class-card"><div class="class-icon icon-blue">{ICON_GAUGE}</div><div><div class="class-label">Classificação</div><div class="class-value">{escape(clean_text(result.classificacao_risco))}</div><div class="class-sub">{escape(risk_subtitle)}</div></div></div>
        <div class="class-card"><div class="class-icon icon-orange">{ICON_SHIELD}</div><div><div class="class-label">Risco</div><div class="class-value">{escape(risk)}</div></div></div>
    </div>
    <div class="metric-grid quick-grid">
        <div class="quick-card"><div class="quick-icon">{ICON_DROPLET}</div><div><div class="quick-label">SpO2 mínima</div><div class="quick-value">{min_spo2}%</div></div></div>
        <div class="quick-card"><div class="quick-icon">{ICON_HEART}</div><div><div class="quick-label">FC pico</div><div class="quick-value">{pico.fc} bpm</div></div></div>
        <div class="quick-card"><div class="quick-icon">{ICON_LUNGS}</div><div><div class="quick-label">Borg pico</div><div class="quick-value">{pico.borg_resp:.1f} / {pico.borg_mmii:.1f}</div></div></div>
        <div class="quick-card"><div class="quick-icon">{ICON_CHECK}</div><div><div class="quick-label">Interrupção</div><div class="quick-value">{'Sim' if patient.interrompeu else 'Não'}</div></div></div>
    </div>
    <div class="insight-grid">
        <div class="insight-card">
            <div class="insight-title">Interpretação clínica</div>
            <div class="insight-main">{escape(factor)}</div>
            <ul>
                <li>Desempenho abaixo do previsto para idade e sexo.</li>
                <li>Sinais de limitação cardiovascular e ventilatória.</li>
                <li>Relação SpO2 e Borg compatível com esforço submáximo.</li>
            </ul>
        </div>
        <div class="insight-card">
            <div class="insight-title">Resumo clínico</div>
            <div class="summary-quote"><div class="insight-text">{escape(summary)}</div></div>
        </div>
    </div>
    """


def result_dashboard_html(patient: PatientData, result, series: pd.DataFrame) -> str:
    payload = build_report_payload(patient, result, series)
    percent = min(max(float(result.percentual_atingido), 0), 100)
    lin_label = f"{br_number(result.lin_principal)} m" if result.lin_principal is not None else "Não definido"
    spo2_values = pd.to_numeric(normalize_timeseries(series)["SpO2"], errors="coerce")
    spo2_values = spo2_values[spo2_values > 0]
    min_spo2 = int(spo2_values.min()) if not spo2_values.empty else 0
    risk = clean_text(result.risco)
    risk_lower = risk.lower()

    if "muito elevado" in risk_lower or "elevadíssimo" in risk_lower:
        risk_subtitle = "Alto risco"
    elif "elevado" in risk_lower:
        risk_subtitle = "Risco elevado"
    elif "moderado" in risk_lower:
        risk_subtitle = "Moderado risco"
    else:
        risk_subtitle = "Baixo risco"

    classification_text = clean_text(result.classificacao_risco).lower()
    if "4" in classification_text:
        risk_marker = 12.5
    elif "3" in classification_text:
        risk_marker = 37.5
    elif "2" in classification_text:
        risk_marker = 62.5
    else:
        risk_marker = 87.5

    summary = clean_text(payload["clinical_summary"]).replace("Resumo clínico: ", "")
    factor = clean_text(result.fator_limitante)
    metrics = payload["metrics"]
    attention_rows = "".join(
        f'<div class="attention-row"><span>{escape(point["label"])}</span>'
        f'<span class="attention-badge {escape(point["type"])}">{escape(point["badge"])}</span></div>'
        for point in payload["attention_points"]
    )

    return f"""
    <div class="result-band">
        <div class="result-band-title"><span class="dashboard-step">6</span> Resultado final do teste</div>
        <div class="result-band-grid">
            <div class="result-tile">
                <div class="result-tile-label">{ICON_SHIELD}<span>Distância percorrida</span></div>
                <div class="result-main-number">{patient.distancia:.0f}<span> m</span></div>
                <div class="result-progress-mini" style="--progress:{percent:.1f}%;"><span></span></div>
                <div style="display:flex;justify-content:space-between;font-size:.74rem;font-weight:850;color:rgba(255,255,255,.85);">
                    <span>Previsto: {br_number(result.dpp_principal, 0)} m</span>
                    <span>LIN: {lin_label}</span>
                </div>
                <div class="result-accent">{br_number(result.percentual_atingido)}% <span style="font-size:.9rem;color:#FFFFFF;">do previsto</span></div>
            </div>
            <div class="result-tile">
                <div class="result-tile-label">{ICON_ACTIVITY}<span>Qualificador funcional</span></div>
                <div class="result-main-number" style="font-size:1.8rem;line-height:1.18;">{escape(clean_text(result.qualificador_funcional))}</div>
            </div>
            <div class="result-tile">
                <div class="result-tile-label">{ICON_GAUGE}<span>Classificação</span></div>
                <div class="result-main-number" style="font-size:2rem;">{escape(clean_text(result.classificacao_risco))}</div>
                <div style="font-size:1.05rem;font-weight:850;color:rgba(255,255,255,.9);">{escape(risk_subtitle)}</div>
                <div class="risk-scale-mini" style="--risk-marker:{risk_marker}%;"><span></span><span></span><span></span><span></span></div>
            </div>
            <div class="result-tile">
                <div class="result-tile-label">{ICON_SHIELD}<span>Risco</span></div>
                <div class="result-main-number" style="font-size:1.65rem;line-height:1.24;">{escape(risk)}</div>
            </div>
        </div>
    </div>
    <div class="dashboard-card">
        <div class="dashboard-title">Monitoramento hemodinâmico <span class="dashboard-title-muted">— Destaques</span></div>
        <div class="monitor-row">
            <div class="monitor-card"><div class="monitor-icon">{ICON_HEART}</div><div><div class="mini-label">FC no pico</div><div class="mini-value">{metrics[0]["value"]}<br><span style="font-size:.78rem;font-weight:750;">{metrics[0]["unit"]}</span></div></div></div>
            <div class="monitor-card"><div class="monitor-icon">{ICON_DROPLET}</div><div><div class="mini-label">SpO2 mínima</div><div class="mini-value">{min_spo2}<br><span style="font-size:.78rem;font-weight:750;">%</span></div></div></div>
            <div class="monitor-card"><div class="monitor-icon">{ICON_LUNGS}</div><div><div class="mini-label">Borg respiratório / MMII</div><div class="mini-value">{metrics[2]["value"]}<br><span style="font-size:.78rem;font-weight:750;">Escala de Borg</span></div></div></div>
            <div class="monitor-card"><div class="monitor-icon">{ICON_HEART}</div><div><div class="mini-label">DP repouso</div><div class="mini-value">{metrics[3]["value"]}<br><span style="font-size:.78rem;font-weight:750;">{metrics[3]["unit"]}</span></div></div></div>
            <div class="monitor-card"><div class="monitor-icon">{ICON_REFRESH}</div><div><div class="mini-label">DP recuperação</div><div class="mini-value">{metrics[4]["value"]}<br><span style="font-size:.78rem;font-weight:750;">{metrics[4]["unit"]}</span></div></div></div>
        </div>
    </div>
    <div class="insight-grid" style="margin-top:16px;">
        <div class="insight-card">
            <div class="insight-title">Interpretação clínica</div>
            <div class="insight-main">{escape(factor)}</div>
            <ul>
                <li>Desempenho analisado em relação ao previsto para idade e sexo.</li>
                <li>Relação entre FC, SpO2 e Borg considerada no padrão funcional.</li>
                <li>Interpretação deve ser associada ao contexto clínico do paciente.</li>
            </ul>
            <div class="attention-list">{attention_rows}</div>
        </div>
        <div class="insight-card">
            <div class="insight-title">Resumo clínico</div>
            <div class="insight-text">{escape(summary)}</div>
        </div>
    </div>
    """


def integrated_recovery_html(patient: PatientData, series: pd.DataFrame) -> str:
    analysis = build_integrated_recovery_analysis(patient, series)

    def value(key: str, suffix: str = "", digits: int = 2) -> str:
        return escape(format_analysis_value(analysis[key], suffix, digits))

    interpretations = "".join(f"<li>{escape(item)}</li>" for item in analysis["interpretations"][:4])
    return f"""
    <div class="dashboard-card" style="margin-top:16px;">
        <div class="dashboard-title">Recuperação do Duplo Produto após o TC6M <span class="dashboard-title-muted">— análise complementar</span></div>
        <div class="monitor-row" style="grid-template-columns: repeat(4, minmax(0, 1fr));">
            <div class="monitor-card"><div class="monitor-icon">{ICON_HEART}</div><div><div class="mini-label">DP repouso</div><div class="mini-value">{value("dp_repouso", " bpm.mmHg", 0)}</div></div></div>
            <div class="monitor-card"><div class="monitor-icon">{ICON_ACTIVITY}</div><div><div class="mini-label">DP 1 min</div><div class="mini-value">{value("dp_1", " bpm.mmHg", 0)}</div></div></div>
            <div class="monitor-card"><div class="monitor-icon">{ICON_REFRESH}</div><div><div class="mini-label">DP 6 min</div><div class="mini-value">{value("dp_6", " bpm.mmHg", 0)}</div></div></div>
            <div class="monitor-card"><div class="monitor-icon">{ICON_GAUGE}</div><div><div class="mini-label">Recuperação DP 6 min</div><div class="mini-value">{value("recovery_percent_6", " %")}</div></div></div>
        </div>
        <div class="insight-grid" style="margin-top:16px;">
            <div class="insight-card">
                <div class="insight-title">Custo cardiovascular estimado por metro</div>
                <div class="insight-main">{value("cost_dp_per_m", " DP/m")}</div>
                <div class="insight-text">Relaciona o aumento do Duplo Produto no pós-teste imediato com a distância obtida no TC6M.</div>
            </div>
            <div class="insight-card">
                <div class="insight-title">Velocidade média no TC6M</div>
                <div class="insight-main">{value("velocity_ms", " m/s")}</div>
                <div class="insight-text">Ritmo médio: <strong>{value("pace_m_min", " m/min")}</strong>.</div>
            </div>
        </div>
        <div class="insight-grid" style="margin-top:16px;">
            <div class="insight-card">
                <div class="insight-title">Velocidade normalizada pelo comprimento do membro</div>
                <div class="insight-main">{value("normalized_velocity", "", 3)}</div>
                <div class="insight-text">Comprimento informado: <strong>{value("limb_length_m", " m")}</strong>. Análise biomecânica complementar, sem critério diagnóstico isolado.</div>
            </div>
            <div class="insight-card">
                <div class="insight-title">Interpretação cautelosa</div>
                <ul>{interpretations}</ul>
            </div>
        </div>
        <div class="soft-note" style="margin-top:14px;">{escape(analysis["notice"])}</div>
    </div>
    """


def curve_metrics(series: pd.DataFrame) -> dict[str, object]:
    clean = normalize_timeseries(series)
    exercise = clean.iloc[1:7].copy()
    recovery = clean.iloc[7:].copy()
    repouso = clean.iloc[0]

    fc_peak = int(exercise["FC"].max())
    fc_delta = int(fc_peak - int(repouso["FC"]))
    spo2_values = pd.to_numeric(exercise["SpO2"], errors="coerce")
    spo2_values = spo2_values[spo2_values > 0]
    spo2_min = int(spo2_values.min()) if not spo2_values.empty else 0
    spo2_drop = int(int(repouso["SpO2"]) - spo2_min) if spo2_min else 0
    borg_resp_peak = float(exercise["Borg Respiratório"].max())
    borg_mmii_peak = float(exercise["Borg MMII"].max())
    borg_gap = abs(borg_resp_peak - borg_mmii_peak)
    recovery_last = recovery.iloc[-1] if not recovery.empty else clean.iloc[-1]

    if borg_gap <= 1:
        pattern = "Esforço global/misto"
        reading = "Percepção de esforço mista"
    elif borg_resp_peak > borg_mmii_peak:
        pattern = "Predomínio respiratório"
        reading = "Predomínio ventilatório"
    else:
        pattern = "Predomínio periférico"
        reading = "Predomínio muscular"

    return {
        "fc_peak": fc_peak,
        "fc_delta": fc_delta,
        "spo2_min": spo2_min,
        "spo2_drop": spo2_drop,
        "borg_resp_peak": borg_resp_peak,
        "borg_mmii_peak": borg_mmii_peak,
        "pattern": pattern,
        "reading": reading,
        "recovery_fc": int(recovery_last["FC"]),
        "recovery_spo2": int(recovery_last["SpO2"]),
    }


def achados_panel_html(series: pd.DataFrame) -> str:
    metrics = curve_metrics(series)
    findings = build_curve_findings(series)
    finding_icons = [
        (ICON_TREND_UP, "tone-green"),
        (ICON_TREND_DOWN, "tone-blue"),
        (ICON_USERS, "tone-orange"),
        (ICON_REFRESH, "tone-purple"),
    ]
    rows = "".join(
        f'<div class="finding-row">'
        f'<div class="finding-row-icon {tone}">{icon}</div>'
        f'<div class="finding-row-text">{escape(text)}</div>'
        f"</div>"
        for text, (icon, tone) in zip(findings, finding_icons)
    )
    return f"""
    <div class="analytics-card">
        <div class="analytics-grid">
            <div class="finding-metric">
                <div class="finding-icon tone-green">{ICON_HEART}</div>
                <div class="finding-label">Δ FC</div>
                <div class="finding-value">+{metrics["fc_delta"]}</div>
                <div class="finding-unit">bpm</div>
            </div>
            <div class="finding-metric">
                <div class="finding-icon tone-blue">{ICON_DROPLET}</div>
                <div class="finding-label">Queda de SpO2</div>
                <div class="finding-value">-{metrics["spo2_drop"]}</div>
                <div class="finding-unit">pp</div>
            </div>
            <div class="finding-metric">
                <div class="finding-icon tone-orange">{ICON_USERS}</div>
                <div class="finding-label">Borg pico</div>
                <div class="finding-value" style="font-size:1.8rem;color:#8A2F0B;">{metrics["borg_resp_peak"]:.1f} / {metrics["borg_mmii_peak"]:.1f}</div>
            </div>
            <div class="finding-metric">
                <div class="finding-icon tone-green">{ICON_TARGET}</div>
                <div class="finding-label">Padrão</div>
                <div class="finding-value" style="font-size:1.45rem;">{escape(str(metrics["pattern"]))}</div>
            </div>
        </div>
        <div class="analytics-card" style="box-shadow:none;margin-top:0;">
            <div class="analytics-title"><span class="analytics-title-icon">{ICON_CLIPBOARD}</span>Resumo dos achados</div>
            <div class="finding-list">{rows}</div>
        </div>
        <div class="analytics-card" style="box-shadow:none;">
            <div class="analytics-title"><span class="analytics-title-icon">{ICON_HEART}</span>Leitura clínica</div>
            <div class="reading-grid">
                <div class="reading-pill"><span class="finding-row-icon tone-green">{ICON_HEART}</span><span class="reading-text">Resposta cronotrópica importante</span></div>
                <div class="reading-pill"><span class="finding-row-icon tone-blue">{ICON_DROPLET}</span><span class="reading-text">Dessaturação ao esforço</span></div>
                <div class="reading-pill"><span class="finding-row-icon tone-orange">{ICON_USERS}</span><span class="reading-text">{escape(str(metrics["reading"]))}</span></div>
            </div>
        </div>
    </div>
    """


def graph_metrics_html(series: pd.DataFrame) -> str:
    metrics = curve_metrics(series)
    return f"""
    <div class="analytics-card">
        <div class="graph-metrics">
            <div class="graph-metric-card">
                <div class="graph-metric-icon tone-red">{ICON_HEART}</div>
                <div><div class="graph-label">FC pico</div><div class="graph-value">{metrics["fc_peak"]} <span style="font-size:1.1rem;color:#66736F;">bpm</span></div></div>
            </div>
            <div class="graph-metric-card">
                <div class="graph-metric-icon tone-blue">{ICON_DROPLET}</div>
                <div><div class="graph-label">SpO2 mínima</div><div class="graph-value blue">{metrics["spo2_min"]}%</div></div>
            </div>
            <div class="graph-metric-card">
                <div class="graph-metric-icon tone-green">{ICON_REFRESH}</div>
                <div><div class="graph-label">Recuperação</div><div class="graph-value">{metrics["recovery_fc"]} <span style="font-size:1.1rem;">bpm</span> / {metrics["recovery_spo2"]}%</div></div>
            </div>
        </div>
    </div>
    """


def chart_card_start(title: str) -> None:
    st.markdown(
        f'<div class="chart-card"><div class="chart-card-header">'
        f'<div class="chart-title">{escape(title)}</div>'
        f"</div>",
        unsafe_allow_html=True,
    )


def chart_card_end() -> None:
    st.markdown("</div>", unsafe_allow_html=True)


def render_dp_recovery_chart(patient: PatientData, series: pd.DataFrame) -> None:
    figure = build_dp_recovery_figure(patient, series)
    chart_card_start("Recuperação do Duplo Produto após o TC6M")
    if figure is None:
        st.markdown(
            '<div class="chart-note">'
            f'<span class="chart-note-icon">{ICON_INFO}</span>'
            "Dados insuficientes para plotar ao menos dois pontos válidos de Duplo Produto."
            "</div>",
            unsafe_allow_html=True,
        )
    else:
        st.pyplot(figure, use_container_width=True)
    chart_card_end()


def chart_note_html(series: pd.DataFrame) -> str:
    metrics = curve_metrics(series)
    if metrics["borg_resp_peak"] >= metrics["borg_mmii_peak"]:
        message = "Maior esforço no 6º minuto com dissociação leve entre Borg respiratório e MMII."
    else:
        message = "Maior percepção periférica no pico do esforço, com Borg MMII predominante."
    return f'<div class="chart-note"><span class="chart-note-icon">{ICON_CHECK}</span>{escape(message)}</div>'


def build_dashboard_oscillation_figure(series: pd.DataFrame):
    clean = normalize_timeseries(series)
    labels = [str(item).replace("Recuperação ", "Recuperação\n") for item in clean["Tempo"]]
    x_values = list(range(len(clean)))

    fig, ax1 = plt.subplots(figsize=(13.5, 4.6))
    fig.patch.set_facecolor("white")
    ax1.set_facecolor("white")
    ax1.plot(x_values, clean["FC"], color="#E31A1C", marker="o", linewidth=2.4, label="FC (bpm)")
    ax1.set_ylabel("FC (bpm)", color="#E31A1C", fontsize=11, fontweight="bold")
    ax1.tick_params(axis="y", colors="#E31A1C")
    ax1.set_ylim(50, max(150, int(clean["FC"].max()) + 10))
    ax1.grid(True, color="#E5E7EB", linewidth=0.9, alpha=0.9)

    for x, y in zip(x_values, clean["FC"]):
        if y > 0:
            ax1.annotate(
                f"{int(y)}",
                (x, y),
                textcoords="offset points",
                xytext=(0, 11),
                ha="center",
                color="#E31A1C",
                fontsize=9,
                fontweight="bold",
            )

    ax2 = ax1.twinx()
    ax2.plot(x_values, clean["SpO2"], color="#006DD8", marker="s", linewidth=2.2, linestyle="--", label="SpO2 (%)")
    ax2.set_ylabel("SpO2 (%)", color="#006DD8", fontsize=11, fontweight="bold")
    ax2.tick_params(axis="y", colors="#006DD8")
    ax2.set_ylim(max(85, int(clean["SpO2"].min()) - 2), 100)

    for x, y in zip(x_values, clean["SpO2"]):
        if y > 0:
            ax2.annotate(
                f"{int(y)}",
                (x, y),
                textcoords="offset points",
                xytext=(0, -16),
                ha="center",
                color="#006DD8",
                fontsize=9,
                fontweight="bold",
            )

    ax1.set_xticks(x_values)
    ax1.set_xticklabels(labels, fontsize=9)
    lines = ax1.get_lines() + ax2.get_lines()
    ax1.legend(lines, [line.get_label() for line in lines], loc="lower center", bbox_to_anchor=(0.5, -0.24), ncol=2, frameon=False)
    fig.tight_layout()
    return fig


def build_dashboard_effort_figure(series: pd.DataFrame):
    clean = normalize_timeseries(series)
    labels = [str(item).replace("Recuperação ", "Recuperação\n") for item in clean["Tempo"]]
    x_values = list(range(len(clean)))

    fig, ax = plt.subplots(figsize=(13.5, 4.4))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.plot(x_values, clean["Borg Respiratório"], color="#006DD8", marker="o", linewidth=2.4, label="Borg respiratório")
    ax.plot(x_values, clean["Borg MMII"], color="#FF7A1A", marker="s", linewidth=2.2, linestyle="--", label="Borg MMII")
    ax.set_ylim(0, 10)
    ax.set_ylabel("Escala de Borg", fontsize=11, fontweight="bold")
    ax.grid(True, color="#E5E7EB", linewidth=0.9, alpha=0.9)
    ax.set_xticks(x_values)
    ax.set_xticklabels(labels, fontsize=9)

    for series_values, color, offset in [
        (clean["Borg Respiratório"], "#006DD8", 12),
        (clean["Borg MMII"], "#FF7A1A", -18),
    ]:
        for x, y in zip(x_values, series_values):
            ax.annotate(f"{float(y):.0f}", (x, y), textcoords="offset points", xytext=(0, offset), ha="center", color=color, fontsize=9, fontweight="bold")

    exercise = clean.iloc[1:7]
    peak_idx = int(exercise["Borg Respiratório"].idxmax())
    peak_y = float(clean.loc[peak_idx, "Borg Respiratório"])
    ax.annotate(
        "Pico de esforço\nno 6º minuto",
        xy=(peak_idx, peak_y),
        xytext=(peak_idx + 0.7, min(9.2, peak_y + 2.2)),
        arrowprops={"arrowstyle": "->", "color": "#006DD8", "lw": 1.2},
        bbox={"boxstyle": "round,pad=0.35", "fc": "#EEF7FF", "ec": "#75B7FF"},
        color="#0060B8",
        fontsize=9,
        fontweight="bold",
    )
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.27), ncol=2, frameon=False)
    fig.tight_layout()
    return fig


def render_editor() -> None:
    with st.expander("Editar dados do protocolo", expanded=False):
        st.markdown("### Ambiente de teste")
        demo_col, demo_button = st.columns([2, 1])
        with demo_col:
            st.selectbox("Ambiente de teste para preencher automaticamente", TEST_PROFILES, key="perfil_teste")
        with demo_button:
            st.write("")
            if st.button("Preencher ambiente de teste", use_container_width=True):
                fill_demo_profile(st.session_state.perfil_teste)
                try:
                    generate_result()
                except ValueError:
                    pass
                st.rerun()

        st.markdown("### Identificação e antropometria")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.text_input("Nome do paciente", key="nome", on_change=update_evaluation_id)
            st.text_input("Prontuário/ID", key="prontuario", on_change=mark_manual_id)
            st.date_input("Data da avaliação", key="data_avaliacao", on_change=update_evaluation_id)
        with c2:
            st.selectbox("Sexo biológico", ["Masculino", "Feminino"], key="sexo_label", on_change=reset_patient_progress)
            st.number_input("Idade (anos)", min_value=1, max_value=120, step=1, key="idade", on_change=update_evaluation_id)
            st.number_input("Peso (kg)", min_value=1.0, max_value=350.0, step=0.1, key="peso", on_change=reset_patient_progress)
        with c3:
            st.number_input("Altura (cm)", min_value=50.0, max_value=240.0, step=0.1, key="altura_cm", on_change=reset_patient_progress)
            st.number_input(
                "Comprimento do membro inferior (m, opcional)",
                min_value=0.0,
                max_value=1.50,
                step=0.01,
                format="%.2f",
                key="comprimento_membro_inferior_m",
                on_change=reset_patient_progress,
            )
            st.text_input("Avaliador", key="avaliador", on_change=reset_patient_progress)
            st.text_input("Diagnóstico/condição clínica", key="diagnostico", on_change=reset_patient_progress)
            n_col, id_col = st.columns([1.2, 0.8])
            with n_col:
                st.number_input("Nº do teste no dia", min_value=1, max_value=99, step=1, key="numero_teste", on_change=update_evaluation_id)
            with id_col:
                st.write("")
                st.button("Gerar ID", use_container_width=True, on_click=force_evaluation_id)

        if st.session_state.patient_saved:
            st.success("Dados do paciente salvos. A triagem de segurança está liberada.")

        if st.button("Salvar dados do paciente e liberar triagem", type="primary", use_container_width=True, key="save_patient_stage"):
            if not st.session_state.nome.strip():
                st.error("Informe o nome do paciente antes de continuar.")
            else:
                if st.session_state.get("id_auto_ativo", True):
                    st.session_state.pending_force_id = True
                    st.session_state.pending_patient_save = True
                    st.rerun()
                else:
                    st.session_state.patient_saved = True
                    st.session_state.prediction_saved = False
                    st.session_state.execution_saved = False
                    st.session_state.recovery_saved = False

        st.markdown("### Triagem de segurança")
        b1, b2, b3 = st.columns(3)
        with b1:
            if st.button("Sem contraindicações", use_container_width=True):
                set_triage("Sem contraindicações")
                st.rerun()
        with b2:
            if st.button("Contraindicação relativa", use_container_width=True):
                set_triage("Contraindicação relativa")
                st.rerun()
        with b3:
            if st.button("Contraindicação absoluta", use_container_width=True):
                set_triage("Contraindicação absoluta")
                st.rerun()
        st.text_area("Observações da triagem", key="observacao_triagem", on_change=clear_result)
        with st.expander("Contraindicações absolutas"):
            st.write("\n".join(f"- {item}" for item in CONTRAINDICACOES_ABSOLUTAS))
        with st.expander("Contraindicações relativas"):
            st.write("\n".join(f"- {item}" for item in CONTRAINDICACOES_RELATIVAS))

        st.markdown("### Avaliação prévia")
        st.selectbox("Fórmula selecionada", FORMULAS_DPP, key="formula_principal", on_change=clear_result)
        pre_editor = prepare_editor_with_bp(st.session_state.pre_df)
        pre_resp_col = borg_resp_col(pre_editor)
        pre_mmii_col = borg_mmii_col(pre_editor)
        st.session_state.pre_df = restore_pas_pad(
            st.data_editor(
                pre_editor,
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "Tempo": st.column_config.TextColumn("Momento", disabled=True),
                    "FC": st.column_config.NumberColumn("FC (bpm)", min_value=0, max_value=260, step=1),
                    "SpO2": st.column_config.NumberColumn("SpO2 (%)", min_value=0, max_value=100, step=1),
                    "FR": st.column_config.NumberColumn("FR (ipm)", min_value=0, max_value=80, step=1),
                    "PA": st.column_config.TextColumn("PA (mmHg)"),
                    pre_resp_col: st.column_config.NumberColumn("Borg dispneia", min_value=0.0, max_value=10.0, step=0.5),
                    pre_mmii_col: st.column_config.NumberColumn("Borg MMII", min_value=0.0, max_value=10.0, step=0.5),
                },
            )
        )

        st.markdown("### Execução do teste")
        during_resp_col = borg_resp_col(st.session_state.during_df)
        during_mmii_col = borg_mmii_col(st.session_state.during_df)
        st.session_state.during_df = st.data_editor(
            st.session_state.during_df,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "Tempo": st.column_config.TextColumn("Minuto", disabled=True),
                "FC": st.column_config.NumberColumn("FC (bpm)", min_value=0, max_value=260, step=1),
                "SpO2": st.column_config.NumberColumn("SpO2 (%)", min_value=0, max_value=100, step=1),
                during_resp_col: st.column_config.NumberColumn("Borg dispneia", min_value=0.0, max_value=10.0, step=0.5),
                during_mmii_col: st.column_config.NumberColumn("Borg MMII", min_value=0.0, max_value=10.0, step=0.5),
            },
        )

        st.markdown("### Recuperação")
        recovery_editor = prepare_editor_with_bp(st.session_state.recovery_df)
        recovery_resp_col = borg_resp_col(recovery_editor)
        recovery_mmii_col = borg_mmii_col(recovery_editor)
        st.session_state.recovery_df = restore_pas_pad(
            st.data_editor(
                recovery_editor,
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "Tempo": st.column_config.TextColumn("Momento", disabled=True),
                    "FC": st.column_config.NumberColumn("FC (bpm)", min_value=0, max_value=260, step=1),
                    "SpO2": st.column_config.NumberColumn("SpO2 (%)", min_value=0, max_value=100, step=1),
                    "FR": st.column_config.NumberColumn("FR (ipm)", min_value=0, max_value=80, step=1),
                    "PA": st.column_config.TextColumn("PA (mmHg)"),
                    recovery_resp_col: st.column_config.NumberColumn("Borg dispneia", min_value=0.0, max_value=10.0, step=0.5),
                    recovery_mmii_col: st.column_config.NumberColumn("Borg MMII", min_value=0.0, max_value=10.0, step=0.5),
                },
            )
        )

        st.markdown("### Resultado final")
        r1, r2, r3 = st.columns(3)
        with r1:
            st.number_input("Distância percorrida ao final do TC6M (m)", min_value=0.0, max_value=2000.0, step=1.0, key="distancia", on_change=clear_result)
        with r2:
            st.radio("O paciente interrompeu o teste?", ["Não", "Sim"], horizontal=True, key="interrompeu_label", on_change=clear_result)
        with r3:
            if st.session_state.interrompeu_label == "Sim":
                st.number_input("Distância no momento da interrupção (m)", min_value=0.0, max_value=2000.0, step=1.0, key="distancia_interrupcao", on_change=clear_result)
        if st.session_state.interrompeu_label == "Sim":
            st.text_area("Motivo da interrupção", key="motivo_interrupcao", on_change=clear_result)
        if st.button("Gerar resumo final do TC6M", type="primary", use_container_width=True):
            try:
                generate_result()
                st.rerun()
            except ValueError as error:
                st.error(str(error))


def render_editor() -> None:
    """Motor antigo desativado: os controles agora ficam integrados ao layout principal."""

    return None


def render_environment_controls() -> None:
    st.markdown("#### Ambiente de teste")
    c1, c2 = st.columns([3, 1])
    with c1:
        st.selectbox("Preencher automaticamente", TEST_PROFILES, key="perfil_teste")
    with c2:
        st.markdown("<div style='height:1.72rem'></div>", unsafe_allow_html=True)
        if st.button("Preencher teste", use_container_width=True):
            fill_demo_profile(st.session_state.perfil_teste)
            clear_result()
            st.rerun()


def render_identification_stage() -> None:
    apply_pending_identification_actions()

    with st.container(border=True):
        st.markdown(stage_header(1, ICON_USER, "Identificação e antropometria"), unsafe_allow_html=True)
        render_environment_controls()

        row1 = st.columns([2.1, 1.1, 0.75, 0.9, 0.9])
        with row1[0]:
            st.text_input("Nome do paciente", key="nome", on_change=update_evaluation_id)
        with row1[1]:
            st.selectbox("Sexo biológico", ["Masculino", "Feminino"], key="sexo_label", on_change=reset_patient_progress)
        with row1[2]:
            st.number_input("Idade (anos)", min_value=1, max_value=120, step=1, key="idade", on_change=update_evaluation_id)
        with row1[3]:
            st.number_input("Peso (kg)", min_value=1.0, max_value=350.0, step=0.1, key="peso", on_change=reset_patient_progress)
        with row1[4]:
            st.number_input("Altura (cm)", min_value=50.0, max_value=240.0, step=0.1, key="altura_cm", on_change=reset_patient_progress)

        row2 = st.columns([1, 1.25, 1.25, 1.7])
        with row2[0]:
            st.date_input("Data da avaliação", key="data_avaliacao", on_change=update_evaluation_id)
        with row2[1]:
            st.text_input("Prontuário/ID", key="prontuario", on_change=mark_manual_id)
        with row2[2]:
            st.text_input("Avaliador", key="avaliador", on_change=reset_patient_progress)
        with row2[3]:
            st.text_input("Diagnóstico/condição clínica", key="diagnostico", on_change=reset_patient_progress)

        row3 = st.columns([1.25, 0.85, 1.0, 2.0])
        with row3[0]:
            st.number_input(
                "Comprimento do membro inferior (m, opcional)",
                min_value=0.0,
                max_value=1.50,
                step=0.01,
                format="%.2f",
                key="comprimento_membro_inferior_m",
                on_change=reset_patient_progress,
                help="Informe em metros, por exemplo: 0.82, 0.90 ou 0.95. Se ficar vazio/0, a velocidade normalizada não será calculada.",
            )
        with row3[1]:
            st.number_input("Nº do teste no dia", min_value=1, max_value=99, step=1, key="numero_teste", on_change=update_evaluation_id)
        with row3[2]:
            st.markdown("<div style='height:1.72rem'></div>", unsafe_allow_html=True)
            st.button("Gerar ID", use_container_width=True, on_click=force_evaluation_id)
        with row3[3]:
            st.markdown(
                "<div class='soft-note' style='margin-top:1.9rem;font-size:.82rem;line-height:1.25;color:#66736F!important;'>"
                "Campo opcional: usado apenas na velocidade normalizada."
                "</div>",
                unsafe_allow_html=True,
            )

        if st.session_state.patient_saved:
            st.success("Dados do paciente salvos. A triagem de segurança está liberada.")

        if st.button("Salvar dados do paciente e liberar triagem", type="primary", use_container_width=True, key="save_patient_stage"):
            if not st.session_state.nome.strip():
                st.error("Informe o nome do paciente antes de continuar.")
            else:
                if st.session_state.get("id_auto_ativo", True):
                    st.session_state.pending_force_id = True
                    st.session_state.pending_patient_save = True
                    st.rerun()
                else:
                    st.session_state.patient_saved = True
                    st.session_state.prediction_saved = False
                    st.session_state.execution_saved = False
                    st.session_state.recovery_saved = False


def render_triage_stage() -> None:
    with st.container(border=True):
        st.markdown(stage_header(2, ICON_SHIELD_CHECK, "Triagem de segurança"), unsafe_allow_html=True)
        st.markdown('<div class="table-label">Resultado da triagem</div>', unsafe_allow_html=True)
        triage_status = st.session_state.triagem_status
        ok_label = "✓ Sem contraindicações"
        relative_label = "! Contraindicação relativa"
        absolute_label = "! Contraindicação absoluta"
        selected_css = {
            "Sem contraindicações": """
                .st-key-triage_ok_btn button {
                    box-shadow: 0 0 0 3px rgba(14,122,79,.25), 0 12px 24px rgba(14,122,79,.22) !important;
                    transform: translateY(-1px);
                }
                .st-key-triage_relative_btn button, .st-key-triage_absolute_btn button { opacity: .72; }
            """,
            "Contraindicação relativa": """
                .st-key-triage_relative_btn button {
                    box-shadow: 0 0 0 3px rgba(217,139,24,.28), 0 12px 24px rgba(217,139,24,.16) !important;
                    transform: translateY(-1px);
                }
                .st-key-triage_ok_btn button, .st-key-triage_absolute_btn button { opacity: .72; }
            """,
            "Contraindicação absoluta": """
                .st-key-triage_absolute_btn button {
                    box-shadow: 0 0 0 3px rgba(217,74,74,.28), 0 12px 24px rgba(217,74,74,.16) !important;
                    transform: translateY(-1px);
                }
                .st-key-triage_ok_btn button, .st-key-triage_relative_btn button { opacity: .72; }
            """,
        }
        st.markdown(f"<style>{selected_css.get(triage_status, '')}</style>", unsafe_allow_html=True)
        b1, b2, b3 = st.columns(3)
        with b1:
            if st.button(ok_label, use_container_width=True, key="triage_ok_btn"):
                set_triage("Sem contraindicações")
                st.rerun()
        with b2:
            if st.button(relative_label, use_container_width=True, key="triage_relative_btn"):
                set_triage("Contraindicação relativa")
                st.rerun()
        with b3:
            if st.button(absolute_label, use_container_width=True, key="triage_absolute_btn"):
                set_triage("Contraindicação absoluta")
                st.rerun()

        triage_observation = st.session_state.observacao_triagem or "Responda a triagem para liberar a avaliação prévia."
        st.markdown(triage_html(st.session_state.triagem_status, triage_observation), unsafe_allow_html=True)
        st.text_area("Observações da triagem", key="observacao_triagem", on_change=clear_result)


def render_prediction_stage(fc_max: int, fc_submax: int, dpp_principal: float, lin_principal: float) -> None:
    with st.container(border=True):
        st.markdown(stage_header(3, ICON_CALCULATOR, "Avaliação prévia e cálculos preditos"), unsafe_allow_html=True)
        st.selectbox("Fórmula selecionada", FORMULAS_DPP, key="formula_principal", on_change=clear_result)
        st.markdown(
            metric_grid(
                [
                    ("DPP principal", f"{br_number(dpp_principal)} m", False),
                    ("LIN principal", f"{br_number(lin_principal)} m" if lin_principal is not None else "Não definido", False),
                    ("FC máxima prevista", f"{fc_max} bpm", False),
                    ("FC submáxima (85%)", f"{fc_submax} bpm", False),
                    ("Fórmula selecionada", clean_text(st.session_state.formula_principal).split(" - ")[0], True),
                ]
            ),
            unsafe_allow_html=True,
        )
        with st.expander("Ver todas as fórmulas preditas"):
            st.write(f"Enright/Sherrill: {br_number(dpp_enright)} m | LIN: {br_number(lin_enright)} m")
            st.write(f"Iwama et al.: {br_number(dpp_iwama)} m")
            st.write(f"Ben Saad et al.: {br_number(dpp_ben_saad)} m")

        st.markdown("#### Sinais de repouso antes do teste")
        pre_editor = prepare_editor_with_bp(st.session_state.pre_df)
        pre_resp_col = borg_resp_col(pre_editor)
        pre_mmii_col = borg_mmii_col(pre_editor)
        st.session_state.pre_df = restore_pas_pad(
            st.data_editor(
                pre_editor,
                key="pre_editor_main",
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "Tempo": st.column_config.TextColumn("Momento", disabled=True),
                    "FC": st.column_config.NumberColumn("FC (bpm)", min_value=0, max_value=260, step=1),
                    "SpO2": st.column_config.NumberColumn("SpO2 (%)", min_value=0, max_value=100, step=1),
                    "FR": st.column_config.NumberColumn("FR (ipm)", min_value=0, max_value=80, step=1),
                    "PA": st.column_config.TextColumn("PA (mmHg)"),
                    pre_resp_col: st.column_config.NumberColumn("Borg dispneia", min_value=0.0, max_value=10.0, step=0.5),
                    pre_mmii_col: st.column_config.NumberColumn("Borg MMII", min_value=0.0, max_value=10.0, step=0.5),
                },
            )
        )


def render_execution_stage() -> None:
    with st.container(border=True):
        st.markdown(stage_header(4, ICON_FOOTPRINTS, "Execução do teste"), unsafe_allow_html=True)
        st.markdown('<div class="soft-note">Durante a caminhada, registre apenas FC, SpO2 e Borg.</div>', unsafe_allow_html=True)
        during_resp_col = borg_resp_col(st.session_state.during_df)
        during_mmii_col = borg_mmii_col(st.session_state.during_df)
        st.session_state.during_df = st.data_editor(
            st.session_state.during_df,
            key="during_editor_main",
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "Tempo": st.column_config.TextColumn("Minuto", disabled=True),
                "FC": st.column_config.NumberColumn("FC (bpm)", min_value=0, max_value=260, step=1),
                "SpO2": st.column_config.NumberColumn("SpO2 (%)", min_value=0, max_value=100, step=1),
                during_resp_col: st.column_config.NumberColumn("Borg dispneia", min_value=0.0, max_value=10.0, step=0.5),
                during_mmii_col: st.column_config.NumberColumn("Borg MMII", min_value=0.0, max_value=10.0, step=0.5),
            },
        )


def render_recovery_stage() -> None:
    with st.container(border=True):
        st.markdown(stage_header(5, ICON_REFRESH, "Recuperação"), unsafe_allow_html=True)
        st.markdown('<div class="soft-note">Após o teste, registre sinais completos em 1, 3 e 6 minutos.</div>', unsafe_allow_html=True)
        recovery_editor = prepare_editor_with_bp(st.session_state.recovery_df)
        recovery_resp_col = borg_resp_col(recovery_editor)
        recovery_mmii_col = borg_mmii_col(recovery_editor)
        st.session_state.recovery_df = restore_pas_pad(
            st.data_editor(
                recovery_editor,
                key="recovery_editor_main",
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "Tempo": st.column_config.TextColumn("Momento", disabled=True),
                    "FC": st.column_config.NumberColumn("FC (bpm)", min_value=0, max_value=260, step=1),
                    "SpO2": st.column_config.NumberColumn("SpO2 (%)", min_value=0, max_value=100, step=1),
                    "FR": st.column_config.NumberColumn("FR (ipm)", min_value=0, max_value=80, step=1),
                    "PA": st.column_config.TextColumn("PA (mmHg)"),
                    recovery_resp_col: st.column_config.NumberColumn("Borg dispneia", min_value=0.0, max_value=10.0, step=0.5),
                    recovery_mmii_col: st.column_config.NumberColumn("Borg MMII", min_value=0.0, max_value=10.0, step=0.5),
                },
            )
        )


def render_final_test_inputs() -> None:
    with st.container(border=True):
        st.markdown("#### Dados finais do teste")
        r1, r2, r3 = st.columns(3)
        with r1:
            st.number_input("Distância percorrida ao final do TC6M (m)", min_value=0.0, max_value=2000.0, step=1.0, key="distancia", on_change=clear_result)
        with r2:
            st.radio("O paciente interrompeu o teste?", ["Não", "Sim"], horizontal=True, key="interrompeu_label", on_change=clear_result)
        with r3:
            if st.session_state.interrompeu_label == "Sim":
                st.number_input("Distância no momento da interrupção (m)", min_value=0.0, max_value=2000.0, step=1.0, key="distancia_interrupcao", on_change=clear_result)
        if st.session_state.interrompeu_label == "Sim":
            st.text_area("Motivo da interrupção", key="motivo_interrupcao", on_change=clear_result)
        if st.button("Gerar resumo final do TC6M", type="primary", use_container_width=True):
            try:
                generate_result()
                st.rerun()
            except ValueError as error:
                st.error(str(error))


init_state()
sync_navigation_from_query()

inject_css()
render_app_sidebar()
render_header()

st.session_state.contra_abs = st.session_state.triagem_status == "Contraindicação absoluta"
st.session_state.contra_rel = st.session_state.triagem_status == "Contraindicação relativa"

sex = "M" if st.session_state.sexo_label == "Masculino" else "F"
fc_max = calcular_fc_maxima(int(st.session_state.idade))
fc_submax = calcular_fc_submaxima(int(st.session_state.idade))
dpp_principal, lin_principal = calcular_dpp_por_formula(
    st.session_state.formula_principal,
    sex,
    int(st.session_state.idade),
    float(st.session_state.peso),
    float(st.session_state.altura_cm),
)
dpp_enright, lin_enright = calcular_dpp_enright(sex, int(st.session_state.idade), float(st.session_state.peso), float(st.session_state.altura_cm))
dpp_iwama = calcular_dpp_iwama(sex, int(st.session_state.idade))
dpp_ben_saad = calcular_dpp_ben_saad(int(st.session_state.idade), float(st.session_state.peso), float(st.session_state.altura_cm))
current_section = st.session_state.get("nav_section", "avaliacao")

execution_body = (
    '<div class="soft-note">Durante a caminhada, registre apenas FC, SpO2 e Borg.</div>'
    + dataframe_to_table(display_vitals_table(st.session_state.during_df, include_full=False))
)
recovery_body = (
    '<div class="soft-note">Após o teste, registre sinais completos em 1, 3 e 6 minutos.</div>'
    + dataframe_to_table(display_vitals_table(st.session_state.recovery_df, include_full=True))
)

if current_section == "execucao":
    if not st.session_state.get("patient_saved", False):
        st.warning("Salve os dados do paciente na avaliação antes de acessar a execução.")
    elif st.session_state.triagem_status == "Selecione" or st.session_state.contra_abs:
        st.warning("Responda e libere a triagem de segurança antes de acessar a execução.")
    elif not st.session_state.get("prediction_saved", False):
        st.warning("Salve a avaliação prévia antes de acessar a execução.")
    else:
        render_execution_stage()
        if not st.session_state.get("execution_saved", False):
            if st.button("Salvar execução e liberar recuperação", type="primary", use_container_width=True, key="save_execution_stage_sidebar"):
                st.session_state.execution_saved = True
                st.session_state.recovery_saved = False
                clear_result()
                st.rerun()
            st.stop()
        render_recovery_stage()
        if not st.session_state.get("recovery_saved", False):
            if st.button("Salvar recuperação e liberar resultado final", type="primary", use_container_width=True, key="save_recovery_stage_sidebar"):
                st.session_state.recovery_saved = True
                clear_result()
                st.rerun()
            st.stop()
        render_final_test_inputs()
    st.stop()

if current_section == "resultados":
    if st.session_state.resultado_tc6m and st.session_state.paciente_tc6m is not None:
        render_stage_card(
            6,
            ICON_GAUGE,
            "Resultado final",
            result_dashboard_html(
                st.session_state.paciente_tc6m,
                st.session_state.resultado_tc6m,
                st.session_state.serie_tc6m,
            ),
        )
        st.markdown(
            integrated_recovery_html(
                st.session_state.paciente_tc6m,
                st.session_state.serie_tc6m,
            ),
            unsafe_allow_html=True,
        )
    else:
        st.info("Gere o resumo final do TC6M para visualizar os resultados.")
        render_final_test_inputs()
    st.stop()

if current_section == "graficos":
    if st.session_state.resultado_tc6m and st.session_state.paciente_tc6m is not None:
        patient = st.session_state.paciente_tc6m
        series = st.session_state.serie_tc6m
        st.markdown(graph_metrics_html(series), unsafe_allow_html=True)
        chart_card_start("Oscilação cardiorrespiratória durante o TC6M")
        st.pyplot(build_dashboard_oscillation_figure(series), use_container_width=True)
        chart_card_end()
        chart_card_start("Curva de esforço percebido")
        st.pyplot(build_dashboard_effort_figure(series), use_container_width=True)
        st.markdown(chart_note_html(series), unsafe_allow_html=True)
        chart_card_end()
        render_dp_recovery_chart(patient, series)
    else:
        st.info("Gere o resumo final do TC6M para visualizar os gráficos.")
    st.stop()

if current_section == "relatorio":
    if st.session_state.resultado_tc6m and st.session_state.paciente_tc6m is not None:
        patient = st.session_state.paciente_tc6m
        result = st.session_state.resultado_tc6m
        series = st.session_state.serie_tc6m
        excel_bytes = build_excel_bytes(patient, result, series)
        pdf_bytes = build_pdf_bytes(patient, result, series)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Baixar Excel estruturado",
                data=excel_bytes,
                file_name=build_safe_filename(patient.nome, "xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "Baixar PDF clínico",
                data=pdf_bytes,
                file_name=build_safe_filename(patient.nome, "pdf"),
                mime="application/pdf",
                use_container_width=True,
            )
    else:
        st.info("Gere o resumo final do TC6M para liberar os arquivos de relatório.")
    st.stop()

render_identification_stage()

if not st.session_state.get("patient_saved", False):
    st.info("Salve os dados do paciente para liberar a triagem de segurança.")
    st.stop()

render_triage_stage()

if st.session_state.triagem_status == "Selecione":
    st.warning("Trava ativa: responda a triagem de segurança para liberar o protocolo.")
elif st.session_state.contra_abs:
    st.error("Trava de segurança ativada: há contraindicação absoluta. O teste não deve ser iniciado.")

if st.session_state.triagem_status == "Selecione" or st.session_state.contra_abs:
    st.stop()

render_prediction_stage(fc_max, fc_submax, dpp_principal, lin_principal)

if not st.session_state.get("prediction_saved", False):
    if st.button("Salvar avaliação prévia e liberar execução", type="primary", use_container_width=True, key="save_prediction_stage"):
        st.session_state.prediction_saved = True
        st.session_state.execution_saved = False
        st.session_state.recovery_saved = False
        clear_result()
        st.rerun()
    st.info("Depois de conferir a fórmula, FC submáxima e sinais de repouso, salve a avaliação prévia para liberar a execução.")
    st.stop()

render_execution_stage()

if not st.session_state.get("execution_saved", False):
    if st.button("Salvar execução e liberar recuperação", type="primary", use_container_width=True, key="save_execution_stage"):
        st.session_state.execution_saved = True
        st.session_state.recovery_saved = False
        clear_result()
        st.rerun()
    st.info("Registre os dados minuto a minuto e salve a execução para liberar a recuperação.")
    st.stop()

render_recovery_stage()

if not st.session_state.get("recovery_saved", False):
    if st.button("Salvar recuperação e liberar resultado final", type="primary", use_container_width=True, key="save_recovery_stage"):
        st.session_state.recovery_saved = True
        clear_result()
        st.rerun()
    st.info("Registre os sinais de recuperação em 1, 3 e 6 minutos para liberar os dados finais do teste.")
    st.stop()

render_final_test_inputs()

if st.session_state.resultado_tc6m and st.session_state.paciente_tc6m is not None:
    result_body = result_dashboard_html(
        st.session_state.paciente_tc6m,
        st.session_state.resultado_tc6m,
        st.session_state.serie_tc6m,
    )
else:
    result_body = ""
if result_body:
    render_stage_card(6, ICON_GAUGE, "Resultado final", result_body)
    st.markdown(
        integrated_recovery_html(
            st.session_state.paciente_tc6m,
            st.session_state.serie_tc6m,
        ),
        unsafe_allow_html=True,
    )

if st.session_state.resultado_tc6m and st.session_state.paciente_tc6m is not None:
    patient = st.session_state.paciente_tc6m
    result = st.session_state.resultado_tc6m
    series = st.session_state.serie_tc6m
    tab1, tab2, tab3, tab4 = st.tabs(["Achados", "Gráficos", "Exportar", "Dados brutos"])
    with tab1:
        st.markdown(achados_panel_html(series), unsafe_allow_html=True)
    with tab2:
        st.markdown(graph_metrics_html(series), unsafe_allow_html=True)
        chart_card_start("Oscilação cardiorrespiratória durante o TC6M")
        st.pyplot(build_dashboard_oscillation_figure(series), use_container_width=True)
        chart_card_end()
        chart_card_start("Curva de esforço percebido")
        st.pyplot(build_dashboard_effort_figure(series), use_container_width=True)
        st.markdown(chart_note_html(series), unsafe_allow_html=True)
        chart_card_end()
        render_dp_recovery_chart(patient, series)

    with tab3:
        excel_bytes = build_excel_bytes(patient, result, series)
        pdf_bytes = build_pdf_bytes(patient, result, series)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Baixar Excel estruturado",
                data=excel_bytes,
                file_name=build_safe_filename(patient.nome, "xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "Baixar PDF clínico",
                data=pdf_bytes,
                file_name=build_safe_filename(patient.nome, "pdf"),
                mime="application/pdf",
                use_container_width=True,
            )
    with tab4:
        st.dataframe(series, hide_index=True, use_container_width=True)
