from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
import math
from pathlib import Path
from typing import Iterable

import matplotlib.pyplot as plt
import pandas as pd
from openpyxl.styles import Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfgen import canvas

try:
    from reportlab.graphics import renderPDF
    from svglib.svglib import svg2rlg
except Exception:
    renderPDF = None
    svg2rlg = None


FORMULAS_DPP = [
    "Enright e Sherrill (1998) - adultos",
    "Iwama et al. (2009) - adultos",
    "Ben Saad et al. (2009) - crianças/adolescentes",
]

COMPRIMENTOS_CORREDOR_SUPORTADOS = {25.0, 30.0}

FULL_COLUMNS = [
    "Tempo",
    "FC",
    "SpO2",
    "FR",
    "PAS",
    "PAD",
    "Borg Respiratório",
    "Borg MMII",
]

DURING_COLUMNS = [
    "Tempo",
    "FC",
    "SpO2",
    "Borg Respiratório",
    "Borg MMII",
]

NUMERIC_COLUMNS = [
    "FC",
    "SpO2",
    "FR",
    "PAS",
    "PAD",
    "Borg Respiratório",
    "Borg MMII",
]

REPORT_COLORS = {
    "primary": "#0B4238",
    "secondary": "#2F5D55",
    "background": "#EAF3DE",
    "background_deep": "#DDEFE8",
    "text": "#1A1A1A",
    "muted": "#6B6B67",
    "border": "#D8D8D2",
    "fc": "#A32D2D",
    "spo2": "#185FA5",
    "progress": "#BA7517",
    "risk_low": "#639922",
    "risk_moderate": "#BA7517",
    "risk_high": "#E24B4A",
    "risk_very_high": "#A32D2D",
    "ok_bg": "#EAF3DE",
    "ok_text": "#3B6D11",
    "warning_bg": "#FAEEDA",
    "warning_text": "#854F0B",
    "danger_bg": "#FCEBEB",
    "danger_text": "#A32D2D",
}

ICON_DIR = Path(__file__).resolve().parent / "assets" / "icons"

ICON_MAP = {
    "achados_automaticos": ICON_DIR / "achados_automaticos.png",
    "caminhada_tc6m": ICON_DIR / "caminhada_tc6m.png",
    "classificacao_risco": ICON_DIR / "classificacao_risco.png",
    "distancia_percorrida": ICON_DIR / "distancia_percorrida.png",
    "resultado_principal": ICON_DIR / "distancia_percorrida.png",
    "duplo_produto_recuperacao": ICON_DIR / "duplo_produto_recuperacao.png",
    "fator_limitante": ICON_DIR / "fator_limitante.png",
    "graficos_achados": ICON_DIR / "graficos_achados.png",
    "metricas_hemodinamicas": ICON_DIR / "metricas_hemodinamicas.png",
    "interpretacao_integrada": ICON_DIR / "interpretacao_integrada.png",
    "nota_metodologica": ICON_DIR / "nota_metodologica.png",
    "pontos_atencao": ICON_DIR / "pontos_atencao.png",
    "predicoes_comparativas": ICON_DIR / "predicoes_comparativas.png",
    "resumo_clinico": ICON_DIR / "resumo_clinico.png",
    "spo2": ICON_DIR / "spo2.png",
    # Aliases visuais usando apenas a biblioteca oficial disponível.
    "borg_respiratorio": ICON_DIR / "metricas_hemodinamicas.png",
    "dp_repouso": ICON_DIR / "metricas_hemodinamicas.png",
    "dp_recuperacao": ICON_DIR / "duplo_produto_recuperacao.png",
    "velocidade_media": ICON_DIR / "graficos_achados.png",
    "velocidade_normalizada": ICON_DIR / "interpretacao_integrada.png",
}

ICON_ALIASES = {
    "walk": "caminhada_tc6m",
    "shoe": "distancia_percorrida",
    "shield": "classificacao_risco",
    "chart": "graficos_achados",
    "heart": "metricas_hemodinamicas",
    "info": "interpretacao_integrada",
    "alert": "pontos_atencao",
    "summary": "resumo_clinico",
    "search": "achados_automaticos",
}


@dataclass
class PatientData:
    """Dados de identificação, antropometria, triagem e resultado final do TC6M."""

    nome: str
    sexo: str
    idade: int
    peso: float
    altura_cm: float
    distancia: float
    interrompeu: bool
    formula_principal: str
    data_avaliacao: date | None = None
    prontuario: str = ""
    avaliador: str = ""
    diagnostico: str = ""
    comprimento_membro_inferior_m: float | None = None
    comprimento_corredor_m: float = 30.0
    motivo_interrupcao: str = ""
    distancia_interrupcao: float = 0.0
    contraindicacao_absoluta: bool = False
    contraindicacao_relativa: bool = False
    observacao_triagem: str = ""


@dataclass
class VitalSnapshot:
    """Representa uma fase clínica do teste: repouso, pico ou recuperação."""

    tempo: str
    fc: int
    spo2: int
    fr: int
    pas: int
    pad: int
    borg_resp: float
    borg_mmii: float


@dataclass
class TestResult:
    """Agrupa os resultados calculados para exibição, PDF e Excel."""

    formula_principal: str
    dpp_principal: float
    lin_principal: float | None
    dpp_enright: float
    lin_enright: float
    dpp_iwama: float
    dpp_ben_saad: float
    percentual_atingido: float
    qualificador_funcional: str
    classificacao_risco: str
    risco: str
    dp_repouso: int
    dp_pico: int
    dp_recuperacao: int
    fator_limitante: str
    interpretacao: str


def validate_patient_data(data: PatientData) -> None:
    """Valida os campos obrigatórios antes de executar os cálculos clínicos."""

    if not data.nome.strip():
        raise ValueError("Informe o nome do paciente.")
    if data.sexo not in {"M", "F"}:
        raise ValueError("Sexo biológico deve ser masculino ou feminino.")
    if data.idade <= 0:
        raise ValueError("Idade deve ser maior que zero.")
    if data.peso <= 0:
        raise ValueError("Peso deve ser maior que zero.")
    if data.altura_cm <= 0:
        raise ValueError("Altura deve ser maior que zero.")
    if data.distancia < 0:
        raise ValueError("Distância percorrida não pode ser negativa.")
    if data.comprimento_membro_inferior_m is not None and data.comprimento_membro_inferior_m < 0:
        raise ValueError("Comprimento do membro inferior não pode ser negativo.")
    if float(data.comprimento_corredor_m) not in COMPRIMENTOS_CORREDOR_SUPORTADOS:
        raise ValueError("Selecione um protocolo de corredor válido: 30 m ou 25 m.")
    if data.formula_principal not in FORMULAS_DPP:
        raise ValueError("Selecione uma fórmula predita válida.")


def calcular_distancia_por_trechos(
    comprimento_corredor_m: float,
    trechos_completos: int,
    metros_adicionais: float,
) -> float:
    """Calcula a distância real sem modificar as equações preditivas da literatura."""

    comprimento = float(comprimento_corredor_m)
    trechos = int(trechos_completos)
    adicionais = float(metros_adicionais)
    if comprimento not in COMPRIMENTOS_CORREDOR_SUPORTADOS:
        raise ValueError("Selecione um protocolo de corredor válido: 30 m ou 25 m.")
    if trechos < 0:
        raise ValueError("A quantidade de trechos completos não pode ser negativa.")
    if adicionais < 0 or adicionais >= comprimento:
        raise ValueError("Os metros adicionais devem ser menores que o comprimento do corredor.")
    return (trechos * comprimento) + adicionais


def descrever_protocolo_corredor(comprimento_corredor_m: float) -> str:
    """Explica a condição metodológica do percurso usado no TC6M."""

    comprimento = float(comprimento_corredor_m)
    if comprimento == 30.0:
        return "Protocolo padrão ATS/ERS: corredor de 30 m."
    return (
        f"Protocolo adaptado: corredor de {comprimento:.0f} m. "
        "O maior número de retornos pode reduzir a distância percorrida em comparação ao protocolo padrão de 30 m. "
        "As equações preditivas foram mantidas sem correção proporcional."
    )


def calcular_fc_maxima(idade: int) -> int:
    """Calcula FC máxima estimada pela fórmula simples: 220 - idade."""

    return int(round(220 - idade))


def calcular_fc_submaxima(idade: int) -> int:
    """Calcula 85% da FC máxima estimada, usada como referência de segurança."""

    return int(round(calcular_fc_maxima(idade) * 0.85))


def calcular_dpp_enright(sexo: str, idade: int, peso: float, altura_cm: float) -> tuple[float, float]:
    """Calcula DPP e limite inferior de normalidade por Enright e Sherrill."""

    if sexo == "M":
        dpp = (7.57 * altura_cm) - (5.02 * idade) - (1.76 * peso) - 309
        lin = dpp - 153
    elif sexo == "F":
        dpp = (2.11 * altura_cm) - (5.78 * idade) - (2.29 * peso) + 667
        lin = dpp - 139
    else:
        raise ValueError("Sexo biológico deve ser masculino ou feminino.")

    return dpp, lin


def calcular_dpp_iwama(sexo: str, idade: int) -> float:
    """Calcula DPP pela equação de Iwama et al.; homem=1 e mulher=0."""

    genero = 1 if sexo == "M" else 0
    return 622.461 - (1.846 * idade) + (61.503 * genero)


def calcular_dpp_ben_saad(idade: int, peso: float, altura_cm: float) -> float:
    """Calcula DPP por Ben Saad et al., referência para crianças/adolescentes."""

    return (4.63 * altura_cm) - (3.53 * peso) + (10.42 * idade) + 56.32


def calcular_dpp_por_formula(
    formula: str,
    sexo: str,
    idade: int,
    peso: float,
    altura_cm: float,
) -> tuple[float, float | None]:
    """Calcula a DPP principal conforme a fórmula escolhida na interface."""

    if formula == "Enright e Sherrill (1998) - adultos":
        return calcular_dpp_enright(sexo, idade, peso, altura_cm)
    if formula == "Iwama et al. (2009) - adultos":
        return calcular_dpp_iwama(sexo, idade), None
    if formula == "Ben Saad et al. (2009) - crianças/adolescentes":
        return calcular_dpp_ben_saad(idade, peso, altura_cm), None

    raise ValueError("Fórmula predita não reconhecida.")


def obter_qualificador_funcional(distancia_real: float, dpp_principal: float) -> tuple[str, float]:
    """Retorna o nível de déficit funcional e o percentual exato atingido."""

    if dpp_principal <= 0:
        raise ValueError("A DPP principal precisa ser maior que zero.")

    percentual = (distancia_real / dpp_principal) * 100

    if percentual >= 96:
        qualificador = "Nenhum déficit funcional"
    elif 76 <= percentual <= 95:
        qualificador = "Déficit funcional leve"
    elif 51 <= percentual <= 75:
        qualificador = "Déficit funcional moderado"
    elif 5 <= percentual <= 50:
        qualificador = "Déficit funcional grave"
    else:
        qualificador = "Déficit funcional completo"

    return qualificador, percentual


def calcular_duplo_produto(fc: int | float, pas: int | float) -> int:
    """Calcula o duplo produto: frequência cardíaca x pressão arterial sistólica."""

    if float(fc) <= 0 or float(pas) <= 0:
        return 0
    return int(round(float(fc) * float(pas)))


METODOLOGICAL_DP_NOTICE = (
    "O Duplo Produto foi calculado em repouso e nos momentos de recuperação pós-TC6M: "
    "1, 3 e 6 minutos. Esses valores representam estimativas pontuais da carga cardiovascular "
    "antes e após o esforço, não correspondendo ao pico durante a caminhada."
)

NOT_CALCULATED = "Não calculado por ausência de dados válidos."


def _valid_positive(value: object) -> bool:
    """Retorna True apenas para valores numéricos positivos."""

    try:
        return float(value) > 0
    except (TypeError, ValueError):
        return False


def _safe_round(value: float | None, digits: int = 2) -> float | None:
    """Arredonda somente quando o valor existe."""

    if value is None:
        return None
    return round(float(value), digits)


def _dp_from_row(row: pd.Series | None) -> int | None:
    """Calcula DP de uma linha da série apenas se FC e PAS forem válidas."""

    if row is None:
        return None
    fc = row.get("FC", 0)
    pas = row.get("PAS", 0)
    if not _valid_positive(fc) or not _valid_positive(pas):
        return None
    return calcular_duplo_produto(fc, pas)


def _recovery_row(clean: pd.DataFrame, minute: str) -> pd.Series | None:
    """Busca uma linha de recuperação pelo minuto informado."""

    tempo = clean["Tempo"].astype(str).str.lower()
    mask = tempo.str.contains("recupera") & tempo.str.contains(minute)
    if mask.any():
        return clean.loc[mask].iloc[0]
    return None


def build_integrated_recovery_analysis(data: PatientData, timeseries_df: pd.DataFrame) -> dict:
    """Calcula recuperação do DP, custo cardiovascular e velocidade normalizada."""

    clean = normalize_timeseries(timeseries_df)
    repouso = clean.iloc[0] if not clean.empty else None
    rec_1 = _recovery_row(clean, "1")
    rec_3 = _recovery_row(clean, "3")
    rec_6 = _recovery_row(clean, "6")

    dp_repouso = _dp_from_row(repouso)
    dp_1 = _dp_from_row(rec_1)
    dp_3 = _dp_from_row(rec_3)
    dp_6 = _dp_from_row(rec_6)

    delta_dp_1 = dp_1 - dp_repouso if dp_1 is not None and dp_repouso is not None else None
    recovery_dp_1_3 = dp_1 - dp_3 if dp_1 is not None and dp_3 is not None else None
    recovery_dp_1_6 = dp_1 - dp_6 if dp_1 is not None and dp_6 is not None else None
    recovery_percent_6 = ((dp_1 - dp_6) / dp_1) * 100 if dp_1 and dp_6 is not None else None
    return_to_baseline = dp_6 - dp_repouso if dp_6 is not None and dp_repouso is not None else None

    distance = float(data.distancia or 0)
    cost_dp_per_m = (delta_dp_1 / distance) if delta_dp_1 is not None and distance > 0 else None
    velocity_ms = distance / 360 if distance > 0 else None
    pace_m_min = distance / 6 if distance > 0 else None

    limb_length = data.comprimento_membro_inferior_m
    normalized_velocity = None
    if velocity_ms is not None and limb_length is not None and limb_length > 0:
        normalized_velocity = velocity_ms / math.sqrt(9.81 * limb_length)

    dp_values = {
        "Repouso": dp_repouso,
        "1 min": dp_1,
        "3 min": dp_3,
        "6 min": dp_6,
    }

    interpretations: list[str] = []
    if dp_1 is not None and dp_repouso is not None and dp_6 is not None:
        if dp_1 > dp_repouso and abs(dp_6 - dp_repouso) <= max(dp_repouso * 0.15, 1000):
            interpretations.append(
                "Houve elevação do Duplo Produto no pós-teste imediato, com tendência de recuperação ao longo dos 6 minutos."
            )
        elif dp_6 > dp_repouso * 1.15:
            interpretations.append(
                "O Duplo Produto permaneceu elevado após 6 minutos de recuperação, sugerindo recuperação cardiovascular estimada mais lenta. Interpretar em conjunto com sintomas, SpO2, Borg, PA e frequência cardíaca."
            )
        elif delta_dp_1 is not None and delta_dp_1 <= max(dp_repouso * 0.10, 800) and distance < 300:
            interpretations.append(
                "Baixa elevação do Duplo Produto associada a baixa distância pode sugerir esforço submáximo, limitação musculoesquelética, dor, baixa tolerância periférica ou interrupção precoce."
            )
    else:
        interpretations.append("Não foi possível interpretar a recuperação do Duplo Produto por ausência de dados válidos.")

    if cost_dp_per_m is not None:
        if cost_dp_per_m >= 25:
            interpretations.append(
                "O paciente apresentou maior aumento cardiovascular estimado para a distância percorrida, sugerindo maior custo cardiovascular relativo ao desempenho funcional obtido."
            )
        elif cost_dp_per_m <= 8:
            interpretations.append(
                "O aumento cardiovascular estimado por metro foi menor, devendo ser interpretado em conjunto com distância, sintomas e percepção de esforço."
            )
        else:
            interpretations.append(
                "O custo cardiovascular estimado por metro ficou em faixa intermediária e deve ser interpretado junto à distância, sintomas e sinais vitais."
            )
    else:
        interpretations.append("Custo cardiovascular estimado por metro não calculado por ausência de dados válidos.")

    if velocity_ms is not None and pace_m_min is not None:
        interpretations.append("A velocidade média representa o ritmo funcional médio produzido durante os 6 minutos de teste.")
    else:
        interpretations.append("Velocidade média não calculada por ausência de distância válida.")

    if normalized_velocity is not None:
        interpretations.append(
            "A velocidade normalizada ajusta a velocidade média ao comprimento do membro inferior, oferecendo uma análise biomecânica complementar da exigência locomotora relativa."
        )
    else:
        interpretations.append("Velocidade normalizada não calculada por ausência de comprimento do membro inferior válido.")

    return {
        "notice": METODOLOGICAL_DP_NOTICE,
        "dp_values": dp_values,
        "dp_repouso": dp_repouso,
        "dp_1": dp_1,
        "dp_3": dp_3,
        "dp_6": dp_6,
        "delta_dp_1": delta_dp_1,
        "recovery_dp_1_3": recovery_dp_1_3,
        "recovery_dp_1_6": recovery_dp_1_6,
        "recovery_percent_6": _safe_round(recovery_percent_6),
        "return_to_baseline": return_to_baseline,
        "cost_dp_per_m": _safe_round(cost_dp_per_m),
        "velocity_ms": _safe_round(velocity_ms),
        "pace_m_min": _safe_round(pace_m_min),
        "limb_length_m": limb_length if limb_length and limb_length > 0 else None,
        "normalized_velocity": _safe_round(normalized_velocity, 3),
        "interpretations": interpretations,
    }


def build_integrated_recovery_interpretation(analysis: dict) -> str:
    """Texto interpretativo padronizado para a análise integrada no PDF."""

    if analysis.get("dp_1") is None or analysis.get("dp_repouso") is None:
        return (
            "Não foi possível interpretar completamente a recuperação cardiovascular estimada por ausência de dados "
            "válidos de Duplo Produto. Quando disponíveis, os indicadores devem ser interpretados junto à distância, "
            "sintomas, sinais vitais e percepção de esforço."
        )

    base_text = (
        "Houve elevação do Duplo Produto no pós-teste imediato, com tendência de recuperação ao longo dos 6 minutos. "
        "O custo cardiovascular estimado por metro ficou em faixa intermediária e deve ser interpretado junto à "
        "distância, sintomas e sinais vitais. A velocidade média representa o ritmo funcional produzido durante os "
        "6 minutos, enquanto a velocidade normalizada ajusta a velocidade ao comprimento do membro inferior, oferecendo "
        "uma análise biomecânica complementar da exigência locomotora relativa."
    )

    if analysis.get("normalized_velocity") is None:
        return (
            "Houve elevação do Duplo Produto no pós-teste imediato, com tendência de recuperação ao longo dos 6 minutos. "
            "O custo cardiovascular estimado por metro deve ser interpretado junto à distância, sintomas e sinais vitais. "
            "A velocidade média representa o ritmo funcional produzido durante os 6 minutos. A velocidade normalizada não "
            "foi calculada por ausência de comprimento válido do membro inferior."
        )

    return base_text


def format_analysis_value(value: object, suffix: str = "", digits: int = 2) -> str:
    """Formata valores da análise integrada ou retorna mensagem padrão."""

    if value is None:
        return NOT_CALCULATED
    if isinstance(value, int):
        text = f"{value:,}".replace(",", ".")
    elif isinstance(value, float):
        text = f"{value:.{digits}f}".replace(".", ",")
    else:
        text = str(value)
    return f"{text}{suffix}"


def build_dp_recovery_figure(data: PatientData, timeseries_df: pd.DataFrame):
    """Cria o gráfico de linha da recuperação do Duplo Produto."""

    analysis = build_integrated_recovery_analysis(data, timeseries_df)
    valid_points = [(label, value) for label, value in analysis["dp_values"].items() if value is not None]
    if len(valid_points) < 2:
        return None

    labels = [point[0] for point in valid_points]
    values = [point[1] for point in valid_points]
    min_value = min(values)
    max_value = max(values)
    value_range = max(max_value - min_value, max_value * 0.08, 1)
    ax_bottom = max(0, min_value - value_range * 0.18)
    ax_top = max_value + value_range * 0.24

    fig, ax = plt.subplots(figsize=(8.8, 3.4))
    ax.plot(labels, values, marker="o", color=REPORT_COLORS["primary"], linewidth=2.6)
    for label, value in zip(labels, values):
        is_peak = value == max_value
        offset_y = -15 if is_peak else 8
        vertical_align = "top" if is_peak else "bottom"
        ax.annotate(
            f"{value:,}".replace(",", "."),
            (label, value),
            textcoords="offset points",
            xytext=(0, offset_y),
            ha="center",
            va=vertical_align,
            fontsize=8,
            color=REPORT_COLORS["text"],
            bbox={"boxstyle": "round,pad=0.2", "fc": "white", "ec": "none", "alpha": 0.82},
            clip_on=True,
        )
    ax.set_ylim(ax_bottom, ax_top)
    ax.set_ylabel("Duplo Produto (bpm.mmHg)", fontsize=9)
    ax.tick_params(axis="both", labelsize=9)
    ax.grid(True, alpha=0.25)
    fig.tight_layout()
    return fig


def classificar_risco(distancia_real: float, interrompeu: bool) -> tuple[str, str]:
    """Classifica o risco por distância absoluta e interrupção do teste."""

    if interrompeu:
        return "Teste interrompido", "Elevadíssimo risco de morbimortalidade"
    if distancia_real < 300:
        return "Nível 1", "Muito elevado risco de morbimortalidade"
    if 300 <= distancia_real <= 375:
        return "Nível 2", "Elevado risco de morbimortalidade"
    if 375 < distancia_real <= 450:
        return "Nível 3", "Moderado risco de morbimortalidade"
    return "Nível 4", "Baixo risco de mortalidade"


def obter_fator_limitante(borg_resp_pico: float, borg_mmii_pico: float) -> str:
    """Compara Borg respiratório e Borg MMII para sugerir o fator limitante."""

    diferenca = abs(borg_resp_pico - borg_mmii_pico)
    ambos_elevados = borg_resp_pico >= 5 and borg_mmii_pico >= 5

    if ambos_elevados and diferenca <= 1:
        return "Limitação mista"
    if borg_resp_pico > borg_mmii_pico:
        return "Limitação cardiorrespiratória"
    if borg_mmii_pico > borg_resp_pico:
        return "Limitação periférica/muscular"
    return "Sem predominância clara"


def build_default_pre_table() -> pd.DataFrame:
    """Cria a tabela de sinais vitais de repouso, antes do teste."""

    return pd.DataFrame(
        [
            {
                "Tempo": "Antes do teste",
                "FC": 0,
                "SpO2": 0,
                "FR": 0,
                "PAS": 0,
                "PAD": 0,
                "Borg Respiratório": 0.0,
                "Borg MMII": 0.0,
            }
        ]
    )


def build_default_during_table() -> pd.DataFrame:
    """Cria a tabela do período de caminhada: apenas o que é viável medir durante."""

    return pd.DataFrame(
        {
            "Tempo": ["1 min", "2 min", "3 min", "4 min", "5 min", "6 min"],
            "FC": [0] * 6,
            "SpO2": [0] * 6,
            "Borg Respiratório": [0.0] * 6,
            "Borg MMII": [0.0] * 6,
        }
    )


def build_default_recovery_table() -> pd.DataFrame:
    """Cria a tabela de recuperação com sinais vitais completos."""

    return pd.DataFrame(
        {
            "Tempo": ["Recuperação 1 min", "Recuperação 3 min", "Recuperação 6 min"],
            "FC": [0] * 3,
            "SpO2": [0] * 3,
            "FR": [0] * 3,
            "PAS": [0] * 3,
            "PAD": [0] * 3,
            "Borg Respiratório": [0.0] * 3,
            "Borg MMII": [0.0] * 3,
        }
    )


def build_default_timeseries() -> pd.DataFrame:
    """Mantém compatibilidade: retorna a série completa com as três fases."""

    return combine_timeseries(
        build_default_pre_table(),
        build_default_during_table(),
        build_default_recovery_table(),
    )


def normalize_timeseries(df: pd.DataFrame) -> pd.DataFrame:
    """Padroniza colunas e converte os campos numéricos da série temporal."""

    clean = df.copy()

    if "Tempo" not in clean.columns:
        clean.insert(0, "Tempo", [f"Registro {i + 1}" for i in range(len(clean))])

    for column in FULL_COLUMNS:
        if column not in clean.columns:
            clean[column] = 0

    clean["Tempo"] = clean["Tempo"].astype(str)

    for column in NUMERIC_COLUMNS:
        clean[column] = pd.to_numeric(clean[column], errors="coerce").fillna(0)

    for column in ["FC", "SpO2", "FR", "PAS", "PAD"]:
        clean[column] = clean[column].round().astype(int)

    return clean[FULL_COLUMNS]


def combine_timeseries(
    pre_df: pd.DataFrame,
    during_df: pd.DataFrame,
    recovery_df: pd.DataFrame,
) -> pd.DataFrame:
    """Une as tabelas de repouso, durante e recuperação em uma série única."""

    return pd.concat(
        [
            normalize_timeseries(pre_df),
            normalize_timeseries(during_df),
            normalize_timeseries(recovery_df),
        ],
        ignore_index=True,
    )


def _snapshot_from_row(row: pd.Series) -> VitalSnapshot:
    """Transforma uma linha da tabela em objeto clínico de fase."""

    return VitalSnapshot(
        tempo=str(row["Tempo"]),
        fc=int(row["FC"]),
        spo2=int(row["SpO2"]),
        fr=int(row["FR"]),
        pas=int(row["PAS"]),
        pad=int(row["PAD"]),
        borg_resp=float(row["Borg Respiratório"]),
        borg_mmii=float(row["Borg MMII"]),
    )


def get_phase_snapshots(timeseries_df: pd.DataFrame) -> tuple[VitalSnapshot, VitalSnapshot, VitalSnapshot]:
    """Extrai repouso, pico do exercício e recuperação da série temporal."""

    clean = normalize_timeseries(timeseries_df)
    repouso = _snapshot_from_row(clean.iloc[0])

    exercise_rows = clean.iloc[1:7].copy()
    if exercise_rows["FC"].max() > 0:
        peak_index = exercise_rows["FC"].idxmax()
    else:
        peak_index = exercise_rows[["Borg Respiratório", "Borg MMII"]].max(axis=1).idxmax()

    pico = _snapshot_from_row(clean.loc[peak_index])
    recuperacao = _snapshot_from_row(clean.iloc[-1])

    return repouso, pico, recuperacao


def calculate_tc6m_professional(data: PatientData, timeseries_df: pd.DataFrame) -> TestResult:
    """Executa o motor clínico completo do TC6M."""

    validate_patient_data(data)
    clean = normalize_timeseries(timeseries_df)
    repouso, pico, recuperacao = get_phase_snapshots(clean)

    dpp_enright, lin_enright = calcular_dpp_enright(data.sexo, data.idade, data.peso, data.altura_cm)
    dpp_iwama = calcular_dpp_iwama(data.sexo, data.idade)
    dpp_ben_saad = calcular_dpp_ben_saad(data.idade, data.peso, data.altura_cm)
    dpp_principal, lin_principal = calcular_dpp_por_formula(
        data.formula_principal,
        data.sexo,
        data.idade,
        data.peso,
        data.altura_cm,
    )

    qualificador, percentual = obter_qualificador_funcional(data.distancia, dpp_principal)
    classificacao, risco = classificar_risco(data.distancia, data.interrompeu)

    dp_repouso = calcular_duplo_produto(repouso.fc, repouso.pas)
    dp_pico = calcular_duplo_produto(pico.fc, pico.pas)
    dp_recuperacao = calcular_duplo_produto(recuperacao.fc, recuperacao.pas)
    fator_limitante = obter_fator_limitante(pico.borg_resp, pico.borg_mmii)

    interpretacao = build_interpretation(
        data=data,
        percentual=percentual,
        qualificador=qualificador,
        classificacao=classificacao,
        risco=risco,
        fator_limitante=fator_limitante,
        pico=pico,
        dpp_principal=dpp_principal,
    )

    return TestResult(
        formula_principal=data.formula_principal,
        dpp_principal=dpp_principal,
        lin_principal=lin_principal,
        dpp_enright=dpp_enright,
        lin_enright=lin_enright,
        dpp_iwama=dpp_iwama,
        dpp_ben_saad=dpp_ben_saad,
        percentual_atingido=percentual,
        qualificador_funcional=qualificador,
        classificacao_risco=classificacao,
        risco=risco,
        dp_repouso=dp_repouso,
        dp_pico=dp_pico,
        dp_recuperacao=dp_recuperacao,
        fator_limitante=fator_limitante,
        interpretacao=interpretacao,
    )


def build_interpretation(
    data: PatientData,
    percentual: float,
    qualificador: str,
    classificacao: str,
    risco: str,
    fator_limitante: str,
    pico: VitalSnapshot,
    dpp_principal: float,
) -> str:
    """Gera o texto interpretativo automático do relatório final."""

    interrupcao = " Houve interrupção do teste." if data.interrompeu else " Não houve interrupção registrada."
    motivo = f" Motivo: {data.motivo_interrupcao}." if data.motivo_interrupcao.strip() else ""
    protocolo = f" {descrever_protocolo_corredor(data.comprimento_corredor_m)}"

    return (
        f"O paciente percorreu {data.distancia:.2f} m no TC6M. Pela fórmula selecionada "
        f"({data.formula_principal}), a distância predita principal foi de {dpp_principal:.2f} m, "
        f"correspondendo a {percentual:.2f}% do previsto. Qualificador funcional: {qualificador}. "
        f"Classificação por distância: {classificacao}. Risco associado: {risco}. "
        f"No pico registrado durante a caminhada, observou-se FC={pico.fc} bpm, SpO2={pico.spo2}% "
        f"e Borg respiratório/MMII={pico.borg_resp:.1f}/{pico.borg_mmii:.1f}, sugerindo "
        f"{fator_limitante.lower()}.{interrupcao}{motivo}{protocolo}"
    )


def build_patient_dataframe(data: PatientData) -> pd.DataFrame:
    """Organiza identificação, antropometria e triagem para tela, Excel e PDF."""

    return pd.DataFrame(
        {
            "Campo": [
                "Nome",
                "Prontuário/ID",
                "Data da avaliação",
                "Avaliador",
                "Diagnóstico/condição clínica",
                "Sexo biológico",
                "Idade",
                "Peso",
                "Altura",
                "Comprimento do membro inferior",
                "Protocolo do corredor",
                "Contraindicação absoluta",
                "Contraindicação relativa",
                "Observação da triagem",
            ],
            "Valor": [
                format_patient_name(data.nome),
                data.prontuario or "-",
                data.data_avaliacao.strftime("%d/%m/%Y") if data.data_avaliacao else "-",
                data.avaliador or "-",
                data.diagnostico or "-",
                "Masculino" if data.sexo == "M" else "Feminino",
                f"{data.idade} anos",
                f"{data.peso:.1f} kg",
                f"{data.altura_cm:.1f} cm",
                f"{data.comprimento_membro_inferior_m:.2f} m" if data.comprimento_membro_inferior_m else "-",
                descrever_protocolo_corredor(data.comprimento_corredor_m),
                "Sim" if data.contraindicacao_absoluta else "Não",
                "Sim" if data.contraindicacao_relativa else "Não",
                data.observacao_triagem or "-",
            ],
        }
    )


def build_summary_dataframe(data: PatientData, result: TestResult) -> pd.DataFrame:
    """Monta resultados completos em linguagem clínica e por blocos."""

    lin_principal = f"{result.lin_principal:.2f} m" if result.lin_principal is not None else "Não definido para esta fórmula"
    motivo = data.motivo_interrupcao or "-"
    distancia_interrupcao = f"{data.distancia_interrupcao:.2f} m" if data.interrompeu and data.distancia_interrupcao > 0 else "-"

    return pd.DataFrame(
        {
            "Bloco": [
                "Predição",
                "Predição",
                "Predição",
                "Predição",
                "Predição",
                "Predição",
                "Resultado do teste",
                "Resultado do teste",
                "Resultado do teste",
                "Resultado do teste",
                "Resultado do teste",
                "Resultado do teste",
                "Hemodinâmica",
                "Hemodinâmica",
                "Hemodinâmica",
                "Interpretação",
            ],
            "Indicador": [
                "Fórmula principal escolhida",
                "DPP principal",
                "Limite inferior da fórmula principal",
                "DPP Enright/Sherrill",
                "DPP Iwama et al.",
                "DPP Ben Saad et al.",
                "Distância percorrida",
                "Protocolo do corredor",
                "% atingido da DPP principal",
                "Qualificador funcional",
                "Classificação por distância",
                "Interrupção do teste",
                "Duplo produto em repouso",
                "Duplo produto no pico",
                "Duplo produto na recuperação",
                "Fator limitante provável",
            ],
            "Resultado": [
                result.formula_principal,
                f"{result.dpp_principal:.2f} m",
                lin_principal,
                f"{result.dpp_enright:.2f} m | LIN {result.lin_enright:.2f} m",
                f"{result.dpp_iwama:.2f} m",
                f"{result.dpp_ben_saad:.2f} m",
                f"{data.distancia:.2f} m",
                descrever_protocolo_corredor(data.comprimento_corredor_m),
                f"{result.percentual_atingido:.2f} %",
                result.qualificador_funcional,
                f"{result.classificacao_risco} - {result.risco}",
                f"{'Sim' if data.interrompeu else 'Não'} | Distância: {distancia_interrupcao} | Motivo: {motivo}",
                f"{result.dp_repouso} bpm.mmHg" if result.dp_repouso else "Não calculado: falta PAS/FC de repouso",
                f"{result.dp_pico} bpm.mmHg" if result.dp_pico else "Não calculado: PAS de pico não foi registrada",
                f"{result.dp_recuperacao} bpm.mmHg" if result.dp_recuperacao else "Não calculado: falta PAS/FC de recuperação",
                result.fator_limitante,
            ],
        }
    )


def build_oscillation_figure(timeseries_df: pd.DataFrame):
    """Cria gráfico grande de oscilação de FC e SpO2."""

    clean = normalize_timeseries(timeseries_df)
    fig, ax1 = plt.subplots(figsize=(14, 5.8))
    ax1.plot(clean["Tempo"], clean["FC"], marker="o", color=REPORT_COLORS["fc"], linewidth=3, label="FC")
    ax1.set_ylabel("FC (bpm)", color=REPORT_COLORS["fc"], fontsize=12)
    ax1.tick_params(axis="y", labelcolor=REPORT_COLORS["fc"])
    ax1.tick_params(axis="x", rotation=28)
    ax1.grid(True, alpha=0.25)

    ax2 = ax1.twinx()
    ax2.plot(
        clean["Tempo"],
        clean["SpO2"],
        marker="s",
        color=REPORT_COLORS["spo2"],
        linewidth=3,
        linestyle="--",
        label="SpO2",
    )
    ax2.set_ylabel("SpO2 (%)", color=REPORT_COLORS["spo2"], fontsize=12)
    ax2.tick_params(axis="y", labelcolor=REPORT_COLORS["spo2"])

    fig.suptitle("Oscilação cardiorrespiratória durante o TC6M", fontsize=15, fontweight="bold")
    fig.tight_layout()
    return fig


def build_effort_figure(timeseries_df: pd.DataFrame):
    """Cria gráfico grande da curva de esforço percebido."""

    clean = normalize_timeseries(timeseries_df)
    fig, ax = plt.subplots(figsize=(14, 5.8))
    ax.plot(
        clean["Tempo"],
        clean["Borg Respiratório"],
        marker="o",
        color=REPORT_COLORS["spo2"],
        linewidth=3,
        label="Borg respiratório",
    )
    ax.plot(
        clean["Tempo"],
        clean["Borg MMII"],
        marker="s",
        color=REPORT_COLORS["progress"],
        linewidth=3,
        linestyle="--",
        label="Borg MMII",
    )
    ax.set_ylim(0, 10)
    ax.set_ylabel("Escala de Borg", fontsize=12)
    ax.tick_params(axis="x", rotation=28)
    ax.grid(True, alpha=0.25)
    ax.legend(loc="upper left")
    fig.suptitle("Curva de esforço percebido", fontsize=15, fontweight="bold")
    fig.tight_layout()
    return fig


def build_curve_findings(timeseries_df: pd.DataFrame) -> list[str]:
    """Gera achados automáticos simples a partir das curvas de FC, SpO2 e Borg."""

    clean = normalize_timeseries(timeseries_df)
    exercise = clean.iloc[1:7].copy()

    if exercise.empty:
        return ["Não há registros suficientes durante o teste para interpretar as curvas."]

    repouso = clean.iloc[0]
    pico_fc = exercise.loc[exercise["FC"].idxmax()]
    fc_delta = int(pico_fc["FC"] - repouso["FC"])
    spo2_delta = int(repouso["SpO2"] - exercise["SpO2"].min())

    borg_resp_final = float(exercise["Borg Respiratório"].iloc[-1])
    borg_mmii_final = float(exercise["Borg MMII"].iloc[-1])
    diferenca_media_borg = (exercise["Borg Respiratório"] - exercise["Borg MMII"]).abs().mean()
    borg_resp_pico = float(exercise["Borg Respiratório"].max())
    borg_mmii_pico = float(exercise["Borg MMII"].max())

    achados = []

    if fc_delta >= 40:
        achados.append(f"A FC apresentou elevação importante durante o teste, com aumento de {fc_delta} bpm em relação ao repouso.")
    elif fc_delta >= 20:
        achados.append(f"A FC apresentou elevação progressiva moderada, com aumento de {fc_delta} bpm em relação ao repouso.")
    else:
        achados.append(f"A FC apresentou baixa variação durante o teste, com aumento de {fc_delta} bpm em relação ao repouso.")

    if spo2_delta >= 4:
        achados.append(f"Houve queda relevante da SpO2 durante a caminhada ({spo2_delta} pontos percentuais), achado compatível com dessaturação ao esforço.")
    elif spo2_delta >= 1:
        achados.append(f"Houve pequena oscilação da SpO2 durante o esforço ({spo2_delta} ponto(s) percentual(is)).")
    else:
        achados.append("A SpO2 permaneceu estável durante o teste, sem queda relevante registrada.")

    if diferenca_media_borg <= 1:
        achados.append("As curvas de Borg respiratório e Borg MMII caminharam próximas, sugerindo percepção de esforço global/mista.")
    elif borg_resp_pico > borg_mmii_pico:
        achados.append("A curva de Borg respiratório predominou sobre Borg MMII, sugerindo maior limitação ventilatória/cardiorrespiratória.")
    else:
        achados.append("A curva de Borg MMII predominou sobre Borg respiratório, sugerindo maior limitação periférica/muscular.")

    achados.append(
        f"No 6º minuto, Borg respiratório foi {borg_resp_final:.1f} e Borg MMII foi {borg_mmii_final:.1f}."
    )

    return achados


def format_decimal_br(value: float | int, decimals: int = 2) -> str:
    """Formata números no padrão brasileiro para relatório e interface."""

    return f"{float(value):.{decimals}f}".replace(".", ",")


def format_int_br(value: float | int) -> str:
    """Formata inteiro com separador visual simples para métricas grandes."""

    return f"{int(round(float(value))):,}".replace(",", ".")


def get_risk_display(result: TestResult) -> dict[str, str]:
    """Traduz a classificação em rótulo, cor e posição na escala visual."""

    classificacao = result.classificacao_risco
    if "4" in classificacao:
        return {"label": "Baixo", "detail": "Nível 4 - baixo risco", "color": REPORT_COLORS["risk_low"], "index": "4"}
    if "3" in classificacao:
        return {"label": "Moderado", "detail": "Nível 3 - moderado risco", "color": REPORT_COLORS["risk_moderate"], "index": "3"}
    if "2" in classificacao:
        return {"label": "Alto", "detail": "Nível 2 - elevado risco", "color": REPORT_COLORS["risk_high"], "index": "2"}
    return {
        "label": "Muito alto",
        "detail": f"{classificacao} - risco muito elevado",
        "color": REPORT_COLORS["risk_very_high"],
        "index": "1",
    }


def build_attention_points(data: PatientData, result: TestResult, timeseries_df: pd.DataFrame) -> list[dict[str, str]]:
    """Monta pontos críticos do relatório com badges de atenção, monitoramento ou OK."""

    clean = normalize_timeseries(timeseries_df)
    exercise = clean.iloc[1:7].copy()
    repouso, _, _ = get_phase_snapshots(clean)

    valid_spo2 = exercise.loc[exercise["SpO2"] > 0, "SpO2"]
    min_spo2 = int(valid_spo2.min()) if not valid_spo2.empty else 0
    baseline_spo2 = int(repouso.spo2)
    spo2_drop = baseline_spo2 - min_spo2 if baseline_spo2 > 0 and min_spo2 > 0 else 0

    if min_spo2 and (min_spo2 < 94 or spo2_drop >= 4):
        spo2_badge, spo2_type = "Atenção", "warning"
    else:
        spo2_badge, spo2_type = "OK", "ok"

    if result.dp_repouso and result.dp_recuperacao and result.dp_recuperacao > result.dp_repouso:
        dp_badge, dp_type = "Monitorar", "warning"
    elif result.dp_repouso and result.dp_recuperacao:
        dp_badge, dp_type = "OK", "ok"
    else:
        dp_badge, dp_type = "Incompleto", "warning"

    interrupcao_badge, interrupcao_type = ("Ocorreu", "danger") if data.interrompeu else ("Não ocorreu", "ok")

    if data.contraindicacao_absoluta:
        contra_badge, contra_type = "Absoluta", "danger"
    elif data.contraindicacao_relativa:
        contra_badge, contra_type = "Relativa", "warning"
    else:
        contra_badge, contra_type = "Nenhuma", "ok"

    return [
        {"label": "SpO2 < 94% ou queda >= 4%", "badge": spo2_badge, "type": spo2_type},
        {"label": "Duplo produto maior na recuperação", "badge": dp_badge, "type": dp_type},
        {"label": "Interrupção do teste", "badge": interrupcao_badge, "type": interrupcao_type},
        {"label": "Contraindicações", "badge": contra_badge, "type": contra_type},
    ]


def build_factor_limit_description(result: TestResult, timeseries_df: pd.DataFrame) -> str:
    """Gera explicação curta para o fator limitante provável."""

    _, pico, _ = get_phase_snapshots(timeseries_df)
    fator = result.fator_limitante.lower()

    if "mista" in fator:
        return (
            f"Convergência entre esforço respiratório e periférico: FC no pico {pico.fc} bpm, "
            f"SpO2 {pico.spo2}% e Borg respiratório/MMII {pico.borg_resp:.1f}/{pico.borg_mmii:.1f}. "
            "As duas vias parecem contribuir para a limitação funcional."
        )
    if "cardiorrespiratória" in fator:
        return (
            f"Predomínio cardiorrespiratório: Borg respiratório {pico.borg_resp:.1f} foi maior que "
            f"Borg MMII {pico.borg_mmii:.1f}, com FC no pico {pico.fc} bpm e SpO2 {pico.spo2}%."
        )
    if "periférica" in fator:
        return (
            f"Predomínio periférico/muscular: Borg MMII {pico.borg_mmii:.1f} foi maior que "
            f"Borg respiratório {pico.borg_resp:.1f}, sugerindo maior participação de fadiga de membros inferiores."
        )
    return (
        f"Sem predominância clara entre Borg respiratório e Borg MMII no pico "
        f"({pico.borg_resp:.1f}/{pico.borg_mmii:.1f}). Interpretar junto ao quadro clínico."
    )


def build_prediction_note(result: TestResult) -> str:
    """Sinaliza quando uma fórmula comparativa diverge muito da DPP principal."""

    if result.dpp_principal <= 0:
        return ""

    diff_ben_saad = abs(result.dpp_ben_saad - result.dpp_principal) / result.dpp_principal
    if diff_ben_saad >= 0.25:
        return "Ben Saad diverge fortemente da fórmula principal; use apenas no contexto de crianças/adolescentes."
    return "Predições comparativas sem divergência crítica em relação à fórmula principal."


def build_clinical_summary(data: PatientData, result: TestResult, timeseries_df: pd.DataFrame) -> str:
    """Cria resumo clínico objetivo para tela, PDF e prontuário."""

    clean = normalize_timeseries(timeseries_df)
    exercise = clean.iloc[1:7].copy()
    _, pico, _ = get_phase_snapshots(clean)
    sexo_texto = "masculino" if data.sexo == "M" else "feminino"
    valid_spo2 = exercise.loc[exercise["SpO2"] > 0, "SpO2"]
    min_spo2 = int(valid_spo2.min()) if not valid_spo2.empty else pico.spo2
    protocolo = f" {descrever_protocolo_corredor(data.comprimento_corredor_m)}"

    dp_texto = ""
    if result.dp_repouso and result.dp_recuperacao:
        if result.dp_recuperacao > result.dp_repouso:
            dp_texto = (
                " O duplo produto permaneceu maior na recuperação do que no repouso, "
                "sugerindo necessidade de acompanhamento hemodinâmico no pós-teste."
            )
        else:
            dp_texto = " O duplo produto reduziu na recuperação em relação ao repouso."

    interrupcao = ""
    if data.interrompeu:
        interrupcao = " O teste foi interrompido"
        if data.distancia_interrupcao > 0:
            interrupcao += f" aos {data.distancia_interrupcao:.2f} m"
        if data.motivo_interrupcao.strip():
            interrupcao += f" por {data.motivo_interrupcao.strip()}"
        interrupcao += "."

    return (
        f"Resumo clínico: paciente {sexo_texto} de {data.idade} anos percorreu "
        f"{data.distancia:.2f} m no TC6M, atingindo {result.percentual_atingido:.2f}% do previsto "
        f"pela fórmula {result.formula_principal}. O resultado classifica-se como "
        f"{result.qualificador_funcional} e {result.classificacao_risco} - {result.risco.lower()}. "
        f"Durante o esforço, a menor SpO2 registrada foi {min_spo2}% e o pico de Borg respiratório/MMII "
        f"foi {pico.borg_resp:.1f}/{pico.borg_mmii:.1f}, sugerindo {result.fator_limitante.lower()}."
        f"{dp_texto}{interrupcao}{protocolo}"
    )


def build_report_payload(data: PatientData, result: TestResult, timeseries_df: pd.DataFrame) -> dict:
    """Centraliza os dados visuais do relatório profissional."""

    repouso, pico, recuperacao = get_phase_snapshots(timeseries_df)
    progress_width = min(max(result.percentual_atingido, 0), 100)
    lin_label = f"{result.lin_principal:.2f} m" if result.lin_principal is not None else "Não definido"

    return {
        "risk": get_risk_display(result),
        "progress_width": progress_width,
        "lin_label": lin_label,
        "factor_description": build_factor_limit_description(result, timeseries_df),
        "attention_points": build_attention_points(data, result, timeseries_df),
        "prediction_note": build_prediction_note(result),
        "clinical_summary": build_clinical_summary(data, result, timeseries_df),
        "metrics": [
            {"label": "FC no pico", "value": str(pico.fc), "unit": "bpm"},
            {"label": "SpO2 no pico", "value": str(pico.spo2), "unit": "%"},
            {"label": "Borg resp. / MMII", "value": f"{pico.borg_resp:.1f} / {pico.borg_mmii:.1f}", "unit": "Escala de Borg"},
            {"label": "DP repouso", "value": format_int_br(result.dp_repouso) if result.dp_repouso else "-", "unit": "bpm.mmHg"},
            {"label": "DP recuperação", "value": format_int_br(result.dp_recuperacao) if result.dp_recuperacao else "-", "unit": "bpm.mmHg"},
        ],
        "phase": {
            "repouso": repouso,
            "pico": pico,
            "recuperacao": recuperacao,
        },
    }


def format_patient_name(patient_name: str) -> str:
    """Remove prefixos de ambiente de teste quando aparecem no nome final."""

    clean_name = (patient_name or "").strip()
    test_prefixes = [
        "Paciente Teste -",
        "Paciente Teste –",
        "Paciente Teste:",
        "Paciente teste -",
        "Paciente teste –",
        "Paciente teste:",
    ]

    for prefix in test_prefixes:
        if clean_name.startswith(prefix):
            clean_name = clean_name[len(prefix):].strip()
            break

    return clean_name or "Paciente sem identificação"


def _figure_to_png_bytes(fig) -> BytesIO:
    """Converte um gráfico Matplotlib em PNG para inserir no PDF."""

    buffer = BytesIO()
    fig.savefig(buffer, format="png", dpi=180, bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer


def build_excel_bytes(data: PatientData, result: TestResult, timeseries_df: pd.DataFrame) -> bytes:
    """Gera Excel estruturado com identificação, resumo, sinais e interpretação."""

    output = BytesIO()
    clean = normalize_timeseries(timeseries_df)
    patient_df = build_patient_dataframe(data)
    summary_df = build_summary_dataframe(data, result)
    interpretation_df = pd.DataFrame({"Interpretação automatizada": [result.interpretacao]})
    integrated = build_integrated_recovery_analysis(data, timeseries_df)
    integrated_df = pd.DataFrame(
        {
            "Indicador": [
                "DP repouso",
                "DP 1 min",
                "DP 3 min",
                "DP 6 min",
                "Delta DP repouso -> 1 min",
                "Recuperação DP 1 -> 3 min",
                "Recuperação DP 1 -> 6 min",
                "% recuperação DP em 6 min",
                "Retorno ao basal",
                "Custo DP/m",
                "Velocidade média",
                "Ritmo médio",
                "Comprimento do membro inferior",
                "Velocidade normalizada",
            ],
            "Resultado": [
                format_analysis_value(integrated["dp_repouso"], " bpm.mmHg", 0),
                format_analysis_value(integrated["dp_1"], " bpm.mmHg", 0),
                format_analysis_value(integrated["dp_3"], " bpm.mmHg", 0),
                format_analysis_value(integrated["dp_6"], " bpm.mmHg", 0),
                format_analysis_value(integrated["delta_dp_1"], " bpm.mmHg", 0),
                format_analysis_value(integrated["recovery_dp_1_3"], " bpm.mmHg", 0),
                format_analysis_value(integrated["recovery_dp_1_6"], " bpm.mmHg", 0),
                format_analysis_value(integrated["recovery_percent_6"], " %"),
                format_analysis_value(integrated["return_to_baseline"], " bpm.mmHg", 0),
                format_analysis_value(integrated["cost_dp_per_m"], " DP/m"),
                format_analysis_value(integrated["velocity_ms"], " m/s"),
                format_analysis_value(integrated["pace_m_min"], " m/min"),
                format_analysis_value(integrated["limb_length_m"], " m"),
                format_analysis_value(integrated["normalized_velocity"], "", 3),
            ],
        }
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        patient_df.to_excel(writer, sheet_name="Paciente", index=False)
        summary_df.to_excel(writer, sheet_name="Resumo TC6M", index=False)
        clean.to_excel(writer, sheet_name="Sinais vitais", index=False)
        integrated_df.to_excel(writer, sheet_name="Recuperacao DP", index=False)
        interpretation_df.to_excel(writer, sheet_name="Interpretação", index=False)

        workbook = writer.book
        for sheet in workbook.worksheets:
            sheet.freeze_panes = "A2"
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="000000")
                cell.fill = PatternFill(fill_type="solid", fgColor="DDEFE8")
            for column_cells in sheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = min(max_length + 3, 55)

    output.seek(0)
    return output.getvalue()


def _draw_wrapped_text(
    pdf: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    width: float,
    font_size: int = 9,
    color: str | None = None,
) -> float:
    """Desenha texto no PDF com quebra automática de linha."""

    pdf.setFont("Helvetica", font_size)
    pdf.setFillColor(_hex(color or REPORT_COLORS["text"]))
    lines = simpleSplit(text, "Helvetica", font_size, width)
    for line in lines:
        pdf.drawString(x, y, line)
        y -= font_size + 3
    return y


def _clip_text_to_width(pdf: canvas.Canvas, text: str, font_name: str, font_size: float, width: float) -> str:
    """Encurta texto para caber em uma largura fixa no PDF."""

    text = str(text)
    if pdf.stringWidth(text, font_name, font_size) <= width:
        return text

    ellipsis = "..."
    while text and pdf.stringWidth(text + ellipsis, font_name, font_size) > width:
        text = text[:-1]
    return f"{text.rstrip()}{ellipsis}" if text else ellipsis


def _draw_wrapped_text_limited(
    pdf: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    width: float,
    font_size: float = 8,
    color: str | None = None,
    max_lines: int = 4,
) -> float:
    """Desenha texto com limite de linhas para evitar estouro de card."""

    pdf.setFont("Helvetica", font_size)
    pdf.setFillColor(_hex(color or REPORT_COLORS["text"]))
    lines = simpleSplit(str(text), "Helvetica", font_size, width)
    clipped = lines[:max_lines]
    if len(lines) > max_lines and clipped:
        clipped[-1] = _clip_text_to_width(pdf, f"{clipped[-1]}...", "Helvetica", font_size, width)

    for line in clipped:
        pdf.drawString(x, y, line)
        y -= font_size + 3
    return y


def _draw_bullet_list(
    pdf: canvas.Canvas,
    items: Iterable[str],
    x: float,
    y: float,
    width: float,
    font_size: float = 7.4,
    bullet_color: str | None = None,
    max_items: int | None = None,
) -> float:
    """Desenha uma lista clínica compacta com bullets circulares."""

    bullet_fill = _hex(bullet_color or REPORT_COLORS["primary"])
    for index, item in enumerate(items):
        if max_items is not None and index >= max_items:
            break
        pdf.setFillColor(bullet_fill)
        pdf.circle(x + 3, y - 3, 1.6, fill=True, stroke=False)
        y = _draw_wrapped_text(
            pdf,
            str(item),
            x + 12,
            y,
            width - 12,
            font_size,
            REPORT_COLORS["text"],
        )
        y -= 2
    return y


def _draw_table(pdf: canvas.Canvas, rows: Iterable[tuple[str, str]], x: float, y: float, width: float) -> float:
    """Desenha uma tabela simples de duas colunas no PDF clínico."""

    rows = list(rows)
    col1_width = width * 0.50
    row_height = 17
    top_y = y

    pdf.setFont("Helvetica-Bold", 8)
    pdf.setFillColor(colors.HexColor("#DDEFE8"))
    pdf.rect(x, y - row_height + 4, width, row_height, fill=True, stroke=False)
    pdf.setFillColor(colors.black)
    pdf.drawString(x + 5, y - 8, "Campo")
    pdf.drawString(x + col1_width + 5, y - 8, "Resultado")
    y -= row_height

    pdf.setFont("Helvetica", 8)
    for label, value in rows:
        pdf.setStrokeColor(colors.HexColor("#D8DEE6"))
        pdf.line(x, y + 4, x + width, y + 4)
        pdf.drawString(x + 5, y - 8, str(label)[:64])
        pdf.drawString(x + col1_width + 5, y - 8, str(value)[:62])
        y -= row_height

    pdf.line(x + col1_width, top_y + 4, x + col1_width, y + row_height + 4)
    return y


def _hex(hex_color: str):
    """Converte cor hexadecimal em objeto do ReportLab."""

    return colors.HexColor(hex_color)


def _draw_section_label(pdf: canvas.Canvas, text: str, x: float, y: float) -> None:
    """Desenha rótulo pequeno de seção no PDF."""

    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.setFont("Helvetica", 8)
    pdf.drawString(x, y, text.upper())


def _draw_card_box(pdf: canvas.Canvas, x: float, y: float, width: float, height: float, fill: str = "#FFFFFF") -> None:
    """Desenha um card limpo com borda fina."""

    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.setLineWidth(0.6)
    pdf.setFillColor(_hex(fill))
    pdf.roundRect(x, y - height, width, height, 9, fill=True, stroke=True)


def _draw_badge(pdf: canvas.Canvas, text: str, x: float, y: float, badge_type: str = "warning") -> float:
    """Desenha badge visual de status no PDF e retorna a largura usada."""

    palette = {
        "ok": (REPORT_COLORS["ok_bg"], REPORT_COLORS["ok_text"]),
        "warning": (REPORT_COLORS["warning_bg"], REPORT_COLORS["warning_text"]),
        "danger": (REPORT_COLORS["danger_bg"], REPORT_COLORS["danger_text"]),
    }
    bg, fg = palette.get(badge_type, palette["warning"])
    pdf.setFont("Helvetica", 8)
    badge_width = pdf.stringWidth(text, "Helvetica", 8) + 16
    pdf.setFillColor(_hex(bg))
    pdf.roundRect(x, y - 12, badge_width, 15, 7, fill=True, stroke=False)
    pdf.setFillColor(_hex(fg))
    pdf.drawString(x + 8, y - 8, text)
    return badge_width


def _draw_key_value(pdf: canvas.Canvas, key: str, value: str, x: float, y: float, width: float) -> float:
    """Desenha uma linha chave-valor dentro de card."""

    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.setLineWidth(0.4)
    pdf.line(x, y - 3, x + width, y - 3)
    pdf.setFont("Helvetica", 8)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(x, y + 4, key[:34])
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawRightString(x + width, y + 4, value[:32])
    return y - 16


def _draw_progress_bar(pdf: canvas.Canvas, x: float, y: float, width: float, percent: float) -> None:
    """Desenha barra de progresso da distância atingida contra a DPP."""

    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.setFillColor(_hex("#D8DDD9"))
    pdf.roundRect(x, y - 9, width, 9, 5, fill=True, stroke=False)
    fill_width = width * min(max(percent, 0), 100) / 100
    pdf.setFillColor(_hex(REPORT_COLORS["progress"]))
    pdf.roundRect(x, y - 9, fill_width, 9, 5, fill=True, stroke=False)


def _draw_risk_scale(pdf: canvas.Canvas, x: float, y: float, width: float, result: TestResult) -> float:
    """Desenha escala visual de risco no PDF."""

    segments = [
        REPORT_COLORS["risk_low"],
        REPORT_COLORS["risk_moderate"],
        REPORT_COLORS["risk_high"],
        REPORT_COLORS["risk_very_high"],
    ]
    segment_width = (width - 12) / 4
    for index, color in enumerate(segments):
        pdf.setFillColor(_hex(color))
        pdf.roundRect(x + index * (segment_width + 4), y - 2, segment_width, 7, 2, fill=True, stroke=False)

    risk = get_risk_display(result)
    position = {"4": 0, "3": 1, "2": 2}.get(risk["index"], 3)
    indicator_x = x + position * (segment_width + 4) + (segment_width / 2)
    pdf.setStrokeColor(_hex(risk["color"]))
    pdf.setLineWidth(1.2)
    pdf.line(indicator_x, y - 7, indicator_x, y - 19)
    pdf.setFillColor(_hex(risk["color"]))
    pdf.circle(indicator_x, y - 22, 3.5, fill=True, stroke=False)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.setFont("Helvetica", 7.5)
    pdf.drawString(x, y - 32, risk["detail"])
    return y - 54


def _badge_type_for_qualifier(text: str) -> str:
    """Escolhe cor de badge conforme qualificador funcional."""

    lower = text.lower()
    if "nenhum" in lower or "leve" in lower:
        return "ok"
    if "moderado" in lower:
        return "warning"
    return "danger"


def _draw_report_footer(
    pdf: canvas.Canvas,
    page: int,
    total_pages: int,
    margin: float,
    width: float,
    now: str,
    formula: str,
) -> None:
    """Desenha rodapé padronizado no relatório."""

    y = 30
    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.setLineWidth(0.6)
    pdf.line(margin, y + 16, width - margin, y + 16)
    pdf.setFont("Helvetica", 7)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(margin, y, f"Relatório gerado em {now}")
    pdf.drawCentredString(width / 2, y, f"Fórmula principal: {formula}")
    pdf.drawRightString(width - margin, y, f"Página {page} de {total_pages}")
    pdf.drawCentredString(width / 2, y - 10, "Uso restrito à equipe de saúde")


def _draw_pdf_icon(pdf: canvas.Canvas, icon: str, cx: float, cy: float, size: float = 18, color: str | None = None) -> None:
    """Desenha ícones vetoriais simples no PDF sem depender de fonte externa."""

    stroke = _hex(color or REPORT_COLORS["primary"])
    pdf.setStrokeColor(stroke)
    pdf.setFillColor(stroke)
    pdf.setLineWidth(max(size / 11, 1.2))
    half = size / 2

    if icon == "shoe":
        x = cx - half
        y = cy - half * 0.45
        path = pdf.beginPath()
        path.moveTo(x + size * 0.08, y + size * 0.15)
        path.lineTo(x + size * 0.90, y + size * 0.15)
        path.curveTo(x + size * 0.98, y + size * 0.18, x + size * 0.98, y + size * 0.32, x + size * 0.86, y + size * 0.36)
        path.lineTo(x + size * 0.55, y + size * 0.46)
        path.lineTo(x + size * 0.39, y + size * 0.72)
        path.lineTo(x + size * 0.24, y + size * 0.72)
        path.lineTo(x + size * 0.18, y + size * 0.38)
        path.curveTo(x + size * 0.10, y + size * 0.34, x + size * 0.05, y + size * 0.25, x + size * 0.08, y + size * 0.15)
        pdf.drawPath(path, stroke=True, fill=False)
        for lx in (0.38, 0.48, 0.58):
            pdf.line(x + size * lx, y + size * 0.43, x + size * (lx + 0.08), y + size * 0.54)
        pdf.line(x + size * 0.16, y + size * 0.06, x + size * 0.86, y + size * 0.06)
        return

    if icon == "walk":
        pdf.circle(cx, cy + size * 0.30, size * 0.13, fill=False, stroke=True)
        pdf.line(cx, cy + size * 0.16, cx - size * 0.10, cy - size * 0.12)
        pdf.line(cx - size * 0.08, cy + size * 0.04, cx - size * 0.30, cy - size * 0.02)
        pdf.line(cx - size * 0.02, cy - size * 0.10, cx - size * 0.26, cy - size * 0.38)
        pdf.line(cx - size * 0.02, cy - size * 0.10, cx + size * 0.24, cy - size * 0.34)
        pdf.line(cx - size * 0.04, cy + size * 0.08, cx + size * 0.24, cy - size * 0.02)
        return

    if icon == "shield":
        path = pdf.beginPath()
        path.moveTo(cx, cy + half * 0.72)
        path.lineTo(cx + half * 0.62, cy + half * 0.42)
        path.lineTo(cx + half * 0.52, cy - half * 0.18)
        path.curveTo(cx + half * 0.42, cy - half * 0.56, cx + half * 0.16, cy - half * 0.78, cx, cy - half * 0.88)
        path.curveTo(cx - half * 0.16, cy - half * 0.78, cx - half * 0.42, cy - half * 0.56, cx - half * 0.52, cy - half * 0.18)
        path.lineTo(cx - half * 0.62, cy + half * 0.42)
        path.close()
        pdf.drawPath(path, stroke=True, fill=False)
        return

    if icon == "chart":
        x = cx - half * 0.75
        y = cy - half * 0.62
        pdf.line(x, y, x + size * 0.78, y)
        for index, height_ratio in enumerate((0.35, 0.58, 0.85)):
            bx = x + size * (0.12 + index * 0.24)
            pdf.setLineWidth(max(size / 9, 1.4))
            pdf.line(bx, y, bx, y + size * height_ratio * 0.72)
        pdf.setLineWidth(max(size / 14, 1.1))
        pdf.line(x + size * 0.08, y + size * 0.24, x + size * 0.32, y + size * 0.46)
        pdf.line(x + size * 0.32, y + size * 0.46, x + size * 0.54, y + size * 0.44)
        pdf.line(x + size * 0.54, y + size * 0.44, x + size * 0.78, y + size * 0.70)
        return

    if icon == "heart":
        pdf.line(cx - half * 0.75, cy, cx - half * 0.30, cy)
        pdf.line(cx - half * 0.30, cy, cx - half * 0.10, cy + half * 0.42)
        pdf.line(cx - half * 0.10, cy + half * 0.42, cx + half * 0.18, cy - half * 0.50)
        pdf.line(cx + half * 0.18, cy - half * 0.50, cx + half * 0.34, cy)
        pdf.line(cx + half * 0.34, cy, cx + half * 0.76, cy)
        return

    if icon == "alert":
        path = pdf.beginPath()
        path.moveTo(cx, cy + half * 0.75)
        path.lineTo(cx + half * 0.75, cy - half * 0.70)
        path.lineTo(cx - half * 0.75, cy - half * 0.70)
        path.close()
        pdf.drawPath(path, stroke=True, fill=False)
        pdf.line(cx, cy + half * 0.24, cx, cy - half * 0.25)
        pdf.circle(cx, cy - half * 0.48, max(size / 24, 1.0), fill=True, stroke=False)
        return

    if icon == "summary":
        pdf.roundRect(cx - half * 0.55, cy - half * 0.72, size * 0.78, size * 1.18, 2, fill=False, stroke=True)
        pdf.line(cx - half * 0.34, cy + half * 0.12, cx + half * 0.08, cy + half * 0.12)
        pdf.line(cx - half * 0.34, cy - half * 0.14, cx + half * 0.16, cy - half * 0.14)
        pdf.line(cx - half * 0.34, cy - half * 0.40, cx + half * 0.00, cy - half * 0.40)
        return

    if icon == "info":
        pdf.circle(cx, cy, half * 0.72, fill=False, stroke=True)
        pdf.line(cx, cy - half * 0.25, cx, cy + half * 0.22)
        pdf.circle(cx, cy + half * 0.45, max(size / 24, 1.0), fill=True, stroke=False)
        return

    if icon == "search":
        pdf.circle(cx - half * 0.12, cy + half * 0.12, half * 0.45, fill=False, stroke=True)
        pdf.line(cx + half * 0.20, cy - half * 0.20, cx + half * 0.62, cy - half * 0.62)
        return

    pdf.circle(cx, cy, max(size / 5, 2.4), fill=True, stroke=False)


def _resolve_icon_key(icon_key: str) -> str:
    """Normaliza aliases antigos para a biblioteca oficial de ícones."""

    return ICON_ALIASES.get(icon_key, icon_key)


def _draw_svg_icon(pdf: canvas.Canvas, icon_path: Path, cx: float, cy: float, size: float) -> bool:
    """Tenta renderizar SVG no ReportLab usando svglib."""

    if svg2rlg is None or renderPDF is None or not icon_path.exists():
        return False

    try:
        drawing = svg2rlg(str(icon_path))
        if drawing is None or not drawing.width or not drawing.height:
            return False

        scale = size / max(float(drawing.width), float(drawing.height))
        draw_width = float(drawing.width) * scale
        draw_height = float(drawing.height) * scale
        pdf.saveState()
        pdf.translate(cx - draw_width / 2, cy - draw_height / 2)
        pdf.scale(scale, scale)
        renderPDF.draw(drawing, pdf, 0, 0)
        pdf.restoreState()
        return True
    except Exception:
        try:
            pdf.restoreState()
        except Exception:
            pass
        return False


def _draw_png_icon(pdf: canvas.Canvas, icon_path: Path, cx: float, cy: float, size: float) -> bool:
    """Renderiza PNG oficial mantendo proporção e transparência."""

    if not icon_path.exists():
        return False

    try:
        pdf.drawImage(
            ImageReader(str(icon_path)),
            cx - size / 2,
            cy - size / 2,
            width=size,
            height=size,
            preserveAspectRatio=True,
            mask="auto",
        )
        return True
    except Exception:
        return False


def _draw_icon_circle(
    pdf: canvas.Canvas,
    icon_key: str,
    cx: float,
    cy: float,
    circle_size: float = 26,
    icon_size: float | None = None,
    bg_color: str = "#E6F4EF",
    stroke_color: str | None = None,
) -> None:
    """Desenha o círculo suave e o ícone oficial centralizado."""

    resolved = _resolve_icon_key(icon_key)
    icon_size = icon_size or circle_size * 0.62

    icon_path = ICON_MAP.get(resolved)
    if icon_path and icon_path.suffix.lower() == ".png" and _draw_png_icon(pdf, icon_path, cx, cy, circle_size):
        return
    if icon_path and icon_path.suffix.lower() == ".svg" and _draw_svg_icon(pdf, icon_path, cx, cy, icon_size):
        return

    pdf.setFillColor(_hex(bg_color))
    pdf.circle(cx, cy, circle_size / 2, fill=True, stroke=False)

    fallback_alias = {
        "caminhada_tc6m": "walk",
        "distancia_percorrida": "shoe",
        "resultado_principal": "shoe",
        "classificacao_risco": "shield",
        "predicoes_comparativas": "chart",
        "metricas_hemodinamicas": "heart",
        "spo2": "info",
        "borg_respiratorio": "info",
        "dp_repouso": "heart",
        "dp_recuperacao": "chart",
        "fator_limitante": "info",
        "pontos_atencao": "alert",
        "resumo_clinico": "summary",
        "graficos_achados": "chart",
        "achados_automaticos": "search",
        "duplo_produto_recuperacao": "chart",
        "velocidade_media": "chart",
        "velocidade_normalizada": "chart",
        "interpretacao_integrada": "info",
        "nota_metodologica": "info",
    }.get(resolved, resolved)
    _draw_pdf_icon(pdf, fallback_alias, cx, cy, icon_size, stroke_color or REPORT_COLORS["primary"])


def _metric_icon_key(label: str) -> str:
    """Escolhe o ícone oficial para cards de métrica do PDF."""

    lower = label.lower()
    if "spo2" in lower:
        return "spo2"
    if "borg" in lower:
        return "borg_respiratorio"
    if "recuper" in lower:
        return "dp_recuperacao"
    if "dp repouso" in lower:
        return "dp_repouso"
    if "dp" in lower:
        return "metricas_hemodinamicas"
    if "velocidade" in lower or "ritmo" in lower:
        return "velocidade_media"
    return "metricas_hemodinamicas"


def _draw_report_header(
    pdf: canvas.Canvas,
    data: PatientData,
    result: TestResult,
    page: int,
    total_pages: int,
    margin: float,
    width: float,
    height: float,
    now: str,
) -> float:
    """Desenha cabeçalho premium em todas as páginas."""

    patient_name = format_patient_name(data.nome)
    y = height - margin

    _draw_icon_circle(pdf, "caminhada_tc6m", margin + 20, y - 10, circle_size=44, icon_size=34)

    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.setFont("Helvetica-Bold", 17)
    pdf.drawString(margin + 55, y + 2, "Teste de Caminhada de 6 Minutos (TC6M)")
    pdf.setFont("Helvetica", 10)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(margin + 55, y - 17, "Nome do paciente: ")
    pdf.setFont("Helvetica-Bold", 10)
    pdf.setFillColor(_hex(REPORT_COLORS["primary"]))
    pdf.drawString(margin + 142, y - 17, patient_name[:34])

    _draw_badge(
        pdf,
        result.qualificador_funcional[:31],
        width - margin - 138,
        y + 1,
        _badge_type_for_qualifier(result.qualificador_funcional),
    )

    y -= 46
    pdf.setFont("Helvetica", 8)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(margin, y, f"{data.prontuario or 'ID não informado'}")
    pdf.setFillColor(_hex(REPORT_COLORS["primary"]))
    pdf.circle(margin + 130, y + 2, 1.4, fill=True, stroke=False)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(margin + 140, y, "Avaliação cardiorrespiratória funcional")

    y -= 21
    meta = [
        f"Data: {data.data_avaliacao.strftime('%d/%m/%Y') if data.data_avaliacao else now[:10]}",
        f"Avaliador: {data.avaliador or '-'}",
        f"Sexo: {'Masculino' if data.sexo == 'M' else 'Feminino'}",
        f"Idade: {data.idade} anos",
        f"Peso: {data.peso:.1f} kg",
        f"Altura: {data.altura_cm:.1f} cm",
    ]
    x = margin
    meta_widths = [84, 138, 70, 68, 70, 72]
    for item, item_width in zip(meta, meta_widths):
        pdf.setFillColor(_hex(REPORT_COLORS["primary"]))
        pdf.circle(x + 2, y + 2.5, 1.2, fill=True, stroke=False)
        pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
        pdf.setFont("Helvetica", 7.4)
        pdf.drawString(x + 10, y, _clip_text_to_width(pdf, item, "Helvetica", 7.4, item_width - 12))
        x += item_width + 6

    y -= 14
    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.line(margin, y, width - margin, y)
    _draw_report_footer(pdf, page, total_pages, margin, width, now, result.formula_principal)
    return y - 20


def _draw_card_title(pdf: canvas.Canvas, x: float, y: float, icon: str, title: str, color: str | None = None) -> None:
    """Desenha título de card com pequeno ícone vetorial."""

    _draw_icon_circle(
        pdf,
        icon,
        x + 8,
        y + 2,
        circle_size=24,
        icon_size=24,
        bg_color="#FFF7ED" if _resolve_icon_key(icon) == "pontos_atencao" else "#E6F4EF",
        stroke_color=color or REPORT_COLORS["primary"],
    )
    pdf.setFillColor(_hex(REPORT_COLORS["primary"]))
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(x + 31, y, title)


def _draw_metric_box(
    pdf: canvas.Canvas,
    x: float,
    y: float,
    width: float,
    height: float,
    label: str,
    value: str,
    unit: str,
    accent: str = "primary",
    badge: str | None = None,
) -> None:
    """Desenha card pequeno de métrica."""

    _draw_card_box(pdf, x, y, width, height, "#F7FCFA")
    _draw_icon_circle(
        pdf,
        _metric_icon_key(label),
        x + 20,
        y - 26,
        circle_size=34,
        icon_size=34,
        stroke_color=REPORT_COLORS.get(accent, REPORT_COLORS["primary"]),
    )
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.setFont("Helvetica", 7.2)
    pdf.drawString(x + 44, y - 14, label[:24])
    pdf.setFont("Helvetica-Bold", 15)
    pdf.drawString(x + 44, y - 35, value[:16])
    pdf.setFont("Helvetica", 7.4)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(x + 44, y - 49, unit[:18])
    if badge:
        _draw_badge(pdf, badge, x + 10, y - height + 18, "ok" if badge.lower() in {"bom", "boa", "adequado", "referência"} else "warning")


def _draw_kv_row(pdf: canvas.Canvas, x: float, y: float, width: float, label: str, value: str) -> float:
    """Linha chave-valor compacta para listas internas."""

    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.setLineWidth(0.35)
    pdf.line(x, y - 4, x + width, y - 4)
    pdf.setFont("Helvetica", 8.5)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(x, y + 4, label[:34])
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.setFont("Helvetica-Bold", 8.5)
    pdf.drawRightString(x + width, y + 4, value[:34])
    return y - 18


def _build_pdf_bytes_legacy_linear(data: PatientData, result: TestResult, timeseries_df: pd.DataFrame) -> bytes:
    """Gera PDF clínico redesenhado com hierarquia, cards, badges e gráficos."""

    payload = build_report_payload(data, result, timeseries_df)
    output = BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    pdf.setTitle("Relatório Clínico TC6M")
    width, height = A4
    margin = 38
    usable = width - 2 * margin
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    patient_name = format_patient_name(data.nome)

    y = height - margin
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.setFont("Helvetica", 15)
    pdf.drawString(margin, y, "Teste de Caminhada de 6 Minutos (TC6M)")
    _draw_badge(pdf, result.qualificador_funcional, width - margin - 130, y + 2, "warning")

    y -= 18
    pdf.setFont("Helvetica", 10)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(margin, y, f"Nome do paciente: {patient_name}")

    y -= 18
    pdf.setFont("Helvetica", 8)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(margin, y, f"{data.prontuario or 'ID não informado'} - Avaliação cardiorrespiratória funcional")

    y -= 18
    meta = [
        f"Data: {data.data_avaliacao.strftime('%d/%m/%Y') if data.data_avaliacao else now[:10]}",
        f"Avaliador: {data.avaliador or '-'}",
        f"Sexo: {'Masculino' if data.sexo == 'M' else 'Feminino'}",
        f"Idade: {data.idade} anos",
        f"Peso: {data.peso:.1f} kg",
        f"Altura: {data.altura_cm:.1f} cm",
    ]
    pdf.drawString(margin, y, "   |   ".join(meta))
    y -= 13
    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.line(margin, y, width - margin, y)

    y -= 22
    _draw_section_label(pdf, "Resultado principal", margin, y)
    y -= 12
    _draw_card_box(pdf, margin, y, usable, 92, REPORT_COLORS["background_deep"])
    card_top = y
    pdf.setFont("Helvetica", 28)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(margin + 18, card_top - 38, f"{data.distancia:.0f} m")
    pdf.setFont("Helvetica", 8)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(margin + 18, card_top - 54, "Distância percorrida no TC6M")

    bar_x = margin + 190
    bar_w = usable - 220
    pdf.setFont("Helvetica", 7.5)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(bar_x, card_top - 23, "0 m")
    pdf.drawRightString(bar_x + bar_w, card_top - 23, f"Predito: {result.dpp_principal:.2f} m")
    _draw_progress_bar(pdf, bar_x, card_top - 40, bar_w, result.percentual_atingido)
    pdf.drawString(bar_x, card_top - 58, f"LIN: {payload['lin_label']}")
    pdf.drawRightString(bar_x + bar_w, card_top - 58, f"DPP: {result.dpp_principal:.2f} m")
    pdf.setFont("Helvetica", 14)
    pdf.setFillColor(_hex(REPORT_COLORS["progress"]))
    pdf.drawString(bar_x, card_top - 78, f"{result.percentual_atingido:.2f}%")
    pdf.setFont("Helvetica", 8)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(bar_x + 62, card_top - 75, "do previsto")

    y -= 114
    card_w = (usable - 14) / 2
    _draw_card_box(pdf, margin, y, card_w, 126)
    _draw_card_box(pdf, margin + card_w + 14, y, card_w, 126)

    pdf.setFillColor(_hex(REPORT_COLORS["risk_very_high"]))
    pdf.circle(margin + 8, y - 16, 3, fill=True, stroke=False)
    pdf.setFont("Helvetica", 10)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(margin + 14, y - 19, "Classificação de risco")
    risk_y = _draw_risk_scale(pdf, margin + 14, y - 40, card_w - 28, result)
    risk_y = _draw_key_value(pdf, "Qualificador", result.qualificador_funcional, margin + 14, risk_y, card_w - 28)
    _draw_key_value(pdf, "Interrupção", "Sim" if data.interrompeu else "Não", margin + 14, risk_y, card_w - 28)

    pred_x = margin + card_w + 28
    pdf.setFillColor(_hex(REPORT_COLORS["progress"]))
    pdf.circle(pred_x - 6, y - 16, 3, fill=True, stroke=False)
    pdf.setFont("Helvetica", 10)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(pred_x, y - 19, "Predições comparativas")
    pred_y = y - 42
    pred_y = _draw_key_value(pdf, "DPP principal", f"{result.dpp_principal:.2f} m", pred_x, pred_y, card_w - 28)
    pred_y = _draw_key_value(pdf, "LIN principal", payload["lin_label"], pred_x, pred_y, card_w - 28)
    pred_y = _draw_key_value(pdf, "Iwama et al.", f"{result.dpp_iwama:.2f} m", pred_x, pred_y, card_w - 28)
    pred_y = _draw_key_value(pdf, "Ben Saad et al.", f"{result.dpp_ben_saad:.2f} m", pred_x, pred_y, card_w - 28)
    pdf.setFont("Helvetica", 7.5)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    _draw_wrapped_text(pdf, payload["prediction_note"], pred_x, pred_y + 5, card_w - 28, 7.5)

    y -= 148
    _draw_section_label(pdf, "Métricas hemodinâmicas", margin, y)
    y -= 12
    metric_w = (usable - 32) / 5
    for index, metric in enumerate(payload["metrics"]):
        x = margin + index * (metric_w + 8)
        _draw_card_box(pdf, x, y, metric_w, 58, REPORT_COLORS["background_deep"])
        pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
        pdf.setFont("Helvetica", 7)
        pdf.drawString(x + 8, y - 14, metric["label"][:18])
        pdf.setFillColor(_hex(REPORT_COLORS["text"]))
        pdf.setFont("Helvetica", 13)
        pdf.drawString(x + 8, y - 32, metric["value"][:12])
        pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
        pdf.setFont("Helvetica", 7)
        pdf.drawString(x + 8, y - 46, metric["unit"])

    y -= 82
    _draw_card_box(pdf, margin, y, card_w, 116)
    _draw_card_box(pdf, margin + card_w + 14, y, card_w, 116)

    pdf.setFillColor(_hex(REPORT_COLORS["spo2"]))
    pdf.circle(margin + 8, y - 16, 3, fill=True, stroke=False)
    pdf.setFont("Helvetica", 10)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(margin + 14, y - 19, "Fator limitante provável")
    pdf.setFont("Helvetica", 15)
    pdf.drawString(margin + 14, y - 43, result.fator_limitante[:28])
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    _draw_wrapped_text(pdf, payload["factor_description"], margin + 14, y - 62, card_w - 28, 8)

    att_x = margin + card_w + 28
    pdf.setFillColor(_hex(REPORT_COLORS["ok_text"]))
    pdf.circle(att_x - 6, y - 16, 3, fill=True, stroke=False)
    pdf.setFont("Helvetica", 10)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(att_x, y - 19, "Pontos de atenção")
    att_y = y - 42
    for point in payload["attention_points"]:
        pdf.setFont("Helvetica", 8)
        pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
        pdf.drawString(att_x, att_y, point["label"][:31])
        _draw_badge(pdf, point["badge"], att_x + card_w - 92, att_y + 7, point["type"])
        att_y -= 18

    y -= 142
    _draw_section_label(pdf, "Resumo clínico", margin, y)
    y -= 12
    _draw_card_box(pdf, margin, y, usable, 104, REPORT_COLORS["background_deep"])
    _draw_wrapped_text(pdf, payload["clinical_summary"], margin + 16, y - 18, usable - 32, 9.4, REPORT_COLORS["text"])

    pdf.setFont("Helvetica", 7.5)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(
        margin,
        32,
        f"Relatório gerado em {now} - Fórmula principal: {result.formula_principal} - Uso restrito à equipe de saúde",
    )

    pdf.showPage()
    y = height - margin
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.setFont("Helvetica", 15)
    pdf.drawString(margin, y, "Gráficos e achados das curvas")
    y -= 22
    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.line(margin, y, width - margin, y)

    y -= 22
    _draw_section_label(pdf, "Achados automáticos", margin, y)
    y -= 14
    _draw_card_box(pdf, margin, y, usable, 84, REPORT_COLORS["background_deep"])
    finding_y = y - 18
    for finding in build_curve_findings(timeseries_df)[:4]:
        pdf.setFont("Helvetica", 8)
        pdf.setFillColor(_hex(REPORT_COLORS["text"]))
        finding_y = _draw_wrapped_text(pdf, f"- {finding}", margin + 14, finding_y, usable - 28, 8)
        finding_y -= 2

    y -= 112
    oscillation_png = _figure_to_png_bytes(build_oscillation_figure(timeseries_df))
    effort_png = _figure_to_png_bytes(build_effort_figure(timeseries_df))

    _draw_card_box(pdf, margin, y, usable, 252)
    pdf.setFont("Helvetica", 10)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(margin + 14, y - 18, "Oscilação cardiorrespiratória - FC e SpO2")
    pdf.drawImage(ImageReader(oscillation_png), margin + 12, y - 244, width=usable - 24, height=210, preserveAspectRatio=True)

    y -= 276
    _draw_card_box(pdf, margin, y, usable, 238)
    pdf.setFont("Helvetica", 10)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(margin + 14, y - 18, "Curva de esforço percebido - Borg respiratório e MMII")
    pdf.drawImage(ImageReader(effort_png), margin + 12, y - 230, width=usable - 24, height=196, preserveAspectRatio=True)

    pdf.showPage()
    y = height - margin
    integrated = build_integrated_recovery_analysis(data, timeseries_df)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.setFont("Helvetica", 15)
    pdf.drawString(margin, y, "Análise Integrada de Recuperação Cardiovascular e Velocidade")
    y -= 22
    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.line(margin, y, width - margin, y)
    y -= 24

    integrated_rows = [
        ("DP repouso", format_analysis_value(integrated["dp_repouso"], " bpm.mmHg", 0)),
        ("DP 1 min", format_analysis_value(integrated["dp_1"], " bpm.mmHg", 0)),
        ("DP 3 min", format_analysis_value(integrated["dp_3"], " bpm.mmHg", 0)),
        ("DP 6 min", format_analysis_value(integrated["dp_6"], " bpm.mmHg", 0)),
        ("Delta DP repouso -> 1 min", format_analysis_value(integrated["delta_dp_1"], " bpm.mmHg", 0)),
        ("Recuperação DP 1 -> 3 min", format_analysis_value(integrated["recovery_dp_1_3"], " bpm.mmHg", 0)),
        ("Recuperação DP 1 -> 6 min", format_analysis_value(integrated["recovery_dp_1_6"], " bpm.mmHg", 0)),
        ("% recuperação DP em 6 min", format_analysis_value(integrated["recovery_percent_6"], " %")),
        ("Retorno ao basal", format_analysis_value(integrated["return_to_baseline"], " bpm.mmHg", 0)),
        ("Custo DP/m", format_analysis_value(integrated["cost_dp_per_m"], " DP/m")),
        ("Velocidade média", format_analysis_value(integrated["velocity_ms"], " m/s")),
        ("Ritmo médio", format_analysis_value(integrated["pace_m_min"], " m/min")),
        ("Comprimento do membro inferior", format_analysis_value(integrated["limb_length_m"], " m")),
        ("Velocidade normalizada", format_analysis_value(integrated["normalized_velocity"], "", 3)),
    ]
    y = _draw_table(pdf, integrated_rows, margin, y, usable)
    y -= 16

    figure = build_dp_recovery_figure(data, timeseries_df)
    if figure is not None and y > 285:
        graph_png = _figure_to_png_bytes(figure)
        _draw_card_box(pdf, margin, y, usable, 230)
        pdf.setFont("Helvetica", 10)
        pdf.setFillColor(_hex(REPORT_COLORS["text"]))
        pdf.drawString(margin + 14, y - 18, "Recuperação do Duplo Produto após o TC6M")
        pdf.drawImage(ImageReader(graph_png), margin + 14, y - 220, width=usable - 28, height=188, preserveAspectRatio=True)
        y -= 252
    elif figure is None:
        _draw_card_box(pdf, margin, y, usable, 42, REPORT_COLORS["background_deep"])
        _draw_wrapped_text(pdf, "Gráfico não exibido: dados insuficientes para plotar ao menos dois pontos válidos de Duplo Produto.", margin + 14, y - 16, usable - 28, 8.5)
        y -= 60

    if y < 150:
        pdf.showPage()
        y = height - margin

    _draw_section_label(pdf, "Interpretação cautelosa", margin, y)
    y -= 14
    _draw_card_box(pdf, margin, y, usable, 126, REPORT_COLORS["background_deep"])
    interp_y = y - 18
    for item in integrated["interpretations"][:4]:
        interp_y = _draw_wrapped_text(pdf, f"- {item}", margin + 14, interp_y, usable - 28, 8.2, REPORT_COLORS["text"])
        interp_y -= 2
    y -= 146
    _draw_section_label(pdf, "Aviso metodológico", margin, y)
    y -= 14
    _draw_card_box(pdf, margin, y, usable, 54, "#EAF3DE")
    _draw_wrapped_text(pdf, integrated["notice"], margin + 14, y - 18, usable - 28, 8.2, REPORT_COLORS["text"])

    pdf.showPage()
    pdf.save()
    output.seek(0)
    return output.getvalue()


def _build_pdf_bytes_legacy_two_page(data: PatientData, result: TestResult, timeseries_df: pd.DataFrame) -> bytes:
    """Gera PDF clínico premium em duas páginas: resumo executivo e curvas/análise integrada."""

    payload = build_report_payload(data, result, timeseries_df)
    integrated = build_integrated_recovery_analysis(data, timeseries_df)
    output = BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    pdf.setTitle("Relatório Clínico TC6M")

    width, height = A4
    margin = 30
    usable = width - 2 * margin
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    page_total = 2

    # Página 1 - resumo executivo.
    y = _draw_report_header(pdf, data, result, 1, page_total, margin, width, height, now)

    # Card principal.
    _draw_card_box(pdf, margin, y, usable, 132, "#F2FBF8")
    _draw_card_title(pdf, margin + 16, y - 22, "resultado_principal", "RESULTADO PRINCIPAL")
    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.line(margin + 92, y - 56, margin + 92, y - 104)
    pdf.setFillColor(_hex("#E7F5EF"))
    pdf.circle(margin + 48, y - 78, 28, fill=True, stroke=False)
    _draw_icon_circle(pdf, "distancia_percorrida", margin + 48, y - 78, circle_size=56, icon_size=40)
    pdf.setFont("Helvetica-Bold", 38)
    pdf.setFillColor(_hex(REPORT_COLORS["primary"]))
    pdf.drawString(margin + 116, y - 77, f"{data.distancia:.0f} m")
    pdf.setFont("Helvetica", 9)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(margin + 118, y - 95, "Distância percorrida no TC6M")

    bar_x = margin + 260
    bar_w = usable - 290
    pdf.setFont("Helvetica", 8)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(bar_x, y - 54, "0 m")
    pdf.drawRightString(bar_x + bar_w, y - 54, f"Predito: {result.dpp_principal:.2f} m")
    _draw_progress_bar(pdf, bar_x, y - 74, bar_w, result.percentual_atingido)
    pdf.setFillColor(_hex(REPORT_COLORS["progress"]))
    pdf.setFont("Helvetica-Bold", 15)
    pdf.drawString(bar_x, y - 103, f"{result.percentual_atingido:.2f}%")
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.setFont("Helvetica", 8)
    pdf.drawString(bar_x + 72, y - 100, "do previsto")
    pdf.setFillColor(_hex(REPORT_COLORS["progress"]))
    pdf.drawString(bar_x, y - 86, f"LIN: {payload['lin_label']}")
    pdf.drawRightString(bar_x + bar_w, y - 86, f"DPP: {result.dpp_principal:.2f} m")

    y -= 154
    half = (usable - 14) / 2
    _draw_card_box(pdf, margin, y, half, 154)
    _draw_card_box(pdf, margin + half + 14, y, half, 154)

    _draw_card_title(pdf, margin + 18, y - 22, "classificacao_risco", "CLASSIFICAÇÃO DE RISCO", REPORT_COLORS["primary"])
    scale_x = margin + 34
    scale_y = y - 82
    scale_w = half - 68
    label_y = y - 52
    risk_labels = [("Nível 1", "Baixo"), ("Nível 2", "Moderado"), ("Nível 3", "Alto"), ("Nível 4", "Muito alto")]
    segment_w = scale_w / 4
    risk_colors = [REPORT_COLORS["risk_low"], REPORT_COLORS["risk_moderate"], REPORT_COLORS["risk_high"], REPORT_COLORS["risk_very_high"]]
    for index, (label, sub) in enumerate(risk_labels):
        sx = scale_x + index * segment_w
        pdf.setFillColor(_hex(risk_colors[index]))
        pdf.roundRect(sx, scale_y, segment_w - 2, 5, 2, fill=True, stroke=False)
        pdf.setFillColor(_hex(risk_colors[index]))
        pdf.circle(sx + (segment_w / 2), scale_y + 5, 4, fill=True, stroke=False)
        pdf.setFillColor(_hex(REPORT_COLORS["text"]))
        pdf.setFont("Helvetica", 8)
        pdf.drawCentredString(sx + (segment_w / 2), label_y, label)
        pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
        pdf.setFont("Helvetica", 7)
        pdf.drawCentredString(sx + (segment_w / 2), label_y - 12, sub)
    risk = get_risk_display(result)
    position = {"4": 0, "3": 1, "2": 2}.get(risk["index"], 3)
    marker_x = scale_x + position * segment_w + (segment_w / 2)
    pdf.setStrokeColor(_hex(risk["color"]))
    pdf.line(marker_x, scale_y - 4, marker_x, scale_y - 20)
    pdf.setFillColor(_hex(risk["color"]))
    pdf.circle(marker_x, scale_y - 23, 3.2, fill=True, stroke=False)
    row_y = y - 116
    row_y = _draw_kv_row(pdf, margin + 18, row_y, half - 36, "Qualificador", result.qualificador_funcional)
    _draw_kv_row(pdf, margin + 18, row_y, half - 36, "Interrupção", "Sim" if data.interrompeu else "Não")

    pred_x = margin + half + 14
    _draw_card_title(pdf, pred_x + 18, y - 22, "predicoes_comparativas", "PREDIÇÕES COMPARATIVAS", REPORT_COLORS["primary"])
    pred_y = y - 48
    pred_y = _draw_kv_row(pdf, pred_x + 28, pred_y, half - 56, "DPP principal", f"{result.dpp_principal:.2f} m")
    pred_y = _draw_kv_row(pdf, pred_x + 28, pred_y, half - 56, "LIN principal", payload["lin_label"])
    pred_y = _draw_kv_row(pdf, pred_x + 28, pred_y, half - 56, "Iwama et al.", f"{result.dpp_iwama:.2f} m")
    pred_y = _draw_kv_row(pdf, pred_x + 28, pred_y, half - 56, "Ben Saad et al.", f"{result.dpp_ben_saad:.2f} m")
    _draw_wrapped_text(pdf, payload["prediction_note"], pred_x + 28, pred_y + 3, half - 56, 7.4, REPORT_COLORS["muted"])

    y -= 180
    _draw_card_title(pdf, margin, y, "metricas_hemodinamicas", "MÉTRICAS HEMODINÂMICAS")
    y -= 16
    metric_w = (usable - 32) / 5
    for index, metric in enumerate(payload["metrics"]):
        x = margin + index * (metric_w + 8)
        _draw_metric_box(pdf, x, y, metric_w, 66, metric["label"], metric["value"], metric["unit"])

    y -= 88
    _draw_card_box(pdf, margin, y, half, 114, "#F7FBFF")
    _draw_card_box(pdf, margin + half + 14, y, half, 114, "#FFFDF8")
    _draw_card_title(pdf, margin + 18, y - 22, "fator_limitante", "FATOR LIMITANTE PROVÁVEL", REPORT_COLORS["spo2"])
    pdf.setFont("Helvetica-Bold", 14)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(margin + 28, y - 48, result.fator_limitante[:28])
    _draw_wrapped_text(pdf, payload["factor_description"], margin + 28, y - 68, half - 50, 8.1, REPORT_COLORS["text"])

    _draw_card_title(pdf, margin + half + 32, y - 22, "pontos_atencao", "PONTOS DE ATENÇÃO", REPORT_COLORS["progress"])
    att_y = y - 48
    for point in payload["attention_points"][:4]:
        pdf.setFont("Helvetica", 8)
        pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
        pdf.drawString(margin + half + 34, att_y, point["label"][:34])
        _draw_badge(pdf, point["badge"], margin + usable - 82, att_y + 8, point["type"])
        att_y -= 20

    y -= 128
    _draw_card_box(pdf, margin, y, usable, 82, "#F2FBF8")
    _draw_card_title(pdf, margin + 18, y - 20, "resumo_clinico", "RESUMO CLÍNICO")
    pdf.setFont("Helvetica-Bold", 8.4)
    pdf.setFillColor(_hex(REPORT_COLORS["primary"]))
    pdf.drawString(margin + 52, y - 36, "Síntese interpretativa")
    _draw_wrapped_text_limited(
        pdf,
        payload["clinical_summary"],
        margin + 52,
        y - 50,
        usable - 78,
        7.1,
        REPORT_COLORS["text"],
        max_lines=4,
    )

    pdf.showPage()

    # Página 2 - curvas e análise integrada.
    y = _draw_report_header(pdf, data, result, 2, page_total, margin, width, height, now)
    _draw_card_title(pdf, margin, y, "graficos_achados", "Gráficos e achados das curvas")
    y -= 18

    findings = build_curve_findings(timeseries_df)[:4]
    _draw_card_box(pdf, margin, y, usable, 108, "#F2FBF8")
    _draw_card_title(pdf, margin + 16, y - 18, "achados_automaticos", "ACHADOS AUTOMÁTICOS")
    _draw_bullet_list(
        pdf,
        findings,
        margin + 22,
        y - 42,
        usable - 44,
        font_size=7.25,
        bullet_color=REPORT_COLORS["primary"],
        max_items=4,
    )

    y -= 122
    oscillation_png = _figure_to_png_bytes(build_oscillation_figure(timeseries_df))
    effort_png = _figure_to_png_bytes(build_effort_figure(timeseries_df))
    chart_gap = 12
    chart_w = (usable - chart_gap) / 2
    chart_h = 164

    _draw_card_box(pdf, margin, y, chart_w, chart_h)
    _draw_card_title(pdf, margin + 12, y - 18, "graficos_achados", "FC e SpO2")
    pdf.drawImage(
        ImageReader(oscillation_png),
        margin + 10,
        y - chart_h + 10,
        width=chart_w - 20,
        height=chart_h - 40,
        preserveAspectRatio=True,
    )

    chart2_x = margin + chart_w + chart_gap
    _draw_card_box(pdf, chart2_x, y, chart_w, chart_h)
    _draw_card_title(pdf, chart2_x + 12, y - 18, "borg_respiratorio", "Borg respiratório e MMII")
    pdf.drawImage(
        ImageReader(effort_png),
        chart2_x + 10,
        y - chart_h + 10,
        width=chart_w - 20,
        height=chart_h - 40,
        preserveAspectRatio=True,
    )

    y -= chart_h + 14
    _draw_card_title(pdf, margin, y, "duplo_produto_recuperacao", "Análise integrada de recuperação cardiovascular e velocidade")
    y -= 16

    left_w = usable * 0.48
    right_w = usable - left_w - 14
    integrated_h = 210
    _draw_card_box(pdf, margin, y, left_w, integrated_h, "#FFFFFF")
    _draw_card_title(pdf, margin + 12, y - 18, "duplo_produto_recuperacao", "Recuperação do Duplo Produto")
    dp_figure = build_dp_recovery_figure(data, timeseries_df)
    if dp_figure is not None:
        dp_png = _figure_to_png_bytes(dp_figure)
        pdf.drawImage(
            ImageReader(dp_png),
            margin + 12,
            y - integrated_h + 12,
            width=left_w - 24,
            height=integrated_h - 44,
            preserveAspectRatio=True,
        )
    else:
        _draw_wrapped_text(
            pdf,
            "Gráfico não exibido: dados insuficientes para plotar ao menos dois pontos válidos de Duplo Produto.",
            margin + 18,
            y - 48,
            left_w - 36,
            7.4,
            REPORT_COLORS["muted"],
        )

    right_x = margin + left_w + 14
    _draw_card_box(pdf, right_x, y, right_w, integrated_h, "#F7FBFF")
    _draw_card_title(pdf, right_x + 12, y - 18, "interpretacao_integrada", "Leitura integrada", REPORT_COLORS["spo2"])

    metrics = [
        ("DP repouso", format_analysis_value(integrated["dp_repouso"], "", 0), "bpm.mmHg"),
        ("DP 1 min", format_analysis_value(integrated["dp_1"], "", 0), "bpm.mmHg"),
        ("DP 3 min", format_analysis_value(integrated["dp_3"], "", 0), "bpm.mmHg"),
        ("DP 6 min", format_analysis_value(integrated["dp_6"], "", 0), "bpm.mmHg"),
        ("Recup. DP 6 min", format_analysis_value(integrated["recovery_percent_6"], "", 1), "%"),
        ("Custo DP/m", format_analysis_value(integrated["cost_dp_per_m"], "", 2), "DP/m"),
        ("Velocidade média", format_analysis_value(integrated["velocity_ms"], "", 2), "m/s"),
        ("Ritmo médio", format_analysis_value(integrated["pace_m_min"], "", 1), "m/min"),
        ("Velocidade norm.", format_analysis_value(integrated["normalized_velocity"], "", 3), "Complementar"),
        ("Comp. membro inf.", format_analysis_value(integrated["limb_length_m"], "", 2), "m"),
    ]
    metric_y = y - 44
    metric_col_w = (right_w - 34) / 2
    for index, (label, value, unit) in enumerate(metrics):
        row = index // 2
        col = index % 2
        mx = right_x + 12 + col * (metric_col_w + 10)
        my = metric_y - row * 32
        pdf.setFillColor(_hex("#FFFFFF"))
        pdf.roundRect(mx, my - 24, metric_col_w, 28, 7, fill=True, stroke=False)
        pdf.setFont("Helvetica", 5.9)
        pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
        pdf.drawString(mx + 7, my - 8, label[:18])
        pdf.setFont("Helvetica-Bold", 8.8)
        pdf.setFillColor(_hex(REPORT_COLORS["text"]))
        pdf.drawString(mx + 7, my - 20, value[:13])
        pdf.setFont("Helvetica", 5.2)
        pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
        pdf.drawRightString(mx + metric_col_w - 7, my - 20, unit[:12])

    pdf.setFont("Helvetica", 5.6)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(
        right_x + 14,
        y - integrated_h + 16,
        "Velocidade normalizada: análise biomecânica complementar.",
    )

    y -= integrated_h + 8
    interpretation_h = 74
    interp_w = usable * 0.63
    note_w = usable - interp_w - 14
    _draw_card_box(pdf, margin, y, interp_w, interpretation_h, "#F7FBFF")
    _draw_card_title(pdf, margin + 14, y - 17, "interpretacao_integrada", "INTERPRETAÇÃO INTEGRADA", REPORT_COLORS["spo2"])
    _draw_wrapped_text_limited(
        pdf,
        build_integrated_recovery_interpretation(integrated),
        margin + 40,
        y - 36,
        interp_w - 54,
        6.4,
        REPORT_COLORS["text"],
        max_lines=5,
    )

    note_x = margin + interp_w + 14
    _draw_card_box(pdf, note_x, y, note_w, interpretation_h, "#FFFDF8")
    _draw_card_title(pdf, note_x + 12, y - 17, "nota_metodologica", "NOTA METODOLÓGICA", REPORT_COLORS["progress"])
    _draw_wrapped_text_limited(
        pdf,
        integrated["notice"],
        note_x + 22,
        y - 36,
        note_w - 34,
        6.2,
        REPORT_COLORS["text"],
        max_lines=5,
    )

    pdf.save()
    output.seek(0)
    return output.getvalue()


def _compact_pdf_value(value: object, digits: int = 2) -> str:
    """Formata valores para cards pequenos, evitando textos longos demais."""

    if value is None:
        return "Não calc."
    if isinstance(value, int):
        return f"{value:,}".replace(",", ".")
    if isinstance(value, float):
        return f"{value:.{digits}f}".replace(".", ",")
    text = str(value)
    return "Não calc." if text == NOT_CALCULATED else text


def _draw_pdf_metric_tile(
    pdf: canvas.Canvas,
    x: float,
    y: float,
    width: float,
    height: float,
    label: str,
    value: str,
    unit: str,
    icon_key: str = "metricas_hemodinamicas",
    fill: str = "#F7FCFA",
) -> None:
    """Card pequeno e seguro: rótulo curto, número e unidade."""

    _draw_card_box(pdf, x, y, width, height, fill)
    _draw_icon_circle(pdf, icon_key, x + 18, y - 20, circle_size=30, icon_size=30)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.setFont("Helvetica", 6.7)
    pdf.drawString(x + 38, y - 13, _clip_text_to_width(pdf, label, "Helvetica", 6.7, width - 46))
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.setFont("Helvetica-Bold", 13)
    pdf.drawString(x + 10, y - 38, _clip_text_to_width(pdf, value, "Helvetica-Bold", 13, width - 18))
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.setFont("Helvetica", 6.8)
    pdf.drawString(x + 10, y - 52, _clip_text_to_width(pdf, unit, "Helvetica", 6.8, width - 18))


def _draw_pdf_text_card(
    pdf: canvas.Canvas,
    x: float,
    y: float,
    width: float,
    height: float,
    icon_key: str,
    title: str,
    text: str,
    fill: str = "#FFFFFF",
    font_size: float = 8,
) -> None:
    """Card amplo para textos interpretativos, com quebra de linha automática."""

    _draw_card_box(pdf, x, y, width, height, fill)
    _draw_card_title(pdf, x + 14, y - 18, icon_key, title)
    _draw_wrapped_text(pdf, text, x + 20, y - 42, width - 40, font_size, REPORT_COLORS["text"])


def _pdf_cautious_summary(data: PatientData, result: TestResult, timeseries_df: pd.DataFrame) -> str:
    """Resumo curto para PDF sem linguagem forte de morbimortalidade."""

    clean = normalize_timeseries(timeseries_df)
    exercise = clean.iloc[1:7].copy()
    _, pico, _ = get_phase_snapshots(clean)
    sexo_texto = "masculino" if data.sexo == "M" else "feminino"
    valid_spo2 = exercise.loc[exercise["SpO2"] > 0, "SpO2"]
    min_spo2 = int(valid_spo2.min()) if not valid_spo2.empty else pico.spo2
    interrupcao = "sem interrupção registrada" if not data.interrompeu else "com interrupção registrada"
    protocolo = f" {descrever_protocolo_corredor(data.comprimento_corredor_m)}"
    return (
        f"Paciente {sexo_texto} de {data.idade} anos percorreu {data.distancia:.2f} m no TC6M, "
        f"atingindo {result.percentual_atingido:.2f}% do previsto pela fórmula {result.formula_principal}. "
        f"O resultado foi classificado como {result.qualificador_funcional} e {result.classificacao_risco}, "
        "devendo ser interpretado como rastreamento funcional de atenção clínica, sem caráter diagnóstico isolado. "
        f"Durante o esforço, a menor SpO2 registrada foi {min_spo2}% e o pico de Borg respiratório/MMII "
        f"foi {pico.borg_resp:.1f}/{pico.borg_mmii:.1f}, sugerindo {result.fator_limitante.lower()}, {interrupcao}."
        f"{protocolo}"
    )


def _draw_pdf_chart_card(
    pdf: canvas.Canvas,
    x: float,
    y: float,
    width: float,
    height: float,
    title: str,
    png: BytesIO | None,
    icon_key: str = "graficos_achados",
    fallback_text: str = "Dados insuficientes para exibir o gráfico.",
) -> None:
    """Card de gráfico com área ampla e fallback sem quebrar o PDF."""

    _draw_card_box(pdf, x, y, width, height, "#FFFFFF")
    _draw_card_title(pdf, x + 14, y - 18, icon_key, title)
    if png is None:
        _draw_wrapped_text(pdf, fallback_text, x + 22, y - 52, width - 44, 8.5, REPORT_COLORS["muted"])
        return
    pdf.drawImage(
        ImageReader(png),
        x + 14,
        y - height + 12,
        width=width - 28,
        height=height - 46,
        preserveAspectRatio=True,
        anchor="c",
    )


def build_pdf_bytes(data: PatientData, result: TestResult, timeseries_df: pd.DataFrame) -> bytes:
    """Gera PDF clínico premium em 3 páginas com layout fixo e seguro."""

    payload = build_report_payload(data, result, timeseries_df)
    integrated = build_integrated_recovery_analysis(data, timeseries_df)
    output = BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    pdf.setTitle("Relatório Clínico TC6M")

    width, height = A4
    margin = 30
    usable = width - 2 * margin
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    page_total = 3

    # Página 1 — resumo clínico principal.
    y = _draw_report_header(pdf, data, result, 1, page_total, margin, width, height, now)

    _draw_card_box(pdf, margin, y, usable, 128, "#F2FBF8")
    _draw_card_title(pdf, margin + 16, y - 22, "resultado_principal", "RESULTADO PRINCIPAL")
    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.line(margin + 100, y - 54, margin + 100, y - 104)
    _draw_icon_circle(pdf, "distancia_percorrida", margin + 52, y - 78, circle_size=56, icon_size=40)
    pdf.setFont("Helvetica-Bold", 38)
    pdf.setFillColor(_hex(REPORT_COLORS["primary"]))
    pdf.drawString(margin + 120, y - 76, f"{data.distancia:.0f} m")
    pdf.setFont("Helvetica", 8.8)
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.drawString(margin + 122, y - 94, "Distância percorrida no TC6M")

    bar_x = margin + 282
    bar_w = usable - 320
    pdf.setFont("Helvetica", 8)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(bar_x, y - 48, "0 m")
    pdf.drawRightString(bar_x + bar_w, y - 48, f"Predito: {result.dpp_principal:.2f} m")
    _draw_progress_bar(pdf, bar_x, y - 68, bar_w, result.percentual_atingido)
    pdf.setFont("Helvetica", 7.4)
    pdf.setFillColor(_hex(REPORT_COLORS["progress"]))
    pdf.drawString(bar_x, y - 84, f"LIN: {payload['lin_label']}")
    pdf.drawRightString(bar_x + bar_w, y - 84, f"DPP: {result.dpp_principal:.2f} m")
    pdf.setStrokeColor(_hex(REPORT_COLORS["border"]))
    pdf.setLineWidth(0.4)
    pdf.line(bar_x, y - 92, bar_x + bar_w, y - 92)
    pdf.setFont("Helvetica-Bold", 15)
    pdf.drawString(bar_x, y - 110, f"{result.percentual_atingido:.2f}%")
    pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
    pdf.setFont("Helvetica", 8)
    pdf.drawString(bar_x + 72, y - 107, "do previsto")

    y -= 148
    half = (usable - 14) / 2
    _draw_card_box(pdf, margin, y, half, 142)
    _draw_card_box(pdf, margin + half + 14, y, half, 142)

    _draw_card_title(pdf, margin + 18, y - 22, "classificacao_risco", "CLASSIFICAÇÃO FUNCIONAL")
    risk_y = _draw_risk_scale(pdf, margin + 28, y - 50, half - 56, result)
    risk_y = _draw_kv_row(pdf, margin + 18, risk_y - 2, half - 36, "Qualificador", result.qualificador_funcional)
    _draw_kv_row(pdf, margin + 18, risk_y - 2, half - 36, "Interrupção", "Sim" if data.interrompeu else "Não")

    pred_x = margin + half + 14
    _draw_card_title(pdf, pred_x + 18, y - 22, "predicoes_comparativas", "PREDIÇÕES COMPARATIVAS")
    pred_y = y - 48
    pred_y = _draw_kv_row(pdf, pred_x + 28, pred_y, half - 56, "DPP principal", f"{result.dpp_principal:.2f} m")
    pred_y = _draw_kv_row(pdf, pred_x + 28, pred_y, half - 56, "LIN principal", payload["lin_label"])
    pred_y = _draw_kv_row(pdf, pred_x + 28, pred_y, half - 56, "Iwama et al.", f"{result.dpp_iwama:.2f} m")
    pred_y = _draw_kv_row(pdf, pred_x + 28, pred_y, half - 56, "Ben Saad et al.", f"{result.dpp_ben_saad:.2f} m")
    _draw_wrapped_text(pdf, payload["prediction_note"], pred_x + 28, pred_y + 3, half - 56, 7.2, REPORT_COLORS["muted"])

    y -= 166
    _draw_card_title(pdf, margin, y, "metricas_hemodinamicas", "MÉTRICAS HEMODINÂMICAS")
    y -= 16
    metric_w = (usable - 32) / 5
    for index, metric in enumerate(payload["metrics"]):
        x = margin + index * (metric_w + 8)
        _draw_metric_box(pdf, x, y, metric_w, 66, metric["label"], metric["value"], metric["unit"])

    y -= 88
    _draw_card_box(pdf, margin, y, half, 116, "#F7FBFF")
    _draw_card_box(pdf, margin + half + 14, y, half, 116, "#FFFDF8")
    _draw_card_title(pdf, margin + 18, y - 22, "fator_limitante", "FATOR LIMITANTE PROVÁVEL", REPORT_COLORS["spo2"])
    pdf.setFont("Helvetica-Bold", 14)
    pdf.setFillColor(_hex(REPORT_COLORS["text"]))
    pdf.drawString(margin + 28, y - 48, result.fator_limitante[:28])
    _draw_wrapped_text(pdf, payload["factor_description"], margin + 28, y - 68, half - 50, 7.7, REPORT_COLORS["text"])

    attention_x = margin + half + 32
    _draw_card_title(pdf, attention_x, y - 22, "pontos_atencao", "PONTOS DE ATENÇÃO", REPORT_COLORS["progress"])
    att_y = y - 48
    for point in payload["attention_points"][:4]:
        pdf.setFont("Helvetica", 7.8)
        pdf.setFillColor(_hex(REPORT_COLORS["muted"]))
        pdf.drawString(attention_x + 2, att_y, point["label"][:34])
        _draw_badge(pdf, point["badge"], margin + usable - 82, att_y + 8, point["type"])
        att_y -= 20

    y -= 136
    _draw_pdf_text_card(
        pdf,
        margin,
        y,
        usable,
        92,
        "resumo_clinico",
        "RESUMO CLÍNICO",
        _pdf_cautious_summary(data, result, timeseries_df),
        fill="#F2FBF8",
        font_size=7.8,
    )

    pdf.showPage()

    # Página 2 — gráficos e achados das curvas.
    y = _draw_report_header(pdf, data, result, 2, page_total, margin, width, height, now)
    _draw_card_title(pdf, margin, y, "graficos_achados", "GRÁFICOS E ACHADOS DAS CURVAS")
    y -= 18

    findings = build_curve_findings(timeseries_df)[:4]
    _draw_card_box(pdf, margin, y, usable, 102, "#F2FBF8")
    _draw_card_title(pdf, margin + 16, y - 18, "achados_automaticos", "ACHADOS AUTOMÁTICOS")
    _draw_bullet_list(pdf, findings, margin + 24, y - 42, usable - 48, font_size=7.4, max_items=4)

    y -= 118
    oscillation_png = _figure_to_png_bytes(build_oscillation_figure(timeseries_df))
    effort_png = _figure_to_png_bytes(build_effort_figure(timeseries_df))
    _draw_pdf_chart_card(
        pdf,
        margin,
        y,
        usable,
        202,
        "Oscilação cardiorrespiratória — FC e SpO2",
        oscillation_png,
        "graficos_achados",
    )
    y -= 218
    _draw_pdf_chart_card(
        pdf,
        margin,
        y,
        usable,
        202,
        "Curva de esforço percebido — Borg respiratório e MMII",
        effort_png,
        "borg_respiratorio",
    )
    y -= 218
    graph_note = (
        "As curvas devem ser interpretadas em conjunto: a relação entre FC, SpO2 e Borg ajuda a diferenciar "
        "resposta cardiovascular, dessaturação ao esforço e percepção subjetiva de esforço. Os achados sugeridos "
        "não substituem avaliação clínica e devem ser comparados com sintomas, técnica do teste e contexto do paciente."
    )
    _draw_pdf_text_card(
        pdf,
        margin,
        y,
        usable,
        78,
        "interpretacao_integrada",
        "INTERPRETAÇÃO DOS GRÁFICOS",
        graph_note,
        fill="#F7FBFF",
        font_size=7.8,
    )

    pdf.showPage()

    # Página 3 — análise integrada.
    y = _draw_report_header(pdf, data, result, 3, page_total, margin, width, height, now)
    _draw_card_title(pdf, margin, y, "duplo_produto_recuperacao", "ANÁLISE INTEGRADA DE RECUPERAÇÃO CARDIOVASCULAR E VELOCIDADE")
    y -= 18

    dp_figure = build_dp_recovery_figure(data, timeseries_df)
    dp_png = _figure_to_png_bytes(dp_figure) if dp_figure is not None else None
    _draw_pdf_chart_card(
        pdf,
        margin,
        y,
        usable,
        230,
        "Recuperação do Duplo Produto após o TC6M",
        dp_png,
        "duplo_produto_recuperacao",
        "Gráfico não exibido: dados insuficientes para plotar ao menos dois pontos válidos de Duplo Produto.",
    )

    y -= 248
    cards = [
        ("DP repouso", _compact_pdf_value(integrated["dp_repouso"], 0), "bpm.mmHg", "dp_repouso"),
        ("DP 1 min", _compact_pdf_value(integrated["dp_1"], 0), "bpm.mmHg", "dp_recuperacao"),
        ("DP 3 min", _compact_pdf_value(integrated["dp_3"], 0), "bpm.mmHg", "dp_recuperacao"),
        ("DP 6 min", _compact_pdf_value(integrated["dp_6"], 0), "bpm.mmHg", "dp_recuperacao"),
        ("Recup. DP 6 min", _compact_pdf_value(integrated["recovery_percent_6"], 1), "%", "dp_recuperacao"),
        ("Custo DP/m", _compact_pdf_value(integrated["cost_dp_per_m"], 2), "DP/m", "metricas_hemodinamicas"),
        ("Velocidade média", _compact_pdf_value(integrated["velocity_ms"], 2), "m/s", "velocidade_media"),
        ("Ritmo médio", _compact_pdf_value(integrated["pace_m_min"], 1), "m/min", "velocidade_media"),
        ("Velocidade norm.", _compact_pdf_value(integrated["normalized_velocity"], 3), "Complementar", "velocidade_normalizada"),
        ("Comp. membro inf.", _compact_pdf_value(integrated["limb_length_m"], 2), "m", "velocidade_normalizada"),
    ]
    tile_gap = 8
    tile_w = (usable - tile_gap * 4) / 5
    tile_h = 62
    for index, (label, value, unit, icon_key) in enumerate(cards):
        row = index // 5
        col = index % 5
        x = margin + col * (tile_w + tile_gap)
        tile_y = y - row * (tile_h + 10)
        _draw_pdf_metric_tile(pdf, x, tile_y, tile_w, tile_h, label, value, unit, icon_key)

    y -= (tile_h * 2) + 32
    _draw_pdf_text_card(
        pdf,
        margin,
        y,
        usable,
        112,
        "interpretacao_integrada",
        "INTERPRETAÇÃO INTEGRADA",
        build_integrated_recovery_interpretation(integrated),
        fill="#F7FBFF",
        font_size=7.8,
    )
    y -= 130
    _draw_pdf_text_card(
        pdf,
        margin,
        y,
        usable,
        86,
        "nota_metodologica",
        "NOTA METODOLÓGICA",
        integrated["notice"],
        fill="#FFFDF8",
        font_size=7.8,
    )

    pdf.save()
    output.seek(0)
    return output.getvalue()


def build_safe_filename(patient_name: str, extension: str) -> str:
    """Cria nome de arquivo seguro com nome do teste, paciente e data."""

    clean_patient_name = format_patient_name(patient_name)
    clean_name = "".join(char if char.isalnum() else "_" for char in clean_patient_name.strip()).strip("_")
    if not clean_name:
        clean_name = "paciente"

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    return f"TC6M_{clean_name}_{timestamp}.{extension}"
