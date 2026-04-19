from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from typing import Iterable

import matplotlib.pyplot as plt
import pandas as pd
from openpyxl.styles import Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfgen import canvas


FORMULAS_DPP = [
    "Enright e Sherrill (1998) - adultos",
    "Iwama et al. (2009) - adultos",
    "Ben Saad et al. (2009) - crianças/adolescentes",
]

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
    if data.formula_principal not in FORMULAS_DPP:
        raise ValueError("Selecione uma fórmula predita válida.")


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

    return (
        f"O paciente percorreu {data.distancia:.2f} m no TC6M. Pela fórmula selecionada "
        f"({data.formula_principal}), a distância predita principal foi de {dpp_principal:.2f} m, "
        f"correspondendo a {percentual:.2f}% do previsto. Qualificador funcional: {qualificador}. "
        f"Classificação por distância: {classificacao}. Risco associado: {risco}. "
        f"No pico registrado durante a caminhada, observou-se FC={pico.fc} bpm, SpO2={pico.spo2}% "
        f"e Borg respiratório/MMII={pico.borg_resp:.1f}/{pico.borg_mmii:.1f}, sugerindo "
        f"{fator_limitante.lower()}.{interrupcao}{motivo}"
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
                "Contraindicação absoluta",
                "Contraindicação relativa",
                "Observação da triagem",
            ],
            "Valor": [
                data.nome,
                data.prontuario or "-",
                data.data_avaliacao.strftime("%d/%m/%Y") if data.data_avaliacao else "-",
                data.avaliador or "-",
                data.diagnostico or "-",
                "Masculino" if data.sexo == "M" else "Feminino",
                f"{data.idade} anos",
                f"{data.peso:.1f} kg",
                f"{data.altura_cm:.1f} cm",
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
    ax1.plot(clean["Tempo"], clean["FC"], marker="o", color="#8B1E1E", linewidth=3, label="FC")
    ax1.set_ylabel("FC (bpm)", color="#8B1E1E", fontsize=12)
    ax1.tick_params(axis="y", labelcolor="#8B1E1E")
    ax1.tick_params(axis="x", rotation=28)
    ax1.grid(True, alpha=0.25)

    ax2 = ax1.twinx()
    ax2.plot(clean["Tempo"], clean["SpO2"], marker="s", color="#1F5C99", linewidth=3, label="SpO2")
    ax2.set_ylabel("SpO2 (%)", color="#1F5C99", fontsize=12)
    ax2.tick_params(axis="y", labelcolor="#1F5C99")

    fig.suptitle("Oscilação cardiorrespiratória durante o TC6M", fontsize=15, fontweight="bold")
    fig.tight_layout()
    return fig


def build_effort_figure(timeseries_df: pd.DataFrame):
    """Cria gráfico grande da curva de esforço percebido."""

    clean = normalize_timeseries(timeseries_df)
    fig, ax = plt.subplots(figsize=(14, 5.8))
    ax.plot(clean["Tempo"], clean["Borg Respiratório"], marker="o", color="#2B6A4F", linewidth=3, label="Borg respiratório")
    ax.plot(clean["Tempo"], clean["Borg MMII"], marker="s", color="#C26A21", linewidth=3, label="Borg MMII")
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

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        patient_df.to_excel(writer, sheet_name="Paciente", index=False)
        summary_df.to_excel(writer, sheet_name="Resumo TC6M", index=False)
        clean.to_excel(writer, sheet_name="Sinais vitais", index=False)
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


def _draw_wrapped_text(pdf: canvas.Canvas, text: str, x: float, y: float, width: float, font_size: int = 9) -> float:
    """Desenha texto no PDF com quebra automática de linha."""

    pdf.setFont("Helvetica", font_size)
    lines = simpleSplit(text, "Helvetica", font_size, width)
    for line in lines:
        pdf.drawString(x, y, line)
        y -= font_size + 3
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


def build_pdf_bytes(data: PatientData, result: TestResult, timeseries_df: pd.DataFrame) -> bytes:
    """Gera PDF clínico com cabeçalho, resumo, gráficos e interpretação."""

    output = BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    width, height = A4
    margin = 42
    y = height - margin

    pdf.setFillColor(colors.HexColor("#183C34"))
    pdf.rect(0, height - 78, width, 78, fill=True, stroke=False)
    pdf.setFillColor(colors.white)
    pdf.setFont("Helvetica-Bold", 17)
    pdf.drawString(margin, height - 38, "Relatório Clínico - TC6M")
    pdf.setFont("Helvetica", 9)
    pdf.drawString(margin, height - 56, f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    y -= 64
    pdf.setFillColor(colors.black)
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(margin, y, "Identificação e triagem")
    y -= 18

    patient_rows = [(row["Campo"], row["Valor"]) for _, row in build_patient_dataframe(data).iterrows()]
    y = _draw_table(pdf, patient_rows, margin, y, width - 2 * margin)

    y -= 16
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(margin, y, "Resumo técnico")
    y -= 18

    summary_df = build_summary_dataframe(data, result)
    summary_rows = [(f"{row['Bloco']} - {row['Indicador']}", row["Resultado"]) for _, row in summary_df.iterrows()]
    y = _draw_table(pdf, summary_rows, margin, y, width - 2 * margin)

    y -= 15
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(margin, y, "Interpretação automatizada")
    y -= 16
    _draw_wrapped_text(pdf, result.interpretacao, margin, y, width - 2 * margin, font_size=9)

    pdf.showPage()
    y = height - margin
    pdf.setFillColor(colors.black)
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(margin, y, "Gráficos de monitoramento")
    y -= 18

    oscillation_png = _figure_to_png_bytes(build_oscillation_figure(timeseries_df))
    effort_png = _figure_to_png_bytes(build_effort_figure(timeseries_df))

    pdf.drawImage(ImageReader(oscillation_png), margin, y - 230, width=width - 2 * margin, height=215, preserveAspectRatio=True)
    y -= 265
    pdf.drawImage(ImageReader(effort_png), margin, y - 230, width=width - 2 * margin, height=215, preserveAspectRatio=True)

    pdf.showPage()
    pdf.save()
    output.seek(0)
    return output.getvalue()


def build_safe_filename(patient_name: str, extension: str) -> str:
    """Cria nome de arquivo seguro com nome do teste, paciente e data."""

    clean_name = "".join(char if char.isalnum() else "_" for char in patient_name.strip()).strip("_")
    if not clean_name:
        clean_name = "paciente"

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    return f"TC6M_{clean_name}_{timestamp}.{extension}"
