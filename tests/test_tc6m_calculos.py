from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
import math
import unittest

from pypdf import PdfReader

from tc6m import (
    FORMULAS_DPP,
    FULL_COLUMNS,
    ICON_MAP,
    PatientData,
    build_default_during_table,
    build_default_pre_table,
    build_default_recovery_table,
    build_dp_recovery_figure,
    build_integrated_recovery_analysis,
    build_pdf_bytes,
    calcular_duplo_produto,
    calcular_distancia_por_trechos,
    calcular_dpp_enright,
    calcular_fc_maxima,
    calcular_fc_submaxima,
    calculate_tc6m_professional,
    combine_timeseries,
    obter_qualificador_funcional,
)


def sample_timeseries():
    borg_resp = FULL_COLUMNS[6]
    borg_mmii = FULL_COLUMNS[7]

    pre = build_default_pre_table()
    pre.loc[0, ["FC", "SpO2", "FR", "PAS", "PAD", borg_resp, borg_mmii]] = [78, 97, 18, 122, 78, 0, 0]

    during = build_default_during_table()
    during["FC"] = [94, 106, 116, 124, 132, 139]
    during["SpO2"] = [96, 95, 95, 94, 93, 92]
    during[borg_resp] = [1, 2, 3, 4, 5, 6]
    during[borg_mmii] = [1, 2, 3, 4, 5, 5]

    recovery = build_default_recovery_table()
    recovery.loc[:, ["FC", "SpO2", "FR", "PAS", "PAD", borg_resp, borg_mmii]] = [
        [120, 94, 24, 146, 82, 4, 4],
        [98, 96, 20, 132, 80, 2, 2],
        [84, 97, 18, 124, 78, 1, 1],
    ]

    return combine_timeseries(pre, during, recovery)


def sample_patient(**overrides):
    data = {
        "nome": "Homem",
        "sexo": "M",
        "idade": 62,
        "peso": 74.0,
        "altura_cm": 171.0,
        "distancia": 438.0,
        "interrompeu": False,
        "formula_principal": FORMULAS_DPP[0],
        "data_avaliacao": date(2026, 5, 16),
        "prontuario": "TC6M-H-001",
        "avaliador": "Equipe Cardiorrespiratoria",
        "diagnostico": "Avaliacao funcional cardiorrespiratoria",
        "comprimento_membro_inferior_m": 0.90,
    }
    data.update(overrides)
    return PatientData(**data)


class TC6MCalculosTest(unittest.TestCase):
    def test_protocolos_de_corredor_calculam_distancia_real_sem_corrigir_predicao(self):
        self.assertEqual(calcular_distancia_por_trechos(30, 14, 18), 438.0)
        self.assertEqual(calcular_distancia_por_trechos(25, 17, 13), 438.0)

        with self.assertRaises(ValueError):
            calcular_distancia_por_trechos(25, 17, 25)
        with self.assertRaises(ValueError):
            calcular_distancia_por_trechos(12, 30, 0)
        with self.assertRaises(ValueError):
            calcular_distancia_por_trechos(20, 21, 18)

    def test_corredor_adaptado_nao_modifica_equacao_predita(self):
        series = sample_timeseries()
        padrao = calculate_tc6m_professional(sample_patient(comprimento_corredor_m=30), series)
        adaptado = calculate_tc6m_professional(sample_patient(comprimento_corredor_m=25), series)

        self.assertEqual(padrao.dpp_principal, adaptado.dpp_principal)
        self.assertEqual(padrao.percentual_atingido, adaptado.percentual_atingido)

    def test_predicao_percentual_fc_e_qualificador(self):
        dpp, lin = calcular_dpp_enright("M", 62, 74.0, 171.0)

        self.assertAlmostEqual(dpp, 543.99, places=2)
        self.assertAlmostEqual(lin, 390.99, places=2)
        self.assertEqual(calcular_fc_maxima(62), 158)
        self.assertEqual(calcular_fc_submaxima(62), 134)

        qualificador, percentual = obter_qualificador_funcional(438.0, dpp)
        self.assertEqual(qualificador, "Déficit funcional leve")
        self.assertAlmostEqual(percentual, 80.52, places=2)

    def test_duplo_produto_e_recuperacao_integrada(self):
        patient = sample_patient()
        analysis = build_integrated_recovery_analysis(patient, sample_timeseries())

        self.assertEqual(calcular_duplo_produto(78, 122), 9516)
        self.assertEqual(analysis["dp_repouso"], 9516)
        self.assertEqual(analysis["dp_1"], 17520)
        self.assertEqual(analysis["dp_3"], 12936)
        self.assertEqual(analysis["dp_6"], 10416)
        self.assertEqual(analysis["delta_dp_1"], 8004)
        self.assertEqual(analysis["recovery_dp_1_3"], 4584)
        self.assertEqual(analysis["recovery_dp_1_6"], 7104)
        self.assertAlmostEqual(analysis["recovery_percent_6"], 40.55, places=2)
        self.assertAlmostEqual(analysis["cost_dp_per_m"], 18.27, places=2)
        self.assertAlmostEqual(analysis["velocity_ms"], 1.22, places=2)
        self.assertAlmostEqual(analysis["pace_m_min"], 73.0, places=2)
        expected_normalized = (438.0 / 360) / math.sqrt(9.81 * 0.90)
        self.assertAlmostEqual(analysis["normalized_velocity"], round(expected_normalized, 3), places=3)

    def test_dados_ausentes_nao_quebram_analise_grafico_e_pdf(self):
        patient = sample_patient(distancia=180.0, interrompeu=True, comprimento_membro_inferior_m=None)
        series = combine_timeseries(
            build_default_pre_table(),
            build_default_during_table(),
            build_default_recovery_table(),
        )
        result = calculate_tc6m_professional(patient, series)
        analysis = build_integrated_recovery_analysis(patient, series)

        self.assertIsNone(analysis["dp_repouso"])
        self.assertIsNone(analysis["normalized_velocity"])
        self.assertIsNone(build_dp_recovery_figure(patient, series))

        pdf = build_pdf_bytes(patient, result, series)
        self.assertTrue(pdf.startswith(b"%PDF"))
        self.assertEqual(len(PdfReader(BytesIO(pdf)).pages), 3)

    def test_icones_oficiais_existem(self):
        required = {
            "achados_automaticos",
            "caminhada_tc6m",
            "classificacao_risco",
            "distancia_percorrida",
            "duplo_produto_recuperacao",
            "fator_limitante",
            "graficos_achados",
            "metricas_hemodinamicas",
            "interpretacao_integrada",
            "nota_metodologica",
            "pontos_atencao",
            "predicoes_comparativas",
            "resumo_clinico",
            "spo2",
        }
        missing = [key for key in required if key not in ICON_MAP or not Path(ICON_MAP[key]).exists()]
        self.assertEqual(missing, [])

    def test_cenarios_de_distancia_baixa_alta_e_interrupcao_nao_quebram(self):
        series = sample_timeseries()
        scenarios = [
            sample_patient(distancia=120.0),
            sample_patient(distancia=700.0),
            sample_patient(distancia=300.0),
            sample_patient(distancia=375.0),
            sample_patient(distancia=450.0),
            sample_patient(distancia=120.0, interrompeu=True, motivo_interrupcao="fadiga", distancia_interrupcao=120.0),
        ]

        for patient in scenarios:
            with self.subTest(distancia=patient.distancia, interrompeu=patient.interrompeu):
                result = calculate_tc6m_professional(patient, series)
                self.assertGreaterEqual(result.percentual_atingido, 0)
                self.assertTrue(result.classificacao_risco)
                self.assertTrue(result.risco)


if __name__ == "__main__":
    unittest.main()
