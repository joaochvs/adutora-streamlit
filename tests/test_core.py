import unittest
from io import BytesIO
from datetime import time, datetime
import pandas as pd
from core import preparar, gerar_pdf, ler_excel, periodo


class RevisitasTest(unittest.TestCase):
    def test_trechos_dos_ausentes(self):
        df = self.base()
        df["TRECHO"] = [9.0, 10.0, 11.0]
        self.assertEqual(preparar(df).attrs["trechos"], ["9"])
        df.loc[1, "STATUS"] = "Ausente 2"
        self.assertEqual(preparar(df).attrs["trechos"], ["9", "10"])

    def test_ausentes_numerados_e_tracos(self):
        df = pd.DataFrame({"CHAVE_VISITA": ["V1", "V2", "V3", "V4", "V5"],
                           "LOGRADOURO": ["Rua A"]*5,
                           "STATUS": ["AUSENTE 1", " ausente 2 ", "AUSENTE 12", "PENDENTE", "NÃO AUSENTE 1"],
                           "Turno": ["11:25:00", "13:46:00", "09:00:00", "-", "09:00:00"],
                           "Turno2": ["-", "-", "15:00:00", "-", "-"],
                           "Turno4": ["-"]*5, "Turno6": ["-"]*5})
        result = preparar(df)
        self.assertEqual(len(result), 3)
        self.assertEqual(list(result["Próxima visita"]), ["Tarde", "Folder", "Manhã"])
        self.assertEqual(result.iloc[0]["Recomendação"], "2ª visita = Tarde; confirmar dia")
        self.assertEqual(result.iloc[1]["Recomendação"], "3ª visita = Folder (sugestão); confirmar dia")

    def base(self):
        return pd.DataFrame({"CHAVE_VISITA": ["V1", "V2", "V3"], "LOGRADOURO": ["Rua A & B"]*3, "N°": [10.0,20.0,30.0], "COMPLEMENTO": ["Casa 2", "", ""], "STATUS": [" Ausente ", "Pendente", "Atendido"], "Turno": ["Manhã"]*3, "Turno2": ["Tarde", None, None], "Turno4": [None]*3, "Turno6": [None]*3})

    def test_apenas_ausente_e_endereco(self):
        result = preparar(self.base())
        self.assertEqual(len(result), 1)
        self.assertEqual(result.iloc[0]["Endereço"], "Rua A & B, 10, Casa 2")
        self.assertEqual(result.iloc[0]["Última visita"], "Tarde")
        self.assertEqual(result.iloc[0]["Próxima visita"], "Folder")

    def test_ultimo_turno_e_lacunas(self):
        df = self.base()
        df.loc[0, "STATUS"] = "Ausente 1"
        df.loc[0, "Turno6"] = "Manhã"
        self.assertEqual(preparar(df).iloc[0]["Próxima visita"], "Tarde")
        df.loc[0, "Turno6"] = "desconhecido"
        self.assertEqual(preparar(df).iloc[0]["Próxima visita"], "Confirmar")
        for c in ["Turno", "Turno2", "Turno4", "Turno6"]:
            df.loc[0,c] = None
        self.assertEqual(preparar(df).iloc[0]["Última visita"], "Não informado")

    def test_horario_e_sem_datas(self):
        self.assertEqual(periodo(time(9,30)), "Manhã")
        df = self.base()
        df.loc[0, "STATUS"] = "Ausente 1"
        df.loc[0,"Turno2"] = "14:30"
        df["Dia3"] = "data inválida"
        self.assertEqual(preparar(df).iloc[0]["Próxima visita"], "Manhã")
        self.assertEqual(preparar(df).iloc[0]["Última visita"], "14:30 - Tarde")

    def test_excel_e_cabecalho(self):
        from openpyxl import Workbook
        wb = Workbook()
        wb.active.title = "Outra"
        campo = wb.create_sheet("CAMPO")
        campo.append(["Título"])
        campo.append(["CHAVE_VISITA", "LOGRADOURO", "STATUS", "Turno"])
        campo.append(["001", "Rua A", "Ausente", "Manhã"])
        data = BytesIO()
        wb.save(data)
        df, sheet = ler_excel(data.getvalue())
        self.assertEqual(sheet, "CAMPO")
        self.assertEqual(preparar(df).iloc[0]["Próxima visita"], "Tarde")

    def test_historico_e_proximo_dia(self):
        df = self.base().iloc[:1].copy()
        df["Dia"] = ["01/10/2026"]
        df["Dia3"] = ["03/10/2026"]
        df["Turno"] = ["09:15"]
        df["Turno2"] = ["15:34"]
        r = preparar(df).iloc[0]
        self.assertIn("V1: 09:15 - Manhã - 01/10/2026 (quinta-feira)", r["Histórico"])
        self.assertIn("V2: 15:34 - Tarde - 03/10/2026 (sábado)", r["Histórico"])
        self.assertEqual(r["Recomendação"], "3ª visita = Folder (sugestão); 05/10/2026 (segunda-feira)")

    def test_faixas_e_sabado(self):
        df = self.base().iloc[:1].copy()
        df["STATUS"] = "Ausente 1"
        df["Turno2"] = None
        df["TRECHO"] = 8
        df["Dia"] = "quinta"
        self.assertEqual(preparar(df).iloc[0]["Recomendação"], "2ª visita = Tarde (12:00 - 15:00); sexta-feira")
        df["Dia"] = "sexta-feira"
        self.assertEqual(preparar(df).iloc[0]["Recomendação"], "2ª visita = Manhã ou Tarde (08:00 - 11:59 ou 12:00 - 15:00); sábado")
        df["Dia"] = "sábado"
        self.assertIn("segunda-feira", preparar(df).iloc[0]["Recomendação"])

    def test_rural_sem_noite(self):
        for trecho in range(15, 20):
            df = self.base().iloc[:1].copy()
            df["STATUS"] = "Ausente 1"
            df["TRECHO"] = f"Trecho {trecho}"
            df["Turno2"] = "19:30"
            df["Dia3"] = datetime(2026, 10, 1)
            self.assertEqual(preparar(df).iloc[0]["Próxima visita"], "Manhã")
            self.assertNotIn("Noite", preparar(df).iloc[0]["Recomendação"])

    def test_dia_sem_horario_nao_recua(self):
        df = self.base().iloc[:1].copy()
        df["STATUS"] = "Ausente 1"
        df["Dia7"] = "sábado"
        r = preparar(df).iloc[0]
        self.assertEqual(r["Próxima visita"], "Confirmar")
        self.assertIn("segunda-feira", r["Recomendação"])
        self.assertIn("V4:", r["Histórico"])

    def test_pdf_inclusive_sem_ausentes(self):
        result = preparar(self.base())
        self.assertTrue(gerar_pdf(result, "Teste").startswith(b"%PDF"))
        self.assertTrue(gerar_pdf(result.iloc[:0], "Teste").startswith(b"%PDF"))


if __name__ == "__main__":
    unittest.main()
