import unittest
from io import BytesIO
from datetime import time
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
        self.assertEqual(list(result["Próxima visita"]), ["Tarde", "Manhã", "Manhã"])
        self.assertEqual(result.iloc[0]["Recomendação"], "Realizar 2ª visita no período da tarde.")
        self.assertEqual(result.iloc[1]["Recomendação"], "Realizar 3ª visita no período da manhã.")

    def base(self):
        return pd.DataFrame({"CHAVE_VISITA": ["V1", "V2", "V3"], "LOGRADOURO": ["Rua A & B"]*3, "N°": [10.0,20.0,30.0], "COMPLEMENTO": ["Casa 2", "", ""], "STATUS": [" Ausente ", "Pendente", "Atendido"], "Turno": ["Manhã"]*3, "Turno2": ["Tarde", None, None], "Turno4": [None]*3, "Turno6": [None]*3})

    def test_apenas_ausente_e_endereco(self):
        result = preparar(self.base())
        self.assertEqual(len(result), 1)
        self.assertEqual(result.iloc[0]["Endereço"], "Rua A & B, 10, Casa 2")
        self.assertEqual(result.iloc[0]["Última visita"], "Tarde")
        self.assertEqual(result.iloc[0]["Próxima visita"], "Manhã")

    def test_ultimo_turno_e_lacunas(self):
        df = self.base()
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
        df.loc[0,"Turno2"] = "14:30"
        df["Dia3"] = "data inválida"
        self.assertEqual(preparar(df).iloc[0]["Próxima visita"], "Manhã")

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

    def test_pdf_inclusive_sem_ausentes(self):
        result = preparar(self.base())
        self.assertTrue(gerar_pdf(result, "Teste").startswith(b"%PDF"))
        self.assertTrue(gerar_pdf(result.iloc[:0], "Teste").startswith(b"%PDF"))


if __name__ == "__main__":
    unittest.main()
