import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

class GoogleSheetsUpdater:
    def __init__(self, credentials_file, spreadsheet_key, worksheet_index=0):
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
        self.client = gspread.authorize(self.credentials)
        self.spreadsheet = self.client.open_by_key(spreadsheet_key)
        self.worksheet = self.spreadsheet.get_worksheet(worksheet_index)

    def update_data(self, data, start_cell='A4'):
        self.worksheet.update(start_cell, data)

class GradeCalculator:
    def __init__(self, dataframe):
        self.df = dataframe

    def calculate_grades(self, row):
        media = (row['P1'] + row['P2'] + row['P3']) / 3
        misses = row['Faltas']
        total_classes = 60  

        if misses > 0.25 * total_classes:
            return "Reprovado por Falta", 0
        elif media < 50:
            return "Reprovado por Nota", 0
        elif 50 <= media < 70:
            naf = max(0, (100 - media))
            return "Exame Final", int(round(naf))
        else:
            return "Aprovado", 0

    def apply_grades(self):
        self.df[['Situação', 'Nota para Aprovação Final']] = self.df.apply(self.calculate_grades, axis=1, result_type='expand')
        return self.df.values.tolist()

def main():
    credentials_file = "credentials.json"
    spreadsheet_key = "1dNhlz1hW2G0hwOur6MJNskkaLsKc9ai6g4qpG0fCy0A"
    excel_file_path = "Engenharia de Software - Julio Cesar Francisco da Silva.xlsx"

    gs_updater = GoogleSheetsUpdater(credentials_file, spreadsheet_key)
    df = pd.read_excel(excel_file_path, skiprows=2).fillna(0)

    grades_calculator = GradeCalculator(df)
    update_data = grades_calculator.apply_grades()

    gs_updater.update_data(update_data)

if __name__ == "__main__":
    main()
