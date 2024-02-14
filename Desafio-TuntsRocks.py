import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configurando a autenticação para acessar a planilha
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc = gspread.authorize(credentials)

# Abrindo a planilha
spreadsheet_url = 'SUA_URL_DA_PLANILHA_COPIADA'
sh = gc.open_by_url(spreadsheet_url)

# Selecionando a folha de dados
worksheet = sh.get_worksheet(0)

# Obter dados da planilha
data = worksheet.get_all_records()

# Função para calcular a situação dos alunos
def calcular_situacao(aluno):
    media = (aluno['P1'] + aluno['P2'] + aluno['P3']) / 3
    faltas = aluno['Faltas']
    
    if faltas > 0.25 * aluno['Aulas']:
        return 'Reprovado por Falta'
    elif media < 5:
        return 'Reprovado por Nota'
    elif 5 <= media < 7:
        naf = max(0, 2 * (7 - media))
        naf = int(round(naf))
        worksheet.update_cell(aluno['Número'], 6, naf)
        return 'Exame Final'
    else:
        worksheet.update_cell(aluno['Número'], 6, 0)
        return 'Aprovado'

# Atualizar a situação na planilha
for aluno in data:
    situacao = calcular_situacao(aluno)
    worksheet.update_cell(aluno['Número'], 5, situacao)

print("Processo concluído. Verifique a planilha para os resultados.")
