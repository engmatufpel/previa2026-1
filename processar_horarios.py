import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from IPython.display import HTML # Você ainda pode precisar disso se testar em um notebook

# --- DEFINIÇÕES GLOBAIS ---
dias_da_semana = {'2': 'segunda-feira', '3': 'terça-feira', '4': 'quarta-feira', '5': 'quinta-feira', '6': 'sexta-feira'}
horarios_turno = {
    '1': ['08:00-08:50', '08:50-09:40', '10:00-10:50', '10:50-11:40', '11:40-12:30'],
    '2': ['13:30-14:20', '14:20-15:10', '15:10-16:00', '16:00-16:50', '17:10-18:00', '18:00-18:50'],
    '3': ['19:00-19:50', '19:50-20:40', '20:40-21:30', '21:30-22:20', '22:20-23:10']
}
creds_file = 'gcreds.json' # O GitHub Action vai criar este arquivo 

# --- FUNÇÕES DE PROCESSAMENTO ---

def buscar_dados_planilha():
    """Autentica e busca os dados da planilha do Google Sheets."""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        creds = Credentials.from_service_account_file(creds_file, scopes=scopes)
        client = gspread.authorize(creds)
        google_sheet = client.open("planilha-2026-1")
        aba = google_sheet.worksheet("Planilha1")
        dados = aba.get_all_records()
        return pd.DataFrame(dados)
    
    except gspread.exceptions.SpreadsheetNotFound:
        print("ERRO: Planilha não encontrada. Verifique o nome da planilha no Google Sheets.")
        raise
    except gspread.exceptions.WorksheetNotFound:
        print("ERRO: Aba da planilha não encontrada. Verifique o nome da aba (ex: 'Planilha1').")
        raise
    except Exception as e:
        print(f"Ocorreu um erro ao acessar o Google Sheets: {e}")
        raise

# Função para gerar tabela HTML (Original)
def gerar_tabela_html(tabela_horarios, titulo):
    html = f'<h3 style="text-align: center; font-size: 1.5em; font-weight: bold;">{titulo}</h3>'
    html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align: center; width: 100%;">'
    html += '<thead><tr><th>Horário</th><th>Segunda-feira</th><th>Terça-feira</th><th>Quarta-feira</th><th>Quinta-feira</th><th>Sexta-feira</th></tr></thead><tbody>'

    for index, row in tabela_horarios.iterrows():
        if row['horário'] == '---':
            html += '<tr><td colspan="6" style="background-color: #e0e0e0; font-weight: bold;">Intervalo entre Turnos</td></tr>'
        else:
            html += '<tr>' + ''.join(f'<td>{item}</td>' for item in row) + '</tr>'

    html += '</tbody></table>'
    return html

# Função para gerar e exibir tabela de horários (Original)
def gerar_tabela_horarios(df, titulo):
    horarios = pd.DataFrame(df, columns=['horário', 'dia', 'disciplina'])
    tabela_pivot = horarios.pivot_table(index='horário', columns='dia', values='disciplina', aggfunc=lambda x: '<br>'.join(x)).fillna('')

    # Garante todos os dias da semana
    for dia in dias_da_semana.values():
        if dia not in tabela_pivot.columns:
            tabela_pivot[dia] = ''
    tabela_pivot = tabela_pivot[['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira']]

    # Horários completos com intervalos
    turnos_usados = set()
    for linha in df:
        if linha[0] in horarios_turno['1']: turnos_usados.add('1')
        elif linha[0] in horarios_turno['2']: turnos_usados.add('2')
        elif linha[0] in horarios_turno['3']: turnos_usados.add('3')

    horarios_completos = []
    for i, turno in enumerate(sorted(turnos_usados)):
        horarios_completos.extend(horarios_turno[turno])
        if i < len(turnos_usados) - 1:
            horarios_completos.append('---')

    tabela_final = pd.DataFrame({'horário': horarios_completos})
    tabela_final = pd.merge(tabela_final, tabela_pivot, on='horário', how='left').fillna('')

    display(HTML(gerar_tabela_html(tabela_final, titulo)))


# --- INÍCIO DA SEÇÃO NOVA ---

# Função para converter 'semestre' para int de forma segura
def safe_int(val):
    try:
        return int(val)
    except (ValueError, TypeError):
        return None

# Adiciona a coluna 'semestre_int' para filtros confiáveis
planilha['semestre_int'] = planilha['semestre'].apply(safe_int)


# Nova função para gerar a tabela auxiliar de disciplinas (COM CORREÇÃO do Warning)
def gerar_tabela_auxiliar(df_semestre_bruto, colunas_planilha):
    """
    Gera uma tabela HTML auxiliar com os detalhes das disciplinas.
    - df_semestre_bruto: O DataFrame *original* filtrado para o semestre.
    - colunas_planilha: Lista de todas as colunas do DataFrame original.
    """

    # Colunas desejadas e seus novos nomes
    colunas_desejadas = {
        'codigo': 'Código',
        'disciplina': 'Disciplina',
        'turma': 'Turma',
        'professor': 'Professor responsável',
        'departamento': 'Departamento',
        'sala 1': 'Sala',
        'campus': 'Campus'
    }

    # Filtra o dicionário para incluir apenas colunas que REALMENTE existem na planilha
    colunas_para_selecionar = {
        col_orig: col_novo
        for col_orig, col_novo in colunas_desejadas.items()
        if col_orig in colunas_planilha
    }

    if not colunas_para_selecionar:
        # Nenhuma das colunas desejadas foi encontrada, não faz nada
        return

    # Seleciona as colunas existentes
    df_aux = df_semestre_bruto[colunas_para_selecionar.keys()].copy()

    # Remove duplicatas (para ter uma lista limpa de disciplinas/turmas)
    # Usa 'codigo' e 'turma' como chave, se existirem
    subset_duplicatas = ['codigo', 'turma'] if 'codigo' in df_aux.columns and 'turma' in df_aux.columns else list(df_aux.columns)
    df_aux = df_aux.drop_duplicates(subset=subset_duplicatas)

    # Renomeia as colunas para os nomes amigáveis
    df_aux = df_aux.rename(columns=colunas_para_selecionar)

    # Ordena por disciplina, se a coluna "Disciplina" existir
    if 'Disciplina' in df_aux.columns:
        df_aux = df_aux.sort_values(by='Disciplina')

    # Limpa valores NaN (nulos) e converte tudo para string para evitar o warning
    df_aux = df_aux.fillna('---').astype(str)

    # Converte para HTML com estilo
    html_aux = df_aux.to_html(index=False, justify='left', border=1)

    # Aplica estilo CSS para ficar parecido com a tabela de horários
    html_aux = html_aux.replace('<table border="1" class="dataframe">',
                              '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align: left; width: 100%; margin-top: 10px; margin-bottom: 40px;">') # Adiciona margem
    html_aux = html_aux.replace('<th>', '<th style="background-color: #f2f2f2; text-align: left; padding: 5px;">')
    html_aux = html_aux.replace('<td>', '<td style="padding: 5px;">')

    display(HTML(html_aux))

# --- FIM DA SEÇÃO NOVA ---


# --- MODIFICAÇÃO LÓGICA 1 ---
# Dicionários para guardar os dados
semestres_impares = {} # <- Renomeado de 'semestres_pares'
reofertas = []
optativas = []
# --- FIM MODIFICAÇÃO LÓGICA 1 ---


# Processamento das linhas da planilha
for _, row in planilha.iterrows():
    # Usamos a coluna 'semestre_int' que criamos
    semestre = row['semestre_int']

    if pd.isna(semestre):
        continue

    codigo = row.get('codigo', '')
    nome = row.get('disciplina', '')
    turma = row.get('turma', '')
    prof = row.get('professor', '') # O professor ainda é lido, mas não usado no 'texto'

    horarios_disciplina = []
    for i in range(1, 7):
        horario_col = f'horario {i}'
        sala_col = f'sala {i}'

        horario_val = row.get(horario_col)

        if pd.notna(horario_val) and str(horario_val).strip() != '':
            try:
                cod_horario = str(int(horario_val))
            except ValueError:
                continue

            dia = dias_da_semana.get(cod_horario[0], '')
            turno = cod_horario[1]
            aula_idx = int(cod_horario[2]) - 1

            if turno in horarios_turno and 0 <= aula_idx < len(horarios_turno[turno]):
                hora = horarios_turno[turno][aula_idx]
            else:
                continue

            sala = row.get(sala_col, '')

            if pd.isna(sala) or str(sala).strip().lower() == 'nan' or str(sala).strip() == '':
                sala = 'sala indefinida'
            else:
                sala = f'sala {sala}'

            # Texto (sem professor, como solicitado anteriormente)
            texto = f'{codigo} - {nome} {turma} - {sala}'

            horarios_disciplina.append([hora, dia, texto])

    if not horarios_disciplina:
        continue

    # --- MODIFICAÇÃO LÓGICA 2 ---
    # Classificação por tipo (LÓGICA INVERTIDA)
    if semestre % 2 == 1 and semestre < 10: # <-- Captura ÍMPARES
        if semestre not in semestres_impares:
            semestres_impares[semestre] = []
        semestres_impares[semestre].extend(horarios_disciplina)
    elif semestre % 2 == 0 and semestre < 10: # <-- Captura PARES como Reofertas
        reofertas.extend(horarios_disciplina)
    elif semestre == 88:
        optativas.extend(horarios_disciplina)
    # --- FIM MODIFICAÇÃO LÓGICA 2 ---


# --- MODIFICAÇÃO LÓGICA 3 ---
# Geração das tabelas (Lógica invertida)

# Pega os nomes das colunas da planilha UMA VEZ
colunas_da_planilha = list(planilha.columns)

# Loop principal agora itera sobre os SEMESTRES ÍMPARES
for semestre, lista in sorted(semestres_impares.items()):
    # 1. Gera a tabela de horários (como antes)
    gerar_tabela_horarios(lista, f'{semestre}º semestre')

    # 2. Filtra o DataFrame BRUTO para o semestre atual
    df_bruto_semestre = planilha[planilha['semestre_int'] == semestre]

    # 3. Gera a tabela auxiliar
    gerar_tabela_auxiliar(df_bruto_semestre, colunas_da_planilha)


if reofertas:
    # 1. Gera a tabela de horários (como antes)
    gerar_tabela_horarios(reofertas, 'Reofertas')

    # 2. Filtra o DataFrame BRUTO para as reofertas (agora SEMESTRES PARES)
    df_bruto_reofertas = planilha[
        planilha['semestre_int'].apply(lambda x: x is not None and x % 2 == 0 and x < 10)
    ]

    # 3. Gera a tabela auxiliar
    gerar_tabela_auxiliar(df_bruto_reofertas, colunas_da_planilha)


if optativas:
    # 1. Gera a tabela de horários (como antes)
    gerar_tabela_horarios(optativas, 'Optativas')

    # 2. Filtra o DataFrame BRUTO para as optativas
    df_bruto_optativas = planilha[planilha['semestre_int'] == 88]

    # 3. Gera a tabela auxiliar
    gerar_tabela_auxiliar(df_bruto_optativas, colunas_da_planilha)

# --- FIM MODIFICAÇÃO LÓGICA 3 ---
