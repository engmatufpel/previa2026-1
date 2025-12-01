import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- DEFINIÇÕES GLOBAIS ---
dias_da_semana = {'2': 'segunda-feira', '3': 'terça-feira', '4': 'quarta-feira', '5': 'quinta-feira', '6': 'sexta-feira'}
horarios_turno = {
    '1': ['08:00-08:50', '08:50-09:40', '10:00-10:50', '10:50-11:40', '11:40-12:30'],
    '2': ['13:30-14:20', '14:20-15:10', '15:10-16:00', '16:00-16:50', '17:10-18:00', '18:00-18:50'],
    '3': ['19:00-19:50', '19:50-20:40', '20:40-21:30', '21:30-22:20', '22:20-23:10']
}
creds_file = 'gcreds.json'

# --- FUNÇÕES AUXILIARES ---

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
    except Exception as e:
        print(f"Erro ao acessar Google Sheets: {e}")
        raise

def gerar_tabela_html_string(tabela_horarios, titulo):
    """Gera o HTML puro de uma tabela de horários."""
    html = f'<h3 style="text-align: center; font-size: 1.5em; font-weight: bold; margin-top: 30px;">{titulo}</h3>'
    html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align: center; width: 100%;">'
    html += '<thead><tr><th>Horário</th><th>Segunda-feira</th><th>Terça-feira</th><th>Quarta-feira</th><th>Quinta-feira</th><th>Sexta-feira</th></tr></thead><tbody>'

    for index, row in tabela_horarios.iterrows():
        if row['horário'] == '---':
            html += '<tr><td colspan="6" style="background-color: #e0e0e0; font-weight: bold;">Intervalo entre Turnos</td></tr>'
        else:
            html += '<tr>' + ''.join(f'<td>{item}</td>' for item in row) + '</tr>'

    html += '</tbody></table>'
    return html

def processar_tabela_horarios(df, titulo):
    """Processa os dados e retorna a string HTML da tabela de horários."""
    if not df: return ""
    
    horarios = pd.DataFrame(df, columns=['horário', 'dia', 'disciplina'])
    tabela_pivot = horarios.pivot_table(index='horário', columns='dia', values='disciplina', aggfunc=lambda x: '<br>'.join(x)).fillna('')

    for dia in dias_da_semana.values():
        if dia not in tabela_pivot.columns:
            tabela_pivot[dia] = ''
    tabela_pivot = tabela_pivot[['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira']]

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

    return gerar_tabela_html_string(tabela_final, titulo)

def gerar_tabela_auxiliar_string(df_semestre_bruto, colunas_planilha):
    """Gera e retorna a string HTML da tabela auxiliar."""
    colunas_desejadas = {
        'codigo': 'Código',
        'disciplina': 'Disciplina',
        'turma': 'Turma',
        'professor': 'Professor responsável',
        'departamento': 'Departamento',
        'sala 1': 'Sala',
        'campus': 'Campus'
    }

    colunas_para_selecionar = {c: n for c, n in colunas_desejadas.items() if c in colunas_planilha}
    if not colunas_para_selecionar: return ""

    df_aux = df_semestre_bruto[colunas_para_selecionar.keys()].copy()
    
    subset_duplicatas = ['codigo', 'turma'] if 'codigo' in df_aux.columns and 'turma' in df_aux.columns else list(df_aux.columns)
    df_aux = df_aux.drop_duplicates(subset=subset_duplicatas)
    df_aux = df_aux.rename(columns=colunas_para_selecionar)

    if 'Disciplina' in df_aux.columns:
        df_aux = df_aux.sort_values(by='Disciplina')

    df_aux = df_aux.fillna('---').astype(str)
    
    html_aux = df_aux.to_html(index=False, justify='left', border=1)
    
    # Estilização CSS inline para garantir renderização
    html_aux = html_aux.replace('<table border="1" class="dataframe">', 
        '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align: left; width: 100%; margin-top: 10px; margin-bottom: 40px;">')
    html_aux = html_aux.replace('<th>', '<th style="background-color: #f2f2f2; text-align: left; padding: 5px;">')
    html_aux = html_aux.replace('<td>', '<td style="padding: 5px;">')
    
    return html_aux

def safe_int(val):
    try: return int(val)
    except (ValueError, TypeError): return None

# --- FUNÇÃO PRINCIPAL CHAMADA PELO QUARTO ---

def gerar_html_todas_tabelas():
    """Função mestre que orquestra tudo e retorna uma string HTML gigante."""
    
    # 1. Carrega a planilha (CORREÇÃO: agora chamamos a função explicitamente)
    planilha = buscar_dados_planilha()
    
    # 2. Pré-processamento
    planilha['semestre_int'] = planilha['semestre'].apply(safe_int)
    colunas_da_planilha = list(planilha.columns)

    # Estruturas de dados
    semestres_impares = {} 
    reofertas = []
    optativas = []

    # 3. Lógica de Classificação
    for _, row in planilha.iterrows():
        semestre = row['semestre_int']
        if pd.isna(semestre): continue

        codigo = row.get('codigo', '')
        nome = row.get('disciplina', '')
        turma = row.get('turma', '')
        
        # Processa horários
        horarios_disciplina = []
        for i in range(1, 7):
            horario_val = row.get(f'horario {i}')
            if pd.notna(horario_val) and str(horario_val).strip() != '':
                try:
                    cod_horario = str(int(horario_val))
                    dia = dias_da_semana.get(cod_horario[0], '')
                    turno = cod_horario[1]
                    aula_idx = int(cod_horario[2]) - 1
                    
                    if turno in horarios_turno and 0 <= aula_idx < len(horarios_turno[turno]):
                        hora = horarios_turno[turno][aula_idx]
                        sala = row.get(f'sala {i}', '')
                        sala = 'sala indefinida' if (pd.isna(sala) or str(sala).strip().lower() == 'nan' or str(sala).strip() == '') else f'sala {sala}'
                        texto = f'{codigo} - {nome} {turma} - {sala}'
                        horarios_disciplina.append([hora, dia, texto])
                except ValueError:
                    continue

        if not horarios_disciplina: continue

        # Distribuição nos grupos
        if semestre % 2 == 1 and semestre < 10:
            if semestre not in semestres_impares: semestres_impares[semestre] = []
            semestres_impares[semestre].extend(horarios_disciplina)
        elif semestre % 2 == 0 and semestre < 10:
            reofertas.extend(horarios_disciplina)
        elif semestre == 88:
            optativas.extend(horarios_disciplina)

    # 4. Geração do HTML (Acumulando em uma string)
    html_acumulado = ""

    # Semestres Ímpares
    for semestre, lista in sorted(semestres_impares.items()):
        html_acumulado += processar_tabela_horarios(lista, f'{semestre}º semestre')
        df_bruto = planilha[planilha['semestre_int'] == semestre]
        html_acumulado += gerar_tabela_auxiliar_string(df_bruto, colunas_da_planilha)

    # Reofertas
    if reofertas:
        html_acumulado += processar_tabela_horarios(reofertas, 'Reofertas')
        df_bruto = planilha[planilha['semestre_int'].apply(lambda x: x is not None and x % 2 == 0 and x < 10)]
        html_acumulado += gerar_tabela_auxiliar_string(df_bruto, colunas_da_planilha)

    # Optativas
    if optativas:
        html_acumulado += processar_tabela_horarios(optativas, 'Optativas')
        df_bruto = planilha[planilha['semestre_int'] == 88]
        html_acumulado += gerar_tabela_auxiliar_string(df_bruto, colunas_da_planilha)

    return html_acumulado
