import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import sys

# --- CONFIGURAÇÕES GLOBAIS ---
CREDS_FILE = 'gcreds.json'
PLANILHA_NOME = "planilha-2026-1"
ABA_NOME = "Planilha1"

DIAS_DA_SEMANA = {
    '2': 'segunda-feira',
    '3': 'terça-feira',
    '4': 'quarta-feira',
    '5': 'quinta-feira',
    '6': 'sexta-feira'
}

HORARIOS_TURNO = {
    '1': ['08:00-08:50', '08:50-09:40', '10:00-10:50', '10:50-11:40', '11:40-12:30'],
    '2': ['13:30-14:20', '14:20-15:10', '15:10-16:00', '16:00-16:50', '17:10-18:00', '18:00-18:50'],
    '3': ['19:00-19:50', '19:50-20:40', '20:40-21:30', '21:30-22:20', '22:20-23:10']
}

def autenticar_e_obter_dados():
    """
    Autentica no Google Sheets usando Service Account e retorna o DataFrame limpo.
    """
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    try:
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        
        google_sheet = client.open(PLANILHA_NOME)
        aba = google_sheet.worksheet(ABA_NOME)
        
        dados_brutos = aba.get_all_values()

        if not dados_brutos or len(dados_brutos) < 2:
            raise ValueError("Planilha vazia ou sem dados.")

        cabecalho = dados_brutos[0]
        dados_linhas = dados_brutos[1:]

        df = pd.DataFrame(dados_linhas, columns=cabecalho)
        df = df.replace(['nan', 'None', 'NAN', 'NONE'], '')
        
        return df

    except Exception as e:
        print(f"ERRO CRÍTICO ao acessar Google Sheets: {e}")
        sys.exit(1)

def safe_int(val):
    try:
        return int(val)
    except (ValueError, TypeError):
        return None

def gerar_tabela_detalhes_html(df_detalhes):
    """Gera a tabela auxiliar com detalhes das disciplinas (Sala, Professor, etc)."""
    if df_detalhes.empty:
        return ""

    # Seleciona e renomeia colunas para exibição
    colunas_map = {
        'codigo': 'Código',
        'disciplina': 'Disciplina',
        'turma': 'Turma',
        'professor': 'Professor',
        'sala_exibicao': 'Sala / Campus'
    }
    
    # Filtra colunas que existem no DF
    cols_existentes = [c for c in colunas_map.keys() if c in df_detalhes.columns]
    df_show = df_detalhes[cols_existentes].copy()
    df_show = df_show.rename(columns=colunas_map)
    
    # Remove duplicatas baseadas em Código e Turma para a lista
    df_show = df_show.drop_duplicates(subset=['Código', 'Turma'])
    
    if 'Disciplina' in df_show.columns:
        df_show = df_show.sort_values(by='Disciplina')

    html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align: left; width: 100%; margin-top: 10px; margin-bottom: 40px; font-size: 0.9em;">'
    html += '<thead><tr style="background-color: #f2f2f2;">'
    for col in df_show.columns:
        html += f'<th style="padding: 8px;">{col}</th>'
    html += '</tr></thead><tbody>'

    for _, row in df_show.iterrows():
        html += '<tr>'
        for item in row:
            html += f'<td style="padding: 8px;">{item}</td>'
        html += '</tr>'

    html += '</tbody></table>'
    return html

def gerar_grade_horaria_semestre(lista_horarios, turnos_usados):
    """
    Gera a grade visual.
    lista_horarios: lista de listas [hora, dia, texto_celula]
    turnos_usados: set de turnos ('1', '2', '3') presentes neste semestre
    """
    horarios_completos = []
    
    # Define os horários a exibir baseado nos turnos usados naquelas disciplinas
    for turno in sorted(turnos_usados):
        if turno in HORARIOS_TURNO:
            horarios_completos.extend(HORARIOS_TURNO[turno])
            horarios_completos.append("---")

    if not horarios_completos:
        return "<p><i>Nenhum horário registrado.</i></p>"

    if horarios_completos[-1] == "---":
        horarios_completos.pop()

    df_grade = pd.DataFrame(lista_horarios, columns=['horário', 'dia', 'conteudo'])
    df_base = pd.DataFrame({'horário': horarios_completos})

    # Pivota a tabela para formato grade
    # aggfunc junta conteúdos se houver colisão de horário
    pivot = df_grade.pivot_table(
        index='horário', 
        columns='dia', 
        values='conteudo', 
        aggfunc=lambda x: '<br><hr style="margin:2px 0"><br>'.join(x) 
    ).fillna('')

    # Garante dias da semana na ordem correta
    dias_ordenados = ['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira']
    for dia in dias_ordenados:
        if dia not in pivot.columns:
            pivot[dia] = ''
    pivot = pivot[dias_ordenados]

    # Merge para garantir que todos os horários do turno apareçam (inclusive vazios e intervalos)
    tabela_final = pd.merge(df_base, pivot, on='horário', how='left').fillna('')

    # HTML
    html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align: center; width: 100%; margin-bottom: 10px;">'
    html += '<thead><tr style="background-color: #e0e0e0;"><th>Horário</th>'
    for dia in dias_ordenados:
        html += f'<th>{dia}</th>'
    html += '</tr></thead><tbody>'

    for _, row in tabela_final.iterrows():
        if row['horário'] == '---':
            html += f'<tr><td colspan="6" style="background-color: #cccccc; font-weight: bold; font-size: 0.8em;">Intervalo / Troca de Turno</td></tr>'
        else:
            html += '<tr>'
            html += f'<td style="white-space: nowrap; background-color: #f9f9f9; font-weight: bold;">{row["horário"]}</td>'
            for dia in dias_ordenados:
                conteudo = row[dia]
                style = ""
                if conteudo:
                    style = "background-color: #ffffff;" # Célula com aula
                html += f'<td style="{style}">{conteudo}</td>'
            html += '</tr>'
    
    html += '</tbody></table>'
    return html

def gerar_html_todas_tabelas():
    """
    Função principal chamada pelo Quarto.
    """
    df_planilha = autenticar_e_obter_dados()
    
    # Processa coluna semestre
    df_planilha['semestre_int'] = df_planilha['semestre'].apply(safe_int)

    # Estruturas para agrupamento
    # Chave: semestre_int. Valor: { 'horarios': [], 'detalhes': [], 'turnos': set() }
    grupos = {
        'impares': {}, # 1, 3, 5...
        'reofertas': {'horarios': [], 'detalhes': [], 'turnos': set()},
        'optativas': {'horarios': [], 'detalhes': [], 'turnos': set()}
    }

    # --- ITERAÇÃO E PROCESSAMENTO ---
    for _, row in df_planilha.iterrows():
        semestre = row['semestre_int']
        if pd.isna(semestre): continue

        codigo = row.get('codigo', '')
        disciplina = row.get('disciplina', '')
        turma = str(row.get('turma', '')).strip()
        sala_base = row.get('sala 1', '') 
        campus = row.get('campus', '')
        prof = row.get('professor', '')
        
        nome_grade = f"{disciplina} {turma}" if turma else disciplina
        sala_exibicao = f"{sala_base} ({campus})" if campus else sala_base

        # Dados para a tabela de detalhes (Lista abaixo da grade)
        linha_detalhe = {
            'codigo': codigo, 
            'disciplina': disciplina, 
            'turma': turma, 
            'professor': prof,
            'sala_exibicao': sala_exibicao
        }

        # Identifica o grupo
        grupo_alvo = None
        chave_semestre = None

        if semestre == 88:
            grupo_alvo = grupos['optativas']
        elif semestre < 10:
            if semestre % 2 != 0: # Ímpar
                if semestre not in grupos['impares']:
                    grupos['impares'][semestre] = {'horarios': [], 'detalhes': [], 'turnos': set()}
                grupo_alvo = grupos['impares'][semestre]
                chave_semestre = semestre
            else: # Par (Reoferta)
                grupo_alvo = grupos['reofertas']
        
        if not grupo_alvo: continue

        # Adiciona aos detalhes
        grupo_alvo['detalhes'].append(linha_detalhe)

        # Processa Horários (1 a 6) para a Grade
        for i in range(1, 7):
            val_horario = row.get(f'horario {i}')
            if val_horario and str(val_horario).strip():
                try:
                    cod_h = str(int(val_horario))
                    if len(cod_h) < 3: continue
                    
                    dia_cod = cod_h[0]
                    turno_cod = cod_h[1]
                    aula_cod = int(cod_h[2]) - 1

                    if dia_cod in DIAS_DA_SEMANA and turno_cod in HORARIOS_TURNO:
                        lista_h = HORARIOS_TURNO[turno_cod]
                        if 0 <= aula_cod < len(lista_h):
                            hora_real = lista_h[aula_cod]
                            dia_real = DIAS_DA_SEMANA[dia_cod]
                            
                            # FORMATAÇÃO DA CÉLULA (SOLICITADO: Negrito no Código, Nome abaixo, Sem sala)
                            texto_celula = f"<b>{codigo}</b><br><span style='font-size:0.85em'>{nome_grade}</span>"
                            
                            grupo_alvo['horarios'].append([hora_real, dia_real, texto_celula])
                            grupo_alvo['turnos'].add(turno_cod)
                except ValueError:
                    continue

    # --- GERAÇÃO DO HTML FINAL ---
    html = """
    <style>
        h1.titulo-semestre { color: #2c3e50; border-bottom: 2px solid #2c3e50; padding-bottom: 10px; margin-top: 50px; }
        h2.subtitulo { color: #7f8c8d; margin-top: 20px; font-size: 1.2em; }
    </style>
    """

    # 1. Semestres Ímpares (Ordenados)
    for sem in sorted(grupos['impares'].keys()):
        dados = grupos['impares'][sem]
        if not dados['horarios']: continue

        html += f"<h1 class='titulo-semestre'>{sem}º Semestre</h1>"
        
        # Grade Visual
        html += gerar_grade_horaria_semestre(dados['horarios'], dados['turnos'])
        
        # Tabela Detalhada
        df_det = pd.DataFrame(dados['detalhes'])
        html += "<h2 class='subtitulo'>Disciplinas, Salas e Professores</h2>"
        html += gerar_tabela_detalhes_html(df_det)

    # 2. Reofertas
    dados_reofertas = grupos['reofertas']
    if dados_reofertas['horarios']:
        html += "<h1 class='titulo-semestre'>Reofertas (Semestres Pares)</h1>"
        html += gerar_grade_horaria_semestre(dados_reofertas['horarios'], dados_reofertas['turnos'])
        
        if dados_reofertas['detalhes']:
            df_det = pd.DataFrame(dados_reofertas['detalhes'])
            html += "<h2 class='subtitulo'>Disciplinas, Salas e Professores</h2>"
            html += gerar_tabela_detalhes_html(df_det)

    # 3. Optativas
    dados_opt = grupos['optativas']
    if dados_opt['horarios']:
        html += "<h1 class='titulo-semestre'>Optativas</h1>"
        html += gerar_grade_horaria_semestre(dados_opt['horarios'], dados_opt['turnos'])
        
        if dados_opt['detalhes']:
            df_det = pd.DataFrame(dados_opt['detalhes'])
            html += "<h2 class='subtitulo'>Disciplinas, Salas e Professores</h2>"
            html += gerar_tabela_detalhes_html(df_det)

    return html

if __name__ == "__main__":
    # Teste local
    print(gerar_html_todas_tabelas())
