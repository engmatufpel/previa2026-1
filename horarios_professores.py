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
        # Autenticação via arquivo JSON (padrão para GitHub Actions)
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        
        google_sheet = client.open(PLANILHA_NOME)
        aba = google_sheet.worksheet(ABA_NOME)
        
        # Lê tudo como string para evitar erros de tipagem
        dados_brutos = aba.get_all_values()

        if not dados_brutos or len(dados_brutos) < 2:
            raise ValueError("Planilha vazia ou sem dados.")

        cabecalho = dados_brutos[0]
        dados_linhas = dados_brutos[1:]

        # Cria DataFrame forçando string em tudo
        df = pd.DataFrame(dados_linhas, columns=cabecalho)
        
        # Limpeza de valores nulos literais
        df = df.replace(['nan', 'None', 'NAN', 'NONE'], '')
        
        return df

    except Exception as e:
        print(f"ERRO CRÍTICO ao acessar Google Sheets: {e}")
        sys.exit(1) # Encerra o script com erro para o GitHub Actions pegar

def gerar_tabela_html_string(df, titulo_tabela):
    """Gera o HTML de uma tabela genérica."""
    html = f'<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align: center; width: 100%; margin-bottom: 20px;">'
    
    # Cabeçalho
    html += '<thead><tr style="background-color: #f2f2f2;">'
    for col in df.columns:
        html += f'<th>{col}</th>'
    html += '</tr></thead><tbody>'

    # Linhas
    for _, row in df.iterrows():
        html += '<tr>'
        for item in row:
            html += f'<td>{item}</td>'
        html += '</tr>'

    html += '</tbody></table>'
    return html

def gerar_grade_horaria_html(horarios_df, professor, turnos_professor):
    """Gera a grade horária formatada (com intervalos) para o professor."""
    horarios_completos = []

    # Coleta todos os horários possíveis baseados nos turnos que o professor trabalha
    for turno in sorted(turnos_professor):
        if turno in HORARIOS_TURNO:
            horarios_completos.extend(HORARIOS_TURNO[turno])
            horarios_completos.append("---") # Marcador de intervalo

    if not horarios_completos:
        return "<p><i>Nenhum horário cadastrado para este professor.</i></p>"

    # Remove o último separador se existir
    if horarios_completos and horarios_completos[-1] == "---":
        horarios_completos.pop()

    horarios_completos_df = pd.DataFrame({'horário': horarios_completos})
    
    # Pivota a tabela: Linhas = Horários, Colunas = Dias
    horarios_pivot = horarios_df.pivot_table(
        index='horário', 
        columns='dia', 
        values='disciplina', 
        aggfunc=lambda x: '<br>'.join(x)
    ).fillna('')

    # Garante que todos os dias da semana apareçam
    dias_ordenados = ['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira']
    for dia in dias_ordenados:
        if dia not in horarios_pivot.columns:
            horarios_pivot[dia] = ''

    horarios_pivot = horarios_pivot[dias_ordenados]

    # Junta com a lista completa de horários para mostrar buracos e intervalos
    tabela_final = pd.merge(horarios_completos_df, horarios_pivot, on='horário', how='left').fillna('')
    
    # Formatação visual para o intervalo
    html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align: center; width: 100%; margin-bottom: 40px;">'
    html += '<thead><tr style="background-color: #e0e0e0;"><th>Horário</th>'
    for dia in dias_ordenados:
        html += f'<th>{dia}</th>'
    html += '</tr></thead><tbody>'

    for _, row in tabela_final.iterrows():
        if row['horário'] == '---':
            html += f'<tr><td colspan="6" style="background-color: #cccccc; font-weight: bold; font-size: 0.8em;">Intervalo / Troca de Turno</td></tr>'
        else:
            html += '<tr>'
            html += f'<td>{row["horário"]}</td>'
            for dia in dias_ordenados:
                html += f'<td>{row[dia]}</td>'
            html += '</tr>'
    
    html += '</tbody></table>'
    return html

def processar_dados_e_gerar_html():
    """
    Função principal que orquestra a leitura, processamento e geração do HTML.
    Retorna uma string contendo todo o HTML.
    """
    #print("Iniciando processamento de dados...")
    df_planilha = autenticar_e_obter_dados()
    
    tabelas_professores = {}
    horarios_professores = {}
    turnos_professores = {}

    # --- PROCESSAMENTO ---
    #print(f"Processando {len(df_planilha)} linhas...")
    for _, row in df_planilha.iterrows():
        # Extração segura de dados
        codigo = row.get('codigo', '')
        disciplina = row.get('disciplina', '')
        professor = row.get('professor', '')
        turma = str(row.get('turma', '')).strip()
        
        nome_exibicao = f"{disciplina} {turma}" if turma else disciplina
        
        # Créditos
        try:
            creditos = int(row.get('creditos', 0))
        except (ValueError, TypeError):
            creditos = 0
            
        # Alunos
        alunos_val = row.get('alunos')
        if not alunos_val or str(alunos_val).strip() == '':
            num_alunos = 0
        else:
            # Conta vírgulas para estimar número de alunos (lógica original)
            matriculas = str(alunos_val).strip().split(',')
            num_alunos = len([m for m in matriculas if m.strip()])

        # Monta objeto da disciplina
        dados_disciplina = {
            'Código': codigo,
            'Disciplina': nome_exibicao,
            'Créditos': creditos,
            'Alunos (est.)': num_alunos
        }

        # Agrupa por professor
        if professor not in tabelas_professores:
            tabelas_professores[professor] = []
        tabelas_professores[professor].append(dados_disciplina)

        # Processa Horários (1 a 6)
        for i in range(1, 7):
            val_horario = row.get(f'horario {i}')
            
            if val_horario and str(val_horario).strip():
                try:
                    cod_h = str(int(val_horario)) # Remove .0 se existir
                    
                    if len(cod_h) < 3: continue 

                    dia_cod = cod_h[0]
                    turno_cod = cod_h[1]
                    aula_cod = int(cod_h[2]) - 1 # Index 0-based

                    if dia_cod in DIAS_DA_SEMANA and turno_cod in HORARIOS_TURNO:
                        lista_horarios = HORARIOS_TURNO[turno_cod]
                        if 0 <= aula_cod < len(lista_horarios):
                            hora_real = lista_horarios[aula_cod]
                            dia_real = DIAS_DA_SEMANA[dia_cod]
                            
                            texto_celula = f"<b>{codigo}</b><br>{nome_exibicao}<br><span style='font-size:0.8em'>{row.get(f'sala {i}', 'Sala Indef.')}</span>"
                            
                            item_grade = [hora_real, dia_real, texto_celula]

                            if professor not in horarios_professores:
                                horarios_professores[professor] = []
                            horarios_professores[professor].append(item_grade)

                            if professor not in turnos_professores:
                                turnos_professores[professor] = set()
                            turnos_professores[professor].add(turno_cod)

                except ValueError:
                    continue

    # --- GERAÇÃO DO HTML ---
    #print("Gerando HTML final...")
    html_acumulado = """
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; }
            h1 { color: #333; border-bottom: 2px solid #333; padding-bottom: 10px; margin-top: 50px; }
            h2 { color: #555; margin-top: 20px; }
            table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
            th { background-color: #f2f2f2; }
            tr:nth-child(even) { background-color: #f9f9f9; }
        </style>
    </head>
    <body>
    """

    professores_ordenados = sorted([p for p in tabelas_professores.keys() if p.strip()])

    for prof in professores_ordenados:
        html_acumulado += f"<h1 id='{prof.replace(' ', '_')}'>{prof}</h1>"
        
        # 1. Tabela de Carga Horária / Disciplinas
        df_disc = pd.DataFrame(tabelas_professores[prof])
        
        # Linha de Total
        total_creditos = df_disc['Créditos'].sum()
        df_total = pd.DataFrame([{
            'Código': '', 
            'Disciplina': '<b>TOTAL DE CRÉDITOS</b>', 
            'Créditos': f'<b>{total_creditos}</b>', 
            'Alunos (est.)': ''
        }])
        
        # Concatena e converte para HTML
        df_final = pd.concat([df_disc, df_total], ignore_index=True)
        html_acumulado += "<h2>Disciplinas e Carga</h2>"
        html_acumulado += gerar_tabela_html_string(df_final, "")

        # 2. Grade Horária
        html_acumulado += "<h2>Grade Horária</h2>"
        if prof in horarios_professores:
            df_horarios = pd.DataFrame(horarios_professores[prof], columns=['horário', 'dia', 'disciplina'])
            turnos = turnos_professores.get(prof, set())
            html_acumulado += gerar_grade_horaria_html(df_horarios, prof, turnos)
        else:
            html_acumulado += "<p><i>Sem horários alocados.</i></p>"
        
        html_acumulado += "<hr>"

    html_acumulado += "</body></html>"
    return html_acumulado

# --- BLOCO DE EXECUÇÃO PRINCIPAL ---
if __name__ == "__main__":
    # Quando rodado diretamente (python gerar_tabelas_professores.py), ele cria um arquivo.
    html_final = processar_dados_e_gerar_html()
    
    nome_arquivo = "relatorio_professores.html"
    with open(nome_arquivo, "w", encoding="utf-8") as f:
        f.write(html_final)
    
    print(f"Sucesso! Arquivo '{nome_arquivo}' gerado.")
