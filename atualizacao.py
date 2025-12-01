# @title
from datetime import datetime
from zoneinfo import ZoneInfo # Import necessário para lidar com fusos horários

# Função para exibir a data, hora e dia da semana da última execução
def exibir_data_hora_execucao():
    # Define o fuso horário de Brasília
    fuso_brasil = ZoneInfo('America/Sao_Paulo')
    
    # Obtém o momento atual já com o fuso horário correto
    agora = datetime.now(fuso_brasil)

    # Dicionário para mapear o número do dia da semana com o nome em português
    dias_semana = {
        0: 'Segunda-feira',
        1: 'Terça-feira',
        2: 'Quarta-feira',
        3: 'Quinta-feira',
        4: 'Sexta-feira',
        5: 'Sábado',
        6: 'Domingo'
    }

    dia_semana = dias_semana[agora.weekday()]  # Obter o nome do dia da semana
    print(f"Última atualização: {agora.strftime('%d/%m/%Y %H:%M:%S')} ({dia_semana})")

# Exibir a data, hora e dia da semana no output
exibir_data_hora_execucao()
