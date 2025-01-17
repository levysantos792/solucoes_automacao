from robot.api.deco import keyword
from datetime import datetime, timedelta
from pytz import timezone

@keyword("Formatar Data SAP")
def formatar_data_para_sap(data: str):
    data = data.replace("/", ".")
    
    return data[0:10]

@keyword("Formatar Quantidade Para SAP")
def formatar_qtde_sap(qtde: str):
    return qtde[0:5].replace(".", ",")

@keyword("Formatar DateTime para SAP")
def formatar_datetime_para_SAP(data: datetime):
    return data.strftime("%d.%m.%Y")

@keyword("Formatar Data e Hora para SAP")
def formatar_data_hora_para_sap(data_hora: datetime) -> tuple:
    data_formatada = data_hora.strftime("%d.%m.%Y")
    hora_formatada = data_hora.strftime("%H:%M:%S")
    return data_formatada, hora_formatada

# Lista de feriados nacionais em 2025 no formato DD.MM.YYYY
FERIADOS = [
    "01.01.2025",  # Confraternização Universal (Ano Novo)
    "03.03.2025",  # Carnaval (Segunda-feira)
    "04.03.2025",  # Carnaval (Terça-feira)
    "21.04.2025",  # Tiradentes
    "01.05.2025",  # Dia do Trabalho
    "19.06.2025",  # Corpus Christi
    "07.09.2025",  # Independência do Brasil
    "12.10.2025",  # Nossa Senhora Aparecida
    "02.11.2025",  # Finados
    "15.11.2025",  # Proclamação da República
    "25.12.2025",  # Natal
]
FERIADOS = [datetime.strptime(data, "%d.%m.%Y").date() for data in FERIADOS]

@keyword("Próximo Dia Útil")
def proximo_dia_util(data: str = None) -> str:
    if data:
        date = datetime.strptime(data, "%d.%m.%Y").date()
    else:
        tz = timezone("America/Sao_Paulo")
        date = datetime.now(tz).date()

    while date.weekday() in (5, 6) or date in FERIADOS:
        date += timedelta(days=1)
    return date.strftime("%d.%m.%Y")

@keyword("Subtrair Horas")
def subtrair_horas(hora_inicial: str, hora_final: str) -> int:
    """
    Subtrai duas horas no formato HH:MM:SS e retorna a diferença em horas inteiras.

    :param hora_inicial: Hora inicial no formato "HH:MM:SS".
    :param hora_final: Hora final no formato "HH:MM:SS".
    :return: Diferença em horas inteiras.
    """
    formato = "%H:%M:%S"
    t1 = datetime.strptime(hora_inicial, formato)
    t2 = datetime.strptime(hora_final, formato)

    # Calcula a diferença em segundos
    diferenca_segundos = (t1 - t2).total_seconds()

    # Converte para horas inteiras
    diferenca_horas = abs(int(diferenca_segundos // 3600))
    return diferenca_horas


@keyword("Subtrair Horas Formato Completo")
def subtrair_horas_formato_completo(hora_inicial: str, hora_final: str) -> int:
#Subtrai duas horas no formato HH:MM:SS e retorna a diferença em horas e minutos.
# :param hora_inicial: Hora inicial no formato "HH:MM:SS".
# :param hora_final: Hora final no formato "HH:MM:SS".
# :return: Diferença no formato "HH:MM".
    formato = "%H:%M:%S"
    t1 = datetime.strptime(hora_inicial, formato)
    t2 = datetime.strptime(hora_final, formato)

    # Calcula a diferença em segundos
    diferenca_segundos = abs((t1 - t2).total_seconds())

    # Converte para horas e minutos
    diferenca_horas = int(diferenca_segundos // 3600)
    diferenca_minutos = int((diferenca_segundos % 3600) // 60)

    return f"{diferenca_horas:02}:{diferenca_minutos:02}"