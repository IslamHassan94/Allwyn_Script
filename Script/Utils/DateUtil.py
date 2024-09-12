from datetime import datetime, timedelta


def getTodaysDate():
    from datetime import datetime
    today = datetime.today()
    day = today.strftime("%d/%m/%y")
    return str(day)


def getTodaysDateInSerialFormat():
    from datetime import datetime
    today = datetime.today()
    day = today.strftime('%Y%m%d')
    return str(day)


def getCiclo():
    if 1 <= int(getTodaysDate()) <= 8:
        return 'Ciclo 1'
    elif 8 < int(getTodaysDate()) <= 15:
        return 'Ciclo 8'
    elif 15 < int(getTodaysDate()) <= 22:
        return 'Ciclo 15'
    else:
        return 'Ciclo 22'


def get_tomorrows_date():
    # Get today's date
    today = datetime.today()
    # Calculate tomorrow's date
    tomorrow = today + timedelta(days=1)
    # Format the date as 'day-month-year'
    formatted_date = tomorrow.strftime('%d-%b-%y')
    return formatted_date


def get_yesterday_date():
    yesterday = datetime.now() - timedelta(days=1)
    return yesterday.strftime('%d/%m/%Y')
