from datetime import datetime, timedelta


def getTodaysDate():
    from datetime import datetime
    today = datetime.today()
    day = today.strftime('%d')
    print(day)


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


if __name__ == '__main__':
    print(get_tomorrows_date())
