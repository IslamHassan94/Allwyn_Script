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


def get_yesterday_date():
    yesterday = datetime.now() - timedelta(days=1)
    return yesterday.strftime('%d/%m/%Y')
