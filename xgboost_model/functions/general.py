import datetime as dt
import getpass

def get_greeting():
    username = getpass.getuser()
    now=dt.datetime.now()
    current_hour = now.hour
    current_day = now.strftime('%A')

    if current_hour < 12:
        greeting = 'Good morning'
    elif 12 <= current_hour < 17:
        greeting = 'Good afternoon'
    else:
        greeting = 'Good evening'
    return f"{greeting} {username}, happy {current_day}"