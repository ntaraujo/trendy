import locale
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

def msg(text):
    print(f'#### TRENDY ---- {text}')

def simple_to_datetime(date):
    from datetime import datetime
    return datetime(int(date[4:]), int(date[2:4]), int(date[:2]))

def datetime_to_simple(date):
    return date.strftime('%d%m%Y')

def iso_to_simple(date):
    return ''.join(date.split('-')[::-1])

def capitalized_month(date):
    return date.strftime('%B').upper()

if __name__ == '__main__':
    print(capitalized_month(simple_to_datetime('03033445')))
