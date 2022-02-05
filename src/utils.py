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

def pasted_to_list(text):
    return text.splitlines()

from gooey.python_bindings import argparse_to_json

def action_to_json(action, widget, options):
    dropdown_types = {'Listbox', 'Dropdown', 'Counter'}
    if action.required:
        # Text fields get a default check that user input is present
        # and not just spaces, dropdown types get a simplified
        # is-it-present style check
        validator = ('user_input and not user_input.isspace()'
                     if widget not in dropdown_types
                     else 'user_input')
        error_msg = 'Esse campo é obrigatório'
    else:
        # not required; do nothing;
        validator = 'True'
        error_msg = ''

    base = argparse_to_json.merge(argparse_to_json.item_default, {
        'validator': {
            'type': 'ExpressionValidator',
            'test': validator,
            'message': error_msg
        },
    })

    if (options.get(action.dest) or {}).get('initial_value') != None:
        value = options[action.dest]['initial_value']
        options[action.dest]['initial_value'] = argparse_to_json.handle_initial_values(action, widget, value)
    default = argparse_to_json.handle_initial_values(action, widget, action.default)
    if default == argparse_to_json.argparse.SUPPRESS:
        default = None



    final_options = argparse_to_json.merge(base, options.get(action.dest) or {})
    argparse_to_json.validate_gooey_options(action, widget, final_options)

    return {
        'id': action.option_strings[0] if action.option_strings else action.dest,
        'type': widget,
        'cli_type': argparse_to_json.choose_cli_type(action),
        'required': action.required,
        'data': {
            'display_name': action.metavar or action.dest,
            'help': action.help,
            'required': action.required,
            'nargs': action.nargs or '',
            'commands': action.option_strings,
            'choices': list(map(str, action.choices)) if action.choices else [],
            'default': default,
            'dest': action.dest,
        },
        'options': final_options
    }

argparse_to_json.action_to_json = action_to_json

def retry(func, times=3, wait=1):
    from time import sleep
    def new_func(*args, **kwargs):
        for _ in range(times):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                msg(f"Erro suprimido. Tentando novamente. Mensagem:\n{e}")
                sleep(wait)
        return func(*args, **kwargs)
    return new_func

example_args = {
    'Códigos': """969611
1000560
611379
980420
""",
    'Nomes': """""",
    'Datas': """03122021
03012022
03022022
""",
    'Senha*': "SENHA_TOTVS",
    'Rede': 'diversa',
    'Dinâmica': 'C:\\Users\\Nathan\\Downloads\\NOVA DINÂMICA.xlsx'
    }

if __name__ == '__main__':
    print(capitalized_month(simple_to_datetime('03033445')))
