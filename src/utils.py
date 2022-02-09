from collections import deque
import locale
from gooey.python_bindings import argparse_to_json
import appdirs
import os
import sys
import configparser
import shutil

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

    if (options.get(action.dest) or {}).get('initial_value') is not None:
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


cache = configparser.ConfigParser()
cache['default'] = {}
cache_file_path = os.path.join(appdirs.user_cache_dir(), 'trendy_cache.ini')


def save_cache():
    with open(cache_file_path, 'w') as cache_file:
        cache.write(cache_file)


def load_cache():
    if os.path.exists(cache_file_path):
        cache.read(cache_file_path)


data_dir_path = appdirs.user_data_dir()


def global_path(local_path, basename=None):
    basename = basename or os.path.basename(local_path)
    path = os.path.join(data_dir_path, basename)
    if not os.path.exists(path):
        shutil.copy(local_path, path)
    return path


def compiled():
    return getattr(sys, 'frozen', False)

scheduled_call_queue = deque()
running_scheduled = False

def schedule(function, *args, **kwargs):
    if running_scheduled:
        function(*args, **kwargs)
    else:
        scheduled_call_queue.append((function, args, kwargs))

def run_scheduled():
    global running_scheduled
    running_scheduled = True
    for function, args, kwargs in scheduled_call_queue:
        function(*args, **kwargs)
    scheduled_call_queue.clear()
    running_scheduled = False

def scheduled(function):
    def new_function(*args, **kwargs):
        schedule(function, *args, **kwargs)
    return new_function


example_args = {
    'cods_cliente': """969611
1000560
611379
980420
""",
    'nomes_cliente': """""",
    'prevs_emb': """03122021
03012022
03022022
""",
    'senha_totvs': "SENHA_TOTVS",
    'dinamica': '/Users/macbookpro/Desktop/NOVA DINÂMICA.xlsx',
    # 'dinamica': 'C:\\Users\\Nathan\\Downloads\\NOVA DINÂMICA.xlsx',
    'nome_rede': 'diversa'
}

if __name__ == '__main__':
    print(msg("nova dinâmica"))
