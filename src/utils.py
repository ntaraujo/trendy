import configparser
import locale
import os
import shutil
import sys
from collections import deque
from filecmp import cmp as filecmp
from time import sleep

import appdirs
import chardet
from gooey import local_resource_path
from gooey.python_bindings import argparse_to_json

default_open = open


def encoding_of(file_path):
    if not os.path.exists(file_path):
        return 'utf-8'

    with default_open(file_path, 'rb') as rawdata:
        result = chardet.detect(rawdata.read(100000))['encoding']
    return 'utf-8' if result == 'ascii' else result


def open(*args, **kwargs):
    if len(args) < 4 and 'encoding' not in kwargs:
        kwargs['encoding'] = encoding_of(args[0] if args else kwargs['file'])

    return default_open(*args, **kwargs)


app_dir_path = os.path.join(appdirs.user_data_dir(), "Trendy")

if not os.path.exists(app_dir_path):
    os.mkdir(app_dir_path)

cache_dir_path = os.path.join(app_dir_path, "cache")

if not os.path.exists(cache_dir_path):
    os.mkdir(cache_dir_path)


def system_remove(what):
    try:
        os.remove(what)
    except:
        shutil.rmtree(what, ignore_errors=True)


log_file = open(os.path.join(app_dir_path, 'last-log.txt'), 'w', encoding='utf-8')

locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')


def msg(text):
    ascii_compatible = text.encode('ascii', 'replace').decode()
    print(f'#### TRENDY ---- {ascii_compatible}')
    log_file.write(text + '\n')


_current_progress = 0
_total_progress = 0


def total_progress(number=None):
    global _total_progress

    if number is None:
        global _current_progress

        _current_progress = 0
        _total_progress = 0
    else:
        _total_progress += number

    msg(f'Total de iterações esperadas: {_total_progress}')


def _progress(amount):
    global _current_progress
    _current_progress += amount
    print(f'Progresso: {_current_progress}/{_total_progress}')


def progress(func=None, amount=1):
    def new_func(*args, **kwargs):
        _progress(amount)
        return func(*args, **kwargs)

    return new_func if func else _progress(amount)


debugger_active = getattr(sys, 'gettrace', lambda: None)() is not None
if debugger_active:
    msg("Rodando em modo de depuração")


def simple_to_datetime(date):
    from datetime import datetime
    return datetime(int(date[4:]), int(date[2:4]), int(date[:2]))


def datetime_to_simple(date):
    return date.strftime('%d%m%Y')


def iso_to_simple(date):
    return ''.join(date.split('-')[::-1])


pt_br_months = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro"
]


def capitalized_month(date):
    return pt_br_months[date.month - 1].upper()


def pasted_to_list(text):
    return [line.strip() for line in text.splitlines() if line.strip()]


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


def retry(func, exc=None, times=3, wait=1, on_debug=False):
    if not on_debug and debugger_active:
        return func

    if exc is None:
        exc = Exception

    def new_func(*args, **kwargs):
        for _ in range(times):
            try:
                return func(*args, **kwargs)
            except exc as e:
                msg(f"Erro suprimido. Tentando novamente\n{type(e).__name__}\n{e}")
                sleep(wait)
        return func(*args, **kwargs)

    return new_func


data = configparser.ConfigParser()
data['default'] = {}
data_file_path = os.path.join(app_dir_path, 'data.ini')


def save_data():
    msg(f'Salvando dados em "{data_file_path}"')

    with open(data_file_path, 'w', encoding='utf-8') as data_file:
        data.write(data_file)


def load_data():
    msg(f'Carregando dados de "{data_file_path}"')

    if os.path.exists(data_file_path):
        data.read(data_file_path, encoding='utf-8')


def global_path(local_path, basename=None):
    local_path = local_resource_path(local_path)
    basename = basename or os.path.basename(local_path)
    path = os.path.join(app_dir_path, basename)
    if os.path.exists(path) and not filecmp(local_path, path):
        msg(f'Atualizando arquivo em "{path}"')
        os.remove(path)
    if not os.path.exists(path):
        shutil.copy(local_path, path)
    return path


compiled = getattr(sys, 'frozen', False)

scheduled_call_queue = deque()
running_scheduled = False


def schedule(function, *args, **kwargs):
    if running_scheduled or not compiled:
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


def common_start(*strings):
    def _iter():
        for letters in zip(*strings):
            letters_iter = iter(letters)
            first = next(letters_iter)
            if all(first == others for others in letters_iter):
                yield first
            else:
                return

    return ''.join(_iter())


def system_open(what):
    import subprocess
    import platform

    if platform.system() == 'Darwin':
        subprocess.call(('open', what))
    elif platform.system() == 'Windows':
        os.startfile(what)
    else:
        subprocess.call(('xdg-open', what))


class ExampleArgs:
    cods_cliente = """969611
1000560
611379
980420
"""
    nomes_cliente = """"""
    prevs_emb = """03032022
"""
    senha_totvs = "SENHA_TOTVS"
    dinamica = '/Users/macbookpro/Desktop/NOVA DINÂMICA.xlsx'
    # 'dinamica = 'C:\\Users\\Nathan\\Downloads\\NOVA DINÂMICA.xlsx'
    nome_rede = 'kammu'
    cac = '/Users/macbookpro/Desktop/dev/trendy/examples/cac.csv'
    oc = '4500655584'
    romaneio_inicio = '/Users/macbookpro/Desktop/dev/trendy/examples/romaneio-inicio.xls'
    arquivo_oc = '/Users/macbookpro/Desktop/dev/trendy/examples/arquivo-oc.csv'
    arquivo_wgpd513 = '/Users/macbookpro/Desktop/dev/trendy/examples/wgpd513.xlsx'
    arquivo_devolucao = '/Users/macbookpro/Desktop/dev/trendy/examples/devolucao.xlsx'
    abrir_pasta = True
    limpar_cache = True
    limpar_dados = True


if __name__ == '__main__':
    print(msg("nova dinâmica"))
