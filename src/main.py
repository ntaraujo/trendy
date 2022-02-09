from gooey import Gooey, GooeyParser, local_resource_path
from automators.web import Web
from automators.excel import Excel
import signal
from utils import cache, save_cache, load_cache, msg, compiled, run_scheduled, log_file
import codecs
import sys

if sys.stdout.encoding != 'UTF-8':
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'UTF-8':
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

web = Web()
excel = Excel()


# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

@Gooey(
    language="portuguese",
    program_name="Trendy",
    image_dir=local_resource_path("gooey-images"),
    language_dir=local_resource_path("gooey-languages"),
    navigation="SIDEBAR",
    sidebar_title="Ações",
    # show_success_modal=False,  # working different on mac
    show_sidebar=True,
    shutdown_signal=signal.SIGTERM,
    advanced=True,
    tabbed_groups=True,
    # requires_shell=False,  # not working on mac
    clear_before_run=True
)
def main():
    load_cache()

    parser = GooeyParser(description="Escreva um X nos campos que não devem ser usados")
    subs = parser.add_subparsers(dest='action')

    def argument(group_, name, **kwargs):
        cache_name = name.replace('--', '')
        default = cache['default'].get(cache_name, 'None')
        if default == 'None':
            default = None
        group_.add_argument(name, default=default, **kwargs)

    def sub_parser(name):
        return subs.add_parser(name)

    def group(parser_, name, **kwargs):
        return parser_.add_argument_group(name, **kwargs)

    posicao_parser = sub_parser('Posição')

    posicao_basic_group = group(posicao_parser, 'Básico', gooey_options={'columns': 1})
    argument(posicao_basic_group, 'senha_totvs', widget='PasswordField', help="A senha de acesso ao TOTVS")
    argument(posicao_basic_group, '--nome_rede',
             help="O nome da rede da qual o programa fará a posição. Ela será buscada no "
                  "arquivo da dinâmica")
    argument(posicao_basic_group, '--dinamica', widget='FileChooser',
             help='A dinâmica é o arquivo com os códigos e nomes de cada cliente')

    posicao_advanced_group = group(posicao_parser, 'Avançado')
    argument(posicao_advanced_group, '--prevs_emb', widget='Textarea', gooey_options={'height': 100},
             help="As datas de previsão de embarque, separadas por ENTER")
    argument(posicao_advanced_group, '--cods_cliente', widget='Textarea', gooey_options={'height': 100},
             help="Os códigos de cada cliente, separados por ENTER")
    argument(posicao_advanced_group, '--nomes_cliente', widget='Textarea', gooey_options={'height': 100},
             help="Os respectivos nomes para cada cliente, separados por ENTER")

    args = parser.parse_args()
    cache['default'] = {key: str(value) for key, value in args.__dict__.items()}
    args.__dict__ = {key: None if value == 'X' else value.strip() for key, value in args.__dict__.items()}
    save_cache()

    def run(action):
        try:
            action(args, web, excel)
        except Exception as e:
            import traceback

            if not web.opened:
                msg("INFO: Navegador não foi aberto")
            elif compiled:
                web.print('last-error-trendy.png')
                web.close()
            try:
                run_scheduled()
            except Exception as e:
                log_file.write(traceback.format_exc() + '\n')
                raise e
            else:
                log_file.write(traceback.format_exc() + '\n')
                raise e

    if args.action == 'Posição':
        from actions.posicao import Posicao
        run(Posicao)


if __name__ == '__main__':
    main()
