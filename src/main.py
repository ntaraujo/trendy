# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

from gooey import Gooey, GooeyParser, local_resource_path
from automators.web import Web
from automators.excel import Excel
import signal
from utils import cache, save_cache, load_cache, msg, compiled, run_scheduled, log_file

# TODO separate classes: ExcelWriter and ExcelReader
# TODO separate classes: WebWriter and WebReader
# TODO remove passwords from git history
# TODO executable on mac os big sur
web = Web()
excel = Excel()


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
        default = cache['default'].get(cache_name, 'X')
        if default == 'None':
            default = 'X'
        group_.add_argument(name, default=default, **kwargs)

    def sub_parser(name):
        return subs.add_parser(name)

    def group(parser_, name, **kwargs):
        return parser_.add_argument_group(name, **kwargs)


    posicao_parser = sub_parser('Posição')

    posicao_basic_group = group(posicao_parser, 'Básico', gooey_options={'columns': 1})
    argument(posicao_basic_group, '--nome_rede',
             help="O nome da rede da qual o programa fará a posição. Ela será buscada no "
                  "arquivo da dinâmica")
    argument(posicao_basic_group, '--dinamica', widget='FileChooser',
             help='A dinâmica é o arquivo com os códigos e nomes de cada cliente')
    argument(posicao_basic_group, 'senha_totvs', widget='PasswordField', help="A senha de acesso ao TOTVS")

    posicao_advanced_group = group(posicao_parser, 'Avançado')
    argument(posicao_advanced_group, '--prevs_emb', widget='Textarea', gooey_options={'height': 100},
             help="As datas de previsão de embarque, separadas por ENTER")
    argument(posicao_advanced_group, '--cods_cliente', widget='Textarea', gooey_options={'height': 100},
             help="Os códigos de cada cliente, separados por ENTER")
    argument(posicao_advanced_group, '--nomes_cliente', widget='Textarea', gooey_options={'height': 100},
             help="Os respectivos nomes para cada cliente, separados por ENTER")


    titulos_parser = sub_parser('Títulos')

    titulos_basic_group = group(titulos_parser, 'Básico', gooey_options={'columns': 1})
    argument(titulos_basic_group, '--nome_rede',
             help="O nome da rede da qual o programa fará o relatório de títulos. Ela será buscada no "
                  "arquivo da dinâmica")
    argument(titulos_basic_group, '--dinamica', widget='FileChooser',
             help='A dinâmica é o arquivo com os códigos e nomes de cada cliente')
    argument(titulos_basic_group, 'senha_totvs', widget='PasswordField', help="A senha de acesso ao TOTVS")

    titulos_advanced_group = group(titulos_parser, 'Avançado')
    argument(titulos_advanced_group, '--cods_cliente', widget='Textarea', gooey_options={'height': 100},
             help="Os códigos de cada cliente, separados por ENTER")
    argument(titulos_advanced_group, '--nomes_cliente', widget='Textarea', gooey_options={'height': 100},
             help="Os respectivos nomes para cada cliente, separados por ENTER")


    romaneio1_parser = sub_parser('RomaneioInício')

    romaneio1_basic_group = group(romaneio1_parser, 'Básico', gooey_options={'columns': 1})
    argument(romaneio1_basic_group, 'cac', widget='FileChooser', help='Arquivo CAC no formato .csv')
    argument(romaneio1_basic_group, 'senha_totvs', widget='PasswordField', help="A senha de acesso ao TOTVS")


    romaneio2_parser = sub_parser('RomaneioFim')

    romaneio2_basic_group = group(romaneio2_parser, 'Básico', gooey_options={'columns': 1})
    argument(romaneio2_basic_group, 'romaneio_inicio', widget='FileChooser', help='Arquivo gerado pela ação RomaneioInício')
    argument(romaneio2_basic_group, 'arquivo_oc', widget='FileChooser', help='Arquivo com a OC na primeira coluna em .csv (etiqueta)')


    romaneio_parser = sub_parser('RomaneioCompleto')

    romaneio_basic_group = group(romaneio_parser, 'Básico', gooey_options={'columns': 1})
    argument(romaneio_basic_group, 'oc', help="Número da ordem de compra")
    argument(romaneio_basic_group, 'cac', widget='FileChooser', help='Arquivo CAC no formato .csv')
    argument(romaneio_basic_group, 'etiqueta', widget='FileChooser', help='Arquivo da etiqueta no formato .csv')
    argument(romaneio_basic_group, 'senha_totvs', widget='PasswordField', help="A senha de acesso ao TOTVS")


    args = parser.parse_args()
    cache['default'] = {key: str(value) for key, value in args.__dict__.items()}
    args.__dict__ = {key: None if value == 'X' else value.strip() for key, value in args.__dict__.items()}
    save_cache()

    def run(action):
        try:
            action(args, web, excel)
        except Exception as e:

            if not web.opened:
                msg("INFO: Navegador não foi aberto")
            elif compiled:
                web.print('last-error.png')
                web.close()
            try:
                run_scheduled()
            except Exception as e:
                log_file.write(f'{type(e).__name__}\n{e}\n')
                raise e
            else:
                log_file.write(f'{type(e).__name__}\n{e}\n')
                raise e

    if args.action == 'Posição':
        from actions.posicao import Posicao
        run(Posicao)
    elif args.action == 'Títulos':
        from actions.titulos import Titulos
        run(Titulos)
    elif args.action == 'RomaneioInício':
        from actions.romaneio_inicio import RomaneioInicio
        run(RomaneioInicio)
    elif args.action == 'RomaneioFim':
        from actions.romaneio_fim import RomaneioFim
        run(RomaneioFim)
    elif args.action == 'RomaneioCompleto':
        from actions.romaneio_completo import RomaneioCompleto
        run(RomaneioCompleto)


if __name__ == '__main__':
    main()
