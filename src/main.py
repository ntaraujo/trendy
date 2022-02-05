from gooey import Gooey, GooeyParser, local_resource_path
from automators.web import Web
from automators.excel import Excel
import signal

web = Web()
excel = Excel()


@Gooey(
    language="portuguese",
    program_name="Trendy",
    image_dir=local_resource_path("gooey-images"),
    language_dir=local_resource_path("gooey-languages"),
    navigation="SIDEBAR",
    sidebar_title="Ações",
    show_success_modal=False,
    show_sidebar=True,
    shutdown_signal=signal.CTRL_C_EVENT,
    advanced=True,
    tabbed_groups=True,
    requires_shell=False,
    clear_before_run=True
)
def main():
    parser = GooeyParser(description="Aplicativo de automação com planilhas e sistemas online")
    subs = parser.add_subparsers(dest='action')

    posicao_parser = subs.add_parser('Posição')

    posicao_basic_group = posicao_parser.add_argument_group('Básico', gooey_options={'columns': 1})
    posicao_basic_group.add_argument('Senha*', widget='PasswordField', help="A senha de acesso ao TOTVS")
    posicao_basic_group.add_argument('--Rede',
                                     help="O nome da rede da qual o programa fará a posição. Ela será buscada no "
                                          "arquivo da dinâmica")
    posicao_basic_group.add_argument('--Dinâmica', widget='FileChooser',
                                     help='A dinâmica é o arquivo com os códigos e nomes de cada cliente')

    posicao_advanced_group = posicao_parser.add_argument_group('Avançado')
    posicao_advanced_group.add_argument('--Datas', widget='Textarea', gooey_options={'height': 100},
                                        help="As datas de previsão de embarque, separadas por ENTER")
    posicao_advanced_group.add_argument('--Códigos', widget='Textarea', gooey_options={'height': 100},
                                        help="Os códigos de cada cliente, separados por ENTER")
    posicao_advanced_group.add_argument('--Nomes', widget='Textarea', gooey_options={'height': 100},
                                        help="Os respectivos nomes para cada cliente, separados por ENTER")

    args = parser.parse_args().__dict__

    def run(action):
        try:
            action(args, web, excel)
        except Exception as e:
            # uncomment when packaging
            # web.close()
            raise e

    if args['action'] == 'Posição':
        from actions.posicao import Posicao
        run(Posicao)


if __name__ == '__main__':
    main()
