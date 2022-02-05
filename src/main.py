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
    parser.parse_args()

if __name__ == '__main__':
    main()
