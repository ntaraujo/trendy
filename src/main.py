from gooey import Gooey, GooeyParser
from web import Web
from excel import Excel

web = Web()
excel = Excel()

@Gooey(dump_build_config=True, program_name="Trendy")
def main():
    parser = GooeyParser(description="Aplicativo de automação para planilhas e relatórios")
    parser.parse_args()

if __name__ == '__main__':
    main()
