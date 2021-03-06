# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from actions.base_action import RedeAction
from utils import msg, run_scheduled, progress


class Titulos(RedeAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)
        self.make_sheet_value = 7

        if not self.web.opened:
            self.web.open()
        self.web.totvs_access()
        if not self.web.totvs_logged:
            password = self.args.senha_totvs
            self.web.totvs_login(password)
        self.web.totvs_fav_program_access(3, 2)

        self.make_workbook()

        run_scheduled()
        self.web.close()

    @staticmethod
    def filter_table(complete_table):
        msg("Filtrando tabelas")

        table = []
        vl_count = 0
        for line in complete_table[1:]:
            if "AN" not in line[2] and "SIM" in line[7]:
                table.append(
                    [cell for index, cell in enumerate(line) if index not in (0, 6, 7, 8, 9, 10, 11, 12, 14, 15)])
                vl_count += float(line[13].replace('.', '').replace(',', '.'))
        table = sorted(table, key=lambda x: float(x[-1]), reverse=True)
        table.insert(0, [
            'TOTAL', None, None, None, None, f'{vl_count:_.2f}'.replace('.', ',').replace('_', '.'), None, None, None
        ])
        table.insert(0, [
            'Est', 'Esp', 'Série', 'Documento', '/P', 'Total Saldo', 'Emissão', 'Dt Vcto', 'Dias'
        ])
        return table

    def make_sheet(self, cod_cliente, nome_cliente):
        msg(f'Construindo o Relatório de Títulos da loja "{nome_cliente}"')

        self.web.totvs_fav_clientes_va_para(cod_cliente)
        progress()
        main_window = self.web.totvs_fav_clientes_documentos(cod_cliente)
        progress()
        self.web.totvs_fav_clientes_filtro()
        progress()

        table = self.filter_table(self.web.totvs_fav_clientes_complete_table())
        progress()

        self.excel.insert(table[:2])
        self.excel.on_back_range(
            2, 9,
            self.excel.center,
            (self.excel.font_color, (255, 0, 0))
        )
        progress()

        self.excel.insert(table[2:])
        self.excel.on_back_range(
            len(table) - 1, 9,
            self.excel.center
        )
        progress()

        self.excel.run("titulos_general_format")

        self.web.driver.close()
        self.web.driver.switch_to.window(main_window)
        progress()


if __name__ == '__main__':
    from automators import web, excel
    from utils import ExampleArgs

    Titulos(
        ExampleArgs(),
        web.Web(),
        excel.Excel()
    )
