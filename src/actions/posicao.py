# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from actions.base_action import RedeAction
from utils import msg, pasted_to_list, retry, run_scheduled, capitalized_month, simple_to_datetime, progress


class Posicao(RedeAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)

        self.prevs_emb = pasted_to_list(self.args.prevs_emb or '')
        if not self.prevs_emb:
            from datetime import datetime
            from utils import datetime_to_simple

            month = datetime.now().month + 1
            year = datetime.now().year

            for _ in range(4):
                new_date = datetime_to_simple(datetime(year, month, 3))
                msg(f'Data a ser procurada: {new_date}')
                self.prevs_emb.insert(0, new_date)
                if month == 1:
                    month = 12
                    year -= 1
                else:
                    month -= 1

        self.make_sheet_value = len(self.prevs_emb)

        if not self.web.opened:
            self.web.open()
        self.web.totvs_access()
        if not self.web.totvs_logged:
            password = self.args.senha_totvs
            self.web.totvs_login(password)
        self.web.totvs_fav_program_access(3, 18)

        retry(self.web.switch_to_frame)("Fr_work")

        self.make_workbook()

        run_scheduled()
        self.web.close()

    @staticmethod
    def filter_table(complete_table):
        msg("Filtrando tabelas")

        table = [['Pedido', 'Status', 'Est', 'NF', 'Dt Saída', 'Modelo', 'Descrição', 'Qt Pares', 'Vl Líq', 'Nr Ordem'],
                 ['TOTAL',  None,     None,  None, None,       None,     None,        None,       None,     None]]
        qt_count = 0
        vl_count = 0
        for line in complete_table[2:]:
            if "Cancelado" not in line[1]:
                table.append([cell for index, cell in enumerate(line) if index not in (2, 4, 5, 6, 14, 15, 16)])
                qt_count += int(line[11])
                vl_count += float(line[12].replace('.', '').replace(',', '.'))
        table[1][7] = qt_count  # noqa
        table[1][8] = f'{vl_count:_.2f}'.replace('.', ',').replace('_', '.')  # noqa
        return table

    def make_sheet(self, cod_cliente, nome_cliente):
        msg(f'Construindo a Posição da loja "{nome_cliente}"')

        self.excel.insert(nome_cliente)
        self.excel.on_back_range(
            1, 10,
            self.excel.bold,
            self.excel.center,
            (self.excel.font_color, (255, 0, 0)),
            self.excel.merge_across
        )

        for prev_emb in self.prevs_emb:
            progress()

            self.web.totvs_fav_pedidos_fill(cod_cliente, prev_emb, "01012000")
            table = self.filter_table(self.web.totvs_fav_pedidos_complete_table())

            self.excel.insert([[None], ["PEDIDO " + capitalized_month(simple_to_datetime(prev_emb))]])
            self.excel.on_back_range(
                2, 10,
                self.excel.bold,
                self.excel.center,
                self.excel.merge_across
            )
            self.excel.insert(table)

        self.excel.run("posicao_general_format")


if __name__ == '__main__':
    from automators import web, excel
    from utils import ExampleArgs

    Posicao(
        ExampleArgs(),
        web.Web(),
        excel.Excel()
    )
