# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from actions.base_action import BaseAction
from utils import run_scheduled, data_dir_path, total_progress, progress
from sortedcontainers import SortedDict
import os


class PedidoAberto(BaseAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)

        dinamica_sheet = self.excel.xls_workbook(self.args.dinamica).active
        len_dinamica = dinamica_sheet.max_row

        arquivo_sheet = self.excel.xls_workbook(self.args.arquivo_wgpd513).active
        len_arquivo = arquivo_sheet.max_row

        total_progress(len_dinamica + 2 * len_arquivo)

        table_dict = SortedDict()

        cod_dict = {
            progress(str)(row[10]).strip(): row for row in dinamica_sheet.values if row[10]
        }

        arquivo_iter = arquivo_sheet.values
        next(arquivo_iter)
        progress()

        len_table_dict = 0

        for row in arquivo_iter:
            progress()

            if not row or not any(row):
                break

            if "20212441" in row[20] or "Cancelado" in row[14]:
                continue

            cod = str(row[4]).strip()
            try:
                rede, fantasia = cod_dict[cod][8:10]
            except KeyError:
                rede = fantasia = "=NA()"
            vend = row[1]

            if vend not in table_dict:
                table_dict[vend] = SortedDict()
            if rede not in table_dict[vend]:
                table_dict[vend][rede] = SortedDict()
            if fantasia not in table_dict[vend][rede]:
                table_dict[vend][rede][fantasia] = []
            new = [col for i, col in enumerate(row) if i in (
                1, 3, 4, 5, 14, 28, 31, 33, 34, 35, 38, 39, 41, 42, 44, 45, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56
            )]
            table_dict[vend][rede][fantasia].append(new[:3] + [rede, fantasia] + new[3:])
            len_table_dict += 1

        total_progress(len_table_dict - len_arquivo)

        title = [
            "Vend", "Nr. Pedido", "Cód. Cliente", "REDE", "FANTASIA", "Razão Social", "Situação", "Cod. Produção",
            "Ordem Compra", "Prev. Fat", "Emis. Nff", "Nr. Nota", "Qtde", "Vlr. Líq.", "Descrição Item",
            "Descrição Cor", "Vlr. Unit.", "CCF", "Prazo Adicional Pedido", "Prazo Adicional Dias NF",
            "#Acordo Comercial - 1 e 700", "#Comercialização E-commerce", "#Condição Pagamento", "#Desc.Cliente",
            "#Desconto item", "#MELISSA - Franquias", "#Melissa Item B2B", "#Pontualidade NF"
        ]
        blank = [None for _ in range(28)]

        pedido = self.excel.xls_workbook()

        sheet = pedido.active

        row_total_end = 1

        for vend in table_dict.values():
            for rede in vend.values():
                for fantasia in rede.values():
                    sheet.append(title)
                    self.excel.xls_on_back_range(sheet, 1, 28, self.excel.xls_bold, (self.excel.xls_color, ("B7DEE8",)))
                    row_total_start = row_total_end + 1

                    # qt_count = 0
                    # vl_count = 0
                    for row in fantasia:
                        progress()
                        sheet.append(row)
                        # qt_count += int(row[12])
                        # vl = row[13]
                        # vl_count += vl if type(vl) in (float, int) else float(vl.replace('.', '').replace(',', '.'))

                    row_total_end = sheet.max_row

                    total = blank.copy()
                    total[0] = "TOTAL"
                    total[12] = f'=SUM(M{row_total_start}:M{row_total_end})'
                    total[13] = f'=SUM(N{row_total_start}:N{row_total_end})'
                    sheet.append(total)
                    self.excel.xls_on_back_range(sheet, 1, 28, self.excel.xls_bold)

                    sheet.append(blank)

                    row_total_end += 3

        temp_path = os.path.join(data_dir_path, "temp-excel.xlsx")
        pedido.save(temp_path)

        self.excel.open_app()
        self.excel.open_macros()
        self.excel.open(temp_path)
        self.excel.run("pedido_aberto_general_format")

        run_scheduled()


if __name__ == '__main__':
    from automators import web, excel
    from utils import ExampleArgs

    PedidoAberto(
        ExampleArgs(),
        web.Web(),
        excel.Excel()
    )
