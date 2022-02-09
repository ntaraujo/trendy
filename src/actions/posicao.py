if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from utils import msg, pasted_to_list, run_scheduled


class Posicao:
    def __init__(self, args, web, excel):

        self.args = args
        self.web = web
        self.excel = excel
        # TODO access like args.cods_cliente
        self.cods_clientes = pasted_to_list(self.args['cods_cliente'] or '')
        self.nomes_clientes = pasted_to_list(self.args['nomes_cliente'] or '')

        if not self.cods_clientes or not self.nomes_clientes:
            for status, rede, nome_cliente, cod_cliente in self.excel.file_vertical_search(
                    self.args['nome_rede'].upper(), self.args['dinamica'], 9, 5, 9, 10, 11):
                if cod_cliente is None or cod_cliente == '':
                    msg(f'CLIENTE SEM CÓDIGO: "{rede}" --> "{nome_cliente}"')
                    continue
                elif status is None or "INATIVO" in status.upper():
                    continue

                msg(f'Cliente encontrado: "{rede}" --> "{nome_cliente}" --> "{cod_cliente}"')
                self.cods_clientes.append(cod_cliente)
                self.nomes_clientes.append(nome_cliente)

        self.prevs_emb = pasted_to_list(self.args['prevs_emb'] or '')
        if not self.prevs_emb:
            from datetime import datetime
            from utils import datetime_to_simple

            this_month = datetime.now().month
            this_year = datetime.now().year

            for _ in range(3):
                new_date = datetime_to_simple(datetime(this_year, this_month, 3))
                msg(f'Data a ser procurada: {new_date}')
                self.prevs_emb.append(new_date)
                if this_month == 1:
                    this_month = 12
                    this_year -= 1
                else:
                    this_month -= 1

        if not self.web.opened:
            self.web.open()
        self.web.totvs_access()
        if not self.web.totvs_logged:
            password = self.args['senha_totvs']
            self.web.totvs_login(password)
        self.web.totvs_fav_pedidos()

        self.excel.open_app()
        self.excel.open_macros()

        self.make_workbook()

        run_scheduled()

        web.close()

    def make_workbook(self):
        msg('Construindo a Posição da rede')

        self.excel.open()

        for cod_cliente, nome_cliente in zip(self.cods_clientes, self.nomes_clientes):
            self.excel.new_sheet(nome_cliente)
            self.make_sheet(cod_cliente, nome_cliente)

        self.excel.delete_sheet(0)

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
        table[1][8] = f'{vl_count:_.2f}'.replace('_', '.')  # noqa
        return table

    def make_sheet(self, cod_cliente, nome_cliente):
        msg(f'Construindo a Posição da loja "{nome_cliente}"')

        from utils import capitalized_month, simple_to_datetime

        self.excel.insert(nome_cliente)
        self.excel.on_back_range(
            1, 10,
            self.excel.bold,
            self.excel.center,
            (self.excel.color, (255, 0, 0)),
            self.excel.merge_across
        )

        for prev_emb in self.prevs_emb:
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
    from utils import example_args

    Posicao(
        example_args,
        web.Web(),
        excel.Excel()
    )
