if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from utils import msg, pasted_to_list


class Posicao:
    def __init__(self, args, web, excel):

        self.args = args
        self.web = web
        self.excel = excel

        self.cods_clientes = pasted_to_list(args['cods_cliente'] or '')
        self.nomes_clientes = pasted_to_list(args['nomes_cliente'] or '')

        if not self.cods_clientes or not self.nomes_clientes:
            for rede, nome_cliente, cod_cliente in excel.file_vertical_search(args['nome_rede'].upper(), args['dinamica'], 9,
                                                                              9, 10, 11):
                if cod_cliente is None or cod_cliente == '':
                    msg(f'CLIENTE SEM CÓDIGO: "{rede}" --> "{nome_cliente}"')
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

        if not web.opened:
            web.open()
        web.totvs_access()
        if not web.totvs_logged:
            password = args['senha_totvs']
            web.totvs_login(password)
        web.totvs_fav_pedidos()

        excel.open_macros()

        self.make_workbook()

        web.close()

    def make_workbook(self):
        msg('Construindo a Posição da rede')

        args, web, excel = self.args, self.web, self.excel

        excel.open()

        for cod_cliente, nome_cliente in zip(self.cods_clientes, self.nomes_clientes):
            excel.new_sheet(nome_cliente)
            self.make_sheet(cod_cliente, nome_cliente)

        excel.file.sheets[0].delete()

    @staticmethod
    def filter_table(complete_table):
        msg("Filtrando tabelas")

        table = [['Pedido', 'Status', 'Est', 'NF', 'Dt Saída', 'Modelo', 'Descrição', 'Qt Pares', 'Nr Ordem'],
                 ['TOTAL', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', None, '\xa0']]
        count = 0
        for line in complete_table[2:]:
            if "Cancelado" not in line[1]:
                table.append([cell for index, cell in enumerate(line) if index not in (2, 4, 5, 6, 12, 14, 15, 16)])
                count += int(line[11])
        table[1][7] = count  # noqa
        return table

    def make_sheet(self, cod_cliente, nome_cliente):
        msg(f'Construindo a Posição da loja "{nome_cliente}"')

        from utils import capitalized_month, simple_to_datetime

        excel = self.excel
        web = self.web

        excel.insert(nome_cliente)
        inserted = excel.back_range(1, 9)
        excel.bold(inserted)
        excel.center(inserted)
        excel.color(inserted, 255, 0, 0)
        excel.merge_across(inserted)

        for prev_emb in self.prevs_emb:
            web.totvs_fav_pedidos_fill(cod_cliente, prev_emb, "01012000")
            table = self.filter_table(web.totvs_fav_pedidos_complete_table())

            excel.insert([[None], ["PEDIDO " + capitalized_month(simple_to_datetime(prev_emb))]])
            inserted = excel.back_range(2, 9)
            excel.bold(inserted)
            excel.center(inserted)
            excel.merge_across(inserted)
            excel.insert(table)

        excel.run("posicao_general_format")


if __name__ == '__main__':
    from automators import web, excel
    from utils import example_args

    Posicao(
        example_args,
        web.Web(),
        excel.Excel()
    )
