from utils import pasted_to_list, msg

class BaseAction:
    def __init__(self, args, web, excel):
        self.args = args
        self.web = web
        self.excel = excel


class ExcelAction(BaseAction):

    def make_workbook(self):
        raise NotImplementedError()

    @staticmethod
    def filter_table(complete_table):
        raise NotImplementedError()

    def make_sheet(self, cod_cliente, nome_cliente):
        raise NotImplementedError()


class RedeAction(ExcelAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)

        self.cods_clientes = pasted_to_list(self.args.cods_cliente or '')
        self.nomes_clientes = pasted_to_list(self.args.nomes_cliente or '')

        if not self.cods_clientes or not self.nomes_clientes:
            for status, rede, nome_cliente, cod_cliente in self.excel.file_vertical_search(
                    self.args.nome_rede.upper(), self.args.dinamica, 9, 5, 9, 10, 11):
                if cod_cliente is None or cod_cliente == '':
                    msg(f'CLIENTE SEM CÓDIGO: "{rede}" --> "{nome_cliente}"')
                    continue
                elif status is None or "INATIVO" in status.upper():
                    continue

                msg(f'Cliente encontrado: "{rede}" --> "{nome_cliente}" --> "{cod_cliente}"')
                self.cods_clientes.append(cod_cliente)
                self.nomes_clientes.append(nome_cliente)

    def make_workbook(self):
        msg('Construindo a pasta de trabalho da rede')

        self.excel.open_app()
        self.excel.open_macros()
        self.excel.open()

        for cod_cliente, nome_cliente in zip(self.cods_clientes, self.nomes_clientes):
            self.excel.new_sheet(nome_cliente)
            self.make_sheet(cod_cliente, nome_cliente)

        self.excel.delete_sheet(0)
