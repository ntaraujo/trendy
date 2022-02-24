# debug helper for vscode
import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from actions.base_action import WebToExcelAction
from utils import retry, run_scheduled, total_progress, progress, msg
import os
import re
from difflib import get_close_matches


class EspelhoDevolucao(WebToExcelAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)

        dinamica_sheet = self.excel.xls_workbook(self.args.dinamica).active
        len_dinamica = dinamica_sheet.max_row

        arquivo_sheet = self.excel.xls_workbook(self.args.arquivo_devolucao).active
        len_arquivo = arquivo_sheet.max_row

        total_progress((len_dinamica + len_arquivo) * 2)

        arquivo_iter = arquivo_sheet.values
        for _ in range(3):
            next(arquivo_iter)
            progress()

        razao = next(arquivo_iter)
        progress()
        razao = re.match(r' *RAZÃO SOCIAL: *(.*)', razao[0])[1]
        razao = re.sub(r' +|\n+|\r+', ' ', razao).upper().strip()
        razao = get_close_matches(razao, (progress(row[7].strip)() for row in dinamica_sheet.values if type(row[7])==str), 1)[0]
        msg(f"Razão encontrada: {razao}")
        possible_cod_clientes = []
        possible_lojas = []

        for _loja, _cod_cliente in self.excel.xls_file_vertical_search(razao, dinamica_sheet, 8, 10, 11):
            possible_lojas.append(_loja)
            possible_cod_clientes.append(_cod_cliente)
        total_progress(len(possible_lojas))

        loja = re.sub(r' +|\n+|\r+', ' ', re.match(r' *NOME DA LOJA: *(.*)', next(arquivo_iter)[0])[1]).upper().strip()
        progress()
        loja = get_close_matches(loja, possible_lojas, 1, 0)[0]
        msg(f"Loja encontrada: {loja}")

        for _loja, cod_cliente in zip(possible_lojas, possible_cod_clientes):
            if _loja == loja:
                msg(f"Código encontrado: {cod_cliente}")
                break

        for _ in range(6):
            next(arquivo_iter)
            progress()
        
        cod_produtos = []
        qtds = []
        descs = []

        for row in arquivo_iter:
            progress()

            if not row or not any(row):
                break

            cod_produtos.append(str(row[0]).strip())
            qtds.append(str(row[2]).strip())
            descs.append(row[3].strip() if row[3] else None)

        if not self.web.opened:
            self.web.open()
        self.web.totvs_access()
        if not self.web.totvs_logged:
            password = self.args.senha_totvs
            self.web.totvs_login(password)
        self.web.totvs_fav_program_access(3, 17)

        retry(self.web.switch_to_frame)("Fr_work")

        self.web.totvs_fav_notas1_fill(cod_cliente, ','.join(cod_produtos))

        table = self.web.totvs_fav_notas1_complete_table()

        total_progress(len(table) - 1)

        cod_product_dict = {}

        for row, link in zip(table[1:], self.web.totvs_table_links):
            progress()

            row.append(link)

            cod = row[8]
            if cod in cod_product_dict:
                cod_product_dict[cod].append(row)
            else:
                cod_product_dict[cod] = [row]

        # self.excel.open_app()
        # self.excel.open(global_path("resources/espelho.xlsx"))

        run_scheduled()


if __name__ == '__main__':
    from automators import web, excel
    from utils import ExampleArgs

    EspelhoDevolucao(
        ExampleArgs(),
        web.Web(),
        excel.Excel()
    )
