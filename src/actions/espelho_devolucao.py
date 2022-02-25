# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from actions.base_action import WebToExcelAction
from utils import retry, run_scheduled, simple_to_datetime, total_progress, progress, msg, global_path, common_start
import os
import re
from difflib import get_close_matches
from datetime import datetime
from dateutil.relativedelta import relativedelta


class EspelhoDevolucao(WebToExcelAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)

        dinamica_sheet = self.excel.xls_workbook(self.args.dinamica).active
        len_dinamica = dinamica_sheet.max_row

        arquivo_sheet = self.excel.xls_workbook(self.args.arquivo_devolucao).active
        len_arquivo = arquivo_sheet.max_row

        espelho_precos_sheet = self.excel.xls_workbook(global_path("resources/espelho.xlsx"))['Lista de Preço']
        len_espelho_precos = espelho_precos_sheet.max_row

        total_progress((len_dinamica + len_arquivo) * 2 + len_espelho_precos)

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
        
        descs = []
        cod_produtos = []
        qtds = []
        vl_units = []
        nfs = []
        fabricas = []
        obss = []

        for row in arquivo_iter:
            progress()

            if not row or not row[0] or not row[2]:
                break

            cod_produtos.append(str(row[0]).strip())
            qtds.append(str(row[2]).strip())

        if not self.web.opened:
            self.web.open()
        self.web.totvs_access()
        if not self.web.totvs_logged:
            password = self.args.senha_totvs
            self.web.totvs_login(password)
        self.web.totvs_fav_program_access(3, 17)

        retry(self.web.switch_to_frame)("Fr_work")

        self.web.totvs_fav_notas1_fill(cod_cliente, ','.join(cod_produtos))

        totvs_table = self.web.totvs_fav_notas1_complete_table()

        total_progress(len(totvs_table) * 2 - 2)

        cod_produto_dicts_dict = {}

        assert len(totvs_table) == len(self.web.totvs_table_links) + 1

        for row, link in zip(totvs_table[1:], self.web.totvs_table_links):
            progress()

            totvs_dict = {
                'fabrica': row[0],
                'nf': row[2],
                'dt_saida': simple_to_datetime(row[5].replace('/', '')) if '?' not in row[5] else None,
                'dt_entrada': simple_to_datetime(row[6].replace('/', '')) if '?' not in row[6] else None,
                'vl_unit': row[10],
                'total': int(row[11]),
                'saldo': int(row[12]),
                'link': link
            }

            _cod_produto = row[8]
            if _cod_produto in cod_produto_dicts_dict:
                cod_produto_dicts_dict[_cod_produto].append(totvs_dict)
            else:
                cod_produto_dicts_dict[_cod_produto] = [totvs_dict]
        
        assert set(cod_produtos) == set(cod_produto_dicts_dict)
        
        espelho_precos_iter = espelho_precos_sheet.values
        next(espelho_precos_iter)
        progress()
        cod_produto_desc_dict = {progress(str)(row[0]).strip():row[1] for row in espelho_precos_iter}

        three_months = datetime.today() + relativedelta(months=-3)

        for cod_produto in cod_produtos:
            dicts = cod_produto_dicts_dict[cod_produto]

            best_dict_index = None
            best_dict_date = datetime(1, 1, 1)

            latest_dict_index = None
            latest_dict_date = datetime(1, 1, 1)

            latest_saida_dict_index = None
            latest_saida_dict_date = datetime(1, 1, 1)

            for index, totvs_dict in enumerate(dicts):
                progress()

                dt_entrada = totvs_dict['dt_entrada']
                if dt_entrada and dt_entrada > latest_dict_date:
                    latest_dict_date = dt_entrada
                    latest_dict_index = index
                if dt_entrada and dt_entrada < three_months and dt_entrada > best_dict_date:
                    best_dict_date = dt_entrada
                    best_dict_index = index
                dt_saida = totvs_dict['dt_saida']
                if dt_saida and dt_entrada > latest_saida_dict_date:
                    latest_saida_dict_date = dt_saida
                    latest_saida_dict_index = index
            
            if best_dict_index is not None:
                best_dict = dicts[best_dict_index]
            elif latest_dict_index is not None:
                best_dict = dicts[latest_dict_index]
            elif latest_saida_dict_index is not None:
                best_dict = dicts[latest_saida_dict_index]
            else:
                best_dict = dicts[-1]
            
            desc = cod_produto_desc_dict.get(cod_produto, None)
            obs = ''
            need_ipi = cod_produto.startswith("34")

            if not desc or need_ipi:
                main_window = self.web.prepare_for_new_window()
                self.web.driver.execute_script(best_dict['link'])
                new_window = self.web.get_new_window()
                self.web.driver.switch_to.window(new_window)

                notas2_table = self.web.totvs_fav_notas2_itens_table()[2:-1]

                self.web.driver.close()
                self.web.driver.switch_to.window(main_window)

            if not desc:
                possible_descs = [re.sub(r' (BB|INF|AD) .*$', r' \1', row[2]) for row in notas2_table]
                desc = common_start(*possible_descs)
            if need_ipi:
                first_ipi_total = float(notas2_table[0][9].replace('.', '').replace(',', '.'))
                first_qtd = int(notas2_table[0][4])
                obs = f'IPI R$ {first_ipi_total/first_qtd}'
            
            descs.append(desc)
            vl_units.append(best_dict['vl_unit'])
            nfs.append(best_dict['nf'])
            fabricas.append(best_dict['fabrica'])
            obss.append(obs)


        self.excel.open_app()
        self.excel.open(global_path("resources/espelho.xlsx"))

        self.excel.assign('A9', descs)
        self.excel.assign('B9', cod_produtos)
        self.excel.assign('C9', qtds)
        self.excel.assign('D9', vl_units)
        self.excel.assign('F9', nfs)
        self.excel.assign('I9', fabricas)
        self.excel.assign('J9', obss)

        run_scheduled()


if __name__ == '__main__':
    from automators import web, excel
    from utils import ExampleArgs

    EspelhoDevolucao(
        ExampleArgs(),
        web.Web(),
        excel.Excel()
    )
