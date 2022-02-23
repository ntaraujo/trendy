# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from actions.base_action import WebToExcelAction
from utils import run_scheduled, common_start, global_path, total_progress, progress
import re


class RomaneioInicio(WebToExcelAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)

        cac = list(self.excel.get_csv_reader(self.args.cac, reader_kwargs={"delimiter": ";"}))

        cnpj = cac[2][0][1:]
        nf = cac[5][0]
        serie = cac[6][0]
        estabelecimento = "20" if nf[0] == "5" else "21"

        cod_ind = 1
        tam_ind = 5
        cor_ind = 4
        qtd_ind = 6
        cac_start_at = 9

        if not self.web.opened:
            self.web.open()
        self.web.totvs_access()
        if not self.web.totvs_logged:
            password = self.args.senha_totvs
            self.web.totvs_login(password)
        self.web.totvs_fav_program_access(3, 16)
        self.web.totvs_fav_notas_va_para(estabelecimento, serie, nf)
        self.web.totvs_fav_notas_items(nf)

        nf_itens_start_at = 2
        nf_itens_end_at = -1
        nf_itens = self.web.totvs_fav_notas_complete_table()[nf_itens_start_at:nf_itens_end_at]
        desc_ind = 2
        preco_ind = 5

        total_progress(len(cac) - cac_start_at + len(nf_itens) * 2 + 4)

        nf_itens_dict = {progress(line[1].strip)()[:10]: line for line in nf_itens}
        descs = [progress(re.match)(r'(.*) *[A-Z]+ *$', line[desc_ind])[1] for line in nf_itens]
        desc = common_start(*descs).strip()

        table = []
        for line in cac[cac_start_at:]:
            cod = line[cod_ind].strip()
            cod_ref = cod[:10]

            cor = line[cor_ind].strip()
            tam = line[tam_ind].strip()
            qtd = line[qtd_ind].strip()
            preco = nf_itens_dict[cod_ref][preco_ind].strip()

            table.append((cod, None, desc, cor, tam, qtd, preco))
            progress()
        
        self.excel.open_app()
        self.excel.open(global_path("resources/romaneio.xls"))
        progress()

        self.excel.assign('D9', cnpj)
        progress()
        self.excel.assign('G8', [[nf], [serie]])
        progress()
        self.excel.assign('C14', table)
        progress()

        run_scheduled()
        self.web.close()


if __name__ == '__main__':
    from automators import web, excel
    from utils import ExampleArgs

    RomaneioInicio(
        ExampleArgs(),
        web.Web(),
        excel.Excel()
    )
