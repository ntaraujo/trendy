# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from actions.base_action import WebToExcelAction
from utils import run_scheduled, common_start, global_path
import re


class Romaneio(WebToExcelAction):
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
        nf_itens = self.web.totvs_fav_notas_complete_table()
        desc_ind = 2
        preco_ind = 5
        nf_itens_start_at = 2
        nf_itens_end_at = -1

        nf_itens_dict = {line[1].strip()[:10]: line for line in nf_itens[nf_itens_start_at:nf_itens_end_at]}
        descs = [re.match(r'(.*) *[A-Z]+ *$', line[desc_ind])[1] for line in nf_itens[nf_itens_start_at:nf_itens_end_at]]
        desc = common_start(*descs).strip()

        oc = self.args.oc

        etiqueta = list(self.excel.get_csv_reader(self.args.etiqueta, reader_kwargs={"delimiter": ";"}))
        material_ind = 1
        etiqueta_start_at = 1

        etiqueta_dict = {line[6].strip()+line[9].strip(): line for line in etiqueta[etiqueta_start_at:]}

        table = []
        for line in cac[cac_start_at:]:
            cod = line[cod_ind].strip()
            cod_ref = cod[:10]
            material = etiqueta_dict[cod_ref][material_ind].strip()

            cor = line[cor_ind].strip()
            tam = line[tam_ind].strip()
            qtd = line[qtd_ind].strip()
            preco = nf_itens_dict[cod_ref][preco_ind].strip()
            total = float(preco.replace('.', '').replace(',', '.')) * int(qtd)
            total = f'{total:_.2f}'.replace('.', ',').replace('_', '.')

            table.append((cod, material, desc, cor, tam, qtd, preco, total, oc))
        
        self.excel.open_app()
        self.excel.open(global_path("resources/romaneio.xls"))

        sh = self.excel.file.sheets.active
        sh['D9'].value = cnpj
        sh['G8'].value = [[nf], [serie], [oc]]
        sh['C14'].value = table

        run_scheduled()
        self.web.close()


if __name__ == '__main__':
    from automators import web, excel
    from utils import ExampleArgs

    Romaneio(
        ExampleArgs(),
        web.Web(),
        excel.Excel()
    )
