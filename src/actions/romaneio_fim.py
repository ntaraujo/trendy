# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from actions.base_action import BaseAction
from utils import run_scheduled, global_path
import utils

utils.running_scheduled = True


class RomaneioFim(BaseAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)
    
        arquivo_oc = list(self.excel.get_csv_reader(self.args.arquivo_oc, reader_kwargs={"delimiter": ";"}))
        material_ind = 2
        oc_ind = 0
        arquivo_oc_start_at = 1

        arquivo_oc_dict = {line[7].strip()+line[10].strip(): line for line in arquivo_oc[arquivo_oc_start_at:]}

        self.excel.open_app()
        self.excel.open(global_path(self.args.romaneio_inicio))
        cod_refs = [cod[:10].strip() for cod in self.excel.file.sheets.active.range('C14:C100').value if cod]

        materials = []
        ocs = []

        for cod_ref in cod_refs:
            o, _, m = arquivo_oc_dict[cod_ref][oc_ind:material_ind+1]
            materials.append((m,))
            ocs.append((o,))

        self.excel.assign('D14', materials)
        self.excel.assign('K14', ocs)
        self.excel.assign('G10', ocs[-1])

        run_scheduled()


if __name__ == '__main__':
    from automators import web, excel
    from utils import ExampleArgs

    RomaneioFim(
        ExampleArgs(),
        web.Web(),
        excel.Excel()
    )
