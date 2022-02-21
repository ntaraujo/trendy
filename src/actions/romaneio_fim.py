# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from actions.base_action import BaseAction
from utils import run_scheduled, global_path
import utils


class RomaneioFim(BaseAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)
    
        arquivo_oc = list(self.excel.get_csv_reader(self.args.arquivo_oc, reader_kwargs={"delimiter": ";"}))
        material_ind = 2
        oc_ind = 0
        arquivo_oc_start_at = 1

        arquivo_oc_dict = {line[7].strip()+line[10].strip(): line for line in arquivo_oc[arquivo_oc_start_at:]}

        utils.running_scheduled = True

        self.excel.open_app()
        self.excel.open(global_path(self.args.romaneio_inicio))

        sheet_range = self.excel.file.sheets.active.range

        for line in range(14, 101):
            cod = sheet_range((line, 3)).value

            if not cod:
                break

            cod_ref = cod[:10].strip()
            material = arquivo_oc_dict[cod_ref][material_ind].strip()
            oc = arquivo_oc_dict[cod_ref][oc_ind].strip()

            self.excel.assign((line, 4), material)
            self.excel.assign((line, 11), oc)
        
        self.excel.assign('G10', oc)

        run_scheduled()


if __name__ == '__main__':
    from automators import web, excel
    from utils import ExampleArgs

    RomaneioFim(
        ExampleArgs(),
        web.Web(),
        excel.Excel()
    )
