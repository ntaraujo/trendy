from actions.base_action import BaseAction
from utils import app_dir_path, cache_dir_path, system_open, system_remove


class Opcoes(BaseAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)

        args = self.args.__dict__
        if args['Abrir pasta do aplicativo']:
            system_open(app_dir_path)
        elif args['Apagar o cache']:
            system_remove(cache_dir_path)
        elif args['Apagar meus dados']:
            system_remove(app_dir_path)
