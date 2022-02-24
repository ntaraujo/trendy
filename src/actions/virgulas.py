from actions.base_action import BaseAction
from utils import cache_dir_path, system_open, pasted_to_list
import os


class Virgulas(BaseAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)

        text = ','.join(pasted_to_list(self.args.texto))
        path = os.path.join(cache_dir_path, "Text.txt")
        with open(path, 'w') as file:
            file.write(text)
        system_open(path)
