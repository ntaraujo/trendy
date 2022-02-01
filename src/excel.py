from utils import msg

class Excel():
    def __init__(self):
        self.macros_file = None
        self.file = None
    
    def open_macros(self):
        msg("Abrindo macros")

        import xlwings as xw
        from os import path

        self.macros_file = xw.Book(path.abspath(path.join(path.dirname(__file__), "macros.xlsb")))
    
    def open(self, path=None):
        if path is None:
            msg("Abrindo nova pasta do Excel")
        else:
            from os.path import abspath
            path = abspath(path)
            msg(f'Abrindo pasta do Excel em "{path}"')            

        import xlwings as xw

        self.file = xw.Book(path)
    
    def run(self, macro, *args, **kwargs):
        msg(f'Executando a macro "{macro}"')

        return self.macros_file.macro(macro)(*args, **kwargs)

if __name__ == "__main__":
    excel = Excel()
    excel.open()
    excel.open_macros()
    excel.run("test")
