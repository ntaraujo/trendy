# debug helper for vscode
# import os; os.chdir('/Users/macbookpro/Desktop/dev/trendy/src')

if __name__ == '__main__':
    from gooey import local_resource_path
    from sys import path as sys_path

    sys_path.insert(0, local_resource_path(""))

from actions.base_action import BaseAction
from utils import run_scheduled, global_path, insort

class PedidoAberto(BaseAction):
    def __init__(self, args, web, excel):
        super().__init__(args, web, excel)

        title = ["Vend",	"Nr. Pedido",	"Cód. Cliente",	"REDE",	"FANTASIA",	"Razão Social",	"Situação",	"Cod. Produção",	"Ordem Compra",	"Prev. Fat",	"Emis. Nff",	"Nr. Nota",	"Qtde",	 "Vlr. Líq.", 	"Descrição Item",	"Descrição Cor",	"Vlr. Unit.",	"CCF",	"Prazo Adicional Pedido",	"Prazo Adicional Dias", "NF",	"#Acordo Comercial - 1 e 700",	"#Comercialização E-commerce",	"#Condição Pagamento",	"#Desc.Cliente",	"#Desconto item",	"#MELISSA - Franquias",	"#Melissa Item B2B",	"#Pontualidade NF"]
        table = []

        arquivo_iter = self.excel.xls_workbook(self.args.arquivo_wgpd513).active.values
        next(arquivo_iter)

        for row in arquivo_iter:
            if not row:
                break

            if "20212441" in row[20] or "Cancelado" in row[14]:
                continue
            
            

        run_scheduled()


if __name__ == '__main__':
    from automators import web, excel
    from utils import ExampleArgs

    PedidoAberto(
        ExampleArgs(),
        web.Web(),
        excel.Excel()
    )
