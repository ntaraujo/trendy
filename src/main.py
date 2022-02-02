from gooey import Gooey, GooeyParser
from web import Web
from excel import Excel

web = Web()
excel = Excel()

@Gooey(dump_build_config=True, program_name="Trendy")
def main():
    parser = GooeyParser(description="Aplicativo de automação para planilhas e relatórios")
    parser.parse_args()

def posicao(cod_cliente, nome_cliente, prevs_emb, implantacao_ini):
    from utils import capitalized_month, simple_to_datetime

    if not web.opened:
        web.open()
    web.totvs_access()
    if not web.totvs_logged:
        web.totvs_login()
    web.totvs_fav_pedidos()

    excel.insert(nome_cliente)

    for prev_emb in prevs_emb:
        web.totvs_fav_pedidos_fill(cod_cliente, prev_emb, implantacao_ini)
        table = web.totvs_fav_pedidos_complete_table()

        excel.insert("PEDIDO " + capitalized_month(simple_to_datetime(prev_emb)))
        excel.insert(table)

def posicoes(cods_clientes, nomes_clientes, prevs_emb, implantacao_ini, file_path=None):
    excel.open(file_path)
    # excel.open_macros()

    for cod_cliente, nome_cliente in zip(cods_clientes, nomes_clientes):
        excel.new_sheet(nome_cliente)
        posicao(cod_cliente, nome_cliente, prevs_emb, implantacao_ini)
    
    excel.file.sheets[0].delete()
    
    # excel.save()

if __name__ == '__main__':
    # main()
    posicoes(("1000595",), ("SUNSET",), ("03012022",), "16022000")
