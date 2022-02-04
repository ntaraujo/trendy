from utils import msg, pasted_to_list

class Posicao():
    def __init__(self, web, excel, cods_clientes, nomes_clientes, prevs_emb, implantacao_ini):
        msg('Construindo a Posição da rede')

        if not web.opened:
            web.open()
        web.totvs_access()
        if not web.totvs_logged:
            web.totvs_login()
        web.totvs_fav_pedidos()

        excel.open_macros()
        excel.open()

        for cod_cliente, nome_cliente in zip(cods_clientes, nomes_clientes):
            excel.new_sheet(nome_cliente)
            self.make_sheet(web, excel, cod_cliente, nome_cliente, prevs_emb, implantacao_ini)
        
        excel.file.sheets[0].delete()

    @staticmethod
    def filter_table(complete_table):
        msg("Filtrando tabelas")

        table = [['Pedido', 'Status', 'Est', 'NF', 'Dt Saída', 'Modelo', 'Descrição', 'Qt Pares', 'Nr Ordem'], ['TOTAL', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', None, '\xa0']]
        count = 0
        for line in complete_table[2:]:
            if "Cancelado" not in line[1]:
                table.append([cell for index, cell in enumerate(line) if index not in (2,4,5,6,12,14,15,16)])
                count += int(line[11])
        table[1][7] = count
        return table

    def make_sheet(self, web, excel, cod_cliente, nome_cliente, prevs_emb, implantacao_ini):
        msg(f'Construindo a Posição da loja "{nome_cliente}"')

        from utils import capitalized_month, simple_to_datetime

        excel.insert(nome_cliente)
        inserted = excel.back_range(1, 9)
        excel.bold(inserted)
        excel.center(inserted)
        excel.color(inserted, 255, 0, 0)
        excel.merge_across(inserted)

        for prev_emb in prevs_emb:
            web.totvs_fav_pedidos_fill(cod_cliente, prev_emb, implantacao_ini)
            table = self.filter_table(web.totvs_fav_pedidos_complete_table())

            excel.insert([[None],["PEDIDO " + capitalized_month(simple_to_datetime(prev_emb))]])
            inserted = excel.back_range(2, 9)
            excel.bold(inserted)
            excel.center(inserted)
            excel.merge_across(inserted)
            excel.insert(table)
        
        excel.run("posicao_general_format")

if __name__ == '__main__':
    from automators import web, excel

    Posicao(
        web.Web(),
        excel.Excel(),
        pasted_to_list("""969611
1000560
611379
980420
"""),
        pasted_to_list("""3 MENINAS (RUA) - LJ
A CASA DAS 3 MENINAS (ABC) - ANETE
A CASA DAS 3 MENINAS (PLAZA) - DEBORAH
CLUBE MELISSA - GRAND PLAZA SANTO ANDRÉ
"""),
        pasted_to_list("""03122021
03012022
03022022
"""),
        "16022000"
        )
