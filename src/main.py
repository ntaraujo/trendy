from gooey import Gooey, GooeyParser, local_resource_path
from web import Web
from excel import Excel

web = Web()
excel = Excel()

@Gooey(language="gooey-lang", program_name="Trendy", image_dir=local_resource_path("gooey-images"), language_dir=local_resource_path(""), navigation="TABBED", sidebar_title="Ações")
def main():
    parser = GooeyParser(description="Aplicativo de automação para planilhas e relatórios")
    parser.parse_args()

def posicao_table(complete_table):
    table = [['Pedido', 'Status', 'Est', 'NF', 'Dt Saída', 'Modelo', 'Descrição', 'Qt Pares', 'Nr Ordem'], ['TOTAL', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', None, '\xa0']]
    count = 0
    for line in complete_table[2:]:
        if "Cancelado" not in line[1]:
            table.append([cell for index, cell in enumerate(line) if index not in (2,4,5,6,12,14,15,16)])
            count += int(line[11])
    table[1][7] = count
    return table

def posicao(cod_cliente, nome_cliente, prevs_emb, implantacao_ini):
    from utils import capitalized_month, simple_to_datetime

    if not web.opened:
        web.open()
    web.totvs_access()
    if not web.totvs_logged:
        web.totvs_login()
    web.totvs_fav_pedidos()

    excel.insert(nome_cliente)
    inserted = excel.back_range(1, 9)
    excel.bold(inserted)
    excel.center(inserted)
    excel.color(inserted, 255, 0, 0)
    excel.merge_across(inserted)

    for prev_emb in prevs_emb:
        web.totvs_fav_pedidos_fill(cod_cliente, prev_emb, implantacao_ini)
        table = posicao_table(web.totvs_fav_pedidos_complete_table())

        excel.insert([[None],["PEDIDO " + capitalized_month(simple_to_datetime(prev_emb))]])
        inserted = excel.back_range(2, 9)
        excel.bold(inserted)
        excel.center(inserted)
        excel.merge_across(inserted)
        excel.insert(table)
    
    excel.run("posicao_general_format")

def posicoes(cods_clientes, nomes_clientes, prevs_emb, implantacao_ini, file_path=None):
    excel.open_macros()
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
    # print(posicao_table([['Pedido', 'Status', 'Tipo Ped.', 'Est', 'Dt Refer', 'Prev Emb', 'Pré-Data', 'NF', 'Dt Saída', 'Modelo', 'Descrição', 'Qt Pares', 'Vl Líq', 'Nr Ordem', 'CR', 'CO', 'EF'], ['TOTAL', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', '\xa0', '768', '52.537,08', '\xa0', '\xa0', '\xa0', '\xa0'], ['490766801', 'Faturado', 'Normal', '20', '26/01/2022', '03/01/2022', '?', '5226220', '29/01/2022', '33423', 'MELISSA PAPETE WIDE AD', '24', '2.451,84', '\xa0', 'L', 'L', 'L'], ['490766802', 'Faturado', 'Normal', '21', '19/01/2022', '03/01/2022', '?', '1529257', '19/01/2022', '33427', 'MELISSA SHINY HEEL AD', '24', '2.838,96', '\xa0', 'L', 'L', 'L'], ['490766803', 'Faturado', 'Normal', '20', '19/01/2022', '03/01/2022', '?', '5222890', '19/01/2022', '33429', 'MELISSA SHINY AD', '36', '1.935,36', '\xa0', 'L', 'L', 'L'], ['490766804', 'Faturado', 'Normal', '20', '19/01/2022', '03/01/2022', '?', '5222889', '19/01/2022', '33431', 'MELISSA BRIGHTNESS AD', '36', '3.290,40', '\xa0', 'L', 'L', 'L'], ['490766805', 'Faturado', 'Normal', '21', '17/01/2022', '03/01/2022', '?', '1528851', '18/01/2022', '33521', 'MINI MELISSA POSSESSION SHINY INF', '24', '1.935,60', '\xa0', 'L', 'L', 'L'], ['490766806', 'Faturado', 'Normal', '21', '17/01/2022', '03/01/2022', '?', '1528541', '18/01/2022', '33522', 'MINI MELISSA POSSESSION SHINY BB', '24', '1.419,36', '\xa0', 'L', 'L', 'L'], ['490766807', 'Faturado', 'Normal', '20', '17/01/2022', '03/01/2022', '?', '5220779', '18/01/2022', '33528', 'MELISSA SUN LONG BEACH AD', '36', '1.161,00', '\xa0', 'L', 'L', 'L'], ['490766808', 'Faturado', 'Normal', '20', '19/01/2022', '03/01/2022', '?', '5222986', '20/01/2022', '33530', 'MELISSA SUN RODEO AD', '30', '1.451,70', '\xa0', 'L', 'L', 'L'], ['490766809', 'Programado', 'Normal', '40', '26/02/2022', '03/01/2022', '?', '\xa0', '?', '33531', 'MELISSA FLIP FLOP FREE AD', '24', '2.193,60', '\xa0', 'L', 'L', 'L'], ['490766810', 'Faturado', 'Normal', '21', '17/01/2022', '03/01/2022', '?', '1528808', '18/01/2022', '33538', 'MELISSA SOLAR II + BOBO AD', '12', '1.225,92', '\xa0', 'L', 'L', 'L'], ['490766811', 'Faturado', 'Normal', '21', '17/01/2022', '03/01/2022', '?', '1528705', '18/01/2022', '33539', 'MELISSA HARMONIC CHROME IX AD', '36', '2.129,04', '\xa0', 'L', 'L', 'L'], ['490766812', 'Faturado', 'Normal', '21', '24/01/2022', '03/01/2022', '?', '1532188', '26/01/2022', '33542', 'MELISSA MULE III AD', '12', '1.612,92', '\xa0', 'L', 'L', 'L'], ['490766813', 'Faturado', 'Normal', '21', '17/01/2022', '03/01/2022', '?', '1528832', '18/01/2022', '33546', 'MINI MELISSA MAR SANDAL JELLY POP INF', '24', '1.935,60', '\xa0', 'L', 'L', 'L'], ['490766814', 'Faturado', 'Normal', '20', '24/01/2022', '03/01/2022', '?', '5225125', '25/01/2022', '33547', 'MELISSA BIKINI STRIPE AD', '24', '1.290,48', '\xa0', 'L', 'L', 'L'], ['490766815', 'Faturado', 'Normal', '20', '26/01/2022', '03/01/2022', '?', '5226223', '29/01/2022', '33557', 'MELISSA SUN CITY WALK AD', '18', '871,02', '\xa0', 'L', 'L', 'L'], ['490766816', 'Faturado', 'Normal', '21', '17/01/2022', '03/01/2022', '?', '1528680', '18/01/2022', '33559', 'MINI MELISSA DORA III BB', '18', '967,86', '\xa0', 'L', 'L', 'L'], ['490766817', 'Faturado', 'Normal', '21', '21/01/2022', '03/01/2022', '21/01/2022', '1530912', '21/01/2022', '33571', 'MELISSA THE REAL JELLY SANDAL AD', '12', '838,80', '\xa0', 'L', 'L', 'L'], ['490766818', 'Faturado', 'Normal', '21', '17/01/2022', '03/01/2022', '?', '1528793', '18/01/2022', '33580', 'MINI MELISSA SUNNY BB', '30', '1.613,10', '\xa0', 'L', 'L', 'L'], ['490766819', 'Faturado', 'Normal', '21', '17/01/2022', '03/01/2022', '?', '1528611', '18/01/2022', '33587', 'MELISSA FUNKY AD', '24', '3.484,08', '\xa0', 'L', 'L', 'L'], ['490766820', 'Faturado', 'Normal', '20', '24/01/2022', '03/01/2022', '?', '5225139', '25/01/2022', '33614', 'MELISSA FLIP FLOP SLIM III AD', '18', '1.258,02', '\xa0', 'L', 'L', 'L'], ['490766821', 'Faturado', 'Normal', '20', '24/01/2022', '03/01/2022', '?', '5225124', '25/01/2022', '33617', 'MINI MELISSA COSMIC SANDAL INF', '30', '2.580,90', '\xa0', 'L', 'L', 'L'], ['490766822', 'Faturado', 'Normal', '21', '24/01/2022', '03/01/2022', '?', '1532239', '24/01/2022', '33634', 'MELISSA SEDUCTION VI AD', '12', '967,80', '\xa0', 'L', 'L', 'L'], ['490766823', 'Faturado', 'Normal', '21', '21/01/2022', '03/01/2022', '21/01/2022', '1530728', '21/01/2022', '33646', 'MELISSA THE REAL JELLY SLIDE AD', '12', '774,24', '\xa0', 'L', 'L', 'L'], ['490766824', 'Faturado', 'Normal', '20', '24/01/2022', '03/01/2022', '?', '5225077', '24/01/2022', '33656', 'MELISSA DARE STRAP + CAMILA COUTINHO AD', '30', '2.903,10', '\xa0', 'L', 'L', 'L'], ['490766825', 'Faturado', 'Normal', '21', '17/01/2022', '03/01/2022', '?', '1528769', '18/01/2022', '33657', 'MELISSA T-BAR STRAP + CAMILA COUTINHO AD', '24', '1.548,48', '\xa0', 'L', 'L', 'L'], ['490766826', 'Faturado', 'Normal', '21', '17/01/2022', '03/01/2022', '?', '1528575', '18/01/2022', '33682', 'MINI MELISSA ULTRAGIRL SWEET X BB', '24', '1.548,48', '\xa0', 'L', 'L', 'L'], ['490766827', 'Faturado', 'Normal', '20', '17/01/2022', '03/01/2022', '?', '5220778', '18/01/2022', '33694', 'MELISSA SUN VENICE SHINY AD', '36', '1.355,04', '\xa0', 'L', 'L', 'L'], ['490766828', 'Faturado', 'Normal', '21   ', '24/01/2022', '03/01/2022', '?', '1532341', '26/01/2022', '33771', 'MELISSA AIRBUBBLE FLIP FLOP AD', '24', '1.677,60', '\xa0', 'L', 'L', 'L'], ['490766829', 'Programado', 'Normal', '40', '26/02/2022', '03/01/2022', '?', '\xa0', '?', '33772', 'MELISSA FREE PLATFORM AD', '18', '2.129,22', '\xa0', 'L', 'L', 'L'], ['490767001', 'Faturado', 'Normal', '20', '28/12/2021', '03/01/2022', '?', '5213549', '30/12/2021', '34102', 'MINIATURA MELISSA CORACAO XIII SP', '60', '561,00', '\xa0', 'L', 'L', 'L'], ['490767002', 'Faturado', 'Normal', '20', '24/01/2022', '03/01/2022', '?', '5225144', '25/01/2022', '34305', 'MELISSA SUN SANTA MONICA II', '12', '586,56', '\xa0', 'L', 'L', 'L']]))
