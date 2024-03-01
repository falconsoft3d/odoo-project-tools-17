import xmlrpc.client
import json
import time
from itertools import chain
import logging
import base64
from io import BytesIO
import xlrd

# Servidor Demo
demo_host = 'creekside.bim20.com'
demo_port = 8069
demo_db = 'db16-mexico'
demo_user = 'admin'
demo_password = 'asas7778884443xddSSdd'


demo_url = 'http://%s:%d/xmlrpc/2/' % (demo_host, demo_port)
demo_common_proxy = xmlrpc.client.ServerProxy(demo_url + 'common')
object_proxy = xmlrpc.client.ServerProxy(demo_url + 'object')
demo_uid = demo_common_proxy.login(demo_db, demo_user, demo_password)


# File XLS con proveedores

if demo_uid:
    print('Conectado al servidor demo')


def _create_proveedor(estado):
    if estado is True:
        print('Cargando datos')
        file = '3000_datos_productos_REV.xls'

        # Leer archivo XLS con las columnas del producto de odoo
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheet_by_index(0)
        print('Leyendo archivo: ', file)
        for row in range(1, sheet.nrows):
            if row > 100000000000:
                continue
            
            
            name = sheet.cell(row, 1).value
            tipo = sheet.cell(row, 2).value
            default_code = sheet.cell(row, 3).value
            unidad_medida = sheet.cell(row, 4).value
            list_price = sheet.cell(row, 6).value
            standard_price = sheet.cell(row, 7).value
            categoria = sheet.cell(row, 8).value
            sub_categoria = sheet.cell(row, 8).value
            
            
            
            if tipo == 'ALMACENABLE':
                type = 'product'
            else:
                type = 'service'

            
            categoryp_id = object_proxy.execute_kw(demo_db, demo_uid, demo_password, 'product.category', 'search_read', [[['name', '=', categoria]]])
            if not categoryp_id:
                categoryp_id = object_proxy.execute_kw(demo_db, demo_uid, demo_password, 'product.category', 'create', [{'name': categoria}])
                if categoryp_id:
                    print('Categoria padre creada: ', categoryp_id)
            else:
                categoryp_id = categoryp_id[0]['id']
            
            
            category_id = object_proxy.execute_kw(demo_db, demo_uid, demo_password, 'product.category', 'search_read', [[['name', '=', sub_categoria]]])
            if not category_id:
                category_id = object_proxy.execute_kw(demo_db, demo_uid, demo_password, 'product.category', 'create', 
                                                      [
                                                          {
                                                           'name': sub_categoria,
                                                           'parent_id': categoryp_id,
                                                           }
                                                       ]
                                                      )
                if category_id:
                    print('Categoria creada: ', sub_categoria)
            else:
                category_id = category_id[0]['id']

            


            print("...........")
            print(row)
            print(sub_categoria)
            print(category_id)

            product = object_proxy.execute_kw(demo_db, demo_uid, demo_password, 'product.template', 'search_read', [[['default_code', '=', default_code]]])
            if not product:
                vals = {
                    'name': name,
                    'type': type,
                    'default_code': default_code,
                    'list_price': list_price,
                    'standard_price': standard_price,
                    'categ_id': category_id,
                }
                product_id = object_proxy.execute_kw(demo_db, demo_uid, demo_password, 'product.template', 'create', [vals])
                if product_id:
                    print(str(row) + '- Producto creado: ', default_code)

        print('Cargando archivo: ', file)

def main():
    print("Ha comenzado el proceso")
    _create_proveedor(True)
    print('Ha finalizado la carga tabla')
main()
