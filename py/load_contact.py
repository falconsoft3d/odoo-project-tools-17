# -- coding: utf-8 --
#************#
# Modificar precios
#************#

import xmlrpc.client
import json
import time
from itertools import chain
import logging
import base64
from io import BytesIO
import xlrd

# Servidor Demo
demo_host = '127.0.0.1'
demo_port = 8069
demo_db = 'db16-spain'
demo_user = 'admin'
demo_password = 'x1234567890'


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
        file = '2000_datos_clientes_proveedores.xls'

        # Leer archivo XLS con las columnas del producto de odoo
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheet_by_index(0)
        print('Leyendo archivo: ', file)
        for row in range(1, sheet.nrows):
            if row > 100000000000:
                continue
            
            number = sheet.cell(row, 0).value
            nif = sheet.cell(row, 1).value
            name = sheet.cell(row, 2).value
            street = sheet.cell(row, 3).value
            estado = sheet.cell(row, 4).value
            ciudad = sheet.cell(row, 5).value
            country = sheet.cell(row, 6).value

            print(number)
            print(nif)

            
            country_id = False
            if country == 'MEXICO':
                country_id = 156

            
            email = sheet.cell(row, 7).value
            phone = sheet.cell(row, 10).value
            
            state_id = False
            if estado == 'SINALOA':
                state_id  = '508'
            elif estado == 'SONORA':
                state_id  = '510'
            elif estado == 'JALISCO':
                state_id  = '498'
            elif estado == 'NUEVO LEON':
                state_id  = '503'
            elif estado == 'COLIMA':
                state_id  = '489'
                
            
            vals = {
                'name': name,
                'street': street,
                'city': ciudad,
                'email': email,
                'phone': phone,
                'country_id': country_id,
                'state_id': state_id,
            }
            print(vals)

            # Vemos si existe ese proveedor
            proveedor = object_proxy.execute_kw(demo_db, demo_uid, demo_password, 'res.partner', 'search_read', [[['comment', '=', number]]])
            if not proveedor:
                vals.update({
                    'vat': nif,
                })
                proveedor_id = object_proxy.execute_kw(demo_db, demo_uid, demo_password, 'res.partner', 'create', [vals])
                if proveedor_id:
                    print(str(row) + '- Proveedor creado: ', nif)
            else:
                print('Proveedor ya existe: ', nif)

        print('Cargando archivo: ', file)

def main():
    print("Ha comenzado el proceso")
    _create_proveedor(True)
    print('Ha finalizado la carga tabla')
main()
