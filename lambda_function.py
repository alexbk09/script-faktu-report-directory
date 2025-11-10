import json
from openpyxl import Workbook
import boto3
from io import BytesIO
import os
import pymysql
from datetime import datetime, date
import time
# Configura el cliente de S3
s3_client = boto3.client('s3', region_name='us-east-1')

# Configura la conexión a la base de datos
def get_db_connection():
    return pymysql.connect(
        host=os.getenv('DB_HOST'),
        user=os.getenv('DB_USERNAME'),
        password=os.getenv('DB_PASSWORD'),
        database=os.getenv('DB_DATABASE'),
        port=int(os.getenv('DB_PORT'))
    ) 

def change_status_download(status, path):
    connection = get_db_connection()
    cursor = connection.cursor(pymysql.cursors.DictCursor)
    query = "UPDATE download_file SET status = %s WHERE path = %s "
    cursor.execute(query, (status, path))
    connection.commit()
    cursor.close()
    connection.close()


def upload_file_s3(path, bucket, object_name=None):
    # Subir el archivo a S3
    try:
        connection = get_db_connection()
        cursor = connection.cursor(pymysql.cursors.DictCursor)
        with open('/tmp/reporte_directory.xlsx', 'rb') as file:
            s3_client.upload_fileobj(
                file,
                os.getenv('S3_BUCKET'),
                path
            )
        change_status_download('success', path)
        print("Archivo subido exitosamente a S3")
    except Exception as e:
        print(f"Error al subir el archivo a S3: {str(e)}")
        change_status_download('error', path)
    finally:
        # Cerrar la conexión a la base de datos
        cursor.close()
        connection.close()

def handler(event, context):
    try:
        print("Evento recibido:", event)
        connection = get_db_connection()
        cursor = connection.cursor(pymysql.cursors.DictCursor)

        
        # obtener los parametros del evento
        searchContact = event.get('searchContact', '')
        sector = event.get('sector', '')
        types = event.get('type', '')
        unit = event.get('unit', '')
        search = event.get('search', '')
        path = event.get('path', '')

        # Conectar a la base de datos
        connection = get_db_connection()
        cursor = connection.cursor(pymysql.cursors.DictCursor)

        where_clauses = []
        if searchContact:
            where_clauses.append(f"directory_udc_contact.name >= '{searchContact}'")
        if sector:
            where_clauses.append(f"directory.sector >= '{sector}'")
        if types:
            where_clauses.append(f"directory.type >= '{types}'")
        if unit:
            where_clauses.append(f"directory_udc.unidad = '{unit}'")
        if search:
            where_clauses.append(f"(directory.company_name = '{search}')")
        # Ejecutar una consulta SQL para obtener los datos
        query = """
            SELECT 
                directory.*,
                directory_udc.unidad,
                directory_udc.cod_unidad,
                directory_udc.region_unidad,
                directory_udc.id as 'udc_id',
                directory_udc_contact.id as 'udc_contact_id',
                directory_udc_contact.email as 'udc_contact_email',
                directory_udc_contact.phone as 'udc_contact_phone',
                directory_udc_contact.name as 'udc_contact_name',
                directory_udc_contact.departament as 'departament',
                directory_udc_contact.cargo as 'cargo'
            FROM 
                directory
            LEFT JOIN
                directory_udc ON directory_udc.directory_id = directory.id
            LEFT JOIN 
                directory_udc_contact ON directory_udc_contact.udc_id = directory_udc.id
            """
        if where_clauses:
            query += " WHERE " + " AND ".join(where_clauses)
        query += " ORDER BY directory.created_at DESC;"
        print("Consulta SQL:", query)
        cursor.execute(query)
        results = cursor.fetchall()

        # Crear un nuevo libro de Excel
        wb = Workbook()
        ws = wb.active

        # Escribir las cabeceras
        headers = [
            'Tipo','Sector de la empresa','Nombre','Rut','Telefono','email','Giro','Direccion','Comuna','Web',
            'Cod unidad de compra','Unidad de compra','region unidad de compra',
            'Nombre de contacto','Departamento del contacto','Cargo del contacto','Telefono','email'
        ]

        for col_num, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_num, value=header)
        # Escribir los datos en la hoja de cálculo
        for row_index, row in enumerate(results, start=2):
            ws.cell(row=row_index, column=1, value=row.get('type')) # empresa
            ws.cell(row=row_index, column=2, value=row.get('type_bussiness')) # deudor
            ws.cell(row=row_index, column=3, value=row.get('company_name')) # Asunto
            ws.cell(row=row_index, column=4, value=row.get('rut')) # Deudor
            ws.cell(row=row_index, column=5, value=row.get('phone')) # Rut de deudor
            ws.cell(row=row_index, column=6, value=row.get('email')) # Tipo Documento
            ws.cell(row=row_index, column=7, value=row.get('giro'))  # Tipo Documento
            ws.cell(row=row_index, column=8, value=row.get('address'))  # Tipo Documento
            ws.cell(row=row_index, column=9, value=row.get('comuna'))  # Tipo Documento
            ws.cell(row=row_index, column=10, value=row.get('web'))  # Tipo Documento
            ws.cell(row=row_index, column=11, value=row.get('cod_unidad', 'N/A')) # Tipo Documento
            ws.cell(row=row_index, column=12, value=row.get('unidad', 'N/A')) # Tipo Documento
            ws.cell(row=row_index, column=13, value=row.get('region_unidad', 'N/A')) # Tipo Documento
            ws.cell(row=row_index, column=14, value=row.get('udc_contact_name', 'N/A')) # Tipo Documento
            ws.cell(row=row_index, column=15, value=row.get('departament', 'N/A'))  # Tipo Documento
            ws.cell(row=row_index, column=16, value=row.get('cargo', 'N/A')) # N° Folio
            ws.cell(row=row_index, column=17, value=row.get('udc_contact_phone', 'N/A')) # Fecha ingreso
            ws.cell(row=row_index, column=18, value=row.get('udc_contact_email', 'N/A')) # Tipo de servicio
        # Guardar el archivo Excel localmente
        wb.save('/tmp/reporte_directory.xlsx')
        print("Archivo Excel generado exitosamente: reporte_directory.xlsx")

        upload_file_s3(path, os.getenv('S3_BUCKET'))

        return {
            'statusCode': 200,
            'body': 'Archivo subido exitosamente a S3'
        }
    except Exception as e:
        print(f"Error en el handler: {str(e)}")
        print(f"Línea del error: {e.__traceback__.tb_lineno}")
        change_status_download('error', path)
        return {
            'statusCode': 500,
            'body': f'Error en el handler: {str(e)}'
        }
