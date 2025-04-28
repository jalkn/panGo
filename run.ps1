#!/bin/bash

$GREEN = "Green"
$YELLOW = "Yellow"

function createPeriod {
    Write-Host "🏗️ Creating periodoBR" -ForegroundColor $YELLOW
    # password
    Set-Content -Path "models/period.py" -Value @"
from openpyxl import Workbook
from openpyxl.styles import Font

def create_excel_file():
    # Create a new workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active
    
    # Define the header row
    headers = [
        "Id", "Activo", "Año", "FechaFinDeclaracion", 
        "FechaInicioDeclaracion", "Año declaracion"
    ]
    
    # Write the headers to the first row and make them bold
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = Font(bold=True)
    
    # Data rows
    data = [
        [2, True, "Friday, January 01, 2021", "4/30/2022", "6/1/2021", "2,021"],
        [6, True, "Saturday, January 01, 2022", "3/31/2023", "10/19/2022", "2,022"],
        [7, True, "Sunday, January 01, 2023", "5/12/2024", "11/1/2023", "2,023"],
        [8, True, "Monday, January 01, 2024", "1/1/2025", "10/2/2024", "2,024"]
    ]
    
    # Write data rows
    for row_num, row_data in enumerate(data, 2):  # Start from row 2
        for col_num, cell_value in enumerate(row_data, 1):
            ws.cell(row=row_num, column=col_num, value=cell_value)
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Save the workbook
    filename = "src/periodoBR.xlsx"
    wb.save(filename)
    print(f"Excel file '{filename}' created successfully!")

if __name__ == "__main__":
    create_excel_file()
"@
}

function createPassKey {
    Write-Host "🏗️ Creating passKey" -ForegroundColor $YELLOW
    # password
    Set-Content -Path "models/passKey.py" -Value @"
import msoffcrypto
import openpyxl
import sys
import os
import json
import getpass
from datetime import datetime

def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

def remove_excel_password(input_file, output_file=None, open_password=None, modify_password=None):
    """Handle both open and modify passwords"""
    try:
        with open(input_file, "rb") as file:
            office_file = msoffcrypto.OfficeFile(file)
            
            # Try both passwords if provided
            if open_password or modify_password:
                try:
                    if open_password:
                        office_file.load_key(password=open_password)
                    elif modify_password:
                        office_file.load_key(password=modify_password)
                except Exception as e:
                    print(f"Password error: {str(e)}")
                    return False
            else:
                # Try without password if file isn't encrypted
                try:
                    office_file.load_key(password=None)
                except:
                    print("File is password protected but no password provided")
                    return False
            
            with open(output_file, "wb") as decrypted:
                office_file.decrypt(decrypted)
        
        print(f"File successfully processed. Saved to '{output_file}'")
        return True
        
    except Exception as e:
        if 'password' in str(e).lower():
            raise Exception("La contraseña proporcionada es incorrecta")
        raise

def add_fk_id_estado(input_file, output_file):
    try:
        wb = openpyxl.load_workbook(input_file, read_only=True)  # Use read-only
        ws = wb.active
        
        # Find header row
        headers = [cell.value for cell in ws[1]]
        
        # Add fkIdEstado if needed
        if 'fkIdEstado' not in headers:
            headers.append('fkIdEstado')
            fk_col = len(headers)
        else:
            fk_col = headers.index('fkIdEstado') + 1
        
        # Convert to JSON in chunks
        data = []
        chunk_size = 1000
        print(f"Total rows to process: {ws.max_row}")
        
        for row_num, row in enumerate(ws.iter_rows(min_row=2), start=2):
            if row_num % chunk_size == 0:
                print(f"Processed {row_num} rows")
                
            row_data = {headers[i]: cell.value for i, cell in enumerate(row)}
            row_data['fkIdEstado'] = 1
            data.append(row_data)
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
            
        print(f"Successfully processed {len(data)} rows")
        return True
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return False

if __name__ == "__main__":
    try:
        input("\nPega el archivo de Excel en la carpeta 'src/' y asegúrate de nombrarlo 'dataHistoricaPBI.xlsx'. Presiona Enter cuando esté listo...")

        input_excel_file = "src/dataHistoricaPBI.xlsx"
        if not os.path.exists(input_excel_file):
            log_message(f"ERROR: No se encontró el archivo '{input_excel_file}'")
            log_message("Por favor verifica:")
            log_message("1. Que existe el directorio 'src/'")
            log_message("2. Que el archivo está en 'src/'")
            log_message("3. Que el archivo se llama 'dataHistoricaPBI.xlsx'")
            sys.exit(1)

        output_excel_file = "src/data.xlsx"
        output_json_file = "src/fk1data.json"

        if remove_excel_password(input_excel_file, output_excel_file):
            if add_fk_id_estado(output_excel_file, output_json_file):
                log_message("\nPROCESO COMPLETADO EXITOSAMENTE")
                log_message(f"- Archivo desencriptado: {output_excel_file}")
                log_message(f"- Archivo JSON generado: {output_json_file}")
            else:
                log_message("\nPROCESO PARCIALMENTE COMPLETADO")
                log_message(f"- Archivo desencriptado: {output_excel_file}")
                log_message("- Falló la generación del archivo JSON")
        else:
            log_message("\nPROCESO FALLIDO")
            log_message("- No se pudo desencriptar el archivo de entrada")
    except KeyboardInterrupt:
        log_message("\nOperación cancelada por el usuario")
    except Exception as e:
        log_message(f"\nERROR INESPERADO: {str(e)}")

"@
}

function createCats {
    Write-Host "🏗️ Creating Categories" -ForegroundColor $YELLOW
    
    # Banks
    Set-Content -Path "models/cats.py" -Value @"
import pandas as pd
from datetime import datetime

# Shared constants and functions
TRM_DICT = {
    2020: 3432.50,
    2021: 3981.16,
    2022: 4810.20,
    2023: 4780.38,
    2024: 4409.00
}

CURRENCY_RATES = {
    2020: {
        'EUR': 1.141, 'GBP': 1.280, 'AUD': 0.690, 'CAD': 0.746,
        'HNL': 0.0406, 'AWG': 0.558, 'DOP': 0.0172, 'PAB': 1.000,
        'CLP': 0.00126, 'CRC': 0.00163, 'ARS': 0.0119, 'ANG': 0.558,
        'COP': 0.00026,  'BBD': 0.50, 'MXN': 0.0477, 'BOB': 0.144, 'BSD': 1.00,
        'GYD': 0.0048, 'UYU': 0.025, 'DKK': 0.146, 'KYD': 1.20, 'BMD': 1.00, 
        'VEB': 0.0000000248, 'VES': 0.000000248, 'BRL': 0.187, 'NIO': 0.0278
    },
    2021: {
        'EUR': 1.183, 'GBP': 1.376, 'AUD': 0.727, 'CAD': 0.797,
        'HNL': 0.0415, 'AWG': 0.558, 'DOP': 0.0176, 'PAB': 1.000,
        'CLP': 0.00118, 'CRC': 0.00156, 'ARS': 0.00973, 'ANG': 0.558,
        'COP': 0.00027, 'BBD': 0.50, 'MXN': 0.0492, 'BOB': 0.141, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.024, 'DKK': 0.155, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0.00000000002, 'VES': 0.00000002, 'BRL': 0.192, 'NIO': 0.0285
    },
    2022: {
        'EUR': 1.051, 'GBP': 1.209, 'AUD': 0.688, 'CAD': 0.764,
        'HNL': 0.0408, 'AWG': 0.558, 'DOP': 0.0181, 'PAB': 1.000,
        'CLP': 0.00117, 'CRC': 0.00155, 'ARS': 0.00597, 'ANG': 0.558,
        'COP': 0.00021, 'BBD': 0.50, 'MXN': 0.0497, 'BOB': 0.141, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.025, 'DKK': 0.141, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0, 'VES': 0.000000001, 'BRL': 0.196, 'NIO': 0.0267
    },
    2023: {
        'EUR': 1.096, 'GBP': 1.264, 'AUD': 0.676, 'CAD': 0.741,
        'HNL': 0.0406, 'AWG': 0.558, 'DOP': 0.0177, 'PAB': 1.000,
        'CLP': 0.00121, 'CRC': 0.00187, 'ARS': 0.00275, 'ANG': 0.558,
        'COP': 0.00022, 'BBD': 0.50, 'MXN': 0.0564, 'BOB': 0.143, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.025, 'DKK': 0.148, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0, 'VES': 0.000000001, 'BRL': 0.194, 'NIO': 0.0267
    },
    2024: {
        'EUR': 1.093, 'GBP': 1.267, 'AUD': 0.674, 'CAD': 0.742,
        'HNL': 0.0405, 'AWG': 0.558, 'DOP': 0.0170, 'PAB': 1.000,
        'CLP': 0.00111, 'CRC': 0.00192, 'ARS': 0.00121, 'ANG': 0.558,
        'COP': 0.00022, 'BBD': 0.50, 'MXN': 0.0547, 'BOB': 0.142, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.024, 'DKK': 0.147, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0, 'VES': 0.000000001, 'BRL': 0.190, 'NIO': 0.0260 }
}

def get_trm(year):
    """Gets TRM for a given year from the dictionary"""
    return TRM_DICT.get(year)

def get_exchange_rate(currency_code, year):
    """Gets exchange rate for a given currency and year from the dictionary"""
    year_rates = CURRENCY_RATES.get(year)
    if year_rates:
        return year_rates.get(currency_code)
    return None

def get_currency_code(moneda_text):
    """Extracts the currency code from the 'Texto Moneda' field"""
    currency_mapping = {
        'HNL -Lempira hondureño': 'HNL',
        'EUR - Euro': 'EUR',
        'AWG - Florín holandés o de Aruba': 'AWG',
        'DOP - Peso dominicano': 'DOP',
        'PAB -Balboa panameña': 'PAB', 
        'CLP - Peso chileno': 'CLP',
        'CRC - Colón costarricense': 'CRC',
        'ARS - Peso argentino': 'ARS',
        'AUD - Dólar australiano': 'AUD',
        'ANG - Florín holandés': 'ANG',
        'CAD -Dólar canadiense': 'CAD',
        'GBP - Libra esterlina': 'GBP',
        'USD - Dolar estadounidense': 'USD',
        'COP - Peso colombiano': 'COP',
        'BBD - Dólar de Barbados o Baja': 'BBD',
        'MXN - Peso mexicano': 'MXN',
        'BOB - Boliviano': 'BOB',
        'BSD - Dolar bahameño': 'BSD',
        'GYD - Dólar guyanés': 'GYD',
        'UYU - Peso uruguayo': 'UYU',
        'DKK - Corona danesa': 'DKK',
        'KYD - Dólar de las Caimanes': 'KYD',
        'BMD - Dólar de las Bermudas': 'BMD',
        'VEB - Bolívar venezolano': 'VEB',  
        'VES - Bolívar soberano': 'VES',  
        'BRL - Real brasilero': 'BRL',  
        'NIO - Córdoba nicaragüense': 'NIO',
    }
    return currency_mapping.get(moneda_text)

def get_valid_year(row, periodo_df):
    """Extracts a valid year, handling missing values and format variations."""
    try:
        fkIdPeriodo = pd.to_numeric(row['fkIdPeriodo'], errors='coerce')
        if pd.isna(fkIdPeriodo):  # Handle missing fkIdPeriodo
            print(f"Warning: Missing fkIdPeriodo at index {row.name}. Skipping row.")
            return None

        matching_row = periodo_df[periodo_df['Id'] == fkIdPeriodo]
        if matching_row.empty:
            print(f"Warning: No matching Id found in periodoBR.xlsx for fkIdPeriodo {fkIdPeriodo} at index {row.name}. Skipping row.")
            return None

        year_str = matching_row['Año'].iloc[0]

        try:
            year = int(year_str)  # Try direct conversion to integer
            return year
        except (ValueError, TypeError):
            try:
                year = pd.to_datetime(year_str, errors='coerce').year  # Try datetime conversion, handle errors gracefully
                if pd.isna(year):  # check for NaT which occurs when conversion fails.
                    raise ValueError  # If conversion failed re-raise a ValueError.
                return year

            except ValueError:
                print(f"Warning: Invalid year format '{year_str}' for fkIdPeriodo {fkIdPeriodo} at index {row.name}. Skipping row.")
                return None

    except Exception as e:
        print(f"Error in get_valid_year for fkIdPeriodo {fkIdPeriodo} at index {row.name}: {e}")
        return None

def analyze_banks(file_path, output_file_path, periodo_file_path):
    """Analyze bank account data"""
    df = pd.read_excel(file_path)
    periodo_df = pd.read_excel(periodo_file_path)

    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario',
        'Nombre', 'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Banco - Entidad', 'Banco - Tipo Cuenta', 'Texto Moneda',
        'Banco - fkIdPaís', 'Banco - Nombre País',
        'Banco - Saldo', 'Banco - Comentario'
    ]
    
    banks_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Banco', maintain_columns].copy()
    banks_df = banks_df[banks_df['fkIdEstado'] != 1]
    
    banks_df['Banco - Saldo COP'] = 0.0
    banks_df['TRM Aplicada'] = None
    banks_df['Tasa USD'] = None
    banks_df['Año Declaración'] = None 
    
    for index, row in banks_df.iterrows():
        try:
            year = get_valid_year(row, periodo_df)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index} and fkIdPeriodo {row['fkIdPeriodo']}. Skipping row.")
                banks_df.loc[index, 'Año Declaración'] = "Año no encontrado"
                continue 
                
            banks_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                banks_df.loc[index, 'Banco - Saldo COP'] = float(row['Banco - Saldo'])
                banks_df.loc[index, 'TRM Aplicada'] = 1.0
                banks_df.loc[index, 'Tasa USD'] = None
                continue
                
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Banco - Saldo']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    banks_df.loc[index, 'Banco - Saldo COP'] = cop_amount
                    banks_df.loc[index, 'TRM Aplicada'] = trm
                    banks_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            banks_df.loc[index, 'Año Declaración'] = "Error de procesamiento"
            continue
    
    banks_df.to_excel(output_file_path, index=False)

def analyze_debts(file_path, output_file_path, periodo_file_path):
    """Analyze debts data"""
    df = pd.read_excel(file_path)
    periodo_df = pd.read_excel(periodo_file_path)

    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Pasivos - Entidad Personas',
        'Pasivos - Tipo Obligación', 'fkIdMoneda', 'Texto Moneda',
        'Pasivos - Valor', 'Pasivos - Comentario', 'Pasivos - Valor COP'
    ]
    
    debts_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Pasivo', maintain_columns].copy()
    debts_df = debts_df[debts_df['fkIdEstado'] != 1]
    
    debts_df['Pasivos - Valor COP'] = 0.0
    debts_df['TRM Aplicada'] = None
    debts_df['Tasa USD'] = None
    debts_df['Año Declaración'] = None 
    
    for index, row in debts_df.iterrows():
        try:
            year = get_valid_year(row, periodo_df)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
            
            debts_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                debts_df.loc[index, 'Pasivos - Valor COP'] = float(row['Pasivos - Valor'])
                debts_df.loc[index, 'TRM Aplicada'] = 1.0
                debts_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Pasivos - Valor']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    debts_df.loc[index, 'Pasivos - Valor COP'] = cop_amount
                    debts_df.loc[index, 'TRM Aplicada'] = trm
                    debts_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            debts_df.loc[index, 'Año Declaración'] = "Error de procesamiento"
            continue

    debts_df.to_excel(output_file_path, index=False)

def analyze_goods(file_path, output_file_path, periodo_file_path):
    """Analyze goods/patrimony data"""
    df = pd.read_excel(file_path)
    periodo_df = pd.read_excel(periodo_file_path)

    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Patrimonio - Activo', 'Patrimonio - % Propiedad',
        'Patrimonio - Propietario', 'Patrimonio - Valor Comercial',
        'Patrimonio - Comentario',
        'Patrimonio - Valor Comercial COP', 'Texto Moneda'
    ]
    
    goods_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Patrimonio', maintain_columns].copy()
    goods_df = goods_df[goods_df['fkIdEstado'] != 1]
    
    goods_df['Patrimonio - Valor COP'] = 0.0
    goods_df['TRM Aplicada'] = None
    goods_df['Tasa USD'] = None
    goods_df['Año Declaración'] = None 
    
    for index, row in goods_df.iterrows():
        try:
            year = get_valid_year(row, periodo_df)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
                
            goods_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                goods_df.loc[index, 'Patrimonio - Valor COP'] = float(row['Patrimonio - Valor Comercial'])
                goods_df.loc[index, 'TRM Aplicada'] = 1.0
                goods_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Patrimonio - Valor Comercial']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    goods_df.loc[index, 'Patrimonio - Valor COP'] = cop_amount
                    goods_df.loc[index, 'TRM Aplicada'] = trm
                    goods_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            continue
        
    goods_df['Patrimonio - Valor Corregido'] = goods_df['Patrimonio - Valor COP'] * (goods_df['Patrimonio - % Propiedad'] / 100)
    
    # Rename columns for consistency
    rename_dict = {
        'Patrimonio - Valor Corregido': 'Bienes - Valor Corregido',
        'Patrimonio - Valor Comercial COP': 'Bienes - Valor Comercial COP',
        'Patrimonio - Comentario': 'Bienes - Comentario',
        'Patrimonio - Valor Comercial': 'Bienes - Valor Comercial',
        'Patrimonio - Propietario': 'Bienes - Propietario',
        'Patrimonio - % Propiedad': 'Bienes - % Propiedad',
        'Patrimonio - Activo': 'Bienes - Activo',
        'Patrimonio - Valor COP': 'Bienes - Valor COP'
    }
    goods_df = goods_df.rename(columns=rename_dict)
    
    goods_df.to_excel(output_file_path, index=False)

def analyze_incomes(file_path, output_file_path, periodo_file_path):
    """Analyze income data"""
    df = pd.read_excel(file_path)
    periodo_df = pd.read_excel(periodo_file_path)

    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Ingresos - fkIdConcepto', 'Ingresos - Texto Concepto',
        'Ingresos - Valor', 'Ingresos - Comentario', 'Ingresos - Otros',
        'Ingresos - Valor_COP', 'Texto Moneda'
    ]

    incomes_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Ingreso', maintain_columns].copy()
    incomes_df = incomes_df[incomes_df['fkIdEstado'] != 1]
    
    incomes_df['Ingresos - Valor COP'] = 0.0
    incomes_df['TRM Aplicada'] = None
    incomes_df['Tasa USD'] = None
    incomes_df['Año Declaración'] = None 
    
    for index, row in incomes_df.iterrows():
        try:
            year = get_valid_year(row, periodo_df)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
            
            incomes_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                incomes_df.loc[index, 'Ingresos - Valor COP'] = float(row['Ingresos - Valor'])
                incomes_df.loc[index, 'TRM Aplicada'] = 1.0
                incomes_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Ingresos - Valor']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    incomes_df.loc[index, 'Ingresos - Valor COP'] = cop_amount
                    incomes_df.loc[index, 'TRM Aplicada'] = trm
                    incomes_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            continue
    
    incomes_df.to_excel(output_file_path, index=False)

def analyze_investments(file_path, output_file_path, periodo_file_path):
    """Analyze investment data"""
    df = pd.read_excel(file_path)
    periodo_df = pd.read_excel(periodo_file_path)

    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Inversiones - Tipo Inversión', 'Inversiones - Entidad',
        'Inversiones - Valor', 'Inversiones - Comentario',
        'Inversiones - Valor COP', 'Texto Moneda'
    ]
    
    invest_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Inversión', maintain_columns].copy()
    invest_df = invest_df[invest_df['fkIdEstado'] != 1]
    
    invest_df['Inversiones - Valor COP'] = 0.0
    invest_df['TRM Aplicada'] = None
    invest_df['Tasa USD'] = None
    invest_df['Año Declaración'] = None 
    
    for index, row in invest_df.iterrows():
        try:
            year = get_valid_year(row, periodo_df)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
            
            invest_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                invest_df.loc[index, 'Inversiones - Valor COP'] = float(row['Inversiones - Valor'])
                invest_df.loc[index, 'TRM Aplicada'] = 1.0
                invest_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Inversiones - Valor']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    invest_df.loc[index, 'Inversiones - Valor COP'] = cop_amount
                    invest_df.loc[index, 'TRM Aplicada'] = trm
                    invest_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            continue
    
    invest_df.to_excel(output_file_path, index=False)

def run_all_analyses():
    """Run all analysis functions with their respective file paths"""
    file_path = 'src/data.xlsx'
    periodo_file_path = 'src/periodoBR.xlsx'
    
    analyze_banks(file_path, 'tables/cats/banks.xlsx', periodo_file_path)
    analyze_debts(file_path, 'tables/cats/debts.xlsx', periodo_file_path)
    analyze_goods(file_path, 'tables/cats/goods.xlsx', periodo_file_path)
    analyze_incomes(file_path, 'tables/cats/incomes.xlsx', periodo_file_path)
    analyze_investments(file_path, 'tables/cats/investments.xlsx', periodo_file_path)

if __name__ == "__main__":
    run_all_analyses()
"@
}

function createNets {
    Write-Host "🏗️ Creating Nets" -ForegroundColor $YELLOW
    # Banks
    Set-Content -Path "models/nets.py" -Value @"
import pandas as pd

# Common columns used across all analyses
COMMON_COLUMNS = [
    'Usuario', 'Nombre', 'Compañía', 'Cargo',
    'fkIdPeriodo', 'fkIdEstado',
    'Año Creación', 'Año Envío',
    'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
    'Año Declaración'
]

# Base groupby columns for summaries
BASE_GROUPBY = ['Usuario', 'Nombre', 'Compañía', 'Cargo', 'fkIdPeriodo', 'Año Declaración', 'Año Creación']

def analyze_banks(file_path, output_file_path):
    """Analyze bank accounts data"""
    df = pd.read_excel(file_path)

    # Specific columns for banks
    bank_columns = [
        'Banco - Entidad', 'Banco - Tipo Cuenta',
        'Banco - fkIdPaís', 'Banco - Nombre País',
        'Banco - Saldo', 'Banco - Comentario',
        'Banco - Saldo COP'
    ]
    
    df = df[COMMON_COLUMNS + bank_columns]
    
    # Create a temporary combination column for counting
    df_temp = df.copy()
    df_temp['Bank_Account_Combo'] = df['Banco - Entidad'] + "|" + df['Banco - Tipo Cuenta']
    
    # Perform all aggregations
    summary = df_temp.groupby(BASE_GROUPBY).agg(
        **{
            'Cant_Bancos': pd.NamedAgg(column='Banco - Entidad', aggfunc='nunique'),
            'Cant_Cuentas': pd.NamedAgg(column='Bank_Account_Combo', aggfunc='nunique'),
            'Banco - Saldo COP': pd.NamedAgg(column='Banco - Saldo COP', aggfunc='sum')
        }
    ).reset_index()

    summary.to_excel(output_file_path, index=False)
    return summary

def analyze_debts(file_path, output_file_path):
    """Analyze debts data"""
    df = pd.read_excel(file_path)

    # Specific columns for debts
    debt_columns = [
        'Pasivos - Entidad Personas', 'Pasivos - Tipo Obligación', 
        'Pasivos - Valor', 'Pasivos - Comentario',
        'Pasivos - Valor COP', 'Texto Moneda'
    ]
    
    df = df[COMMON_COLUMNS + debt_columns]
    
    # Calculate total Pasivos and count occurrences
    summary = df.groupby(BASE_GROUPBY).agg({      
        'Pasivos - Valor COP': 'sum',
        'Pasivos - Entidad Personas': 'count'
    }).reset_index()

    # Rename columns for clarity
    summary = summary.rename(columns={
        'Pasivos - Entidad Personas': 'Cant_Deudas',
        'Pasivos - Valor COP': 'Total Pasivos'
    })

    summary.to_excel(output_file_path, index=False)
    return summary

def analyze_goods(file_path, output_file_path):
    """Analyze goods/assets data"""
    df = pd.read_excel(file_path)
    
    # Specific columns for goods
    goods_columns = [
        'Bienes - Activo', 'Bienes - % Propiedad',
        'Bienes - Propietario', 'Bienes - Valor Comercial',
        'Bienes - Comentario', 'Bienes - Valor Comercial COP',
        'Bienes - Valor Corregido'
    ]
    
    df = df[COMMON_COLUMNS + goods_columns]

    summary = df.groupby(BASE_GROUPBY).agg({
        'Bienes - Valor Corregido': 'sum',
        'Bienes - Activo': 'count' 
    }).reset_index()

    # Rename columns for clarity
    summary = summary.rename(columns={
        'Bienes - Activo': 'Cant_Bienes',
        'Bienes - Valor Corregido': 'Total Bienes'
    })

    summary.to_excel(output_file_path, index=False) 
    return summary

def analyze_incomes(file_path, output_file_path):
    """Analyze income data"""
    df = pd.read_excel(file_path)
    
    # Specific columns for incomes
    income_columns = [
        'Ingresos - fkIdConcepto', 'Ingresos - Texto Concepto',
        'Ingresos - Valor', 'Ingresos - Comentario',
        'Ingresos - Otros', 'Ingresos - Valor COP',
        'Texto Moneda'
    ]

    df = df[COMMON_COLUMNS + income_columns]
    
    # Calculate Ingresos and count occurrences
    summary = df.groupby(BASE_GROUPBY).agg({
        'Ingresos - Valor COP': 'sum',
        'Ingresos - Texto Concepto': 'count'
    }).reset_index()

    # Rename columns for clarity
    summary = summary.rename(columns={
        'Ingresos - Texto Concepto': 'Cant_Ingresos',
        'Ingresos - Valor COP': 'Total Ingresos'
    })

    summary.to_excel(output_file_path, index=False)
    return summary

def analyze_investments(file_path, output_file_path):
    """Analyze investments data"""
    df = pd.read_excel(file_path)
    
    # Specific columns for investments
    invest_columns = [
        'Inversiones - Tipo Inversión', 'Inversiones - Entidad',
        'Inversiones - Valor', 'Inversiones - Comentario',
        'Inversiones - Valor COP', 'Texto Moneda'
    ]
    
    df = df[COMMON_COLUMNS + invest_columns]
    
    # Calculate total Inversiones and count occurrences
    summary = df.groupby(BASE_GROUPBY + ['Inversiones - Tipo Inversión']).agg( 
        {'Inversiones - Valor COP': 'sum',
         'Inversiones - Tipo Inversión': 'count'}
    ).rename(columns={
        'Inversiones - Tipo Inversión': 'Cant_Inversiones',
        'Inversiones - Valor COP': 'Total Inversiones'
    }).reset_index()
    
    summary.to_excel(output_file_path, index=False)
    return summary 

def calculate_assets(banks_file, goods_file, invests_file, output_file):
    """Calculate total assets by combining banks, goods and investments"""
    banks = pd.read_excel(banks_file)
    goods = pd.read_excel(goods_file)
    invests = pd.read_excel(invests_file)

    # Group investments by base columns (summing across types)
    invests_grouped = invests.groupby(BASE_GROUPBY).agg({
        'Total Inversiones': 'sum',
        'Cant_Inversiones': 'sum'
    }).reset_index()

    # Merge all three dataframes
    merged = pd.merge(goods, banks, on=BASE_GROUPBY, how='outer')
    merged = pd.merge(merged, invests_grouped, on=BASE_GROUPBY, how='outer')
    merged.fillna(0, inplace=True)

    # Calculate total assets
    merged['Total Activos'] = (
        merged['Total Bienes'] + 
        merged['Banco - Saldo COP'] + 
        merged['Total Inversiones']
    )

    # Reorder and rename columns
    final_columns = BASE_GROUPBY + [
        'Total Bienes', 'Cant_Bienes',
        'Banco - Saldo COP', 'Cant_Bancos', 'Cant_Cuentas',
        'Total Inversiones', 'Cant_Inversiones',
        'Total Activos'
    ]
    merged = merged[final_columns]

    merged.to_excel(output_file, index=False)
    return merged

def calculate_net_worth(debts_file, assets_file, output_file):
    """Calculate net worth by combining assets and debts"""
    debts = pd.read_excel(debts_file)
    assets = pd.read_excel(assets_file)

    # Merge the summaries
    merged = pd.merge(
        assets, 
        debts, 
        on=BASE_GROUPBY, 
        how='outer'
    )
    merged.fillna(0, inplace=True)
    
    # Calculate net worth
    merged['Total Patrimonio'] = merged['Total Activos'] - merged['Total Pasivos']
    
    # Final column order
    final_columns = BASE_GROUPBY + [
        'Total Activos',
        'Cant_Bienes',
        'Cant_Bancos',
        'Cant_Cuentas',
        'Cant_Inversiones',
        'Total Pasivos',
        'Cant_Deudas',
        'Total Patrimonio'
    ]
    merged = merged[final_columns]
    
    merged.to_excel(output_file, index=False)
    return merged

def run_all_analyses():
    """Run all analyses in sequence with default file paths"""
    # Individual analyses
    bank_summary = analyze_banks(
        'tables/cats/banks.xlsx',
        'tables/nets/bankNets.xlsx'
    )
    
    debt_summary = analyze_debts(
        'tables/cats/debts.xlsx',
        'tables/nets/debtNets.xlsx'
    )
    
    goods_summary = analyze_goods(
        'tables/cats/goods.xlsx',
        'tables/nets/goodNets.xlsx'
    )
    
    income_summary = analyze_incomes(
        'tables/cats/incomes.xlsx',
        'tables/nets/incomeNets.xlsx'
    )
    
    invest_summary = analyze_investments(
        'tables/cats/investments.xlsx',
        'tables/nets/investNets.xlsx'
    )
    
    # Combined analyses
    assets_summary = calculate_assets(
        'tables/nets/bankNets.xlsx',
        'tables/nets/goodNets.xlsx',
        'tables/nets/investNets.xlsx',
        'tables/nets/assetNets.xlsx'
    )
    
    net_worth_summary = calculate_net_worth(
        'tables/nets/debtNets.xlsx',
        'tables/nets/assetNets.xlsx',
        'tables/nets/worthNets.xlsx'
    )
    
    return {
        'bank_summary': bank_summary,
        'debt_summary': debt_summary,
        'goods_summary': goods_summary,
        'income_summary': income_summary,
        'invest_summary': invest_summary,
        'assets_summary': assets_summary,
        'net_worth_summary': net_worth_summary
    }

if __name__ == '__main__':
    # Run all analyses when script is executed
    results = run_all_analyses()
    print("All nets analyses completed successfully!")
"@
}

function createTrends {
    Write-Host "🏗️ Creating Trends" -ForegroundColor $YELLOW

# trends
   
Set-Content -Path "models/trends.py" -Value @"
import pandas as pd

def get_trend_symbol(value):
    """Determine the trend symbol based on the percentage change."""
    try:
        value_float = float(value.strip('%')) / 100
        if pd.isna(value_float):
            return "➡️"
        elif value_float > 0.1:  # more than 10% increase
            return "📈"
        elif value_float < -0.1:  # more than 10% decrease
            return "📉"
        else:
            return "➡️"  # relatively stable
    except Exception:
        return "➡️"

def calculate_variation(df, column):
    """Calculate absolute and relative variations for a specific column."""
    df = df.sort_values(by=['Usuario', 'Año Declaración'])
    
    absolute_col = f'{column} Var. Abs.'
    relative_col = f'{column} Var. Rel.'
    
    df[absolute_col] = df.groupby('Usuario')[column].diff()
    
    df[relative_col] = (
        df.groupby('Usuario')[column]
        .ffill()
        .pct_change(fill_method=None) * 100
    )
    
    df[relative_col] = df[relative_col].apply(lambda x: f"{x:.2f}%" if not pd.isna(x) else "0.00%")
    
    return df

def embed_trend_symbols(df, columns):
    """Add trend symbols to variation columns."""
    for col in columns:
        absolute_col = f'{col} Var. Abs.'
        relative_col = f'{col} Var. Rel.'
        
        if absolute_col in df.columns:
            df[absolute_col] = df.apply(
                lambda row: f"{row[absolute_col]:.2f} {get_trend_symbol(row[relative_col])}" 
                if pd.notna(row[absolute_col]) else "N/A ➡️",
                axis=1
            )
        
        if relative_col in df.columns:
            df[relative_col] = df.apply(
                lambda row: f"{row[relative_col]} {get_trend_symbol(row[relative_col])}", 
                axis=1
            )
    
    return df

def calculate_leverage(df):
    """Calculate financial leverage."""
    df['Apalancamiento'] = (df['Patrimonio'] / df['Activos']) * 100
    return df

def calculate_debt_level(df):
    """Calculate debt level."""
    df['Endeudamiento'] = (df['Pasivos'] / df['Activos']) * 100
    return df

def process_asset_data(df_assets):
    """Process asset data with variations and trends."""
    df_assets_grouped = df_assets.groupby(['Usuario', 'Año Declaración']).agg(
        BancoSaldo=('Banco - Saldo COP', 'sum'),
        Bienes=('Total Bienes', 'sum'),
        Inversiones=('Total Inversiones', 'sum')
    ).reset_index()

    for column in ['BancoSaldo', 'Bienes', 'Inversiones']:
        df_assets_grouped = calculate_variation(df_assets_grouped, column)
    
    df_assets_grouped = embed_trend_symbols(df_assets_grouped, ['BancoSaldo', 'Bienes', 'Inversiones'])
    return df_assets_grouped

def process_income_data(df_income):
    """Process income data with variations and trends."""
    df_income_grouped = df_income.groupby(['Usuario', 'Año Declaración']).agg(
        Ingresos=('Total Ingresos', 'sum'),
        Cant_Ingresos=('Cant_Ingresos', 'sum')
    ).reset_index()

    df_income_grouped = calculate_variation(df_income_grouped, 'Ingresos')
    df_income_grouped = embed_trend_symbols(df_income_grouped, ['Ingresos'])
    return df_income_grouped

def calculate_yearly_variations(df):
    """Calculate yearly variations for all columns."""
    df = df.sort_values(['Usuario', 'Año Declaración'])
    
    columns_to_analyze = [
        'Activos', 'Pasivos', 'Patrimonio', 
        'Apalancamiento', 'Endeudamiento',
        'BancoSaldo', 'Bienes', 'Inversiones', 'Ingresos',
        'Cant_Ingresos'
    ]
    
    new_columns = {}
    
    for column in [col for col in columns_to_analyze if col in df.columns]:
        grouped = df.groupby('Usuario')[column]
        
        for year in [2021, 2022, 2023, 2024]:
            abs_col = f'{year} {column} Var. Abs.'
            new_columns[abs_col] = grouped.diff()
            
            rel_col = f'{year} {column} Var. Rel.'
            pct_change = grouped.pct_change(fill_method=None) * 100
            new_columns[rel_col] = pct_change.apply(
                lambda x: f"{x:.2f}%" if not pd.isna(x) else "0.00%"
            )
    
    df = pd.concat([df, pd.DataFrame(new_columns)], axis=1)
    
    for column in [col for col in columns_to_analyze if col in df.columns]:
        for year in [2021, 2022, 2023, 2024]:
            abs_col = f'{year} {column} Var. Abs.'
            rel_col = f'{year} {column} Var. Rel.'
            
            if abs_col in df.columns:
                df[abs_col] = df.apply(
                    lambda row: f"{row[abs_col]:.2f} {get_trend_symbol(row[rel_col])}" 
                    if pd.notna(row[abs_col]) else "N/A ➡️",
                    axis=1
                )
            if rel_col in df.columns:
                df[rel_col] = df.apply(
                    lambda row: f"{row[rel_col]} {get_trend_symbol(row[rel_col])}", 
                    axis=1
                )
    
    return df

def save_results(df, excel_filename="tables/trends/trends.xlsx", json_filename=None):
    """Save results to Excel and optionally JSON."""
    try:
        df.to_excel(excel_filename, index=False)
        print(f"Data saved to {excel_filename}")
        
        if json_filename:
            df.to_json(json_filename, orient='records', indent=4, force_ascii=False)
            print(f"Data saved to {json_filename}")
    except Exception as e:
        print(f"Error saving file: {e}")

def main():
    """Main function to process all data and generate analysis files."""
    try:
        # Process worth data
        df_worth = pd.read_excel("tables/nets/worthNets.xlsx")
        df_worth = df_worth.rename(columns={
            'Total Activos': 'Activos',
            'Total Pasivos': 'Pasivos',
            'Total Patrimonio': 'Patrimonio'
        })
        
        df_worth = calculate_leverage(df_worth)
        df_worth = calculate_debt_level(df_worth)
        
        for column in ['Activos', 'Pasivos', 'Patrimonio', 'Apalancamiento', 'Endeudamiento']:
            df_worth = calculate_variation(df_worth, column)
        
        df_worth = embed_trend_symbols(df_worth, ['Activos', 'Pasivos', 'Patrimonio', 'Apalancamiento', 'Endeudamiento'])
        
        # Process asset data
        df_assets = pd.read_excel("tables/nets/assetNets.xlsx")
        df_assets_processed = process_asset_data(df_assets)
        
        # Process income data
        df_income = pd.read_excel("tables/nets/incomeNets.xlsx")
        df_income_processed = process_income_data(df_income)
        
        # Merge all data
        df_combined = pd.merge(df_worth, df_assets_processed, on=['Usuario', 'Año Declaración'], how='left')
        df_combined = pd.merge(df_combined, df_income_processed, on=['Usuario', 'Año Declaración'], how='left')
        
        # Save basic trends
        save_results(df_combined, "tables/trends/trends.xlsx")
        
        # Calculate and save yearly variations
        df_yearly = calculate_yearly_variations(df_combined)
        save_results(df_yearly, "tables/trends/overTrends.xlsx", "src/data.json")
        
    except FileNotFoundError as e:
        print(f"Error: Required file not found - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
"@
}

function createApp {
    Write-Host "🏗️ Creating App" -ForegroundColor $YELLOW
    # app.py
    Set-Content -Path "app.py" -Value @"
import http.server
import socketserver
import os
import webbrowser
import threading
import json
from models.passKey import remove_excel_password, add_fk_id_estado
from models.cats import run_all_analyses as run_cats_analyses
from models.nets import run_all_analyses as run_nets_analyses
from models.trends import main as run_trends_analysis

class CustomHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()
    
    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()
    
    def do_POST(self):
        if self.path == '/upload':
            try:
                content_length = int(self.headers['Content-Length'])
                content_type = self.headers['Content-Type']
                
                if 'multipart/form-data' not in content_type:
                    self.send_response(400)
                    self.end_headers()
                    self.wfile.write(json.dumps({'success': False, 'message': 'Only multipart form data is supported'}).encode())
                    return
                
                # Read the raw POST data
                post_data = self.rfile.read(content_length)
                
                # Find the boundary from the content type
                boundary = content_type.split("boundary=")[1].encode()
                
                # Split the multipart data
                parts = post_data.split(b'--' + boundary)
                
                file_data = None
                open_password = ''
                modify_password = ''
                for part in parts:
                    if b'filename="' in part:
                        # This part contains the file
                        header, content = part.split(b'\r\n\r\n', 1)
                        file_data = content.rstrip(b'\r\n')
                    elif b'name="openPassword"' in part:
                        # This part contains the open password
                        header, content = part.split(b'\r\n\r\n', 1)
                        open_password = content.rstrip(b'\r\n').decode('utf-8')
                    elif b'name="modifyPassword"' in part:
                        # This part contains the modify password
                        header, content = part.split(b'\r\n\r\n', 1)
                        modify_password = content.rstrip(b'\r\n').decode('utf-8')
                
                if not file_data:
                    self.send_response(400)
                    self.end_headers()
                    self.wfile.write(json.dumps({'success': False, 'message': 'No file found in upload'}).encode())
                    return
                
                # Save the uploaded file
                input_path = 'src/dataHistoricaPBI.xlsx'
                with open(input_path, 'wb') as f:
                    f.write(file_data)
                
                # Process the file
                output_excel = 'src/data.xlsx'
                output_json = 'src/fk1data.json'
                
                # Remove password if needed
                remove_excel_password(input_path, output_excel, open_password, modify_password)
                
                # Add fkIdEstado and convert to JSON
                add_fk_id_estado(output_excel, output_json)
                
                # Run all analyses
                run_cats_analyses()
                run_nets_analyses()
                run_trends_analysis()
                
                # Read the processed data to include in response
                with open(output_json, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                response_data = {
                    'success': True,
                    'message': 'File processed successfully',
                    'file': output_json,
                    'stats': {
                        'rows_processed': len(data),
                        'columns': len(data[0].keys()) if data else 0
                    }
                }
                
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(response_data).encode())
                
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'success': False,
                    'message': str(e),
                    'error_type': type(e).__name__
                }).encode())

def start_server(port=8000):
    """Starts a simple HTTP server in a background thread."""
    Handler = CustomHTTPRequestHandler

    with socketserver.TCPServer(("", port), Handler) as httpd:
        print(f"Serving at port {port}")
        httpd.serve_forever()

def main():
    """Automates the process of generating tables and opening index.html in a browser."""

    print("Hold CTRL and click http://localhost:8000/")

    # Get the absolute path to index.html
    index_path = os.path.abspath("index.html")

    # Check if index.html exists
    if os.path.exists(index_path):
        # Start the HTTP server in a separate thread
        server_thread = threading.Thread(target=start_server)
        server_thread.daemon = True  # Allow the main thread to exit even if the server is running
        server_thread.start()

        # Open index.html in the default browser
        webbrowser.open("http://localhost:8000")  # Open via the server URL
    else:
        print("Error: index.html not found in the current directory.")
        return # exit the program if index.html doesn't exist

    # Keep the main thread alive (optional) or perform other tasks
    input("Press Enter to stop the server and exit...\n")  #Wait for user input to terminate

if __name__ == "__main__":
    main()
"@
}

function createConflictScript {
    Write-Host "🏗️ Creating Conflict Script" -ForegroundColor $YELLOW
    
    Set-Content -Path "models/conflicts.py" -Value @"
import pandas as pd
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers

def extract_specific_columns(input_file, output_file, custom_headers=None):
    
    try:
        # Setup output directory
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        
        # Read raw data (no automatic parsing)
        df = pd.read_excel(input_file, header=None)
        
        # Column selection (first 11 + specified extras)
        base_cols = list(range(11))  # Columns 0-10 (A-K)
        extra_cols = [12,14,16,18,20,22,24,26,28]
        selected_cols = [col for col in base_cols + extra_cols if col < df.shape[1]]
        
        # Extract data with headers
        result = df.iloc[3:, selected_cols].copy()
        result.columns = df.iloc[2, selected_cols].values
        
        # Apply custom headers if provided
        if custom_headers is not None:
            if len(custom_headers) != len(result.columns):
                raise ValueError(f"Custom headers count ({len(custom_headers)}) doesn't match column count ({len(result.columns)})")
            result.columns = custom_headers
        
        # Merge C,D,E,F → C (indices 2,3,4,5)
        if all(c in selected_cols for c in [2,3,4,5]):
            result.iloc[:, 2] = result.iloc[:, 2:6].astype(str).apply(' '.join, axis=1)
            result.drop(result.columns[3:6], axis=1, inplace=True)
            selected_cols = [c for c in selected_cols if c not in [3,4,5]] 
            
        # Capitalize "Nombre" column AFTER merging
        if "Nombre" in result.columns:
            result["Nombre"] = result["Nombre"].str.title()
            
        # Special handling for Column J (input index 9)
        if 9 in selected_cols:
            j_pos = selected_cols.index(9)  # Find its position in output
            date_col = result.columns[j_pos]
            
            # Convert with European date format
            result[date_col] = pd.to_datetime(
                result[date_col],
                dayfirst=True,
                errors='coerce'
            )
            
            # Save with Excel formatting
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                result.to_excel(writer, index=False)
                
                # Get the worksheet and format the date column
                worksheet = writer.sheets['Sheet1']
                date_col_letter = get_column_letter(j_pos + 1)
                
                # Apply date format to all cells in the column
                for cell in worksheet[date_col_letter]:
                    if cell.row == 1:  # Skip header
                        continue
                    cell.number_format = 'DD/MM/YYYY'
                
                # Auto-adjust columns
                for idx, col in enumerate(result.columns):
                    col_letter = get_column_letter(idx+1)
                    worksheet.column_dimensions[col_letter].width = max(
                        len(str(col))+2,
                        result[col].astype(str).str.len().max()+2
                    )
            
            print(f"Success! Output saved to: {output_file}")
        
        else:
            print("Warning: Column J not found in selected columns")
    
    except Exception as e:
        print(f"Error: {str(e)}")

# Example usage with custom headers
custom_headers = [
    "ID", "# Documento", "Nombre", "1er Nombre", "1er Apellido", 
    "2do Apellido", "Compañía", "Cargo", "Email", "Fecha de Inicio", 
    "Q1", "Q2", "Q3", "Q4", "Q5",
    "Q6", "Q7", "Q8", "Q9", "Q10"
]

extract_specific_columns(
    input_file="src/conflictos.xls",
    output_file="tables/conflicts.xlsx",
    custom_headers=custom_headers
)
"@
}

function createIDScript {
    Write-Host "🏗️ Creating ID Script" -ForegroundColor $YELLOW
    
    Set-Content -Path "models/ids.py" -Value @"
import pandas as pd
from pathlib import Path
import warnings

def process_excel_files(activos_file, retirados_file, output_file, column_names):
   
    # Suppress SettingWithCopyWarning
    warnings.simplefilter(action='ignore', category=pd.errors.SettingWithCopyWarning)
    
    try:
        # Create output directory if needed
        Path(output_file).parent.mkdir(parents=True, exist_ok=True)
        
        # Function to load and standardize data from either file
        def load_data(file_path, estado_value):
            # Read Excel file
            df = pd.read_excel(file_path, header=None, engine='openpyxl')
            
            # Validate minimum rows
            if len(df) < 3:
                raise ValueError(f"Input file {file_path} has less than 3 rows")
                
            # Extract headers and data
            headers = df.iloc[2].tolist()
            data = df.iloc[3:].copy()
            data.columns = headers
            
            # Select required columns using .loc to avoid warnings
            col_indices = {
                'B': 1, 'I': 8, 'J': 9, 'K': 10,
                'L': 11, 'O': 14, 'S': 18
            }
            
            # Verify columns exist
            missing_cols = [col for col, idx in col_indices.items() if idx >= len(headers)]
            if missing_cols:
                raise ValueError(f"Missing columns in {file_path}: {missing_cols}")
            
            # Create new DataFrame with selected columns to avoid view issues
            selected_data = pd.DataFrame({
                column_names['B']: data.iloc[:, col_indices['B']],
                column_names['I']: data.iloc[:, col_indices['I']],
                column_names['J']: data.iloc[:, col_indices['J']],
                column_names['K']: data.iloc[:, col_indices['K']],
                column_names['L']: data.iloc[:, col_indices['L']],
                column_names['O']: data.iloc[:, col_indices['O']],
                column_names['S']: data.iloc[:, col_indices['S']]
            })
            
            # Concatenate name components safely
            concat_cols = [column_names[col] for col in ['I', 'J', 'K', 'L']]
            nombres = (
                selected_data[concat_cols]
                .fillna('')
                .astype(str)
                .apply(lambda x: ' '.join(x).strip(), axis=1)
                .str.lower()
                .str.title()
            )
            
            # Create final DataFrame with all columns
            result = pd.DataFrame({
                column_names['B']: selected_data[column_names['B']],
                'Nombre': nombres,
                column_names['O']: selected_data[column_names['O']],
                column_names['S']: selected_data[column_names['S']],
                'Estado': estado_value
            })
            
            return result
        
        # Load both datasets
        df_activos = load_data(activos_file, 'Activo')
        df_retirados = load_data(retirados_file, 'Retirado')
        
        # Combine datasets
        combined_df = pd.concat([df_activos, df_retirados], ignore_index=True)
        
        # Select final output columns
        output_cols = [
            column_names['B'],  # # Documento
            'Nombre',          # Full name
            column_names['O'],  # Cargo
            column_names['S'],  # Compañía
            'Estado'           # Activo/Retirado
        ]
        
        # Save to Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            combined_df[output_cols].to_excel(writer, index=False)
            
        print(f"Successfully processed and combined files:\n"
              f"Active employees: {activos_file}\n"
              f"Retired employees: {retirados_file}\n"
              f"Output: {output_file}\n"
              f"Total records: {len(combined_df)} (Active: {len(df_activos)}, Retired: {len(df_retirados)})")
        
    except Exception as e:
        print(f"Error processing files: {str(e)}")
        raise

if __name__ == "__main__":
    # Configuration
    CONFIG = {
        'activos_file': "src/ACTIVOS.xlsx",      # Current active employees
        'retirados_file': "src/RETIRADOS.xlsx",  # New retired employees
        'output_file': "tables/IDS.xlsx",
        'column_names': {
            'B': '# Documento',
            'I': 'Nombre1',
            'J': 'Nombre2', 
            'K': 'Nombre3',
            'L': 'Nombre4',
            'O': 'Cargo',
            'S': 'Compañía'
        }
    }
    
    # Run processing
    process_excel_files(
        activos_file=CONFIG['activos_file'],
        retirados_file=CONFIG['retirados_file'],
        output_file=CONFIG['output_file'],
        column_names=CONFIG['column_names']
    )
"@
}

function createIDtrends {
    Write-Host "🏗️ Creating ID Trends" -ForegroundColor $YELLOW
    
    Set-Content -Path "models/idTrends.py" -Value @"
import pandas as pd
from pathlib import Path

def merge_trends_data(ids_file, trends_file, output_file):
    """
    Merge trends data with employee IDs using first-letter matching:
    - Keeps all records from trends.xlsx
    - Adds # Documento from IDS.xlsx
    - Matches on first letters of Nombre, Cargo, and Compañía
    """
    try:
        # Create output directory if needed
        Path(output_file).parent.mkdir(parents=True, exist_ok=True)
        
        # Read both files
        df_ids = pd.read_excel(ids_file, engine='openpyxl')
        df_trends = pd.read_excel(trends_file, engine='openpyxl')
        
        # Create first-letter merge keys
        for df in [df_ids, df_trends]:
            df['merge_key'] = (
                df['Nombre'].str[0].str.lower() + '|' +
                df['Cargo'].str[0].str.lower() + '|' +
                df['Compañía'].str[0].str.lower()
            )
        
        # Perform left join (keep all trends records)
        merged_df = pd.merge(
            left=df_trends,
            right=df_ids[['merge_key', '# Documento']],
            how='left',
            on='merge_key'
        )
        
        # Handle potential duplicates - keep first match
        merged_df = merged_df.drop_duplicates(subset=df_trends.columns.tolist(), keep='first')
        
        # Clean up - drop the merge key and reorder columns
        merged_df = merged_df.drop(columns=['merge_key'])
        cols = ['# Documento'] + [col for col in merged_df.columns if col != '# Documento']
        merged_df = merged_df[cols]
        
        # Save to Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False)
            
        print(f"Successfully merged files using first-letter matching:\n"
              f"IDS file: {ids_file}\n"
              f"Trends file: {trends_file}\n"
              f"Output: {output_file}\n"
              f"Total records: {len(merged_df)}\n"
              f"Records with matched ID: {merged_df['# Documento'].notna().sum()}")
        
    except Exception as e:
        print(f"Error merging files: {str(e)}")
        raise

if __name__ == "__main__":
    # Configuration
    CONFIG = {
        'ids_file': "tables/IDS.xlsx",  # Current combined file
        'trends_file': "tables/trends/trends.xlsx",       # New trends data
        'output_file': "tables/idTrends.xlsx"  # Output file
    }
    
    # Run merging
    merge_trends_data(
        ids_file=CONFIG['ids_file'],
        trends_file=CONFIG['trends_file'],
        output_file=CONFIG['output_file']
    )
"@
}

function createINtrends {
    Write-Host "🏗️ Merging Trends/Conflicts" -ForegroundColor $YELLOW
    
    Set-Content -Path "models/inTrends.py" -Value @"
import pandas as pd
from pathlib import Path

def join_conflicts_data(trends_file, conflicts_file, output_file, how='left'):
    """
    Join trends data with conflicts data using different join types:
    - Uses # Documento as the join key
    - Keeps all columns from both files
    - Handles duplicate column names
    - Converts join key to consistent type (string)
    - Supports join types: 'left', 'right', 'inner', 'outer'
    - Saves as inTrends.xlsx
    """
    try:
        # Create output directory if needed
        Path(output_file).parent.mkdir(parents=True, exist_ok=True)
        
        # Read both files
        df_trends = pd.read_excel(trends_file, engine='openpyxl')
        df_conflicts = pd.read_excel(conflicts_file, engine='openpyxl')
        
        # Check if # Documento exists in both files
        if '# Documento' not in df_trends.columns or '# Documento' not in df_conflicts.columns:
            raise ValueError("# Documento column missing in one or both input files")
        
        # Convert '# Documento' columns to same type (string)
        df_trends['# Documento'] = df_trends['# Documento'].astype(str)
        df_conflicts['# Documento'] = df_conflicts['# Documento'].astype(str)
        
        # Perform the join operation
        joined_df = pd.merge(
            left=df_trends,
            right=df_conflicts,
            how=how,  # Type of join: 'left', 'right', 'inner', 'outer'
            on='# Documento',
            suffixes=('_trends', '_conflicts')
        )
        
        # Save to Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            joined_df.to_excel(writer, index=False)
            
        print(f"Successfully joined files with '{how}' join:\n"
              f"Trends file: {trends_file}\n"
              f"Conflicts file: {conflicts_file}\n"
              f"Output: {output_file}\n"
              f"Total records: {len(joined_df)}\n"
              f"Records from trends: {len(df_trends)}\n"
              f"Records from conflicts: {len(df_conflicts)}\n"
              f"Join type: {how}")
        
    except Exception as e:
        print(f"Error joining files: {str(e)}")
        raise

if __name__ == "__main__":
    # Configuration
    CONFIG = {
        'trends_file': "tables/idTrends.xlsx",  # Current trends file
        'conflicts_file': "tables/conflicts.xlsx",      # Conflicts data
        'output_file': "tables/inTrends.xlsx",        # Output file
        'join_type': "left"  # Can be 'left', 'right', 'inner', 'outer'
    }
    
    # Run joining
    join_conflicts_data(
        trends_file=CONFIG['trends_file'],
        conflicts_file=CONFIG['conflicts_file'],
        output_file=CONFIG['output_file'],
        how=CONFIG['join_type']
    )
"@
}

function createIndex {
Write-Host "🏗️ Creating HTML" -ForegroundColor $YELLOW
    # html
    Set-Content -Path "index.html" -Value @"
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>A R P A</title>
    <link rel="stylesheet" href="static/style.css">
    <link rel="shortcut icon" href="favicon.png" type="image/x-icon">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
</head>
<body>
    <div class="topnav-container">
        <div class="logoIN"></div>
        <div class="nomPag">A R P A</div>
    </div>
    
    <div class="tab-container">
        <button class="tab active" data-tab="bienes-rentas">Bienes y Rentas</button>
        <!--<button class="tab" data-tab="transactions">Extractos</button>-->
    </div>
    
    <div id="bienes-rentas" class="tab-content active">
        <div class="filter-form">
            <label for="excelUpload" class="file-upload-label">
              <span class="file-upload-button">Cargar archivo Excel</span>
              <input type="file"
                     id="excelUpload"
                     accept=".xlsx,.xls"
                     aria-describedby="fileUploadHelp"
                     class="file-upload-input">
            </label>
            <span id="fileUploadStatus" aria-live="polite" class="file-upload-status"></span>
    
                <div id="passwordContainer" style="display: none;">
                    <div style="display: flex; gap: 15px; width: 100%;">
                        <div class="password-input-group" style="flex: 1;">
                            <input type="password"
                                id="excelOpenPassword"
                                placeholder="Contraseña de apertura"
                                class="password-input">
                            <span class="toggle-password" onclick="togglePassword('excelOpenPassword')">👁️ </span>
                        </div>
                        <div class="password-input-group" style="flex: 1;">
                            <input type="password"
                                id="excelModifyPassword"
                                placeholder="Contraseña de modificación"
                                class="password-input">
                            <span class="toggle-password" onclick="togglePassword('excelModifyPassword')">👁️ </span>
                        </div><button id="analyzeButton">Analizar Archivo</button>
                    </div>
                </div>
    
    
                <button onclick="exportToExcel()" style="margin-left: auto; background-color:rgb(0, 176, 15);" class="fa fa-file-excel-o"> Exportar a Excel</button>
                <button onclick="exportFrozenColumnsToExcel()" style="background-color:rgb(0, 176, 15);" class="fa fa-file-excel-o"> Exportar columnas congeladas</button>
            <div id="passwordError" class="error-message"></div>
        </div>
    </div>
    <!--
    <div id="transactions" class="tab-content" style="display: none;">
        <h2>Extractos</h2>
        <p>Este es la tabla para los extratos.</p>
    </div>
    -->
    <div id="loadingBarContainer" style="display: none;">
        <div id="loadingBar"></div>
        <div id="loadingText">Analizando archivo...</div>
    </div>
    <!-- SheetJS for Excel export 
    <div class="filter-form">
        <div class="filter-buttons">
            <button onclick="applyPredeterminedFilter('Patrimonio', '>', '3000000000')">Patrimonio > ,000M</button>
            <button onclick="applyPredeterminedFilter('Patrimonio Var. Rel.', '>', '30')">Patrimonio Var. Rel. > 30%</button>
            <button onclick="applyPredeterminedFilter('Cant_Bienes', '<', '0')">Cant. Bienes < 0</button>
            <button onclick="applyPredeterminedFilter('Endeudamiento Var. Rel.', '>', '50')">Endeudamiento Var. Rel. > 50%</button>
            <button onclick="applyPredeterminedFilter('Ingresos', '>', '50000000')">Ingresos > </button>
            <button onclick="applyPredeterminedFilter('Cant_Deudas', '>=', '5')">Cant. Deudas ≥ 5</button>
            <button onclick="applyPredeterminedFilter('Cant_Bienes', '>=', '6')">Cant. Bienes ≥ 6</button>
        </div> 
    </div>-->
    <div class="filter-form">
        <select id="column" aria-label="Seleccionar columna para filtrar" title="Columna para filtrar">
            <option value="">-- Selecciona columna --</option>
            <optgroup label="Información Personal">
                <option value="Nombre">Nombre</option>
                <option value="Año Declaración">Año Declaración</option>
                <option value="Compañía">Compañía</option>
                <option value="Cargo">Cargo</option>
                <option value="Usuario">Usuario</option>
            </optgroup>
            <optgroup label="Valores Principales">
                <option value="Activos">Activos</option>
                <option value="Pasivos">Pasivos</option>
                <option value="Patrimonio">Patrimonio</option>
                <option value="Apalancamiento">Apalancamiento</option>
                <option value="Endeudamiento">Endeudamiento</option>
                <option value="Cant_Deudas">Cant. Deudas</option>
                <option value="BancoSaldo">Saldo Bancario</option>
                <option value="Cant_Bancos">Cant. Bancos</option>
                <option value="Bienes">Bienes</option>
                <option value="Cant_Bienes">Cant. Bienes</option>
                <option value="Inversiones">Inversiones</option>
                <option value="Cant_Inversiones">Cant. Inversiones</option>
                <option value="Ingresos">Ingresos</option>
                <option value="Cant_Ingresos">Cant. Ingresos</option>
            </optgroup>
            <optgroup label="Variaciones Absolutas">
                <option value="Activos Var. Abs.">Activos Var. Abs.</option>
                <option value="Pasivos Var. Abs.">Pasivos Var. Abs.</option>
                <option value="Patrimonio Var. Abs.">Patrimonio Var. Abs.</option>
                <option value="Apalancamiento Var. Abs.">Apalancamiento Var. Abs.</option>
                <option value="Endeudamiento Var. Abs.">Endeudamiento Var. Abs.</option>
                <option value="BancoSaldo Var. Abs.">BancoSaldo Var. Abs.</option>
                <option value="Bienes Var. Abs.">Bienes Var. Abs.</option>
                <option value="Inversiones Var. Abs.">Inversiones Var. Abs.</option>
                <option value="Ingresos Var. Abs.">Ingresos Var. Abs.</option>
            </optgroup>
            <optgroup label="Variaciones Relativas">
                <option value="Activos Var. Rel.">Activos Var. Rel.</option>
                <option value="Pasivos Var. Rel.">Pasivos Var. Rel.</option>
                <option value="Patrimonio Var. Rel.">Patrimonio Var. Rel.</option>
                <option value="Apalancamiento Var. Rel.">Apalancamiento Var. Rel.</option>
                <option value="Endeudamiento Var. Rel.">Endeudamiento Var. Rel.</option>
                <option value="BancoSaldo Var. Rel.">BancoSaldo Var. Rel.</option>
                <option value="Bienes Var. Rel.">Bienes Var. Rel.</option>
                <option value="Inversiones Var. Rel.">Inversiones Var. Rel.</option>
                <option value="Ingresos Var. Rel.">Ingresos Var. Rel.</option>
            </optgroup>
        </select>
        
        <select id="operator" aria-label="Seleccionar operador de filtro" title="Operador de filtro">
          <option value=">">Mayor que</option>
          <option value="<">Menor que</option>
          <option value="=">Igual a</option>
          <option value=">=">Mayor o igual</option>
          <option value="<=">Menor o igual</option>
          <option value="between">Entre</option>
          <option value="contains">Contiene</option>
        </select>
        
        <input type="text" id="value1" placeholder="Valor">
        <input type="text" id="value2" placeholder="y" style="display: none;">
        
        <button onclick="addFilter()"><i class="fa fa-filter"></i> Agregar Filtro</button>
        <button onclick="clearFilters()" style="background-color: #dc3545; color: white;">Limpiar Filtros</button>
        
        <div style="margin-left: auto; background-color:rgb(0, 176, 15);">
            <button onclick="changeDataSource()">Estado</button>
            <select id="dataSource" aria-label="Seleccionar fuente de datos" title="Fuente de datos">
                <option value="src/data.json">Completo</option>
                <option value="src/fk1Data.json">Incompleto</option>
            </select>
        </div>
    </div>
    
    <div id="filters"></div>
    
    <div class="table-scroll-container">
        <table id="results">
            <thead>
                <tr class="column-controls">

                    <th>
                        <div class="column-controls-container">
                            <button class="freeze-btn" onclick="toggleFreezeColumn(0)">
                                <i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i>
                            </button>
                            <button class="freeze-btn" onclick="showColumnStats('Nombre')">
                                <i class="material-icons" style="font-size:18px">equalizer</i>
                            </button>
                            <button onclick="quickFilter('Nombre')"><span class="sort-icon">↕</span></button>
                        </div>
                        <input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(0, this.value)">
                    </th>

                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(1)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Año Declaración')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Año Declaración')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(1, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(2)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Compañía')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Compañía')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(2, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(3)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Cargo')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Cargo')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(3, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(4)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Usuario')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Usuario')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(4, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(5)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Activos')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Activos')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(5, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(6)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Pasivos')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Pasivos')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(6, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(7)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Patrimonio')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Patrimonio')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(7, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(8)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Apalancamiento')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Apalancamiento')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(8, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(9)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Endeudamiento')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Endeudamiento')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(9, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(10)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Cant_Deudas')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Cant_Deudas')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(10, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(11)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('BancoSaldo')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('BancoSaldo')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(11, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(12)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Cant_Bancos')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Cant_Bancos')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(12, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(13)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Bienes')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Bienes')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(13, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(14)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Cant_Bienes')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Cant_Bienes')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(14, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(15)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Inversiones')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Inversiones')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(15, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(16)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Cant_Inversiones')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Cant_Inversiones')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(16, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(17)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Ingresos')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Ingresos')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(17, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(18)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Cant_Ingresos')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Cant_Ingresos')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(18, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(19)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Activos Var. Abs.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Activos Var. Abs.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(19, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(20)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Pasivos Var. Abs.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Pasivos Var. Abs.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(20, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(21)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Patrimonio Var. Abs.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Patrimonio Var. Abs.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(21, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(22)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Apalancamiento Var. Abs.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Apalancamiento Var. Abs.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(22, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(23)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Endeudamiento Var. Abs.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Endeudamiento Var. Abs.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(23, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(24)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('BancoSaldo Var. Abs.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('BancoSaldo Var. Abs.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(24, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(25)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Bienes Var. Abs.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Bienes Var. Abs.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(25, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(26)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Inversiones Var. Abs.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Inversiones Var. Abs.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(26, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(27)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Ingresos Var. Abs.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Ingresos Var. Abs.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(27, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(28)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Activos Var. Rel.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Activos Var. Rel.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(28, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(29)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Pasivos Var. Rel.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Pasivos Var. Rel.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(29, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(30)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Patrimonio Var. Rel.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Patrimonio Var. Rel.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(30, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(31)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Apalancamiento Var. Rel.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Apalancamiento Var. Rel.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(31, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(32)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Endeudamiento Var. Rel.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Endeudamiento Var. Rel.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(32, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(33)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('BancoSaldo Var. Rel.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('BancoSaldo Var. Rel.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(33, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(34)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Bienes Var. Rel.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Bienes Var. Rel.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(34, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(35)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Inversiones Var. Rel.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Inversiones Var. Rel.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(35, this.value)"></th>
                    <th><div class="column-controls-container"><button class="freeze-btn" onclick="toggleFreezeColumn(36)"><i class="glyphicon glyphicon-pushpin" style="font-size:18px;"></i></button><button class="freeze-btn" onclick="showColumnStats('Ingresos Var. Rel.')"><i class="material-icons" style="font-size:18px">equalizer</i></button><button onclick="quickFilter('Ingresos Var. Rel.')"><span class="sort-icon">↕</span></button></div><input type="range" class="width-slider" min="50" max="300" value="120" oninput="resizeColumn(36, this.value)"></th>    
                </tr>
                <tr>
                    <th>
                        <button>Nombre</button>
                    </th>
                
                    <th><button>Año Declaración</button></th>
                    <th><button>Compañía</button></th>
                    <th><button>Cargo</button></th>
                    <th><button>Usuario</button></th>
                    <th><button>Activos</button></th>
                    <th><button>Pasivos</button></th>
                    <th><button>Patrimonio</button></th>
                    <th><button>Apalancamiento</button></th>
                    <th><button>Endeudamiento</button></th>
                    <th><button>Cant_Deudas</button></th>
                    <th><button>BancoSaldo</button></th>
                    <th><button>Cant_Bancos</button></th>
                    <th><button>Bienes</button></th>
                    <th><button>Cant_Bienes</button></th>
                    <th><button>Inversiones</button></th>
                    <th><button>Cant_Inversiones</button></th>
                    <th><button>Ingresos</button></th>
                    <th><button>Cant_Ingresos</button></th>
                    <th><button>Activos Var. Abs.</button></th>
                    <th><button>Pasivos Var. Abs.</button></th>
                    <th><button>Patrimonio Var. Abs.</button></th>
                    <th><button>Apalancamiento Var. Abs.</button></th>
                    <th><button>Endeudamiento Var. Abs.</button></th>
                    <th><button>BancoSaldo Var. Abs.</button></th>
                    <th><button>Bienes Var. Abs.</button></th>
                    <th><button>Inversiones Var. Abs.</button></th>
                    <th><button>Ingresos Var. Abs.</button></th>
                    <th><button>Activos Var. Rel.</button></th>
                    <th><button>Pasivos Var. Rel.</button></th>
                    <th><button>Patrimonio Var. Rel.</button></th>
                    <th><button>Apalancamiento Var. Rel.</button></th>
                    <th><button>Endeudamiento Var. Rel.</button></th>
                    <th><button>BancoSaldo Var. Rel.</button></th>
                    <th><button>Bienes Var. Rel.</button></th>
                    <th><button>Inversiones Var. Rel.</button></th>
                    <th><button>Ingresos Var. Rel.</button></th>
                    <th><button><i class="material-icons" style="font-size:24px">trending_up</i></button></th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td colspan="35" class="loading">Cargando datos...</td>
                </tr>
            </tbody>
        </table>
    </div>

    <!-- SheetJS for Excel export -->
    <script src="https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js"></script>
    <script src="static/script.js"></script>
</body>
</html>
"@

Write-Host "🏗️ Creating CSS" -ForegroundColor $YELLOW
    # css
    Set-Content -Path "static/style.css" -Value @"
@import url('https://fonts.googleapis.com/css2?family=Open+Sans&display=swap');

* {
    font-family: 'Open Sans', sans-serif;
    box-sizing: border-box;
}

body {
    margin: 0;
    padding: 20px;
    background-color: #f8f9fa;
}

.topnav-container {
    display: flex;
    align-items: center;
    margin-bottom: 20px;
}

.logoIN {
    cursor: pointer;
    width: 40px;
    height: 40px;
    background-color: #0b00a2;
    border-radius: 8px;
    display: inline-flex;
    position: relative;
}

.logoIN::before {
    content: "";
    width: 40px;
    height: 40px;
    border-radius: 50%;
    position: absolute;
    top: 30%;
    left: 70%;
    transform: translate(-50%, -50%);
    background-image: linear-gradient(to right, 
        #ffffff 2px, transparent 1.5px,
        transparent 1.5px, #ffffff 1.5px,
        #ffffff 2px, transparent 1.5px);
    background-size: 4px 100%; 
}

.nomPag {
    margin-left: 10px;
    color: #0b00a2;
    font-weight: bold;
    font-size: 1,5rem;
}

.tab-container {
    display: flex;
    border-bottom: 1px solid #ccc;
}

.tab {
    padding: 10px 15px;
    border: none;
    background-color: #f0f0f0;
    cursor: pointer;
    border-bottom: none; /* Remove bottom border */
    margin-bottom: -1px; /* Adjust margin to cover the border */
}

.tab.active {
    background-color: #fff;
    border-bottom: 1px solid #fff;
    border: 1px solid #ccc; /* Add border around the active tab */
    border-bottom: none;
}

.tab-content {
    padding: 15px;
    /* border: 1px solid #ccc; Removed to avoid double border*/
}

h1 {
    color: #333;
    margin-top: 0;
}

.filter-form, .predetermined-filters {
    background: white;
    padding: 15px;
    border-radius: 8px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    margin-bottom: 20px;
}

.filter-form {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px;
}

.filter-form select, 
.filter-form input, 
.filter-form button,
.password-input {
    padding: 8px 12px;
    border: 1px solid #ddd;
    border-radius: 4px;
    transition: all 0.3s;
}

.filter-form button, 
.apply-filter-btn {
    background-color: #0b00a2;
    color: white;
    border: none;
    cursor: pointer;
}

.filter-form button:hover, 
.apply-filter-btn:hover {
    background-color: #09007a;
}

.filter-form select:focus, 
.filter-form select.highlighted,
.password-input:focus {
    border: 2px solid #0b00a2;
    background-color: #f0f5ff;
    outline: none;
    box-shadow: 0 0 0 2px rgba(11, 0, 162, 0.2);
}

#filters {
    margin-bottom: 20px;
}

.filter-tag {
    display: inline-block;
    background: #e9ecef;
    padding: 5px 10px;
    border-radius: 20px;
    margin-right: 10px;
    margin-bottom: 10px;
}

.filter-tag button {
    background: none;
    border: none;
    color: #dc3545;
    margin-left: 5px;
    cursor: pointer;
}

/* Table container */
.table-scroll-container {
    position: relative;
    height: calc(100vh - 300px);
    overflow: auto;
    margin-top: 20px;
    border: 1px solid #ddd;
    border-radius: 8px;
    background: white;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

/* Table structure */
#results {
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    min-width: 1200px;
    background: white;
}

/* Header rows */
#results thead {
    position: sticky;
    top: 0;
    z-index: 1000;
}

/* Column controls row (first header row) */
.column-controls {
    position: sticky;
    top: 0;
    z-index: 1010;
    background-color: #f8f9fa;
}

/* Regular headers (second header row) */
#results thead tr:not(.column-controls) {
    position: sticky;
    top: 40px;
    z-index: 1000;
    background-color: #f8f9fa;
}

/* Frozen columns */
.frozen-column {
    position: sticky;
    background-color: white;
    z-index: 20;
}

/* Frozen column controls need highest priority */
.column-controls th.frozen-column {
    z-index: 1030 !important;
}

#results th {
    position: sticky;
    top: 0;
    background-color: #f8f9fa;
}

#results tr:hover, 
#results tr:hover td {
    background-color: #f1f1f1;
}

#results td:last-child {
    position: sticky;
    right: 0;
    z-index: 5;
    background-color: white;
}

#results tr:hover td:last-child {
    background-color: #f1f1f1;
}

#results th button {
    padding: 12px;
    background: none;
    border: none;
    width: 100%;
    text-align: left;
    font-weight: bold;
    cursor: pointer;
}

#results th button:hover {
    color: #0b00a2;
}

#results th button .sort-icon {
    margin-left: 5px;
    opacity: 0.5;
}

#results th button:hover .sort-icon,
#results th.sorted-asc button .sort-icon,
#results th.sorted-desc button .sort-icon {
    opacity: 1;
}

#results th.sorted-asc button,
#results th.sorted-desc button {
    color: #0b00a2;
    font-weight: bold;
}

.glyphicon, .trend-icon {
    font-size: 25px;
    position: sticky; 
    color: #bcbcbc;
    margin-bottom: 0;
}

.glyphicon:hover {
    color: #1e00ff;
}

.trend-icon {
    font-size: 1.2em;
    margin-left: 3px;
}

.loading {
    text-align: center;
    padding: 20px;
    font-style: italic;
    color: #6c757d;
}

.frozen-column.highlighted-column {
    z-index: 25 !important;
    background-color: #e6f0ff !important;
}

.highlighted-column {
    background-color: #f0f5ff !important;
    font-weight: bold;
}

/* Modal Styles */
.modal-overlay {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(0,0,0,0.5);
    display: flex;
    justify-content: center;
    align-items: center;
    z-index: 1000;
}

.modal-content {
    background: white;
    padding: 20px;
    border-radius: 8px;
    width: 90%;
    max-width: 800px;
    max-height: 90vh;
    overflow-y: auto;
    box-shadow: 0 4px 8px rgba(0,0,0,0.2);
}

.modal-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 20px;
    padding-bottom: 10px;
    border-bottom: 1px solid #eee;
}

.close-button {
    background: none;
    border: none;
    font-size: 24px;
    cursor: pointer;
    color: #666;
}

.detail-section {
    margin-bottom: 25px;
    padding-bottom: 15px;
    border-bottom: 1px solid #f0f0f0;
}

.detail-section:last-child {
    border-bottom: none;
}

.detail-grid, 
.variation-grid,
.stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
    gap: 12px;
}

.detail-item, 
.variation-item,
.stat-item {
    padding: 8px;
    background: #f9f9f9;
    border-radius: 4px;
}

.detail-item strong, 
.variation-item strong,
.stat-item strong {
    color: #333;
}

.modal-body {
    display: flex;
    flex-direction: column;
    gap: 20px;
}

.detail-section h3,
.common-values-section h3 {
    margin-top: 0;
    color: #0b00a2;
    border-bottom: 1px solid #dee2e6;
    padding-bottom: 8px;
}

.variation-item {
    border-left: 3px solid #0b00a2;
}

.filter-buttons {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}

.filter-buttons button,
.action-btn {
    background-color: #e9ecef;
    color: #495057;
    border: 1px solid #ced4da;
    padding: 5px 10px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 0.9rem;
    transition: all 0.2s;
}

.filter-buttons button:hover,
.action-btn:hover {
    background-color: #dee2e6;
}

.filter-buttons button.active {
    background-color: #0b00a2;
    color: white;
    border-color: #0b00a2;
}

#loadingBarContainer {
    width: 100%;
    background-color: #f1f1f1;
    padding: 3px;
    border-radius: 5px;
    margin: 10px 0;
}

#loadingBar {
    width: 0%;
    height: 20px;
    background-color: #00a231;
    border-radius: 3px;
    transition: width 0.3s;
    text-align: center;
    line-height: 20px;
    color: white;
}

#loadingText {
    text-align: center;
    margin-top: 5px;
    font-size: 0.9rem;
    color: #333;
}

#passwordContainer {
    display: flex;
    flex-direction: row;
    gap: 15px;
    align-items: center;
    width: 100%;
}

.password-input-group {
    position: relative;
    min-width: 0;
    flex: 1;
}

.toggle-password {
    position: absolute;
    right: 10px;
    top: 50%;
    transform: translateY(-50%);
    cursor: pointer;
}

.toggle-password:hover {
    opacity: 1;
}

.password-strength {
    margin-top: 5px;
    height: 4px;
    background-color: #eee;
    border-radius: 2px;
    overflow: hidden;
}

.password-strength-bar {
    height: 100%;
    width: 0%;
    transition: width 0.3s, background-color 0.3s;
}

@keyframes shake {
    0%, 100% { transform: translateX(0); }
    20%, 60% { transform: translateX(-5px); }
    40%, 80% { transform: translateX(5px); }
}

.shake {
    animation: shake 0.6s;
}

/* File Upload Styles */
.file-upload-container {
    margin: 10px 0;
}

.file-upload-label {
    display: inline-block;
    cursor: pointer;
}

.file-upload-button {
    display: inline-block;
    padding: 8px 12px;
    background-color: rgb(0, 176, 15);
    color: white;
    border-radius: 4px;
    transition: background-color 0.3s;
}

.file-upload-button:hover {
    background-color: #09007a;
}

.file-upload-input {
    position: absolute;
    width: 1px;
    height: 1px;
    padding: 0;
    margin: -1px;
    overflow: hidden;
    clip: rect(0, 0, 0, 0);
    border: 0;
}

.file-upload-help {
    display: block;
    font-size: 0.8rem;
    color: #666;
    margin-top: 4px;
}

.file-upload-status {
    display: block;
    margin-top: 4px;
    font-size: 0.9rem;
}

/* Column Stats Modal Styles */
.stats-grid {
    margin-bottom: 20px;
}

.stat-item {
    border-left: 4px solid #0b00a2;
}

.stat-item span {
    font-size: 1.1rem;
    font-weight: bold;
    color: #0b00a2;
}

.common-values-section {
    margin: 25px 0;
}

.common-values-grid {
    display: grid;
    grid-template-columns: 1fr;
    gap: 8px;
}

.common-value-item {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px;
    background: #f8f9fa;
    border-radius: 4px;
}

.common-value-item .value {
    flex: 2;
    font-weight: bold;
}

.common-value-item .count {
    flex: 1;
    text-align: right;
    color: #6c757d;
    margin: 0 15px;
}

.quick-filter-actions {
    display: flex;
    gap: 10px;
    margin-top: 20px;
    flex-wrap: wrap;
}

.action-btn {
    padding: 8px 12px;
}

.action-btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
}

/* Tab Styles */
.modal-tabs {
    display: flex;
    border-bottom: 1px solid #ddd;
    margin-bottom: 15px;
}

.tab-btn {
    padding: 10px 20px;
    background: none;
    border: none;
    border-bottom: 3px solid transparent;
    cursor: pointer;
    font-weight: bold;
    color: #666;
}

.tab-btn.active {
    color: #0b00a2;
    border-bottom-color: #0b00a2;
}

.tab-content {
    display: none;
}

.tab-content.active {
    display: block;
}

/* Values Tab Styles */
.values-search {
    margin-bottom: 15px;
    display: flex;
    align-items: center;
    gap: 10px;
}

#values-search-input {
    flex: 1;
    padding: 8px 12px;
    border: 1px solid #ddd;
    border-radius: 4px;
}

.values-count {
    color: #666;
    font-size: 0.9rem;
}

.values-container {
    max-height: 400px;
    overflow-y: auto;
    margin-bottom: 15px;
    border: 1px solid #eee;
    border-radius: 4px;
}

.value-item {
    display: flex;
    align-items: center;
    padding: 8px 12px;
    border-bottom: 1px solid #f0f0f0;
}

.value-item:nth-child(even) {
    background-color: #f9f9f9;
}

.value-item .value {
    flex: 2;
    padding-right: 10px;
}

.value-item .count {
    flex: 1;
    text-align: right;
    color: #6c757d;
    font-size: 0.9rem;
    padding-right: 15px;
}

.values-pagination {
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 15px;
    margin-top: 10px;
}

.values-pagination button {
    padding: 5px 10px;
    background: #f8f9fa;
    border: 1px solid #ddd;
    border-radius: 4px;
    cursor: pointer;
}

.values-pagination button:disabled {
    opacity: 0.5;
    cursor: not-allowed;
}

.values-pagination button:hover:not(:disabled) {
    background: #e9ecef;
}

/* Column controls styling */
.column-controls th {
    padding: 5px !important;
    height: 40px;
    vertical-align: middle;
}

.column-controls-container {
    display: flex;
    flex-direction: row;
    align-items: center;
}

.freeze-btn {
    background: none;
    border: none;
    cursor: pointer;
    font-size: 14px;
    padding: 2px;
    color: #bcbcbc;
}

.freeze-btn:hover {
    color: #0b00a2;
}

.freeze-btn.active, 
.freeze-btn.active .glyphicon {
    color: #ff0000 !important;
}

.width-slider {
    width: 80px;
}

/* Headers need higher z-index */
#results thead th.frozen-column {
    z-index: 120;
}

/* Body cells */
#results tbody td.frozen-column {
    z-index: 10;
}

/* Special styling for the column controls row */
.column-controls th.frozen-column {
    z-index: 130 !important;
}

/* Ensure the slider stays with the column */
.width-slider {
    position: sticky;
    left: 0;
    z-index: 10;
}

/* Adjust left positions for multiple frozen columns */
.frozen-column ~ .frozen-column {
    left: auto;
    box-shadow: none;
}

@media (max-width: 768px) {
    .filter-form select, 
    .filter-form input, 
    .filter-form button,
    #analyzeButton {
        width: 100%;
        margin-right: 0;
    }
    
    #results thead {
        display: none;
    }
    
    #results tr {
        display: block;
        margin-bottom: 15px;
        border: 1px solid #ddd;
        border-radius: 8px;
    }
    
    #results td {
        display: flex;
        justify-content: space-between;
        align-items: center;
        text-align: right;
        padding-left: 50%;
        position: relative;
    }
    
    #results td::before {
        content: attr(data-label);
        position: absolute;
        left: 12px;
        font-weight: bold;
    }
    
    .detail-grid, 
    .variation-grid,
    .stats-grid {
        grid-template-columns: 1fr;
    }
    
    #passwordContainer {
        flex-direction: column;
        gap: 10px;
    }
    
    .password-input-group {
        width: 100%;
    }
}
"@

Write-Host "🏗️ Creating Javascript" -ForegroundColor $YELLOW
    #javascript
Set-Content -Path "static/script.js" -Value @'
// Global variables
let allData = [];
let filteredData = [];
let lastSelectedColumn = '';
let currentFilterColumn = '';
let currentSortColumn = '';
let sortDirection = 'asc';
let processingData = false;
const filters = [];
let currentDataSource = 'src/data.json';
let selectedFile = null;
let frozenColumns = [];

// DOM elements
const operatorSelect = document.getElementById('operator');
const value2Input = document.getElementById('value2');

// Global variables for values pagination
let currentValuesPage = 1;
const valuesPerPage = 50;
let currentValuesSearch = '';
let currentValuesColumn = '';

// Initializing the application
document.addEventListener('DOMContentLoaded', () => {
    loadData();
    setupEventListeners();
});

function renderValuesPage() {
    const container = document.getElementById('values-container');
    if (!container) return;
    
    const columnValues = allData.map(item => item[currentValuesColumn]);
    const valueCounts = {};
    columnValues.forEach(v => {
        if (v !== undefined && v !== null) {
            const val = typeof v === 'string' ? v.trim() : v;
            valueCounts[val] = (valueCounts[val] || 0) + 1;
        }
    });
    
    let allValues = Object.keys(valueCounts)
        .sort((a, b) => {
            const isNumeric = !isNaN(parseFloat(a)) && !isNaN(parseFloat(b));
            if (isNumeric) {
                return parseFloat(a) - parseFloat(b);
            }
            return a.localeCompare(b);
        });
    
    // Apply search filter
    if (currentValuesSearch) {
        const searchTerm = currentValuesSearch.toLowerCase();
        allValues = allValues.filter(v => 
            String(v).toLowerCase().includes(searchTerm)
        );
    }
    
    const totalValues = allValues.length;
    const totalPages = Math.ceil(totalValues / valuesPerPage);
    const startIdx = (currentValuesPage - 1) * valuesPerPage;
    const endIdx = Math.min(startIdx + valuesPerPage, totalValues);
    const pageValues = allValues.slice(startIdx, endIdx);
    
    // Update UI
    container.innerHTML = pageValues.map(value => `
        <div class="value-item">
            <span class="value">${formatValueForDisplay(value)}</span>
            <span class="count">${valueCounts[value]} (${Math.round((valueCounts[value] / columnValues.length) * 100)}%)</span>
            <button onclick="applyCommonValueFilter('${currentValuesColumn}', '${value.replace(/'/g, "\\'")}')" 
                    class="apply-filter-btn">
                Filtrar
            </button>
        </div>
    `).join('');
    
    document.getElementById('values-showing').textContent = `${startIdx + 1}-${endIdx}`;
    document.getElementById('values-page-info').textContent = `Página ${currentValuesPage} de ${totalPages}`;
    document.getElementById('values-prev').disabled = currentValuesPage <= 1;
    document.getElementById('values-next').disabled = currentValuesPage >= totalPages;
}

function formatValueForDisplay(value) {
    if (value === null || value === undefined) return 'N/A';
    if (typeof value === 'string') return value;
    if (Math.abs(value) >= 1000000) {
        return '$' + (value / 1000000).toFixed(2) + 'M';
    }
    return new Intl.NumberFormat('es-CO').format(value);
}

function searchValues(columnName) {
    currentValuesSearch = document.getElementById('values-search-input').value;
    currentValuesPage = 1;
    currentValuesColumn = columnName;
    renderValuesPage();
}

function navigateValuesPage(direction) {
    currentValuesPage += direction;
    renderValuesPage();
}

function switchTab(tabId, button) {
    // Hide all tab contents
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Deactivate all tab buttons
    document.querySelectorAll('.modal-tabs .tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // Activate selected tab
    document.getElementById(`${tabId}-tab`).classList.add('active');
    button.classList.add('active');
    
    // If switching to values tab and not yet loaded
    if (tabId === 'values' && document.getElementById('values-container').innerHTML === '') {
        renderValuesPage();
    }
}


// Load JSON data
async function loadData() {
    try {
        const timestamp = new Date().getTime();
        document.querySelector('#results tbody').innerHTML = `
            <tr>
                <td colspan="35" class="loading">Cargando datos...</td>
            </tr>
        `;

        const response = await fetch(`${currentDataSource}?t=${timestamp}`);
        if (!response.ok) throw new Error(`Error al cargar ${currentDataSource}`);
        
        allData = await response.json();
        filteredData = [...allData];
        renderTable();
        
    } catch (error) {
        console.error('Error:', error);
        document.querySelector('#results tbody').innerHTML = `
            <tr>
                <td colspan="35">Carga el archivo excel para generar el análisis de datos</td>
            </tr>
        `;
    }
}

    function togglePassword(inputId) {
        const input = document.getElementById(inputId);
        const toggle = input.nextElementSibling;
        
        if (input.type === 'password') {
            input.type = 'text';
            toggle.textContent = '🙈';
        } else {
            input.type = 'password';
            toggle.textContent = '👁️';
        }
    }

    function showPasswordError(message) {
        const errorElement = document.getElementById('passwordError');
        errorElement.textContent = message;
        
        // Add shake animation to password inputs
        document.querySelectorAll('.password-input').forEach(input => {
            input.classList.add('shake');
            setTimeout(() => input.classList.remove('shake'), 600);
        });
        
        // Highlight inputs in red
        document.querySelectorAll('.password-input').forEach(input => {
            input.style.borderColor = '#dc3545';
            setTimeout(() => {
                input.style.borderColor = input === document.activeElement ? '#0b00a2' : '#ddd';
            }, 2000);
        });
    }

    function clearPasswordError() {
        document.getElementById('passwordError').textContent = '';
        document.querySelectorAll('.password-input').forEach(input => {
            input.style.borderColor = input === document.activeElement ? '#0b00a2' : '#ddd';
        });
    }

    // Update your analyzeButton click handler to handle password errors
    document.getElementById('analyzeButton').addEventListener('click', async function() {
        if (!selectedFile) {
            alert('Por favor seleccione un archivo primero');
            return;
        }
        
        clearPasswordError();
        
        const openPassword = document.getElementById('excelOpenPassword').value;
        const modifyPassword = document.getElementById('excelModifyPassword').value;
        const statusElement = document.getElementById('fileUploadStatus');
        const loadingBarContainer = document.getElementById('loadingBarContainer');
        const loadingBar = document.getElementById('loadingBar');
        
        try {
            loadingBarContainer.style.display = 'block';
            loadingBar.style.width = '10%';
            statusElement.textContent = 'Preparando análisis...';
            statusElement.style.color = '#0b00a2';
            
            const result = await processExcelFile(selectedFile, openPassword, modifyPassword);
            
            loadingBar.style.width = '100%';
            statusElement.textContent = 'Finalizando...';
            
            if (result.success) {
                statusElement.textContent = 'Análisis completado correctamente! Los datos se han actualizado.';
                statusElement.style.color = 'green';
                
                document.getElementById('excelUpload').value = '';
                document.getElementById('excelOpenPassword').value = '';
                document.getElementById('excelModifyPassword').value = '';
                selectedFile = null;
                
                setTimeout(() => {
                    loadingBarContainer.style.display = 'none';
                    loadingBar.style.width = '0%';
                }, 1000);
                
                await loadData();
            } else {
                throw new Error(result.message || 'Error desconocido al procesar el archivo');
            }
        } catch (error) {
            console.error('Error details:', error);
            loadingBarContainer.style.display = 'none';
            loadingBar.style.width = '0%';
            
            // This is where we simplify the error message
            if (error.message.toLowerCase().includes('intenta de nuevo') || 
                error.message.toLowerCase().includes('password') || 
                error.message.toLowerCase().includes('contraseña')) {
                showPasswordError('Verifica tu contraseña');
                statusElement.textContent = 'Intenta de nuevo';
            } else {
                statusElement.textContent = `Error: ${error.message}`;
            }
            statusElement.style.color = 'red';
        }
    });
    
    // Add this new function to handle file processing
    async function processExcelFile(file, openPassword, modifyPassword) {
        const formData = new FormData();
        formData.append('file', file);
        if (openPassword) formData.append('openPassword', openPassword);
        if (modifyPassword) formData.append('modifyPassword', modifyPassword);
    
        try {
            const response = await fetch('http://localhost:8000/upload', {
                method: 'POST',
                body: formData,
                headers: {
                    'Accept': 'application/json'
                }
            });
    
            if (!response.ok) {
                const errorData = await response.json();
                if (errorData.message && errorData.message.toLowerCase().includes('contraseña')) {
                    throw new Error('Intenta de nuevo');
                }
                throw new Error(errorData.message || 'Error desconocido');
            }
    
            return await response.json();
        } catch (error) {
            console.error("Fetch error:", error);
            throw error; // This will be caught by the calling function
        }
    }
    

function changeDataSource() {
    const dataSourceSelect = document.getElementById('dataSource');
    currentDataSource = dataSourceSelect.value;
    
    // Clear existing filters
    filters.length = 0;
    renderFilters();
    
    // Reset sort
    currentSortColumn = '';
    sortDirection = 'asc';
    
    // Reload data
    loadData();
    
    // Clear any highlights
    document.querySelectorAll('.highlighted-column').forEach(el => {
        el.classList.remove('highlighted-column');
    });
    document.getElementById('column').classList.remove('highlighted');
    currentFilterColumn = '';
    lastSelectedColumn = '';
    
    // Reset column dropdown
    document.getElementById('column').selectedIndex = 0;
}

// Set up event listeners
function setupEventListeners() {
    operatorSelect.addEventListener('change', toggleValue2Input);
    
    document.getElementById('column').addEventListener('change', function() {
        currentFilterColumn = this.value;
        if (this.value) {
            this.classList.add('highlighted');
        } else {
            this.classList.remove('highlighted');
        }
        
        // Auto-focus the value input for quick filtering
        if (this.value && lastSelectedColumn !== this.value) {
            document.getElementById('value1').focus();
        }
        lastSelectedColumn = this.value;
    });
    // listener to excelupload
    document.getElementById('excelUpload').addEventListener('change', function(e) {
        const statusElement = document.getElementById('fileUploadStatus');
        const passwordContainer = document.getElementById('passwordContainer');
        
        if (this.files.length > 0) {
            selectedFile = this.files[0];
            statusElement.textContent = `Archivo seleccionado: ${selectedFile.name}`;
            statusElement.style.color = '#0b00a2';
            passwordContainer.style.display = 'block'; // Show both password fields
        } else {
            selectedFile = null;
            statusElement.textContent = '';
            passwordContainer.style.display = 'none'; // Hide both password fields
        }
    });
}

// excel upload fucntion
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    // Show loading state
    processingData = true;
    document.querySelector('#results tbody').innerHTML = `
        <tr>
            <td colspan="37" class="loading">Procesando archivo Excel, por favor espere...</td>
        </tr>
    `;

    try {
        const formData = new FormData();
        formData.append('file', file);

        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error('Error en el análisis del archivo');
        }

        const result = await response.json();
        
        if (result.success) {
            // Reload the data after processing
            await loadData();
            alert('Archivo procesado correctamente. Los datos han sido actualizados.');
        } else {
            throw new Error(result.message || 'Error desconocido al procesar el archivo');
        }
    } catch (error) {
        console.error('Error:', error);
        document.querySelector('#results tbody').innerHTML = `
            <tr>
                <td colspan="37">Error al procesar el archivo: ${error.message}</td>
            </tr>
        `;
    } finally {
        processingData = false;
        // Reset the file input
        event.target.value = '';
    }
}

// Toggle second value input for 'between' operator
function toggleValue2Input() {
    value2Input.style.display = operatorSelect.value === 'between' ? 'inline-block' : 'none';
}

// Add a new filter
function addFilter() {
    const column = document.getElementById('column').value;
    const operator = operatorSelect.value;
    const value1 = document.getElementById('value1').value.trim();
    let value2 = '';
    
    if (!column || !operator || !value1) {
        alert('Por favor complete todos los campos del filtro');
        return;
    }
    
    if (operator === 'between') {
        value2 = document.getElementById('value2').value.trim();
        if (!value2) {
            alert('Por favor complete el segundo valor para el filtro "Entre"');
            return;
        }
    }
    
    // Check if we should use string comparison instead of numeric
    const isNumericComparison = ['>', '<', '=', '>=', '<=', 'between'].includes(operator) && 
                               !isNaN(parseFloat(value1)) && 
                               (operator !== 'between' || !isNaN(parseFloat(value2)));
    
    filters.push({ 
        column, 
        operator, 
        value1, 
        value2,
        isNumericComparison  // Add this flag to the filter
    });
    
    renderFilters();
    applyFilters();
    
    // Highlight the column in the table
    highlightColumn(column);
    
    // Keep the column selected in the dropdown
    lastSelectedColumn = column;
    currentFilterColumn = column;
}

function sortTable(columnName, direction) {
    filteredData.sort((a, b) => {
        let valA = a[columnName];
        let valB = b[columnName];
        
        // Handle numeric values
        if (!isNaN(parseFloat(valA)) && !isNaN(parseFloat(valB))) {
            valA = parseFloat(valA);
            valB = parseFloat(valB);
            return direction === 'asc' ? valA - valB : valB - valA;
        }
        
        // Handle string values
        if (typeof valA === 'string' && typeof valB === 'string') {
            return direction === 'asc' 
                ? valA.localeCompare(valB) 
                : valB.localeCompare(valA);
        }
        
        return 0;
    });
    
    renderTable();
    
    // Re-apply column highlight if there's a current filter column
    if (currentFilterColumn) {
        highlightColumn(currentFilterColumn);
    }
}

// Highlight table column
function highlightColumn(columnName) {
    // Remove any existing highlights
    const headers = document.querySelectorAll('#results th');
    const cells = document.querySelectorAll('#results td');
    
    headers.forEach(header => header.classList.remove('highlighted-column'));
    cells.forEach(cell => cell.classList.remove('highlighted-column'));
    
    // Find the column index
    const columnMap = {
        'Usuario': 4,
        'Nombre': 0,
        'Compañía': 2,
        'Cargo': 3,
        'Año Declaración': 1,
        'Activos': 5,
        'Pasivos': 6,
        'Patrimonio': 7,
        'Apalancamiento': 8,
        'Endeudamiento': 9,
        'Cant_Deudas': 10,
        'BancoSaldo': 11,
        'Cant_Bancos': 12,
        'Bienes': 13,
        'Cant_Bienes': 14,
        'Inversiones': 15,
        'Cant_Inversiones': 16,
        'Ingresos': 17,
        'Cant_Ingresos': 18,
        'Activos Var. Abs.': 19,
        'Pasivos Var. Abs.': 20,
        'Patrimonio Var. Abs.': 21,
        'Apalancamiento Var. Abs.': 22,
        'Endeudamiento Var. Abs.': 23,
        'BancoSaldo Var. Abs.': 24,
        'Bienes Var. Abs.': 25,
        'Inversiones Var. Abs.': 26,
        'Ingresos Var. Abs.': 27,
        'Activos Var. Rel.': 28,
        'Pasivos Var. Rel.': 29,
        'Patrimonio Var. Rel.': 30,
        'Apalancamiento Var. Rel.': 31,
        'Endeudamiento Var. Rel.': 32,
        'BancoSaldo Var. Rel.': 33,
        'Bienes Var. Rel.': 34,
        'Inversiones Var. Rel.': 35,
        'Ingresos Var. Rel.': 36
    };
    
    const columnIndex = columnMap[columnName];
    if (columnIndex === undefined) return;
    
    // Highlight header
    if (headers[columnIndex]) {
        headers[columnIndex].classList.add('highlighted-column');
    }
    
    // Highlight cells - including frozen columns
    document.querySelectorAll(`#results tr > *:nth-child(${columnIndex + 1})`).forEach(cell => {
        cell.classList.add('highlighted-column');
    });
    
    // Update current filter column
    currentFilterColumn = columnName;
    lastSelectedColumn = columnName;
}

function removeFilter(index) {
    // Remove the filter at the specified index
    filters.splice(index, 1);
    
    // Re-render the remaining filters
    renderFilters();
    
    // Re-apply the remaining filters
    applyFilters();
    
    // If no filters remain, clear the highlights
    if (filters.length === 0) {
        document.querySelectorAll('.highlighted-column').forEach(el => {
            el.classList.remove('highlighted-column');
        });
        document.getElementById('column').classList.remove('highlighted');
        currentFilterColumn = '';
        lastSelectedColumn = '';
    }
}

// Render active filters
function renderFilters() {
    const filtersContainer = document.getElementById('filters');
    filtersContainer.innerHTML = filters.map((filter, index) => `
        <div class="filter-tag">
            ${filter.column} ${getOperatorSymbol(filter.operator)} ${filter.value1}
            ${filter.operator === 'between' ? ` y ${filter.value2}` : ''}
            <button onclick="removeFilter(${index})">×</button>
        </div>
    `).join('');
}

// Get operator symbol for display
function getOperatorSymbol(operator) {
    const symbols = {
        '>': '>',
        '<': '<',
        '=': '=',
        '>=': '≥',
        '<=': '≤',
        'between': 'entre',
        'contains': 'contiene'
    };
    return symbols[operator] || operator;
}

// Remove a filter
function applyFilters() {
    if (filters.length === 0) {
        filteredData = [...allData];
        renderTable();
        return;
    }
    
    filteredData = allData.filter(item => {
        return filters.every(filter => {
            const itemValue = item[filter.column];
            if (itemValue === undefined || itemValue === null) return false;
            
            // Handle string comparison
            if (!filter.isNumericComparison) {
                const itemStr = String(itemValue).toLowerCase();
                const filterStr = filter.value1.toLowerCase();
                
                switch (filter.operator) {
                    case 'contains': 
                        return itemStr.includes(filterStr);
                    case '=':
                        return itemStr === filterStr;
                    default:
                        return true;
                }
            }
            
            // Handle numeric comparison
            let numericValue;
            if (typeof itemValue === 'string' && itemValue.includes('%')) {
                numericValue = parseFloat(itemValue.replace('%', ''));
            } else {
                numericValue = parseFloat(itemValue);
            }
            
            const filterValue1 = parseFloat(filter.value1);
            const filterValue2 = parseFloat(filter.value2);
            
            switch (filter.operator) {
                case '>': return numericValue > filterValue1;
                case '<': return numericValue < filterValue1;
                case '=': return numericValue === filterValue1;
                case '>=': return numericValue >= filterValue1;
                case '<=': return numericValue <= filterValue1;
                case 'between': 
                    return numericValue >= filterValue1 && numericValue <= filterValue2;
                case 'contains':
                    return String(itemValue).toLowerCase().includes(filter.value1.toLowerCase());
                default: return true;
            }
        });
    });
    
    renderTable();
    
    // Re-apply column highlight if there's a current filter column
    if (currentFilterColumn) {
        highlightColumn(currentFilterColumn);
    }
}

// Handle predetermined filters
function applyPredeterminedFilter(column, operator, value1) {
    // Check if this filter already exists
    const existingIndex = filters.findIndex(f => 
        f.column === column && f.operator === operator && f.value1 === value1
    );
    
    if (existingIndex === -1) {
        // Add new filter
        filters.push({ column, operator, value1 });
        
        // Highlight the button
        const buttons = document.querySelectorAll('.filter-buttons button');
        buttons.forEach(button => {
            if (button.textContent.includes(column) && 
                button.textContent.includes(operator) &&
                button.textContent.includes(value1)) {
                button.classList.add('active');
            }
        });
    } else {
        // Remove existing filter
        filters.splice(existingIndex, 1);
        
        // Remove highlight from button
        const buttons = document.querySelectorAll('.filter-buttons button');
        buttons.forEach(button => {
            if (button.textContent.includes(column) && 
                button.textContent.includes(operator) &&
                button.textContent.includes(value1)) {
                button.classList.remove('active');
            }
        });
    }
    
    renderFilters();
    applyFilters();
    highlightColumn(column);
}

// Clear all filters
function clearFilters() {
    // Clear the filters array
    filters.length = 0;
    
    // Reset the filter form inputs
    document.getElementById('column').value = '';
    document.getElementById('operator').value = '>';
    document.getElementById('value1').value = '';
    document.getElementById('value2').value = '';
    document.getElementById('value2').style.display = 'none';
    
    // Reset the UI
    renderFilters();
    filteredData = [...allData];
    renderTable();
    
    // Clear button highlights
    document.querySelectorAll('.filter-buttons button').forEach(button => {
        button.classList.remove('active');
    });
    
    // Clear column highlights
    document.querySelectorAll('.highlighted-column').forEach(el => {
        el.classList.remove('highlighted-column');
    });
    
    // Reset column dropdown
    document.getElementById('column').classList.remove('highlighted');
    currentFilterColumn = '';
    lastSelectedColumn = '';
    
    // Reset sort indicators
    document.querySelectorAll('#results th .sort-icon').forEach(icon => {
        icon.textContent = '↕';
    });
    document.querySelectorAll('#results th').forEach(th => {
        th.classList.remove('sorted-asc', 'sorted-desc');
    });
    
    // Reset sort state
    currentSortColumn = '';
    sortDirection = 'asc';
}

// Render the data table
function renderTable() {
    const tbody = document.querySelector('#results tbody');
    
    if (filteredData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="37">Generando datos desde el archivo excel</td>
            </tr>
        `;
        return;
    }
    
    tbody.innerHTML = filteredData.map(item => {
        // Function to format cell with color based on value
        const formatCell = (value, isPercentage = false) => {
            if (value === undefined || value === null || value === '') return '';
            
            // Remove trend icons if present
            const cleanValue = String(value).replace(/[📈📉➡️]/g, '').trim();
            
            // Try to parse as number
            const numValue = parseFloat(cleanValue.replace('%', '').replace(/[^\d.-]/g, ''));
            if (isNaN(numValue)) return value;
            
            // Format number
            let formattedValue;
            if (isPercentage) {
                formattedValue = numValue.toFixed(2) + '%';
            } else if (Math.abs(numValue) >= 1000000) {
                formattedValue = '$' + (numValue / 1000000).toFixed(2) + 'M';
            } else {
                formattedValue = new Intl.NumberFormat('es-CO').format(numValue);
            }
            
            // Determine color - only red for negative, black otherwise
            const color = numValue < 0 ? 'color: #dc3545;' : 'color: #000;';
            
            return `<span style="${color}">${formattedValue}</span>`;
        };
        
        return `
            <tr>
                <td>${formatCell(item.Nombre)}</td>
                <td>${formatCell(item['Año Declaración'])}</td>
                <td>${formatCell(item['Compañía'])}</td>
                <td>${formatCell(item.Cargo)}</td>
                <td>${formatCell(item.Usuario)}</td>
                <td>${formatCell(item.Activos)}</td>
                <td>${formatCell(item.Pasivos)}</td>
                <td>${formatCell(item.Patrimonio)}</td>
                <td>${formatCell(item.Apalancamiento, true)}</td>
                <td>${formatCell(item.Endeudamiento, true)}</td>
                <td>${formatCell(item['Cant_Deudas'])}</td>
                <td>${formatCell(item.BancoSaldo)}</td>
                <td>${formatCell(item.Cant_Bancos)}</td>
                <td>${formatCell(item.Bienes)}</td>
                <td>${formatCell(item.Cant_Bienes)}</td>
                <td>${formatCell(item.Inversiones)}</td>
                <td>${formatCell(item.Cant_Inversiones)}</td>
                <td>${formatCell(item.Ingresos)}</td>
                <td>${formatCell(item.Cant_Ingresos)}</td>
                <td>${formatCell(item['Activos Var. Abs.'])}</td>
                <td>${formatCell(item['Pasivos Var. Abs.'])}</td>
                <td>${formatCell(item['Patrimonio Var. Abs.'])}</td>
                <td>${formatCell(item['Apalancamiento Var. Abs.'], true)}</td>
                <td>${formatCell(item['Endeudamiento Var. Abs.'], true)}</td>
                <td>${formatCell(item['BancoSaldo Var. Abs.'])}</td>
                <td>${formatCell(item['Bienes Var. Abs.'])}</td>
                <td>${formatCell(item['Inversiones Var. Abs.'])}</td>
                <td>${formatCell(item['Ingresos Var. Abs.'])}</td>
                <td>${formatCell(item['Activos Var. Rel.'], true)}</td>
                <td>${formatCell(item['Pasivos Var. Rel.'], true)}</td>
                <td>${formatCell(item['Patrimonio Var. Rel.'], true)}</td>
                <td>${formatCell(item['Apalancamiento Var. Rel.'], true)}</td>
                <td>${formatCell(item['Endeudamiento Var. Rel.'], true)}</td>
                <td>${formatCell(item['BancoSaldo Var. Rel.'], true)}</td>
                <td>${formatCell(item['Bienes Var. Rel.'], true)}</td>
                <td>${formatCell(item['Inversiones Var. Rel.'], true)}</td>
                <td>${formatCell(item['Ingresos Var. Rel.'], true)}</td>
                <td style="position: sticky; right: 0; background-color: white;">
                    <button onclick="viewDetails('${item.Usuario}', ${item['Año Declaración']})" style="background-color: #0b00a2; color: white; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer;
                    ">Ver</button>
                </td>
            </tr>
        `;
    }).join('');

    // After rendering, update frozen columns
    if (frozenColumns.length > 0) {
        updateFrozenColumns();
    }
    
    // Ensure proper z-index stacking
    document.querySelectorAll('#results thead th').forEach(th => {
        th.style.zIndex = '20';
    });
}

// Track frozen columns

function toggleFreezeColumn(columnIndex) {
    const columnCells = document.querySelectorAll(`#results tr > *:nth-child(${columnIndex + 1})`);
    const freezeBtn = document.querySelector(`th:nth-child(${columnIndex + 1}) .freeze-btn`);
    
    // Get the column name from the header
    const columnName = document.querySelector(`#results thead tr:not(.column-controls) th:nth-child(${columnIndex + 1}) button`).textContent;
    
    if (frozenColumns.includes(columnIndex)) {
        // Unfreeze
        frozenColumns = frozenColumns.filter(col => col !== columnIndex);
        columnCells.forEach(cell => {
            cell.classList.remove('frozen-column');
            cell.style.left = '';
        });
        freezeBtn.classList.remove('active');
        
        // Remove highlight if this was the only frozen column
        if (frozenColumns.length === 0) {
            document.querySelectorAll('.highlighted-column').forEach(el => {
                el.classList.remove('highlighted-column');
            });
        }
    } else {
        // Freeze
        frozenColumns.push(columnIndex);
        frozenColumns.sort((a, b) => a - b); // Keep in order
        freezeBtn.classList.add('active');
        
        // Highlight the column
        highlightColumn(columnName);
    }
    
    updateFrozenColumns();
    
    // Force a reflow to ensure proper rendering
    document.querySelector('.table-scroll-container').style.overflow = 'hidden';
    document.querySelector('.table-scroll-container').offsetHeight;
    document.querySelector('.table-scroll-container').style.overflow = 'auto';
}

function updateFrozenColumns() {
    // First remove all frozen classes and reset positions
    document.querySelectorAll('.frozen-column').forEach(el => {
        el.classList.remove('frozen-column');
        el.style.left = '';
    });

    // Calculate cumulative left positions
    let leftPosition = 0;
    
    frozenColumns.forEach((colIndex, i) => {
        const columnCells = document.querySelectorAll(`
            #results thead th:nth-child(${colIndex + 1}),
            #results tbody td:nth-child(${colIndex + 1})
        `);
        
        const firstCell = columnCells[0];
        const columnWidth = firstCell.offsetWidth;
        
        columnCells.forEach(cell => {
            cell.classList.add('frozen-column');
            cell.style.left = `${leftPosition}px`;
            
            // Set z-index based on cell type
            if (cell.tagName === 'TH') {
                if (cell.parentElement.classList.contains('column-controls')) {
                    cell.style.zIndex = '1030';
                } else {
                    cell.style.zIndex = '1020';
                }
            } else {
                cell.style.zIndex = '10';
            }
        });
        
        leftPosition += columnWidth;
    });
    
    // Adjust the top position of the second header row
    const controlsRow = document.querySelector('.column-controls');
    const secondHeaderRow = document.querySelector('#results thead tr:not(.column-controls)');
    
    if (controlsRow && secondHeaderRow) {
        const controlsHeight = controlsRow.offsetHeight;
        secondHeaderRow.style.top = `${controlsHeight}px`;
    }
}

// View detailed record
function viewDetails(userId, year) {
    // Convert year to number if it's coming as string
    year = typeof year === 'string' ? parseInt(year) : year;
    
    // Find the record with case-insensitive comparison
    const record = allData.find(item => {
        // Handle potential undefined/null values
        const itemUserId = item.Usuario ? item.Usuario.toString().toLowerCase() : '';
        const itemYear = item['Año Declaración'] ? parseInt(item['Año Declaración']) : null;
        
        return itemUserId === userId.toLowerCase() && itemYear === year;
    });

    if (!record) {
        alert(`Registro no encontrado para:\nUsuario: ${userId}\nAño: ${year}`);
        console.error('Record not found:', { userId, year, allData });
        return;
    }

    // Create modal HTML with all available data
    const modalHTML = `
        <div id="detailModal" class="modal-overlay">
            <div class="modal-content">
                <div class="modal-header">
                    <h2>Detalles Completo - ${record.Nombre} (${year})</h2>
                    <div>
                        <button onclick="exportDetailsToExcel()" style="margin-right: 10px; padding: 5px 10px; background-color: #0b00a2; color: white; border: none; border-radius: 4px; cursor: pointer;">
                            Exportar a Excel
                        </button>
                        <button onclick="closeModal()" class="close-button">×</button>
                    </div>
                </div>
                
                <div class="modal-body">
                    ${renderDetailSection('Información Básica', [
                        { label: 'Nombre', value: record.Nombre },
                        { label: 'Usuario', value: record.Usuario },
                        { label: 'Compañía', value: record['Compañía'] },
                        { label: 'Cargo', value: record.Cargo },
                        { label: 'Año Declaración', value: record['Año Declaración'] },
                        { label: 'Año Creación', value: record['Año Creación'] }
                    ])}
                    
                    ${renderFinancialSection('Resumen Financiero', record)}
                    
                    ${renderVariationSection('Variaciones Recientes', record)}
                    
                    <!-- Yearly Variations Sections -->
                    ${renderYearlyVariationSection('Variaciones con 2021', record, '2021')}
                    ${renderYearlyVariationSection('Variaciones con 2022', record, '2022')}
                    ${renderYearlyVariationSection('Variaciones con 2023', record, '2023')}
                    ${renderYearlyVariationSection('Variaciones con 2024', record, '2024')}
                    
                    <!-- Yearly Count Variations Sections -->
                    ${renderYearlyCountVariationSection('Variaciones de Cantidad con 2021', record, '2021')}
                    ${renderYearlyCountVariationSection('Variaciones de Cantidad con 2022', record, '2022')}
                    ${renderYearlyCountVariationSection('Variaciones de Cantidad con 2023', record, '2023')}
                    ${renderYearlyCountVariationSection('Variaciones de Cantidad con 2024', record, '2024')}
                </div>
            </div>
        </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHTML);
}

// Helper function to render detail sections
function renderDetailSection(title, fields) {
    const filteredFields = fields.filter(field => field.value !== undefined && field.value !== null);
    
    if (filteredFields.length === 0) return '';
    
    return `
        <div class="detail-section">
            <h3>${title}</h3>
            <div class="detail-grid">
                ${filteredFields.map(field => `
                    <div class="detail-item">
                        <strong>${field.label}:</strong>
                        <span>${field.value}</span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Helper function to render financial section
function renderFinancialSection(title, record) {
    const financialFields = [
        { label: 'Activos', value: formatNumber(record.Activos) },
        { label: 'Pasivos', value: formatNumber(record.Pasivos) },
        { label: 'Patrimonio', value: formatNumber(record.Patrimonio) },
        { label: 'Apalancamiento', value: record.Apalancamiento ? `${record.Apalancamiento}%` : null },
        { label: 'Endeudamiento', value: record.Endeudamiento ? `${record.Endeudamiento}%` : null },
        { label: 'Saldo Bancario', value: formatNumber(record.BancoSaldo) },
        { label: 'Bienes', value: formatNumber(record.Bienes) },
        { label: 'Inversiones', value: formatNumber(record.Inversiones) },
        { label: 'Ingresos', value: formatNumber(record.Ingresos) }
    ].filter(field => field.value !== undefined && field.value !== null);
    
    return renderDetailSection(title, financialFields);
}

// Helper function to render variation sections
function renderVariationSection(title, record) {
    const variationFields = [
        { label: 'Activos Var. Abs.', value: record['Activos Var. Abs.'] },
        { label: 'Activos Var. Rel.', value: record['Activos Var. Rel.'] },
        { label: 'Pasivos Var. Abs.', value: record['Pasivos Var. Abs.'] },
        { label: 'Pasivos Var. Rel.', value: record['Pasivos Var. Rel.'] },
        { label: 'Patrimonio Var. Abs.', value: record['Patrimonio Var. Abs.'] },
        { label: 'Patrimonio Var. Rel.', value: record['Patrimonio Var. Rel.'] },
        { label: 'Apalancamiento Var. Abs.', value: record['Apalancamiento Var. Abs.'] },
        { label: 'Apalancamiento Var. Rel.', value: record['Apalancamiento Var. Rel.'] },
        { label: 'Endeudamiento Var. Abs.', value: record['Endeudamiento Var. Abs.'] },
        { label: 'Endeudamiento Var. Rel.', value: record['Endeudamiento Var. Rel.'] },
        { label: 'BancoSaldo Var. Abs.', value: record['BancoSaldo Var. Abs.'] },
        { label: 'BancoSaldo Var. Rel.', value: record['BancoSaldo Var. Rel.'] },
        { label: 'Bienes Var. Abs.', value: record['Bienes Var. Abs.'] },
        { label: 'Bienes Var. Rel.', value: record['Bienes Var. Rel.'] },
        { label: 'Inversiones Var. Abs.', value: record['Inversiones Var. Abs.'] },
        { label: 'Inversiones Var. Rel.', value: record['Inversiones Var. Rel.'] },
        { label: 'Ingresos Var. Abs.', value: record['Ingresos Var. Abs.'] },
        { label: 'Ingresos Var. Rel.', value: record['Ingresos Var. Rel.'] }
    ].filter(field => field.value !== undefined && field.value !== null);
    
    if (variationFields.length === 0) return '';
    
    return `
        <div class="detail-section">
            <h3>${title}</h3>
            <div class="variation-grid">
                ${variationFields.map(field => `
                    <div class="variation-item">
                        <strong>${field.label}:</strong>
                        <span>${field.value}</span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Helper function to render yearly variation sections
function renderYearlyVariationSection(title, record, year) {
    const yearlyFields = [
        { label: `Activos Var. Abs. ${year}`, value: record[`${year} Activos Var. Abs.`] },
        { label: `Activos Var. Rel. ${year}`, value: record[`${year} Activos Var. Rel.`] },
        { label: `Pasivos Var. Abs. ${year}`, value: record[`${year} Pasivos Var. Abs.`] },
        { label: `Pasivos Var. Rel. ${year}`, value: record[`${year} Pasivos Var. Rel.`] },
        { label: `Patrimonio Var. Abs. ${year}`, value: record[`${year} Patrimonio Var. Abs.`] },
        { label: `Patrimonio Var. Rel. ${year}`, value: record[`${year} Patrimonio Var. Rel.`] },
        { label: `BancoSaldo Var. Abs. ${year}`, value: record[`${year} BancoSaldo Var. Abs.`] },
        { label: `BancoSaldo Var. Rel. ${year}`, value: record[`${year} BancoSaldo Var. Rel.`] },
        { label: `Bienes Var. Abs. ${year}`, value: record[`${year} Bienes Var. Abs.`] },
        { label: `Bienes Var. Rel. ${year}`, value: record[`${year} Bienes Var. Rel.`] },
        { label: `Inversiones Var. Abs. ${year}`, value: record[`${year} Inversiones Var. Abs.`] },
        { label: `Inversiones Var. Rel. ${year}`, value: record[`${year} Inversiones Var. Rel.`] },
        { label: `Ingresos Var. Abs. ${year}`, value: record[`${year} Ingresos Var. Abs.`] },
        { label: `Ingresos Var. Rel. ${year}`, value: record[`${year} Ingresos Var. Rel.`] }
    ].filter(field => field.value !== undefined && field.value !== null);
    
    if (yearlyFields.length === 0) return '';
    
    return `
        <div class="detail-section">
            <h3>${title}</h3>
            <div class="variation-grid">
                ${yearlyFields.map(field => `
                    <div class="variation-item">
                        <strong>${field.label}:</strong>
                        <span>${field.value}</span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Helper function to render yearly count variations
function renderYearlyCountVariationSection(title, record, year) {
    const countVariationFields = [
        { label: `Cant. Deudas Var. Abs. ${year}`, value: record[`${year} Cant_Deudas Var. Abs.`] },
        { label: `Cant. Deudas Var. Rel. ${year}`, value: record[`${year} Cant_Deudas Var. Rel.`] },
        { label: `Cant. Bancos Var. Abs. ${year}`, value: record[`${year} Cant_Bancos Var. Abs.`] },
        { label: `Cant. Bancos Var. Rel. ${year}`, value: record[`${year} Cant_Bancos Var. Rel.`] },
        { label: `Cant. Bienes Var. Abs. ${year}`, value: record[`${year} Cant_Bienes Var. Abs.`] },
        { label: `Cant. Bienes Var. Rel. ${year}`, value: record[`${year} Cant_Bienes Var. Rel.`] },
        { label: `Cant. Inversiones Var. Abs. ${year}`, value: record[`${year} Cant_Inversiones Var. Abs.`] },
        { label: `Cant. Inversiones Var. Rel. ${year}`, value: record[`${year} Cant_Inversiones Var. Rel.`] },
        { label: `Cant. Ingresos Var. Abs. ${year}`, value: record[`${year} Cant_Ingresos Var. Abs.`] },
        { label: `Cant. Ingresos Var. Rel. ${year}`, value: record[`${year} Cant_Ingresos Var. Rel.`] }
    ].filter(field => field.value !== undefined && field.value !== null);
    
    if (countVariationFields.length === 0) return '';
    
    return `
        <div class="detail-section">
            <h3>${title}</h3>
            <div class="variation-grid">
                ${countVariationFields.map(field => `
                    <div class="variation-item">
                        <strong>${field.label}:</strong>
                        <span>${field.value}</span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Format numbers with thousands separator
function formatNumber(num) {
    if (num === undefined || num === null) return 'N/A';
    return new Intl.NumberFormat('es-CO').format(num);
}

function resizeColumn(columnIndex, width) {
    const columnCells = document.querySelectorAll(`#results tr > *:nth-child(${columnIndex + 1})`);
    columnCells.forEach(cell => {
        cell.style.width = `${width}px`;
        cell.style.minWidth = `${width}px`;
        cell.style.maxWidth = `${width}px`;
    });
    
    // Update frozen columns positions if this column is frozen
    if (frozenColumns.includes(columnIndex)) {
        updateFrozenColumns();
    }
}

// Add this to your setup code
const resizeObserver = new ResizeObserver(() => {
    updateFrozenColumns();
});

// Observe the controls row
const controlsRow = document.querySelector('.column-controls');
if (controlsRow) {
    resizeObserver.observe(controlsRow);
}

function quickFilter(columnName) {
    // Then proceed with the existing sorting functionality
    if (currentSortColumn === columnName) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        currentSortColumn = columnName;
        sortDirection = 'asc';
    }
    
    document.querySelectorAll('#results th').forEach(th => {
        th.classList.remove('sorted-asc', 'sorted-desc');
    });
    
    const columnMap = {
        'Usuario': 4,
        'Nombre': 0,
        'Compañía': 2,
        'Cargo': 3,
        'Año Declaración': 1,
        'Activos': 5,
        'Pasivos': 6,
        'Patrimonio': 7,
        'Apalancamiento': 8,
        'Endeudamiento': 9,
        'Cant_Deudas': 10,
        'BancoSaldo': 11,
        'Cant_Bancos': 12,
        'Bienes': 13,
        'Cant_Bienes': 14,
        'Inversiones': 15,
        'Cant_Inversiones': 16,
        'Ingresos': 17,
        'Cant_Ingresos': 18,
        'Activos Var. Abs.': 19,
        'Pasivos Var. Abs.': 20,
        'Patrimonio Var. Abs.': 21,
        'Apalancamiento Var. Abs.': 22,
        'Endeudamiento Var. Abs.': 23,
        'BancoSaldo Var. Abs.': 24,
        'Bienes Var. Abs.': 25,
        'Inversiones Var. Abs.': 26,
        'Ingresos Var. Abs.': 27,
        'Activos Var. Rel.': 28,
        'Pasivos Var. Rel.': 29,
        'Patrimonio Var. Rel.': 30,
        'Apalancamiento Var. Rel.': 31,
        'Endeudamiento Var. Rel.': 32,
        'BancoSaldo Var. Rel.': 33,
        'Bienes Var. Rel.': 34,
        'Inversiones Var. Rel.': 35,
        'Ingresos Var. Rel.': 36
    };
    
    const columnIndex = columnMap[columnName];
    if (columnIndex !== undefined) {
        const header = document.querySelector(`#results th:nth-child(${columnIndex + 1})`);
        if (header) {
            header.classList.add(`sorted-${sortDirection}`);
            
            const icon = header.querySelector('.sort-icon');
            if (icon) {
                icon.textContent = sortDirection === 'asc' ? '↑' : '↓';
            }
        }
    }
    
    highlightColumn(columnName);
    currentFilterColumn = columnName;
    lastSelectedColumn = columnName;
    
    sortTable(columnName, sortDirection);
}

function showColumnStats(columnName) {
    // Collect all values for this column
    const values = allData.map(item => item[columnName]);
    
    // Calculate basic statistics
    const numericValues = values
        .map(v => typeof v === 'string' ? parseFloat(v.replace(/[^\d.-]/g, '')) : parseFloat(v))
        .filter(v => !isNaN(v));
        
    const isNumeric = numericValues.length > 0;
    
    let stats = {
        count: values.length,
        uniqueCount: new Set(values.filter(v => v !== undefined && v !== null)).size,
        min: null,
        max: null,
        avg: null,
        commonValues: [],
        allUniqueValues: []
    };
    
    if (isNumeric) {
        stats.min = Math.min(...numericValues);
        stats.max = Math.max(...numericValues);
        stats.avg = numericValues.reduce((a, b) => a + b, 0) / numericValues.length;
    }
    
    // Find most common values (top 5)
    const valueCounts = {};
    values.forEach(v => {
        if (v !== undefined && v !== null) {
            const val = typeof v === 'string' ? v.trim() : v;
            valueCounts[val] = (valueCounts[val] || 0) + 1;
        }
    });
    
    stats.commonValues = Object.entries(valueCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
        .map(([value, count]) => ({ value, count }));
    
    // Get all unique values (sorted)
    stats.allUniqueValues = Object.keys(valueCounts)
        .sort((a, b) => {
            if (isNumeric) {
                return parseFloat(a) - parseFloat(b);
            }
            return a.localeCompare(b);
        });
    
    // Format numbers for display
    const formatNumber = (num) => {
        if (num === null || num === undefined) return 'N/A';
        if (typeof num === 'string') return num;
        if (Math.abs(num) >= 1000000) {
            return '$' + (num / 1000000).toFixed(2) + 'M';
        }
        return new Intl.NumberFormat('es-CO').format(num);
    };
    
    // Create modal HTML with tabs
    const modalHTML = `
        <div id="columnStatsModal" class="modal-overlay">
            <div class="modal-content" style="max-width: 800px;">
                <div class="modal-header">
                    <h2>📊 Estadísticas de Columna: ${columnName}</h2>
                    <button onclick="closeModal()" class="close-button">×</button>
                </div>
                
                <div class="modal-tabs">
                    <button class="tab-btn active" onclick="switchTab('stats', this)">Estadísticas</button>
                    <button class="tab-btn" onclick="switchTab('values', this)">Todos los Valores (${stats.allUniqueValues.length})</button>
                </div>
                
                <div class="modal-body">
                    <div id="stats-tab" class="tab-content active">
                        <div class="stats-grid">
                            <div class="stat-item">
                                <strong>Total de valores:</strong>
                                <span>${stats.count}</span>
                            </div>
                            <div class="stat-item">
                                <strong>Valores únicos:</strong>
                                <span>${stats.uniqueCount}</span>
                            </div>
                            ${isNumeric ? `
                            <div class="stat-item">
                                <strong>Promedio:</strong>
                                <span>${formatNumber(stats.avg)}</span>
                            </div>
                            <div class="stat-item">
                                <strong>Mínimo:</strong>
                                <span>${formatNumber(stats.min)}</span>
                            </div>
                            <div class="stat-item">
                                <strong>Máximo:</strong>
                                <span>${formatNumber(stats.max)}</span>
                            </div>
                            ` : ''}
                        </div>
                        
                        <div class="common-values-section">
                            <h3>Valores más comunes</h3>
                            <div class="common-values-grid">
                                ${stats.commonValues.map(item => `
                                    <div class="common-value-item">
                                        <span class="value">${formatNumber(item.value)}</span>
                                        <span class="count">${item.count} (${Math.round((item.count / stats.count) * 100)}%)</span>
                                        <button onclick="applyCommonValueFilter('${columnName}', '${item.value.replace(/'/g, "\\'")}')" 
                                                class="apply-filter-btn">
                                            Filtrar
                                        </button>
                                    </div>
                                `).join('')}
                            </div>
                        </div>
                        
                        <div class="quick-filter-actions">
                            <button onclick="applyMinMaxFilter('${columnName}', 'min')" class="action-btn">
                                Filtrar por mínimo
                            </button>
                            <button onclick="applyMinMaxFilter('${columnName}', 'max')" class="action-btn">
                                Filtrar por máximo
                            </button>
                            <button onclick="applyAvgFilter('${columnName}')" class="action-btn" ${!isNumeric ? 'disabled' : ''}>
                                Filtrar por promedio
                            </button>
                        </div>
                    </div>
                    
                    <div id="values-tab" class="tab-content">
                        <div class="values-search">
                            <input type="text" id="values-search-input" placeholder="Buscar valores..." 
                                   oninput="searchValues('${columnName}')">
                            <div class="values-count">
                                Mostrando <span id="values-showing">0</span> de ${stats.allUniqueValues.length} valores
                            </div>
                        </div>
                        <div class="values-container" id="values-container">
                            <!-- Values will be loaded here with pagination -->
                        </div>
                        <div class="values-pagination">
                            <button id="values-prev" onclick="navigateValuesPage(-1)" disabled>Anterior</button>
                            <span id="values-page-info">Página 1</span>
                            <button id="values-next" onclick="navigateValuesPage(1)">Siguiente</button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    // Add to DOM
    document.body.insertAdjacentHTML('beforeend', modalHTML);
    
    // Initialize values tab
    currentValuesPage = 1;
    currentValuesSearch = '';
    currentValuesColumn = columnName;
    renderValuesPage();
}

function applyCommonValueFilter(columnName, value) {
    try {
        // Determine if the value is numeric
        const isNumeric = !isNaN(parseFloat(value)) && isFinite(value);
        
        // Set the operator based on value type
        const operator = isNumeric ? '=' : 'contains';
        
        // Add filter for this value
        document.getElementById('column').value = columnName;
        document.getElementById('operator').value = operator;
        document.getElementById('value1').value = value;
        
        addFilter();
        closeModal();
    } catch (error) {
        console.error('Error applying common value filter:', error);
        alert('Error al aplicar el filtro. Por favor intente nuevamente.');
    }
}

function applyMinMaxFilter(columnName, type) {
    const columnValues = allData.map(item => parseFloat(item[columnName])).filter(v => !isNaN(v));
    if (columnValues.length === 0) return;
    
    const value = type === 'min' ? Math.min(...columnValues) : Math.max(...columnValues);
    
    document.getElementById('column').value = columnName;
    document.getElementById('operator').value = '=';
    document.getElementById('value1').value = value;
    addFilter();
    closeModal();
}

function applyAvgFilter(columnName) {
    const columnValues = allData.map(item => parseFloat(item[columnName])).filter(v => !isNaN(v));
    if (columnValues.length === 0) return;
    
    const avg = columnValues.reduce((a, b) => a + b, 0) / columnValues.length;
    
    document.getElementById('column').value = columnName;
    document.getElementById('operator').value = '>=';
    document.getElementById('value1').value = avg.toFixed(2);
    addFilter();
    closeModal();
}

function closeModal() {
    const modal = document.getElementById('columnStatsModal') || document.getElementById('detailModal');
    if (modal) modal.remove();
}

// Export details to Excel
function exportDetailsToExcel() {
    const modal = document.getElementById('detailModal');
    if (!modal) return;
    
    // Get all the data from the modal
    const data = [];
    const sections = modal.querySelectorAll('.detail-section');
    
    sections.forEach(section => {
        const sectionTitle = section.querySelector('h3').textContent;
        const items = section.querySelectorAll('.detail-item, .variation-item');
        
        items.forEach(item => {
            const label = item.querySelector('strong')?.textContent.replace(':', '') || '';
            const value = item.querySelector('span')?.textContent || item.textContent.replace(label, '').replace(':', '').trim();
            
            if (label && value) {
                data.push({
                    'Sección': sectionTitle,
                    'Campo': label,
                    'Valor': value
                });
            }
        });
    });
    
    if (data.length === 0) {
        alert('No hay datos para exportar');
        return;
    }
    
    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Detalles");
    
    // Get the user name from the modal title
    const modalTitle = modal.querySelector('.modal-header h2').textContent;
    const fileName = modalTitle.replace('Detalles Completo - ', '').replace(/[\/\\?%*:|"<>]/g, '-') + '.xlsx';
    
    // Export to file
    XLSX.writeFile(workbook, fileName);
}

document.addEventListener('DOMContentLoaded', function() {
    // --- First Listener's Code ---
    const tabs = document.querySelectorAll('.tab');
    const tabContents = document.querySelectorAll('.tab-content');

    tabs.forEach(tab => {
        tab.addEventListener('click', function() {
            // Deactivate all tabs and content
            tabs.forEach(t => t.classList.remove('active'));
            tabContents.forEach(c => c.style.display = 'none');

            // Activate the clicked tab and its content
            this.classList.add('active');
            const tabId = this.dataset.tab;
            document.getElementById(tabId).style.display = 'block';
        });
    });


    // --- Second Listener's Code ---
    // Set initial header positions
    const controlsRow = document.querySelector('.column-controls');
    const secondHeaderRow = document.querySelector('#results thead tr:not(.column-controls)');

    if (controlsRow && secondHeaderRow) {
        const controlsHeight = controlsRow.offsetHeight;
        secondHeaderRow.style.top = `${controlsHeight}px`;
    }

    // Add resize observer
    const resizeObserver = new ResizeObserver(() => {
        updateFrozenColumns(); // Assuming this function is defined elsewhere
    });

    if (controlsRow) {
        resizeObserver.observe(controlsRow);
    }
});

// Export to Excel
function exportToExcel() {
    if (filteredData.length === 0) {
        alert('No hay datos para exportar');
        return;
    }
    
    // Use the original data with trend icons
    const exportData = filteredData.map(item => ({...item}));
    
    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Datos Filtrados");
    
    // Export to file
    XLSX.writeFile(workbook, 'datos_filtrados.xlsx');
}

function exportFrozenColumnsToExcel() {
    if (filteredData.length === 0) {
        alert('No hay datos para exportar');
        return;
    }
    
    if (frozenColumns.length === 0) {
        alert('No hay columnas congeladas para exportar');
        return;
    }
    
    // Create a mapping of column indices to column names
    const columnMap = {
        0: 'Nombre',
        1: 'Año Declaración',
        2: 'Compañía',
        3: 'Cargo',
        4: 'Usuario',
        5: 'Activos',
        6: 'Pasivos',
        7: 'Patrimonio',
        8: 'Apalancamiento',
        9: 'Endeudamiento',
        10: 'Cant_Deudas',
        11: 'BancoSaldo',
        12: 'Cant_Bancos',
        13: 'Bienes',
        14: 'Cant_Bienes',
        15: 'Inversiones',
        16: 'Cant_Inversiones',
        17: 'Ingresos',
        18: 'Cant_Ingresos',
        19: 'Activos Var. Abs.',
        20: 'Pasivos Var. Abs.',
        21: 'Patrimonio Var. Abs.',
        22: 'Apalancamiento Var. Abs.',
        23: 'Endeudamiento Var. Abs.',
        24: 'BancoSaldo Var. Abs.',
        25: 'Bienes Var. Abs.',
        26: 'Inversiones Var. Abs.',
        27: 'Ingresos Var. Abs.',
        28: 'Activos Var. Rel.',
        29: 'Pasivos Var. Rel.',
        30: 'Patrimonio Var. Rel.',
        31: 'Apalancamiento Var. Rel.',
        32: 'Endeudamiento Var. Rel.',
        33: 'BancoSaldo Var. Rel.',
        34: 'Bienes Var. Rel.',
        35: 'Inversiones Var. Rel.',
        36: 'Ingresos Var. Rel.'
    };
    
    // Get the column names for frozen columns
    const frozenColumnNames = frozenColumns.map(colIndex => columnMap[colIndex]);
    
    // Filter the data to include only frozen columns
    const exportData = filteredData.map(item => {
        const filteredItem = {};
        frozenColumnNames.forEach(colName => {
            filteredItem[colName] = item[colName];
        });
        return filteredItem;
    });
    
    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Columnas Congeladas");
    
    // Export to file
    XLSX.writeFile(workbook, 'columnas_congeladas.xlsx');
}
'@
}

function createStructure {
    Write-Host "🏗️ Creating Framework" -ForegroundColor $YELLOW

    # Create Python virtual environment
    python -m venv .venv
    .\.venv\scripts\activate

    # Upgrade pip and install required packages
    python -m pip install --upgrade pip
    python -m pip install pandas python-dotenv openpyxl plotly msoffcrypto-tool pdfplumber warnings pathlib levenshtein

    # Always create subdirectories
    Write-Host "🏗️ Creating directory structure" -ForegroundColor $YELLOW
    $directories = @(
        "src",
        "static",
        "models",
        "tables/cats",
        "tables/nets",
        "tables/trends"
    )
    foreach ($dir in $directories) {
        New-Item -Path $dir -ItemType Directory -Force
    }

}

function migratoDjango {
    Write-Host "🏗️ Creating Django Migration" -ForegroundColor $YELLOW

# Install Python packages
python -m pip install django whitenoise django-bootstrap-v5

# Start Django project
django-admin startproject arpa

# Change directory
cd arpa

# Start new app
python manage.py startapp persons


# Create persons views.py
@"
from django.http import HttpResponse
from django.template import loader
from .models import Member

def members(request):
  mymembers = Member.objects.all().values()
  template = loader.get_template('all_members.html')
  context = {
    'mymembers': mymembers,
  }
  return HttpResponse(template.render(context, request))
  
def details(request, id):
  mymember = Member.objects.get(id=id)
  template = loader.get_template('details.html')
  context = {
    'mymember': mymember,
  }
  return HttpResponse(template.render(context, request))
  
def main(request):
  template = loader.get_template('main.html')
  return HttpResponse(template.render())
"@ | Out-File -FilePath "persons/views.py" -Encoding utf8


# Create persons urls.py
@"
from django.urls import path
from . import views

urlpatterns = [
    path('', views.main, name='main'),
    path('members/', views.members, name='members'),
    path('members/details/<int:id>', views.details, name='details'),
]
"@ | Out-File -FilePath "persons/urls.py" -Encoding utf8


# Create arpa urls.py
@"
from django.contrib import admin
from django.urls import include, path

urlpatterns = [
    path('persons/', include('persons.urls')),
    path('admin/', admin.site.urls),
    path('', include('persons.urls')), 
]
"@ | Out-File -FilePath "arpa/urls.py" -Encoding utf8


# create directory
Write-Host "🏗️ Creating directory structure" -ForegroundColor $YELLOW
$directories = @(
    "persons/templates"
)
foreach ($dir in $directories) {
    New-Item -Path $dir -ItemType Directory -Force
}


# Create master.html
@"
<!DOCTYPE html>
<html>
<head>
  <title>{% block title %}{% endblock %}</title>
</head>
<body>

{% block content %}
{% endblock %}

</body>
</html>
"@ | Out-File -FilePath "persons/templates/master.html" -Encoding utf8


# Create main.html
@"
{% extends "master.html" %}

{% block title %}
  My Tennis Club
{% endblock %}


{% block content %}
  <h1>My Tennis Club</h1>

  <h3>Members</h3>
  
  <p>Check out all our <a href="members/">members</a></p>
  
{% endblock %}
"@ | Out-File -FilePath "persons/templates/main.html" -Encoding utf8


# Create persons.html
@"
<!DOCTYPE html>
<html>
<body>

<h1>Welcome to</h1>
<p>A R P A</p>

</body>
</html>
"@ | Out-File -FilePath "persons/templates/persons.html" -Encoding utf8


# Create all persons.html
@"
{% extends "master.html" %}

{% block title %}
  My Tennis Club - List of all members
{% endblock %}


{% block content %}

  <p><a href="/">HOME</a></p>

  <h1>Members</h1>
  
  <ul>
    {% for x in mymembers %}
      <li><a href="details/{{ x.id }}">{{ x.firstname }} {{ x.lastname }}</a></li>
    {% endfor %}
  </ul>
{% endblock %}
"@ | Out-File -FilePath "persons/templates/all_persons.html" -Encoding utf8


# Create details.html
@"
{% extends "master.html" %}

{% block title %}
  Details about {{ mymember.firstname }} {{ mymember.lastname }}
{% endblock %}


{% block content %}
  <h1>{{ mymember.firstname }} {{ mymember.lastname }}</h1>
  
  <p>Phone {{ mymember.phone }}</p>
  <p>Member since: {{ mymember.joined_date }}</p>
  
  <p>Back to <a href="/members">Members</a></p>
  
{% endblock %}
"@ | Out-File -FilePath "persons/templates/details.html" -Encoding utf8


# Create 404.html
@"
<!DOCTYPE html>
<html>
<title>Wrong address</title>
<body>

<h1>Ooops!</h1>

<h2>I cannot find the file you requested!</h2>

</body>
</html>
"@ | Out-File -FilePath "persons/templates/404.html" -Encoding utf8


# Change Settings, telling django a new app is created
# Path to the Django settings file
$settingsFilePath = ".\arpa\settings.py"  # Adjust this path if needed

# String to insert
$stringToInsert = "    'persons'"

# Line number to insert after (line 39 in your example, but be robust in case the file changes!) 
$lineNumberToInsertAfter = 39


# Read the file content
$fileContent = Get-Content -Path $settingsFilePath

# Check if the string already exists to avoid duplicates.  Crucial!
if ($fileContent -notcontains $stringToInsert) {

    # Insert the string after the specified line.
    $newFileContent = @() # Start with an empty array
    for ($i = 0; $i -lt $fileContent.Count; $i++) {
        $newFileContent += $fileContent[$i] # Add the current line

        # Insert after the desired line number (robust approach)
        if ($i -eq ($lineNumberToInsertAfter - 1)) { # Powershell arrays are zero-based!
            $newFileContent += $stringToInsert
        }
    }

    # Write the modified content back to the file.
    $newFileContent | Set-Content -Path $settingsFilePath

    Write-Host "String '$stringToInsert' inserted into '$settingsFilePath' after line $lineNumberToInsertAfter."
} else {
    Write-Host "String '$stringToInsert' already exists in '$settingsFilePath'. No changes made."
}

python manage.py migrate


# Create models.py
@"
from django.db import models

class Person(models.Model):
  firstname = models.CharField(max_length=255)
  lastname = models.CharField(max_length=255)
  phone = models.IntegerField(null=True)
  joined_date = models.DateField(null=True)
"@ | Out-File -FilePath "persons/models.py" -Encoding utf8

python manage.py makemigrations persons
python manage.py migrate
python manage.py sqlmigrate persons 0001

# Start the server
python manage.py runserver

}

function main {
    Write-Host "🏗️ Setting A R P A" -ForegroundColor $GREEN

    # Call functions to create structure and models
    createStructure
    createPeriod
    createPassKey
    createCats
    createNets
    createTrends
    createIDScript
    createIDtrends
    createINtrends
    createConflictScript
    createApp
    createIndex

    #generate periodoBR
    python models/period.py

    #generate Join between trends and conflicts
    python models/conflicts.py
    python models/ids.py
    python models/idTrends.py
    python models/inTrends.py

    Write-Host "🏗️ The framework is set" -ForegroundColor $YELLOW
    Write-Host "🏗️ Opening index.html in browser..." -ForegroundColor $GREEN
    
    migratoDjango
    cd ..
    #python app.py

}

main