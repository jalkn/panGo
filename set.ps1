function migratoDjango {
    param (
        [string]$ExcelFilePath = $null
    )

    $YELLOW = [ConsoleColor]::Yellow
    $GREEN = [ConsoleColor]::Green

    Write-Host "🚀 Creating ARPA" -ForegroundColor $YELLOW

    # Create Python virtual environment
    python -m venv .venv
    .\.venv\scripts\activate

    # Install required Python packages
    python.exe -m pip install --upgrade pip
    python -m pip install django whitenoise django-bootstrap-v5 openpyxl pandas xlrd>=2.0.1 pdfplumber fitz

    # Create Django project
    django-admin startproject arpa
    cd arpa

    # Create core app
    python manage.py startapp core

# Create models.py with cedula as primary key
Set-Content -Path "core/models.py" -Value @" 
from django.db import models

class Person(models.Model):
    ESTADO_CHOICES = [
        ('Activo', 'Activo'),
        ('Retirado', 'Retirado'),
    ]
    
    cedula = models.CharField(max_length=20, primary_key=True, verbose_name="Cedula")
    nombre_completo = models.CharField(max_length=255, verbose_name="Nombre Completo")
    cargo = models.CharField(max_length=255, verbose_name="Cargo")
    correo = models.EmailField(max_length=255, verbose_name="Correo")
    compania = models.CharField(max_length=255, verbose_name="Compania")
    estado = models.CharField(max_length=20, choices=ESTADO_CHOICES, default='Activo', verbose_name="Estado")
    revisar = models.BooleanField(default=False, verbose_name="Revisar")
    comments = models.TextField(blank=True, null=True, verbose_name="Comentarios")

    def __str__(self):
        return f"{self.cedula} - {self.nombre_completo}"

    class Meta:
        verbose_name = "Persona"
        verbose_name_plural = "Personas"

class FinancialReport(models.Model):
    person = models.ForeignKey(Person, on_delete=models.CASCADE, related_name='financial_reports')
    fkIdPeriodo = models.CharField(max_length=20, blank=True, null=True)
    ano_declaracion = models.CharField(max_length=20, blank=True, null=True)
    año_creacion = models.CharField(max_length=20, blank=True, null=True)
    activos = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    cant_bienes = models.IntegerField(blank=True, null=True)
    cant_bancos = models.IntegerField(blank=True, null=True)
    cant_cuentas = models.IntegerField(blank=True, null=True)
    cant_inversiones = models.IntegerField(blank=True, null=True)
    pasivos = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    cant_deudas = models.IntegerField(blank=True, null=True)
    patrimonio = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    apalancamiento = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    endeudamiento = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    aum_pat_subito = models.CharField(max_length=50, blank=True, null=True)
    activos_var_abs = models.CharField(max_length=50, blank=True, null=True)
    activos_var_rel = models.CharField(max_length=50, blank=True, null=True)
    pasivos_var_abs = models.CharField(max_length=50, blank=True, null=True)
    pasivos_var_rel = models.CharField(max_length=50, blank=True, null=True)
    patrimonio_var_abs = models.CharField(max_length=50, blank=True, null=True)
    patrimonio_var_rel = models.CharField(max_length=50, blank=True, null=True)
    apalancamiento_var_abs = models.CharField(max_length=50, blank=True, null=True)
    apalancamiento_var_rel = models.CharField(max_length=50, blank=True, null=True)
    endeudamiento_var_abs = models.CharField(max_length=50, blank=True, null=True)
    endeudamiento_var_rel = models.CharField(max_length=50, blank=True, null=True)
    banco_saldo = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    bienes = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    inversiones = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    banco_saldo_var_abs = models.CharField(max_length=50, blank=True, null=True)
    banco_saldo_var_rel = models.CharField(max_length=50, blank=True, null=True)
    bienes_var_abs = models.CharField(max_length=50, blank=True, null=True)
    bienes_var_rel = models.CharField(max_length=50, blank=True, null=True)
    inversiones_var_abs = models.CharField(max_length=50, blank=True, null=True)
    inversiones_var_rel = models.CharField(max_length=50, blank=True, null=True)
    ingresos = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    cant_ingresos = models.IntegerField(blank=True, null=True)
    ingresos_var_abs = models.CharField(max_length=50, blank=True, null=True)
    ingresos_var_rel = models.CharField(max_length=50, blank=True, null=True)
    capital = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    last_updated = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Reporte Financiero"
        verbose_name_plural = "Reportes Financieros"
        ordering = ['-ano_declaracion']

    def __str__(self):
        return f"Reporte de {self.person.nombre_completo} ({self.ano_declaracion})"
    
class Conflict(models.Model):
    person = models.ForeignKey(Person, on_delete=models.CASCADE, related_name='conflicts')
    fecha_inicio = models.DateField(verbose_name="Fecha de Inicio", null=True, blank=True)
    q1 = models.BooleanField(verbose_name="Accionista de algún proveedor del grupo", default=False)
    q2 = models.BooleanField(verbose_name="Familiar accionista, proveedor, empleado", default=False)
    q3 = models.BooleanField(verbose_name="Accionista de alguna compania del grupo", default=False)
    q4 = models.BooleanField(verbose_name="Actividades extralaborales", default=False)
    q5 = models.BooleanField(verbose_name="Negocios o bienes con empleados del grupo", default=False)
    q6 = models.BooleanField(verbose_name="Participación en juntas o consejos directivos", default=False)
    q7 = models.BooleanField(verbose_name="Potencial conflicto diferente a los anteriores", default=False)
    q8 = models.BooleanField(verbose_name="Consciente del código de conducta empresarial", default=False)
    q9 = models.BooleanField(verbose_name="Veracidad de la información consignada", default=False)
    q10 = models.BooleanField(verbose_name="Familiar de funcionario público", default=False)
    q11 = models.BooleanField(verbose_name="Relacion con el sector o funcionario público", default=False)
    last_updated = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Conflicto"
        verbose_name_plural = "Conflictos"

    def __str__(self):
        return f"Conflictos de {self.person.nombre_completo}"

class Card(models.Model):
    CARD_TYPE_CHOICES = [
        ('MC', 'Mastercard'),
        ('VI', 'Visa'),
    ]
    
    person = models.ForeignKey(Person, on_delete=models.CASCADE, related_name='cards')
    card_type = models.CharField(max_length=2, choices=CARD_TYPE_CHOICES)
    card_number = models.CharField(max_length=20)
    transaction_date = models.DateField()
    description = models.TextField()
    original_value = models.DecimalField(max_digits=15, decimal_places=2)
    exchange_rate = models.DecimalField(max_digits=10, decimal_places=4, null=True, blank=True)
    charges = models.DecimalField(max_digits=15, decimal_places=2, null=True, blank=True)
    balance = models.DecimalField(max_digits=15, decimal_places=2, null=True, blank=True)
    installments = models.CharField(max_length=20, null=True, blank=True)
    source_file = models.CharField(max_length=255)
    page_number = models.IntegerField()
    last_updated = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Tarjeta"
        verbose_name_plural = "Tarjetas"
        ordering = ['-transaction_date']

    def __str__(self):
        return f"{self.get_card_type_display()} - {self.card_number} - {self.transaction_date}"
"@

# Create views.py with import functionality
Set-Content -Path "core/views.py" -Value @"
from django.http import HttpResponse, HttpResponseRedirect
from django.template import loader
from django.shortcuts import render
from .models import Person, FinancialReport, Conflict, Card
import pandas as pd
from django.contrib import messages
from django.core.paginator import Paginator
from django.db.models import Q
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import os
import glob
from django.contrib.auth.decorators import login_required

@login_required
def main(request):
    persons = Person.objects.all()
    persons = _apply_person_filters_and_sorting(persons, request.GET)
    
    if 'export' in request.GET:
        model_fields = ['cedula', 'nombre_completo', 'cargo', 'correo', 'compania', 'estado', 'revisar', 'comments']
        return export_to_excel(persons, model_fields, 'persons_export')

    """
    Main view showing the list of persons with filtering and pagination
    """
    persons = Person.objects.all()
    persons = _apply_person_filters_and_sorting(persons, request.GET)
    
    # Get dropdown values
    dropdown_values = _get_dropdown_values()
    
    # Pagination
    paginator = Paginator(persons, 1000)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    context = {
        'page_obj': page_obj,
        'persons': page_obj.object_list,
        'persons_count': persons.count(),
        'current_order': request.GET.get('order_by', 'nombre_completo').replace('-', ''),
        'current_direction': request.GET.get('sort_direction', 'asc'),
        'all_params': {k: v for k, v in request.GET.items() if k not in ['order_by', 'sort_direction']},
        **dropdown_values
    }
    return render(request, 'persons.html', context)

def export_to_excel(queryset, model_fields, filename):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = f'attachment; filename="{filename}.xlsx"'
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Export"
    
    # Write headers
    for col_num, field in enumerate(model_fields, 1):
        col_letter = get_column_letter(col_num)
        ws[f'{col_letter}1'] = field
    
    # Write data
    for row_num, obj in enumerate(queryset, 2):
        for col_num, field in enumerate(model_fields, 1):
            col_letter = get_column_letter(col_num)
            value = getattr(obj, field, '')
            ws[f'{col_letter}{row_num}'] = value
    
    wb.save(response)
    return response


def import_period_excel(request):
    """View for importing period data from Excel files"""
    if request.method == 'POST' and request.FILES.get('period_excel_file'):
        excel_file = request.FILES['period_excel_file']
        try:
            # Save the uploaded file to the desired location
            temp_path = "core/src/periodoBR.xlsx"
            with open(temp_path, 'wb+') as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)
            
            messages.success(request, 'Archivo de periodos importado exitosamente!')
        except Exception as e:
            messages.error(request, f'Error procesando archivo de periodos: {str(e)}')
        
        return HttpResponseRedirect('/persons/import/')
    
    return HttpResponseRedirect('/persons/import/')


def _apply_person_filters_and_sorting(queryset, request_params):
    """
    Helper function to apply filters and sorting to a queryset based on request parameters.
    """
    search_query = request_params.get('q', '')
    status_filter = request_params.get('status', '')
    nombre_filter = request_params.get('nombre', '')
    cargo_filter = request_params.get('cargo', '')
    compania_filter = request_params.get('compania', '')
    correo_filter = request_params.get('correo', '')
    order_by = request_params.get('order_by', 'nombre_completo')
    sort_direction = request_params.get('sort_direction', 'asc')
    
    # Apply filters
    if search_query:
        queryset = queryset.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(compania__icontains=search_query) |
            Q(cargo__icontains=search_query) |
            Q(correo__icontains=search_query))
    
    if status_filter:
        queryset = queryset.filter(estado=status_filter)
    if nombre_filter:
        queryset = queryset.filter(nombre_completo__icontains=nombre_filter)
    if cargo_filter:
        queryset = queryset.filter(cargo__icontains=cargo_filter)
    if compania_filter:
        queryset = queryset.filter(compania__icontains=compania_filter)
    if correo_filter:
        queryset = queryset.filter(correo__icontains=correo_filter)
    
    # Apply sorting
    if order_by in ['cedula', 'nombre_completo', 'cargo', 'correo', 'compania', 'estado']:
        if sort_direction == 'desc':
            order_by = f'-{order_by}'
        queryset = queryset.order_by(order_by)
        
    return queryset


def _apply_finance_filters_and_sorting(queryset, request_params):
    """
    Helper function to apply finance-specific filters and sorting to a queryset.
    Handles column-based filtering with operators.
    """
    search_query = request_params.get('q', '')
    column_filter = request_params.get('column', '')
    operator = request_params.get('operator', '')
    value = request_params.get('value', '')
    order_by = request_params.get('order_by', 'nombre_completo')
    sort_direction = request_params.get('sort_direction', 'asc')
    
    # Apply general search filter
    if search_query:
        queryset = queryset.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(compania__icontains=search_query))
    
    # Apply column-based filtering if all required parameters are present
    if column_filter and operator and value:
        try:
            # Handle different operators
            if operator == '>':
                queryset = queryset.filter(**{f'financial_reports__{column_filter}__gt': float(value)})
            elif operator == '<':
                queryset = queryset.filter(**{f'financial_reports__{column_filter}__lt': float(value)})
            elif operator == '=':
                queryset = queryset.filter(**{f'financial_reports__{column_filter}': float(value)})
            elif operator == '>=':
                queryset = queryset.filter(**{f'financial_reports__{column_filter}__gte': float(value)})
            elif operator == '<=':
                queryset = queryset.filter(**{f'financial_reports__{column_filter}__lte': float(value)})
            elif operator == 'between':
                # For between operator, value should be two numbers separated by comma
                values = [v.strip() for v in value.split(',')]
                if len(values) == 2:
                    queryset = queryset.filter(
                        **{f'financial_reports__{column_filter}__gte': float(values[0])}
                    ).filter(
                        **{f'financial_reports__{column_filter}__lte': float(values[1])}
                    )
            elif operator == 'contains':
                queryset = queryset.filter(**{f'financial_reports__{column_filter}__icontains': value})
        except (ValueError, TypeError):
            # If conversion fails, skip this filter
            pass
    
    # Apply sorting - handle both person fields and financial report fields
    if order_by:
        if order_by.startswith('financial_reports__'):
            # Sorting by financial report field
            if sort_direction == 'desc':
                order_by = f'-{order_by}'
            queryset = queryset.order_by(order_by)
        elif order_by in ['cedula', 'nombre_completo', 'compania', 'revisar']:
            # Sorting by person field
            if sort_direction == 'desc':
                order_by = f'-{order_by}'
            queryset = queryset.order_by(order_by)
    
    return queryset


def _apply_conflict_filters_and_sorting(queryset, request_params):
    """
    Helper function to apply conflict-specific filters and sorting to a queryset.
    """
    search_query = request_params.get('q', '')
    column_filter = request_params.get('column', '')
    answer_filter = request_params.get('answer', '')
    value_filter = request_params.get('value', '')
    compania_filter = request_params.get('compania', '')
    order_by = request_params.get('order_by', 'nombre_completo')
    sort_direction = request_params.get('sort_direction', 'asc')
    
    # Apply general search filter
    if search_query:
        queryset = queryset.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(compania__icontains=search_query))
    
    # Apply company filter
    if compania_filter:
        queryset = queryset.filter(compania__icontains=compania_filter)
    
    # Apply column-based filtering if column is selected
    if column_filter:
        # Handle boolean fields (q1-q11)
        if column_filter in [f'q{i}' for i in range(1, 12)]:
            if answer_filter == 'yes':
                queryset = queryset.filter(**{f'conflicts__{column_filter}': True})
            elif answer_filter == 'no':
                queryset = queryset.filter(**{f'conflicts__{column_filter}': False})
        
        # For other fields (if we add them later)
        elif value_filter:
            queryset = queryset.filter(**{f'conflicts__{column_filter}__icontains': value_filter})
    
    # Apply sorting
    if order_by:
        if order_by.startswith('conflicts__'):
            # Sorting by conflict field
            if sort_direction == 'desc':
                order_by = f'-{order_by}'
            queryset = queryset.order_by(order_by)
        elif order_by in ['cedula', 'nombre_completo', 'compania', 'revisar']:
            # Sorting by person field
            if sort_direction == 'desc':
                order_by = f'-{order_by}'
            queryset = queryset.order_by(order_by)
    
    return queryset


def _get_dropdown_values():
    """
    Helper function to get distinct values for dropdown filters
    """
    return {
        'cargos': Person.objects.values_list('cargo', flat=True).distinct().order_by('cargo'),
        'companias': Person.objects.values_list('compania', flat=True).distinct().order_by('compania'),
        'estados': Person.objects.values_list('estado', flat=True).distinct().order_by('estado'),
    }

def process_financial_data():
    """Process financial data from inTrends.xlsx and update FinancialReport model"""
    try:
        in_trends_path = "core/src/inTrends.xlsx"
        
        if not os.path.exists(in_trends_path):
            print("inTrends.xlsx file not found")
            return False
            
        df = pd.read_excel(in_trends_path)
        
        # Ensure Cedula is string type for comparison
        df['Cedula'] = df['Cedula'].astype(str)
        
        def safe_convert(value):
            """Convert a value to decimal, handling special cases"""
            try:
                # Handle numpy/pandas special values
                if pd.isna(value) or pd.isnull(value):
                    return None
                if str(value).lower() in ['-inf', 'inf', 'infinity', '-infinity']:
                    return None
                if str(value).lower() == 'nan':
                    return None
                # Convert to float first to handle scientific notation, then to Decimal
                return float(value)
            except (ValueError, TypeError):
                return None
        
        for _, row in df.iterrows():
            try:
                # Find the person by cedula
                person = Person.objects.filter(cedula=str(row['Cedula'])).first()
                if not person:
                    continue
                    
                # Create or update financial report with proper value handling
                FinancialReport.objects.update_or_create(
                    person=person,
                    fkIdPeriodo=row.get('fkIdPeriodo'),
                    defaults={
                        'ano_declaracion': row.get('Año Declaración'),
                        'año_creacion': row.get('Año Creación'),
                        'activos': safe_convert(row.get('Activos')),
                        'cant_bienes': safe_convert(row.get('Cant_Bienes')),
                        'cant_bancos': safe_convert(row.get('Cant_Bancos')),
                        'cant_cuentas': safe_convert(row.get('Cant_Cuentas')),
                        'cant_inversiones': safe_convert(row.get('Cant_Inversiones')),
                        'pasivos': safe_convert(row.get('Pasivos')),
                        'cant_deudas': safe_convert(row.get('Cant_Deudas')),
                        'patrimonio': safe_convert(row.get('Patrimonio')),
                        'apalancamiento': safe_convert(row.get('Apalancamiento')),
                        'endeudamiento': safe_convert(row.get('Endeudamiento')),
                        'aum_pat_subito': row.get('Aum. Pat. Subito'),
                        'activos_var_abs': row.get('Activos Var. Abs.'),
                        'activos_var_rel': row.get('Activos Var. Rel.'),
                        'pasivos_var_abs': row.get('Pasivos Var. Abs.'),
                        'pasivos_var_rel': row.get('Pasivos Var. Rel.'),
                        'patrimonio_var_abs': row.get('Patrimonio Var. Abs.'),
                        'patrimonio_var_rel': row.get('Patrimonio Var. Rel.'),
                        'apalancamiento_var_abs': row.get('Apalancamiento Var. Abs.'),
                        'apalancamiento_var_rel': row.get('Apalancamiento Var. Rel.'),
                        'endeudamiento_var_abs': row.get('Endeudamiento Var. Abs.'),
                        'endeudamiento_var_rel': row.get('Endeudamiento Var. Rel.'),
                        'banco_saldo': safe_convert(row.get('BancoSaldo')),
                        'bienes': safe_convert(row.get('Bienes')),
                        'inversiones': safe_convert(row.get('Inversiones')),
                        'banco_saldo_var_abs': row.get('BancoSaldo Var. Abs.'),
                        'banco_saldo_var_rel': row.get('BancoSaldo Var. Rel.'),
                        'bienes_var_abs': row.get('Bienes Var. Abs.'),
                        'bienes_var_rel': row.get('Bienes Var. Rel.'),
                        'inversiones_var_abs': row.get('Inversiones Var. Abs.'),
                        'inversiones_var_rel': row.get('Inversiones Var. Rel.'),
                        'ingresos': safe_convert(row.get('Ingresos')),
                        'cant_ingresos': safe_convert(row.get('Cant_Ingresos')),
                        'ingresos_var_abs': row.get('Ingresos Var. Abs.'),
                        'ingresos_var_rel': row.get('Ingresos Var. Rel.'),
                        'capital': safe_convert(row.get('Capital')),
                    }
                )
            except Exception as e:
                print(f"Error processing row for cedula {row['Cedula']}: {str(e)}")
                continue
                
        return True
        
    except Exception as e:
        print(f"Error processing financial data: {str(e)}")
        return False


def details(request, cedula):
    """View showing details for a single person"""
    myperson = Person.objects.get(cedula=cedula)
    financial_reports = FinancialReport.objects.filter(person=myperson).order_by('-ano_declaracion')
    conflicts = Conflict.objects.filter(person=myperson).order_by('-fecha_inicio')
    
    # Process financial data if reports don't exist
    if not financial_reports.exists():
        process_financial_data()
        financial_reports = FinancialReport.objects.filter(person=myperson).order_by('-ano_declaracion')
    
    # Process conflict data if conflicts don't exist
    if not conflicts.exists():
        process_conflict_data()
        conflicts = Conflict.objects.filter(person=myperson).order_by('-fecha_inicio')
    
    return render(request, 'details.html', {
        'myperson': myperson,
        'financial_reports': financial_reports,
        'conflicts': conflicts
    })


def get_analysis_results():
    """Helper function to get analysis results from generated files"""
    import os
    from datetime import datetime
    from pathlib import Path
    
    analysis_files = [
        {'filename': 'Personas.xlsx', 'path': 'core/src/Personas.xlsx'},
        {'filename': 'periodoBR.xlsx', 'path': 'core/src/periodoBR.xlsx'},
        {'filename': 'conflicts.xlsx', 'path': 'core/src/conflicts.xlsx'},
        {'filename': 'banks.xlsx', 'path': 'core/src/banks.xlsx'},
        {'filename': 'debts.xlsx', 'path': 'core/src/debts.xlsx'},
        {'filename': 'goods.xlsx', 'path': 'core/src/goods.xlsx'},
        {'filename': 'incomes.xlsx', 'path': 'core/src/incomes.xlsx'},
        {'filename': 'investments.xlsx', 'path': 'core/src/investments.xlsx'},
        {'filename': 'bankNets.xlsx', 'path': 'core/src/bankNets.xlsx'},
        {'filename': 'debtNets.xlsx', 'path': 'core/src/debtNets.xlsx'},
        {'filename': 'goodNets.xlsx', 'path': 'core/src/goodNets.xlsx'},
        {'filename': 'incomeNets.xlsx', 'path': 'core/src/incomeNets.xlsx'},
        {'filename': 'investNets.xlsx', 'path': 'core/src/investNets.xlsx'},
        {'filename': 'trends.xlsx', 'path': 'core/src/trends.xlsx'},
        {'filename': 'idTrends.xlsx', 'path': 'core/src/idTrends.xlsx'},
        {'filename': 'inTrends.xlsx', 'path': 'core/src/inTrends.xlsx'},  
    ]
    
    results = []
    
    for file_info in analysis_files:
        file_path = Path(file_info['path'])
        result = {
            'filename': file_info['filename'],
            'status': 'pendiente',
            'last_updated': None,
            'records': None
        }
        
        if file_path.exists():
            result['status'] = 'success'
            result['last_updated'] = datetime.fromtimestamp(file_path.stat().st_mtime)
            
            try:
                # Try to count records in Excel files
                if str(file_path).endswith('.xlsx'):
                    import pandas as pd
                    df = pd.read_excel(file_path)
                    result['records'] = len(df)
            except Exception as e:
                result['status'] = 'error'
                result['error'] = str(e)
        
        results.append(result)
    
    return results


def import_persons(request):
    """
    View for importing only personas data from Excel files (saves to Personas.xlsx)
    """
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        try:
            # Read the Excel file to validate columns
            df = pd.read_excel(excel_file)
            
            # Check for required columns
            required_columns = ['NOMBRE COMPLETO', 'Correo', 'Cedula', 'Estado', 'Compania', 'CARGO', 'Activo']
            missing_cols = [col for col in required_columns if col not in df.columns]
            
            if missing_cols:
                messages.error(request, f'El archivo Excel no tiene las columnas requeridas: {", ".join(missing_cols)}', extra_tags='import_persons')
                return HttpResponseRedirect('/persons/import/')
            
            # Save the uploaded file to core/src/Personas.xlsx
            personas_path = "core/src/Personas.xlsx"
            os.makedirs(os.path.dirname(personas_path), exist_ok=True)
            
            # Save the file with only the required columns
            df[required_columns].to_excel(personas_path, index=False)
            
            messages.success(request, 'Archivo de personas importado exitosamente!', extra_tags='import_persons')
        except Exception as e:
            messages.error(request, f'Error procesando archivo: {str(e)}', extra_tags='import_persons')
        
        return HttpResponseRedirect('/persons/import/')
    
    # For GET requests, show the form with analysis results
    analysis_results = get_analysis_results()
    return render(request, 'import.html', {
        'analysis_results': analysis_results
    })

def process_persons_data(request):
    """
    Process data from Personas.xlsx and update Person model
    """
    try:
        # Path to the Personas file
        personas_path = "core/src/Personas.xlsx"
        
        if not os.path.exists(personas_path):
            messages.error(request, 'El archivo Personas.xlsx no existe. Por favor importe los datos primero.')
            return HttpResponseRedirect('/persons/import/')
            
        # Read the Personas file
        df = pd.read_excel(personas_path)
        
        # Convert 'nan' strings to actual NaN values
        df.replace('nan', pd.NA, inplace=True)
        
        # Column mapping from Personas.xlsx to Person model
        column_mapping = {
            'NOMBRE COMPLETO': 'nombre_completo',
            'Correo': 'correo',
            'Cedula': 'cedula',
            'Estado': 'estado',
            'Compania': 'compania',
            'CARGO': 'cargo',
            'Activo': 'revisar'  # Map 'Activo' column to 'revisar' field
        }
        
        # Ensure all required columns exist
        missing_cols = [col for col in column_mapping.keys() if col not in df.columns]
        if missing_cols:
            messages.error(request, f'El archivo Personas.xlsx no tiene las columnas requeridas: {", ".join(missing_cols)}')
            return HttpResponseRedirect('/persons/import/')
        
        # Rename columns to match model
        df.rename(columns=column_mapping, inplace=True)
        
        # Filter only columns we need for Person model
        person_df = df[list(column_mapping.values())].copy()
        
        # Fill empty values and clean data
        person_df.fillna('', inplace=True)
        
        # Convert 'Activo' to boolean for 'revisar' field
        person_df['revisar'] = person_df['revisar'].apply(
            lambda x: bool(x) if pd.notna(x) else False
        )
        
        # Standardize estado values
        person_df['estado'] = person_df['estado'].apply(
            lambda x: 'Activo' if str(x).strip().lower() in ['activo', '1', 'true', 'si'] else 'Retirado'
        )
        
        # Update or create Person records in bulk for better performance
        persons_created = 0
        persons_updated = 0
        
        for _, row in person_df.iterrows():
            obj, created = Person.objects.update_or_create(
                cedula=row['cedula'],
                defaults={
                    'nombre_completo': row['nombre_completo'],
                    'cargo': row['cargo'],
                    'correo': row['correo'],
                    'compania': row['compania'],
                    'estado': row['estado'],
                    'revisar': row['revisar'],
                    'comments': ''  # Initialize comments as empty
                }
            )
            if created:
                persons_created += 1
            else:
                persons_updated += 1
        
        # Prepare success message
        msg_parts = []
        if persons_created or persons_updated:
            msg_parts.append(f"{persons_created} nuevas personas creadas, {persons_updated} actualizadas")
        
        if msg_parts:
            messages.success(request, 'Datos de personas procesados exitosamente! ' + ', '.join(msg_parts) + '.')

    except Exception as e:
        messages.error(request, f'Error procesando datos de personas: {str(e)}')
        # Log the full error for

def process_conflict_data():
    """
    Process data from inTrends.xlsx and update Conflict model
    """
    try:
        in_trends_path = "core/src/inTrends.xlsx"
        
        if not os.path.exists(in_trends_path):
            print("inTrends.xlsx file not found")
            return False
            
        df = pd.read_excel(in_trends_path)
        
        # Ensure Cedula is string type for comparison
        df['Cedula'] = df['Cedula'].astype(str)
        
        for _, row in df.iterrows():
            try:
                # Find the person by cedula
                person = Person.objects.filter(cedula=str(row['Cedula'])).first()
                if not person:
                    continue
                    
                # Create or update conflict record
                Conflict.objects.update_or_create(
                    person=person,
                    defaults={
                        'q1': row.get('q1', False),
                        'q2': row.get('q2', False),
                        'q3': row.get('q3', False),
                        'q4': row.get('q4', False),
                        'q5': row.get('q5', False),
                        'q6': row.get('q6', False),
                        'q7': row.get('q7', False),
                        'q8': row.get('q8', False),
                        'q9': row.get('q9', False),
                        'q10': row.get('q10', False),
                        'q11': row.get('q11', False),
                        'comments': row.get('comments', ''),
                        'fecha_inicio': pd.to_datetime(row.get('fecha_inicio'), errors='coerce'),
                        'fecha_fin': pd.to_datetime(row.get('fecha_fin'), errors='coerce'),
                    }
                )
            except Exception as e:
                print(f"Error processing conflict for cedula {row['Cedula']}: {str(e)}")
                continue
                
        return True
        
    except Exception as e:
        print(f"Error processing conflict data: {str(e)}")
        return False

def import_protected_excel(request):
    """
    View for importing data from password-protected Excel files
    and running the full analysis pipeline
    """
    if request.method == 'POST' and request.FILES.get('protected_excel_file'):
        excel_file = request.FILES['protected_excel_file']
        password = request.POST.get('excel_password', '')
        
        try:
            # Create necessary directories if they don't exist
            os.makedirs('core/src', exist_ok=True)
            
            # Save the uploaded file temporarily
            temp_path = "core/src/dataHistoricaPBI.xlsx"
            with open(temp_path, 'wb+') as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)
            
            output_excel = "core/src/data.xlsx"
            
            # Modified password removal function that works with the view
            def remove_excel_password_browser(input_file, output_file, password):
                try:
                    import msoffcrypto
                    with open(input_file, "rb") as file:
                        office_file = msoffcrypto.OfficeFile(file)
                        if office_file.is_encrypted():
                            if not password:
                                return False, "No password provided"
                            try:
                                office_file.load_key(password=password)
                            except Exception as e:
                                return False, "Incorrect password"
                        else:
                            office_file.load_key(password=None)
                        
                        with open(output_file, "wb") as decrypted:
                            office_file.decrypt(decrypted)
                    return True, "File processed successfully"
                except Exception as e:
                    return False, str(e)
            
            # Remove password and create data.xlsx
            success, message = remove_excel_password_browser(temp_path, output_excel, password)
            
            if not success:
                messages.error(request, f'Error al procesar el archivo protegido: {message}')
                return HttpResponseRedirect('/persons/import/')
            
            # Run the full analysis pipeline
            try:
                # Import the analysis modules
                from core.cats import run_all_analyses as run_cats_analysis
                from core.nets import run_all_analyses as run_nets_analysis
                from core.trends import main as run_trends_analysis
                from core.idTrends import merge_trends_data
                from core.inTrends import merge_conflicts_data  # Add this import
                
                # Ensure periodoBR.xlsx exists
                periodo_file = "core/src/periodoBR.xlsx"
                if not os.path.exists(periodo_file):
                    messages.error(request, 'El archivo de periodos (periodoBR.xlsx) no existe. Por favor cargue primero el archivo de periodos.')
                    return HttpResponseRedirect('/persons/import/')
                
                # Run CATS analysis (generates banks.xlsx, debts.xlsx, etc.)
                run_cats_analysis()
                
                # Run NETS analysis (generates bankNets.xlsx, debtNets.xlsx, etc.)
                run_nets_analysis()
                
                # Run TRENDS analysis (generates trends.xlsx, overTrends.xlsx, data.json)
                run_trends_analysis()
                
                # After running all analyses, now process the idTrends data
                idtrends_file = "core/src/idTrends.xlsx"
                if os.path.exists(idtrends_file):
                    df_idtrends = pd.read_excel(idtrends_file)
                    
                    for _, row in df_idtrends.iterrows():
                        if pd.notna(row['Cedula']):
                            Person.objects.update_or_create(
                                cedula=row['Cedula'],
                                defaults={
                                    'nombre_completo': row.get('Nombre', ''),
                                    'cargo': row.get('Cargo', ''),
                                    'correo': row.get('Correo', ''),
                                    'compania': row.get('Compania', ''),
                                    'estado': row.get('Estado', 'Activo'),
                                    'revisar': False,
                                    'comments': '',
                                }
                            )
                    
                    messages.success(request, 'Datos de personas actualizados desde idTrends.xlsx')
                    
                # Run idTrends analysis - merge trends.xlsx with Personas.xlsx
                personas_file = "core/src/Personas.xlsx"
                trends_file = "core/src/trends.xlsx"
                idtrends_output = "core/src/idTrends.xlsx"
                
                if os.path.exists(personas_file) and os.path.exists(trends_file):
                    merge_trends_data(
                        personas_file=personas_file,
                        trends_file=trends_file,
                        output_file=idtrends_output
                    )
                    messages.success(request, 'Análisis completado: Datos de tendencias fusionados con personas.')
                else:
                    messages.warning(request, 'Análisis completado pero no se pudo fusionar tendencias con personas (archivos faltantes).')
                
                # NEW: Run inTrends analysis - merge idTrends.xlsx with conflicts.xlsx
                conflicts_file = "core/src/conflicts.xlsx"
                intrends_output = "core/src/inTrends.xlsx"
                
                if os.path.exists(idtrends_output) and os.path.exists(conflicts_file):
                    merge_conflicts_data(
                        idtrends_file=idtrends_output,
                        conflicts_file=conflicts_file,
                        output_file=intrends_output
                    )
                    messages.success(request, 'Análisis completado: Datos de conflictos fusionados con tendencias.')
                else:
                    messages.warning(request, 'Análisis completado pero no se pudo fusionar conflictos con tendencias (archivos faltantes).')
                
                messages.success(request, 'Proceso completo! Archivo desencriptado y análisis generados exitosamente.')
            
            except Exception as e:
                messages.error(request, f'Error durante el análisis de datos: {str(e)}')
            
            # Clean up temporary files
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                if os.path.exists(output_excel):  
                    os.remove(output_excel)
            except Exception as e:
                messages.warning(request, f'Advertencia: No se pudieron eliminar algunos archivos temporales: {str(e)}')
            
        except Exception as e:
            messages.error(request, f'Error importing protected file: {str(e)}')
        
        return HttpResponseRedirect('/persons/import/')
    
    return HttpResponseRedirect('/persons/import/')

def import_conflict_excel(request):
    """View for importing conflict data from Excel files"""
    if request.method == 'POST' and request.FILES.get('conflict_excel_file'):
        excel_file = request.FILES['conflict_excel_file']
        try:
            # Save the uploaded file temporarily
            temp_path = "core/src/conflictos.xlsx"
            with open(temp_path, 'wb+') as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)
            
            # Process the file using conflicts.py functionality
            from core.conflicts import extract_specific_columns
            import os
            
            custom_headers = [
                "ID", "Cedula", "Nombre", "1er Nombre", "1er Apellido", 
                "2do Apellido", "Compañía", "Cargo", "Email", "Fecha de Inicio", 
                "Q1", "Q2", "Q3", "Q4", "Q5",
                "Q6", "Q7", "Q8", "Q9", "Q10", "Q11"
            ]
            
            extract_specific_columns(
                input_file=temp_path,
                output_file="core/src/conflicts.xlsx",
                custom_headers=custom_headers
            )
            
            messages.success(request, 'Archivo de conflictos importado exitosamente!')
            
            # Clean up temporary file
            if os.path.exists(temp_path):
                os.remove(temp_path)
            
        except Exception as e:
            messages.error(request, f'Error procesando archivo de conflictos: {str(e)}')
        
        return HttpResponseRedirect('/persons/import/')
    
    return HttpResponseRedirect('/persons/import/')

@login_required
def finance_view(request):
    persons = Person.objects.all().prefetch_related('financial_reports')
    persons = _apply_finance_filters_and_sorting(persons, request.GET)
    
    if 'export' in request.GET:
        model_fields = ['cedula', 'nombre', 'compania', 'ano_declaracion', 'aum_pat_subito', 'activos_var_rel',
                       'pasivos_var_rel', 'patrimonio_var_rel',
                       'apalancamiento_var_rel', 'endeudamiento_var_rel',
                       'banco_saldo_var_rel', 'bienes_var_rel',
                       'inversiones_var_rel', 'comments']
        return export_to_excel(persons, model_fields, 'finance_export')
    
    # Get dropdown values
    dropdown_values = _get_dropdown_values()
    
    # Pagination
    paginator = Paginator(persons, 1000)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    context = {
        'page_obj': page_obj,
        'persons': page_obj.object_list,
        'persons_count': persons.count(),
        'current_order': request.GET.get('order_by', 'nombre_completo').replace('-', ''),
        'current_direction': request.GET.get('sort_direction', 'asc'),
        'all_params': {k: v for k, v in request.GET.items() if k not in ['order_by', 'sort_direction']},
        **dropdown_values
    }
    return render(request, 'finances.html', context)

@login_required
def conflicts_view(request):
    persons = Person.objects.all().prefetch_related('conflicts')
    persons = _apply_conflict_filters_and_sorting(persons, request.GET)
    
    if 'export' in request.GET:
        model_fields = ['cedula', 'nombre', 'compania', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'comments']
        return export_to_excel(persons, model_fields, 'conflicts_export')
    
    # Get dropdown values
    dropdown_values = _get_dropdown_values()
    
    # Pagination
    paginator = Paginator(persons, 1000)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    context = {
        'page_obj': page_obj,
        'persons': page_obj.object_list,
        'persons_count': persons.count(),
        'current_order': request.GET.get('order_by', 'nombre_completo').replace('-', ''),
        'current_direction': request.GET.get('sort_direction', 'asc'),
        'all_params': {k: v for k, v in request.GET.items() if k not in ['order_by', 'sort_direction']},
        **dropdown_values
    }
    return render(request, 'conflicts.html', context)

@login_required
def alerts_view(request):
    """View showing all records marked for review or with comments"""
    persons = Person.objects.filter(Q(revisar=True) | ~Q(comments='')).distinct()
    
    # Apply the same filters as the main view
    persons = _apply_person_filters_and_sorting(persons, request.GET)
    
    if 'export' in request.GET:
        model_fields = ['cedula', 'nombre_completo', 'cargo', 'correo', 'compania', 'estado', 'revisar', 'comments']
        return export_to_excel(persons, model_fields, 'alerts_export')

    # Get dropdown values
    dropdown_values = _get_dropdown_values()
    
    # Pagination
    paginator = Paginator(persons, 1000)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    context = {
        'page_obj': page_obj,
        'persons': page_obj.object_list,
        'persons_count': persons.count(),
        'current_order': request.GET.get('order_by', 'nombre_completo').replace('-', ''),
        'current_direction': request.GET.get('sort_direction', 'asc'),
        'all_params': {k: v for k, v in request.GET.items() if k not in ['order_by', 'sort_direction']},
        **dropdown_values
    }
    return render(request, 'alerts.html', context)

def import_visa_pdfs(request):
    if request.method == 'POST' and request.FILES.getlist('visa_pdf_files'):
        pdf_files = request.FILES.getlist('visa_pdf_files')
        password = request.POST.get('visa_pdf_password', '')
        
        try:
            # Create visa directory if it doesn't exist
            visa_dir = "core/src/visa"
            os.makedirs(visa_dir, exist_ok=True)
            
            # Save password to password.txt if provided
            password_file = os.path.join(visa_dir, "password.txt")
            if password:
                with open(password_file, 'w') as f:
                    f.write(password)
            
            # Save all uploaded PDF files
            pdf_paths = []
            for pdf_file in pdf_files:
                file_path = os.path.join(visa_dir, pdf_file.name)
                with open(file_path, 'wb+') as destination:
                    for chunk in pdf_file.chunks():
                        destination.write(chunk)
                pdf_paths.append(file_path)
                print(f"Saved PDF: {file_path}")  # DEBUG
            
            # Run visa.py script
            from core.visa import main as process_visa
            process_visa()
            
            # Process the generated Excel file and save to Card model
            visa_excel_files = glob.glob(os.path.join(visa_dir, "VISA_*.xlsx"))
            if not visa_excel_files:
                messages.error(request, 'No se encontraron archivos Excel generados por el procesador VISA')
                return HttpResponseRedirect('/persons/import/')
                
            latest_file = max(visa_excel_files, key=os.path.getctime)
            print(f"Processing Visa Excel file: {latest_file}")  # DEBUG
            
            try:
                df = pd.read_excel(latest_file)
                print(f"Found {len(df)} transactions in Excel file")  # DEBUG
            except Exception as e:
                messages.error(request, f'Error leyendo archivo Excel: {str(e)}')
                return HttpResponseRedirect('/persons/import/')
            
            created_count = 0
            no_person_count = 0
            error_count = 0
            person_matches = {}

            for index, row in df.iterrows():
                try:
                    cardholder_name = str(row['Tarjetahabiente']).strip()
                    print(f"\nProcessing transaction {index + 1}/{len(df)} for: {cardholder_name}")  # DEBUG
                    
                    # Skip if we've already tried to match this person
                    if cardholder_name in person_matches:
                        person = person_matches[cardholder_name]
                        if not person:
                            no_person_count += 1
                            print(f"⚠ [Cached] No person found for: {cardholder_name}")
                            continue
                    
                    # Try multiple matching strategies
                    person = None
                    
                    # 1. Exact name match (case insensitive)
                    person = Person.objects.filter(
                        nombre_completo__iexact=cardholder_name
                    ).first()
                    
                    # 2. Contains match with normalized whitespace
                    if not person:
                        person = Person.objects.filter(
                            nombre_completo__iregex=r'\y{}\y'.format(
                                re.escape(cardholder_name)
                            )
                        ).first()
                    
                    # 3. Split name and match parts
                    if not person:
                        name_parts = [p.strip() for p in re.split(r'\s+', cardholder_name) if p.strip()]
                        if len(name_parts) >= 2:
                            # Try matching first name + last name
                            first_name = name_parts[0]
                            last_name = name_parts[-1]
                            
                            # Match first name and any part of last name
                            person = Person.objects.filter(
                                nombre_completo__iregex=rf'\y{re.escape(first_name)}\y.*\y{re.escape(last_name)}\y'
                            ).first()
                            
                            # Try alternative combinations if still not found
                            if not person and len(name_parts) > 2:
                                person = Person.objects.filter(
                                    nombre_completo__iregex=rf'\y{re.escape(first_name)}\y.*\y{re.escape(" ".join(name_parts[-2:]))}\y'
                                ).first()
                    
                    # Cache the result for this cardholder name
                    person_matches[cardholder_name] = person
                    
                    if not person:
                        no_person_count += 1
                        print(f"⚠ No person found for: {cardholder_name}")
                        continue
                    
                    print(f"✓ Found matching person: {person.nombre_completo} (ID: {person.cedula})")
                    
                    # Process numeric values safely
                    def safe_float(value):
                        try:
                            if pd.isna(value):
                                return None
                            if isinstance(value, str):
                                value = value.replace('.', '').replace(',', '.')
                            return float(value)
                        except (ValueError, TypeError):
                            return None
                    
                    # Create card record
                    Card.objects.create(
                        person=person,
                        card_type='VI' if str(row['Número de Tarjeta']).startswith('4') else 'MC',
                        card_number=str(row['Número de Tarjeta']),
                        transaction_date=pd.to_datetime(row['Fecha de Transacción']).date(),
                        description=str(row['Descripción'])[:255],  # Ensure it fits in CharField
                        original_value=safe_float(row['Valor Original']),
                        exchange_rate=safe_float(row['Tasa Pactada']),
                        charges=safe_float(row['Cargos y Abonos']),
                        balance=safe_float(row['Saldo a Diferir']),
                        installments=str(row['Cuotas'])[:20] if pd.notna(row['Cuotas']) else None,
                        source_file=str(row['Archivo'])[:255],
                        page_number=int(row['Página']) if pd.notna(row['Página']) else 0
                    )
                    created_count += 1
                    print("✓ Transaction saved successfully")
                    
                except Exception as e:
                    error_count += 1
                    print(f"⚠ Error processing transaction {index + 1}: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    continue
            
            # Prepare result message
            msg = f"Procesamiento completado: {created_count} transacciones importadas"
            if no_person_count:
                msg += f", {no_person_count} sin persona correspondiente"
            if error_count:
                msg += f", {error_count} con errores"
            
            messages.success(request, msg)
            
            # Print summary to console
            print("\n" + "="*50)
            print("IMPORT SUMMARY")
            print(f"- Total transactions processed: {len(df)}")
            print(f"- Successfully imported: {created_count}")
            print(f"- No person found: {no_person_count}")
            print(f"- Errors: {error_count}")
            print("="*50)
            
            # Clean up files
            try:
                # Delete PDF files
                for pdf_path in pdf_paths:
                    if os.path.exists(pdf_path):
                        os.remove(pdf_path)
                
                # Delete password file if it exists
                if os.path.exists(password_file):
                    os.remove(password_file)
                    
                print("✓ Temporary files cleaned up")
            except Exception as clean_error:
                messages.warning(request, f'Archivos VISA procesados pero hubo un error limpiando: {str(clean_error)}')
                print(f"⚠ Error cleaning files: {str(clean_error)}")
            
        except Exception as e:
            messages.error(request, f'Error procesando archivos VISA: {str(e)}')
            print(f"⚠ Critical error in import_visa_pdfs: {str(e)}")
            import traceback
            traceback.print_exc()
        
        return HttpResponseRedirect('/persons/import/')
    
    return HttpResponseRedirect('/persons/import/')

@login_required
def cards_view(request):
    # Start with all cards and select related person
    cards = Card.objects.all().select_related('person')
    
    # Apply filters
    search_query = request.GET.get('q', '')
    card_type_filter = request.GET.get('card_type', '')
    date_from = request.GET.get('date_from', '')
    date_to = request.GET.get('date_to', '')
    order_by = request.GET.get('order_by', '-transaction_date')  # Default to newest first
    sort_direction = request.GET.get('sort_direction', 'desc')
    
    if search_query:
        cards = cards.filter(
            Q(person__nombre_completo__icontains=search_query) |
            Q(description__icontains=search_query) |
            Q(card_number__icontains=search_query)
        )
    
    if card_type_filter:
        cards = cards.filter(card_type=card_type_filter)
    
    if date_from:
        try:
            cards = cards.filter(transaction_date__gte=date_from)
        except:
            pass
    
    if date_to:
        try:
            cards = cards.filter(transaction_date__lte=date_to)
        except:
            pass
    
    # Apply sorting - handle descending order
    if sort_direction == 'desc' and not order_by.startswith('-'):
        order_by = f'-{order_by}'
    elif sort_direction == 'asc' and order_by.startswith('-'):
        order_by = order_by[1:]
    
    cards = cards.order_by(order_by)
    
    # Pagination
    paginator = Paginator(cards, 25)  # Reduced from 100 to 25 for better UX
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    context = {
        'page_obj': page_obj,
        'cards': page_obj.object_list,  # This should be used in template
        'cards_count': cards.count(),
        'current_order': order_by.replace('-', ''),
        'current_direction': sort_direction,
        'all_params': {k: v for k, v in request.GET.items() if k not in ['order_by', 'sort_direction', 'page']},
    }
    return render(request, 'cards.html', context)
"@

# Create urls.py for core app
Set-Content -Path "core/urls.py" -Value @"
from django.contrib.auth import views as auth_views
from django.urls import path
from . import views
from django.contrib.auth import get_user_model
from django.contrib import messages
from django.shortcuts import render, redirect

def register_superuser(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        email = request.POST.get('email')
        password1 = request.POST.get('password1')
        password2 = request.POST.get('password2')
        
        if password1 != password2:
            messages.error(request, "Passwords don't match")
            return redirect('register')
        
        User = get_user_model()
        if User.objects.filter(username=username).exists():
            messages.error(request, "Username already exists")
            return redirect('register')
        
        try:
            user = User.objects.create_superuser(
                username=username,
                email=email,
                password=password1
            )
            messages.success(request, f"Superuser {username} created successfully!")
            return redirect('login')
        except Exception as e:
            messages.error(request, f"Error creating superuser: {str(e)}")
            return redirect('register')
    
    return render(request, 'registration/register.html')

urlpatterns = [
    path('', views.main, name='main'),
    path('persons/details/<str:cedula>/', views.details, name='details'),
    path('persons/import/', views.import_persons, name='import_persons'),
    path('persons/process/', views.process_persons_data, name='process_persons'),
    path('persons/import-protected/', views.import_protected_excel, name='import_protected_excel'),
    path('persons/import-conflicts/', views.import_conflict_excel, name='import_conflict_excel'),
    path('persons/import-period/', views.import_period_excel, name='import_period_excel'),
    path('finance/', views.finance_view, name='finance_view'),
    path('conflicts/', views.conflicts_view, name='conflicts_view'),
    path('alerts/', views.alerts_view, name='alerts_view'),
    path('logout/', auth_views.LogoutView.as_view(), name='logout'),
    path('register/', register_superuser, name='register'),
    path('visa/import/', views.import_visa_pdfs, name='import_visa_pdfs'),
    path('cards/', views.cards_view, name='cards_view'),
]
"@

    # Create admin.py with enhanced configuration
Set-Content -Path "core/admin.py" -Value @" 
from django.contrib import admin
from .models import Person, FinancialReport, Conflict, Card

def make_active(modeladmin, request, queryset):
    queryset.update(estado='Activo')
make_active.short_description = "Mark selected as Active"

def make_retired(modeladmin, request, queryset):
    queryset.update(estado='Retirado')
make_retired.short_description = "Mark selected as Retired"

def mark_for_check(modeladmin, request, queryset):
    queryset.update(revisar=True)
mark_for_check.short_description = "Mark for check"

def unmark_for_check(modeladmin, request, queryset):
    queryset.update(revisar=False)
unmark_for_check.short_description = "Unmark for check"

class PersonAdmin(admin.ModelAdmin):
    list_display = ("cedula", "nombre_completo", "cargo", "correo", "compania", "estado", "revisar")
    search_fields = ("nombre_completo", "cedula", "comments")
    list_filter = ("estado", "compania", "revisar")
    list_per_page = 25
    ordering = ('nombre_completo',)
    actions = [make_active, make_retired, mark_for_check, unmark_for_check]
    
    fieldsets = (
        (None, {
            'fields': ('cedula', 'nombre_completo', 'cargo')
        }),
        ('Advanced options', {
            'classes': ('collapse',),
            'fields': ('correo', 'compania', 'estado', 'revisar', 'comments'),
        }),
    )
    
class FinancialReportAdmin(admin.ModelAdmin):
    list_display = ('person', 'ano_declaracion', 'patrimonio', 'activos', 'pasivos', 'last_updated')
    list_filter = ('ano_declaracion', 'person__compania', 'person__estado')
    search_fields = ('person__nombre_completo', 'person__cedula')
    list_per_page = 25
    raw_id_fields = ('person',)
    
    fieldsets = (
        (None, {
            'fields': ('person', 'ano_declaracion', 'año_creacion')
        }),
        ('Financial Data', {
            'fields': (
                ('activos', 'pasivos', 'patrimonio'),
                ('apalancamiento', 'endeudamiento'),
                ('banco_saldo', 'bienes', 'inversiones'),
                ('ingresos', 'cant_ingresos'),
                ('aum_pat_subito', 'capital')
            )
        }),
        ('Variations', {
            'classes': ('collapse',),
            'fields': (
                ('activos_var_abs', 'activos_var_rel'),
                ('pasivos_var_abs', 'pasivos_var_rel'),
                ('patrimonio_var_abs', 'patrimonio_var_rel'),
                ('apalancamiento_var_abs', 'apalancamiento_var_rel'),
                ('endeudamiento_var_abs', 'endeudamiento_var_rel'),
                ('banco_saldo_var_abs', 'banco_saldo_var_rel'),
                ('bienes_var_abs', 'bienes_var_rel'),
                ('inversiones_var_abs', 'inversiones_var_rel'),
                ('ingresos_var_abs', 'ingresos_var_rel')
            )
        })
    )

class ConflictAdmin(admin.ModelAdmin):
    list_display = ('person', 'fecha_inicio', 'q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11')
    list_filter = ('q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11')
    search_fields = ('person__nombre_completo', 'person__cedula')
    raw_id_fields = ('person',)
    list_per_page = 25

class CardAdmin(admin.ModelAdmin):
    list_display = ('person', 'get_card_type_display', 'transaction_date', 'description', 'original_value')
    list_filter = ('card_type', 'transaction_date')
    search_fields = ('person__nombre_completo', 'person__cedula', 'description')
    date_hierarchy = 'transaction_date'
    list_per_page = 50

admin.site.register(Person, PersonAdmin)
admin.site.register(FinancialReport, FinancialReportAdmin)
admin.site.register(Conflict, ConflictAdmin)
admin.site.register(Card, CardAdmin)
"@

# Update project urls.py with proper admin configuration
Set-Content -Path "arpa/urls.py" -Value @"
from django.contrib import admin
from django.urls import include, path
from django.contrib.auth import views as auth_views

# Customize default admin interface
admin.site.site_header = 'A R P A'
admin.site.site_title = 'ARPA Admin Portal'
admin.site.index_title = 'Bienvenido a A R P A'

urlpatterns = [
    path('admin/', admin.site.urls),
    path('persons/', include('core.urls')),
    path('accounts/', include('django.contrib.auth.urls')),  # Add this line
    path('', include('core.urls')), 
]
"@

    # Create templates directory structure
    $directories = @(
        "core/src",
        "core/static",
        "core/static/css",
        "core/static/js",
        "core/templates",
        "core/templates/admin",
        "core/templates/registration"
    )
    foreach ($dir in $directories) {
        New-Item -Path $dir -ItemType Directory -Force
    }

# Create cats.py
Set-Content -Path "core/cats.py" -Value @"
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
    file_path = 'core/src/data.xlsx'
    periodo_file_path = 'core/src/periodoBR.xlsx'
    
    analyze_banks(file_path, 'core/src/banks.xlsx', periodo_file_path)
    analyze_debts(file_path, 'core/src/debts.xlsx', periodo_file_path)
    analyze_goods(file_path, 'core/src/goods.xlsx', periodo_file_path)
    analyze_incomes(file_path, 'core/src/incomes.xlsx', periodo_file_path)
    analyze_investments(file_path, 'core/src/investments.xlsx', periodo_file_path)

if __name__ == "__main__":
    run_all_analyses()
"@

# Create nets.py
Set-Content -Path "core/nets.py" -Value @"
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
        'core/src/banks.xlsx',
        'core/src/bankNets.xlsx'
    )
    
    debt_summary = analyze_debts(
        'core/src/debts.xlsx',
        'core/src/debtNets.xlsx'
    )
    
    goods_summary = analyze_goods(
        'core/src/goods.xlsx',
        'core/src/goodNets.xlsx'
    )
    
    income_summary = analyze_incomes(
        'core/src/incomes.xlsx',
        'core/src/incomeNets.xlsx'
    )
    
    invest_summary = analyze_investments(
        'core/src/investments.xlsx',
        'core/src/investNets.xlsx'
    )
    
    # Combined analyses
    assets_summary = calculate_assets(
        'core/src/bankNets.xlsx',
        'core/src/goodNets.xlsx',
        'core/src/investNets.xlsx',
        'core/src/assetNets.xlsx'
    )
    
    net_worth_summary = calculate_net_worth(
        'core/src/debtNets.xlsx',
        'core/src/assetNets.xlsx',
        'core/src/worthNets.xlsx'
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

# Create trends.py
Set-Content -Path "core/trends.py" -Value @"
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

def calculate_sudden_wealth_increase(df):
    """Calculate sudden wealth increase rate (Aum. Pat. Subito) as decimal with 1 decimal place"""
    df = df.sort_values(['Usuario', 'Año Declaración'])
    
    # Calculate total wealth (Activo + Patrimonio)
    df['Capital'] = df['Activos'] + df['Patrimonio']
    
    # Calculate year-to-year change as decimal
    df['Aum. Pat. Subito'] = df.groupby('Usuario')['Capital'].pct_change(fill_method=None)
    
    # Format as decimal (1 place) with trend symbol
    df['Aum. Pat. Subito'] = df['Aum. Pat. Subito'].apply(
        lambda x: f"{x:.1f} {get_trend_symbol(f'{x*100:.1f}%')}" if not pd.isna(x) else "N/A ➡️"
    )
    
    return df

def save_results(df, excel_filename="tables/trends/trends.xlsx"):
    """Save results to Excel."""
    try:
        df.to_excel(excel_filename, index=False)
        print(f"Data saved to {excel_filename}")
    except Exception as e:
        print(f"Error saving file: {e}")

def main():
    """Main function to process all data and generate analysis files."""
    try:
        # Process worth data
        df_worth = pd.read_excel("core/src/worthNets.xlsx")
        df_worth = df_worth.rename(columns={
            'Total Activos': 'Activos',
            'Total Pasivos': 'Pasivos',
            'Total Patrimonio': 'Patrimonio'
        })
        
        df_worth = calculate_leverage(df_worth)
        df_worth = calculate_debt_level(df_worth)
        df_worth = calculate_sudden_wealth_increase(df_worth)
        
        for column in ['Activos', 'Pasivos', 'Patrimonio', 'Apalancamiento', 'Endeudamiento']:
            df_worth = calculate_variation(df_worth, column)
        
        df_worth = embed_trend_symbols(df_worth, ['Activos', 'Pasivos', 'Patrimonio', 'Apalancamiento', 'Endeudamiento'])
        
        # Process asset data
        df_assets = pd.read_excel("core/src/assetNets.xlsx")
        df_assets_processed = process_asset_data(df_assets)
        
        # Process income data
        df_income = pd.read_excel("core/src/incomeNets.xlsx")
        df_income_processed = process_income_data(df_income)
        
        # Merge all data
        df_combined = pd.merge(df_worth, df_assets_processed, on=['Usuario', 'Año Declaración'], how='left')
        df_combined = pd.merge(df_combined, df_income_processed, on=['Usuario', 'Año Declaración'], how='left')
        
        # Save basic trends
        save_results(df_combined, "core/src/trends.xlsx")
        
        # The 'df_yearly' calculation and saving to 'overTrends.xlsx' and 'data.json' have been removed.
        
    except FileNotFoundError as e:
        print(f"Error: Required file not found - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
"@

# Create core/idTrends.py
Set-Content -Path "core/idTrends.py" -Value @"
import pandas as pd
import os
from datetime import datetime

def merge_trends_data(personas_file, trends_file, output_file):
    """
    Merge personas data with trends data and save to output file.
    Returns True if successful, False otherwise.
    """
    try:
        # Load the Excel files
        personas = pd.read_excel(personas_file)
        trends = pd.read_excel(trends_file)

        # Rename columns in trends to match personas where appropriate
        trends = trends.rename(columns={
            'Nombre': 'NOMBRE COMPLETO',
            'Compañía': 'Compania',
            'Cargo': 'CARGO'
        })

        # Perform a full outer join to keep all data from both tables
        merged = pd.merge(
            personas,
            trends,
            on=['NOMBRE COMPLETO', 'CARGO', 'Compania'],
            how='outer',
            indicator=True
        )

        # For rows that only existed in trends, copy the Usuario to Cedula if Cedula is null
        merged['Cedula'] = merged['Cedula'].fillna(merged['Usuario'])

        # Select and order the columns as specified with renamed columns
        final_columns = [
            'Cedula', 'Nombre', 'Cargo', 'Correo', 'Compania', 'Estado',
            'fkIdPeriodo', 'Año Declaración', 'Año Creación', 'Activos', 'Cant_Bienes',
            'Cant_Bancos', 'Cant_Cuentas', 'Cant_Inversiones', 'Pasivos', 'Cant_Deudas',
            'Patrimonio', 'Apalancamiento', 'Endeudamiento', 'Capital', 'Aum. Pat. Subito',
            'Activos Var. Abs.', 'Activos Var. Rel.', 'Pasivos Var. Abs.', 'Pasivos Var. Rel.',
            'Patrimonio Var. Abs.', 'Patrimonio Var. Rel.', 'Apalancamiento Var. Abs.',
            'Apalancamiento Var. Rel.', 'Endeudamiento Var. Abs.', 'Endeudamiento Var. Rel.',
            'BancoSaldo', 'Bienes', 'Inversiones', 'BancoSaldo Var. Abs.', 'BancoSaldo Var. Rel.',
            'Bienes Var. Abs.', 'Bienes Var. Rel.', 'Inversiones Var. Abs.', 'Inversiones Var. Rel.',
            'Ingresos', 'Cant_Ingresos', 'Ingresos Var. Abs.', 'Ingresos Var. Rel.'
        ]

        # Rename the columns in the merged dataframe before selection
        merged = merged.rename(columns={
            'NOMBRE COMPLETO': 'Nombre',
            'CARGO': 'Cargo'
        })

        # Ensure we only keep columns that exist in our merged dataframe
        final_columns = [col for col in final_columns if col in merged.columns]

        final_df = merged[final_columns]

        # Fill null values with 'nan'
        final_df = final_df.fillna('nan')

        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_file), exist_ok=True)

        # Save to new Excel file
        final_df.to_excel(output_file, index=False)
        
        return True

    except Exception as e:
        print(f"Error merging trends data: {str(e)}")
        return False
"@

# Create core/inTrends.py -merge trends with conflictos
Set-Content -Path "core/inTrends.py" -Value @"
import pandas as pd
import os
from datetime import datetime

def merge_conflicts_data(idtrends_file, conflicts_file, output_file):
    """
    Merge idtrends data with conflicts data and save to output file.
    Returns True if successful, False otherwise.
    """
    try:
        # Load the Excel files
        idtrends = pd.read_excel(idtrends_file)
        conflicts = pd.read_excel(conflicts_file)

        # Standardize column names (handle multiple possible email column names)
        email_columns = ['Correo', 'Email', 'Correo Electrónico', 'E-mail', 'Correo_x']
        
        # Find which email column exists in conflicts data
        conflicts_email_col = next((col for col in email_columns if col in conflicts.columns), None)
        if not conflicts_email_col:
            raise ValueError("No valid email column found in conflicts file")
            
        # Rename columns in conflicts to match idtrends
        conflicts = conflicts.rename(columns={
            'Compañía': 'Compania',
            conflicts_email_col: 'Correo'  # Standardize to 'Correo'
        })

        # Ensure idtrends has the Correo column
        if 'Correo' not in idtrends.columns:
            # Try to find email column in idtrends if 'Correo' doesn't exist
            idtrends_email_col = next((col for col in email_columns if col in idtrends.columns), None)
            if idtrends_email_col:
                idtrends = idtrends.rename(columns={idtrends_email_col: 'Correo'})
            else:
                idtrends['Correo'] = ''  # Add empty column if missing

        # Perform a full outer join to keep all data from both tables
        merged = pd.merge(
            idtrends,
            conflicts,
            on=['Cedula', 'Nombre', 'Cargo', 'Compania', 'Correo'],
            how='outer',
            indicator=True
        )

        # Select and order the columns as specified (removed duplicate Correo)
        final_columns = [
            'Cedula', 'Estado', 'Nombre', 'Compania', 'Cargo', 'Correo',
            'fkIdPeriodo', 'Año Declaración', 'Año Creación', 'Activos', 'Cant_Bienes',
            'Cant_Bancos', 'Cant_Cuentas', 'Cant_Inversiones', 'Pasivos', 'Cant_Deudas',
            'Patrimonio', 'Apalancamiento', 'Endeudamiento', 'Aum. Pat. Subito',
            'Activos Var. Abs.', 'Activos Var. Rel.', 'Pasivos Var. Abs.', 'Pasivos Var. Rel.',
            'Patrimonio Var. Abs.', 'Patrimonio Var. Rel.', 'Apalancamiento Var. Abs.',
            'Apalancamiento Var. Rel.', 'Endeudamiento Var. Abs.', 'Endeudamiento Var. Rel.',
            'BancoSaldo', 'Bienes', 'Inversiones', 'BancoSaldo Var. Abs.', 'BancoSaldo Var. Rel.',
            'Bienes Var. Abs.', 'Bienes Var. Rel.', 'Inversiones Var. Abs.', 'Inversiones Var. Rel.',
            'Ingresos', 'Cant_Ingresos', 'Ingresos Var. Abs.', 'Ingresos Var. Rel.',
            'Capital', 'Fecha de Inicio', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11'
        ]

        # Ensure we only keep columns that exist in our merged dataframe
        final_columns = [col for col in final_columns if col in merged.columns]

        final_df = merged[final_columns]

        # Fill null values with empty string instead of 'nan'
        final_df = final_df.fillna('')

        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_file), exist_ok=True)

        # Save to new Excel file
        final_df.to_excel(output_file, index=False)
        
        return True

    except Exception as e:
        print(f"Error merging conflicts data: {str(e)}")
        return False
"@

# Create core/conflicts.py
Set-Content -Path "core/conflicts.py" -Value @"
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
        extra_cols = [12,14,16,18,20,22,24,25,26,28]
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
        
        else:
            print("Warning: Column J not found in selected columns")
    
    except Exception as e:
        print(f"Error: {str(e)}")

# Example usage with custom headers
custom_headers = [
    "ID", "Cedula", "Nombre", "1er Nombre", "1er Apellido", 
    "2do Apellido", "Compañía", "Cargo", "Email", "Fecha de Inicio", 
    "Q1", "Q2", "Q3", "Q4", "Q5",
    "Q6", "Q7", "Q8", "Q9", "Q10", "Q11"
]

extract_specific_columns(
    input_file="core/src/conflictos.xlsx",
    output_file="core/src/conflicts.xlsx",
    custom_headers=custom_headers
)
"@

# Create core/Mastercard.py
Set-Content -Path "core/mc.py" -Value @"
import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime
 
# Change these lines at the top of the file:
pdf_folder = "core/src/GA consolidado Visa"
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
os.makedirs("core/src", exist_ok=True)  # Changed to core/src
output_excel = os.path.join("core/src", f"extractos_resultado_VISA_{timestamp}.xlsx")
pdf_password = ""
password_file = os.path.join(pdf_folder, "password.txt")
if os.path.exists(password_file):
    with open(password_file, 'r') as f:
        pdf_password = f.read().strip()
 
# Columnas incluyendo "Página"
column_names = [
    "Archivo", "Tarjetahabiente", "Número de Tarjeta", "Número de Autorización",
    "Fecha de Transacción", "Descripción", "Valor Original",
    "Tasa Pactada", "Tasa EA Facturada", "Cargos y Abonos",
    "Saldo a Diferir", "Cuotas", "Página"
]
 
# Patrón para detectar transacciones
pattern_transaccion = re.compile(
    r"(\d{6})\s+(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,.]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+/\d+|0\.00)"
)
 
# Patrón para número de tarjeta
pattern_tarjeta = re.compile(r"TARJETA:\s+\*{12}(\d{4})")
 
data_rows = []
pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
 
for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_folder, pdf_file)
    print(f"📄 Procesando: {pdf_file}")
 
    try:
        with pdfplumber.open(pdf_path, password=pdf_password) as pdf:
            tarjetahabiente = ""
            tarjeta = ""
            tiene_transacciones = False
            last_page_number = 1  # predeterminada por si no entra al bucle
 
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if not text:
                    continue
 
                last_page_number = page_number  # actualizar número de página
                lines = text.split("\n")
 
                for idx, line in enumerate(lines):
                    line = line.strip()
 
                    # Buscar cambio de tarjeta y nombre
                    tarjeta_match = pattern_tarjeta.search(line)
                    if tarjeta_match:
                        if tarjetahabiente and tarjeta and not tiene_transacciones:
                            row = [pdf_file, tarjetahabiente, tarjeta, "Sin transacciones", "", "", "", "", "", "", "", "", last_page_number]
                            data_rows.append(row)
 
                        tarjeta = tarjeta_match.group(1)
                        tiene_transacciones = False
 
                        # Capturar nombre desde la línea anterior
                        if idx > 0:
                            posible_nombre = lines[idx - 1].strip()
                            posible_nombre = (
                                posible_nombre
                                .replace("SEÑOR (A):", "")
                                .replace("Señor (A):", "")
                                .replace("SEÑOR:", "")
                                .replace("Señor:", "")
                                .strip()
                                .title()
                            )
                            if len(posible_nombre.split()) >= 2:
                                tarjetahabiente = posible_nombre
                        continue
 
                    # Buscar transacción
                    match = pattern_transaccion.search(' '.join(line.split()))
                    if match and tarjetahabiente and tarjeta:
                        row_data = list(match.groups())
                        row_data.insert(0, tarjeta)
                        row_data.insert(0, tarjetahabiente)
                        row_data.insert(0, pdf_file)
 
                        # Ajustar valores numéricos
                        row_data[6] = row_data[6].replace(".", "#").replace(",", ".").replace("#", ",")
                        row_data[9] = row_data[9].replace(".", "#").replace(",", ".").replace("#", ",")
                        row_data[10] = row_data[10].replace(".", "#").replace(",", ".").replace("#", ",")
 
                        row_data.append(page_number)  # Añadir número de página
                        data_rows.append(row_data)
                        tiene_transacciones = True
 
            # Al final del PDF, si no se registraron transacciones
            if tarjetahabiente and tarjeta and not tiene_transacciones:
                row = [pdf_file, tarjetahabiente, tarjeta, "Sin transacciones", "", "", "", "", "", "", "", "", last_page_number]
                data_rows.append(row)
 
    except Exception as e:
        print(f"⚠️ Error al procesar '{pdf_file}': {e}")
 
# Exportar a Excel
if data_rows:
    df_final = pd.DataFrame(data_rows, columns=column_names)
    df_final.to_excel(output_excel, index=False)
    print(f"\n✅ Archivo generado: {output_excel}")
    print(df_final.head())
else:
    print("\n⚠️ No se extrajo ningún dato.")
"@

# Create core/Visa.py
Set-Content -Path "core/visa.py" -Value @"
import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime

# Configuration
PDF_FOLDER = "core/src/visa"
COLUMN_NAMES = [
    "Archivo", "Tarjetahabiente", "Número de Tarjeta", "Número de Autorización",
    "Fecha de Transacción", "Descripción", "Valor Original",
    "Tasa Pactada", "Tasa EA Facturada", "Cargos y Abonos",
    "Saldo a Diferir", "Cuotas", "Página"
]

# Patterns
TRANSACTION_PATTERN = re.compile(
    r"(\d{6})\s+(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,.]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+/\d+|0\.00)"
)
CARD_PATTERN = re.compile(r"TARJETA:\s+\*{12}(\d{4})")

def get_password():
    """Get password from password.txt if it exists"""
    password_file = os.path.join(PDF_FOLDER, "password.txt")
    if os.path.exists(password_file):
        with open(password_file, 'r') as f:
            return f.read().strip()
    return ""

def clean_value(value):
    """Clean and format numeric values"""
    return value.replace(".", "#").replace(",", ".").replace("#", ",")

def process_pdf(pdf_path, password):
    """Process a single PDF file and extract data"""
    data_rows = []
    try:
        with pdfplumber.open(pdf_path, password=password) as pdf:
            cardholder = ""
            card_number = ""
            has_transactions = False
            last_page = 1

            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if not text:
                    continue

                last_page = page_num
                lines = text.split("\n")

                for idx, line in enumerate(lines):
                    line = line.strip()

                    # Check for card number change
                    card_match = CARD_PATTERN.search(line)
                    if card_match:
                        if cardholder and card_number and not has_transactions:
                            row = [
                                os.path.basename(pdf_path), cardholder, card_number,
                                "Sin transacciones", "", "", "", "", "", "", "", "", last_page
                            ]
                            data_rows.append(row)

                        card_number = card_match.group(1)
                        has_transactions = False

                        # Get cardholder name from previous line
                        if idx > 0:
                            possible_name = lines[idx - 1].strip()
                            possible_name = (
                                possible_name
                                .replace("SEÑOR (A):", "")
                                .replace("Señor (A):", "")
                                .replace("SEÑOR:", "")
                                .replace("Señor:", "")
                                .strip()
                                .title()
                            )
                            if len(possible_name.split()) >= 2:
                                cardholder = possible_name
                        continue

                    # Check for transactions
                    match = TRANSACTION_PATTERN.search(' '.join(line.split()))
                    if match and cardholder and card_number:
                        row_data = list(match.groups())
                        row_data.insert(0, card_number)
                        row_data.insert(0, cardholder)
                        row_data.insert(0, os.path.basename(pdf_path))

                        # Clean numeric values
                        row_data[6] = clean_value(row_data[6])
                        row_data[9] = clean_value(row_data[9])
                        row_data[10] = clean_value(row_data[10])

                        row_data.append(page_num)
                        data_rows.append(row_data)
                        has_transactions = True

            # Add entry if no transactions were found
            if cardholder and card_number and not has_transactions:
                row = [
                    os.path.basename(pdf_path), cardholder, card_number,
                    "Sin transacciones", "", "", "", "", "", "", "", "", last_page
                ]
                data_rows.append(row)

    except Exception as e:
        print(f"⚠ Error processing '{os.path.basename(pdf_path)}': {e}")

    return data_rows

def cleanup_files():
    """Clean up temporary files"""
    try:
        # Delete all PDFs in the folder
        for filename in os.listdir(PDF_FOLDER):
            if filename.lower().endswith('.pdf'):
                os.remove(os.path.join(PDF_FOLDER, filename))
        
        # Delete password file if exists
        password_file = os.path.join(PDF_FOLDER, "password.txt")
        if os.path.exists(password_file):
            os.remove(password_file)
            
        print("✓ Temporary files cleaned up")
    except Exception as e:
        print(f"⚠ Warning: Could not clean all files: {e}")

def main():
    """Main processing function"""
    # Ensure directories exist
    os.makedirs(PDF_FOLDER, exist_ok=True)
    
    # Get password
    password = get_password()
    
    # Get PDF files
    pdf_files = [
        f for f in os.listdir(PDF_FOLDER) 
        if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(PDF_FOLDER, f))
    ]
    
    if not pdf_files:
        print("⚠ No PDF files found in the visa folder")
        return
    
    # Process all PDFs
    all_data = []
    for pdf_file in pdf_files:
        pdf_path = os.path.join(PDF_FOLDER, pdf_file)
        print(f"📄 Processing: {pdf_file}")
        all_data.extend(process_pdf(pdf_path, password))
    
    # Export to Excel
    if all_data:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(PDF_FOLDER, f"VISA_{timestamp}.xlsx")
        
        df = pd.DataFrame(all_data, columns=COLUMN_NAMES)
        df.to_excel(output_file, index=False)
        print(f"\n✓ Excel file generated: {output_file}")
        print(df.head())
        
        # Clean up files
        cleanup_files()
    else:
        print("\n⚠ No data extracted from PDFs")

if __name__ == "__main__":
    main()
"@

# Create custom admin base template
@"
{% extends "admin/base.html" %}

{% block title %}{{ title }} | {{ site_title|default:_('A R P A') }}{% endblock %}

{% block branding %}
<h1 id="site-name"><a href="{% url 'admin:index' %}">{{ site_header|default:_('A R P A') }}</a></h1>
{% endblock %}

{% block nav-global %}{% endblock %}
"@ | Out-File -FilePath "core/templates/admin/base_site.html" -Encoding utf8

# Create master template
@"
<!DOCTYPE html>
<html>
<head>
    <title>{% block title %}ARPA{% endblock %}</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    {% load static %}
    <link rel="stylesheet" href="{% static 'css/style.css' %}">
    <link rel="stylesheet" href="{% static 'css/freeze.css' %}">
</head>
<body>
    {% if user.is_authenticated %}
    <div class="topnav-container">
        <a href="/" style="text-decoration: none;">
            <div class="logoIN"></div>
        </a>
        <div class="navbar-title">{% block navbar_title %}ARPA{% endblock %}</div>
        <div class="navbar-buttons">
            {% block navbar_buttons %}
            <a href="/admin/" class="btn btn-outline-dark" title="Admin">
                <i class="fas fa-wrench"></i>
            </a>
            <a href="/persons/import/" class="btn btn-custom-primary" title="Importar">
                <i class="fas fa-database"></i>
            </a>
            <form method="post" action="{% url 'logout' %}" class="d-inline">
                {% csrf_token %}
                <button type="submit" class="btn btn-custom-primary" title="Cerrar sesión">
                    <i class="fas fa-sign-out-alt"></i>
                </button>
            </form>
            {% endblock %}
        </div>
    </div>
    {% endif %}
    
    <div class="main-container">
        {% if messages %}
            {% for message in messages %}
                <div class="alert alert-{{ message.tags }} alert-dismissible fade show">
                    {{ message }}
                    <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                </div>
            {% endfor %}
        {% endif %}
        
        {% block content %}
        {% endblock %}
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"@ | Out-File -FilePath "core/templates/master.html" -Encoding utf8

#statics css style
@" 
:root {
    --primary-color: #0b00a2;
    --primary-hover: #090086;
    --text-on-primary: white;
}

body {
    margin: 0;
    padding: 20px;
    background-color: #f8f9fa;
}

.topnav-container {
    display: flex;
    align-items: center;
    padding: 0 40px;
    margin-bottom: 20px;
    gap: 15px;
}

.logoIN {
    width: 40px;
    height: 40px;
    background-color: var(--primary-color);
    border-radius: 8px;
    position: relative;
    flex-shrink: 0;
}

.logoIN::before {
    content: "";
    position: absolute;
    width: 100%;
    height: 100%;
    border-radius: 50%;
    top: 30%;
    left: 70%;
    transform: translate(-50%, -50%);
    background-image: linear-gradient(to right, 
        #ffffff 2px, transparent 2px);
    background-size: 4px 100%;
}

.navbar-title {
    color: var(--primary-color);
    font-weight: bold;
    font-size: 1.25rem;
    margin-right: auto;
}

.navbar-buttons {
    display: flex;
    gap: 10px;
}

.btn-custom-primary {
    background-color: white;
    border-color: var(--primary-color);
    color: var(--primary-color);
    padding: 0.5rem 1rem;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-width: 40px;
}

.btn-custom-primary:hover,
.btn-custom-primary:focus {
    background-color: var(--primary-hover);
    border-color: var(--primary-hover);
    color: var(--text-on-primary);
}

.btn-custom-primary i,
.btn-outline-dark i {
    margin-right: 0;
    font-size: 1rem;
    line-height: 1;
    display: inline-block;
    vertical-align: middle;
}

.main-container {
    padding: 0 40px;
}

/* Search filter styles */
.search-filter {
    margin-bottom: 20px;
    max-width: 400px;
}

/* Table row hover effect */
.table-hover tbody tr:hover {
    background-color: rgba(11, 0, 162, 0.05);
}

.btn-my-green {
    background-color: white;
    border-color: rgb(0, 166, 0);
    color: rgb(0, 166, 0);
}

.btn-my-green:hover {
    background-color: darkgreen;
    border-color: darkgreen;
    color: white;
}

.btn-my-green:focus,
.btn-my-green.focus {
    box-shadow: 0 0 0 0.2rem rgba(0, 128, 0, 0.5);
}

.btn-my-green:active,
.btn-my-green.active {
    background-color: darkgreen !important;
    border-color: darkgreen !important;
}

.btn-my-green:disabled,
.btn-my-green.disabled {
    background-color: lightgreen;
    border-color: lightgreen;
    color: #6c757d;
}

/* Card styles */
.card {
    border-radius: 8px;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

/* Table styles */
.table {
    width: 100%;
    margin-bottom: 1rem;
    color: #212529;
}

.table th {
    vertical-align: bottom;
    border-bottom: 2px solid #dee2e6;
}

.table td {
    vertical-align: middle;
}

/* Alert styles */
.alert {
    position: relative;
    padding: 0.75rem 1.25rem;
    margin-bottom: 1rem;
    border: 1px solid transparent;
    border-radius: 0.25rem;
}

/* Badge styles */
.badge {
    display: inline-block;
    padding: 0.35em 0.65em;
    font-size: 0.75em;
    font-weight: 700;
    line-height: 1;
    color: #fff;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    border-radius: 0.25rem;
}

.bg-success {
    background-color:rgb(0, 166, 0) !important;
}

.bg-danger {
    background-color: #dc3545 !important;
}

"@ | Out-File -FilePath "core/static/css/style.css" -Encoding utf8

@"
.loading-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0,0,0,0.5);
    z-index: 9999;
    display: none;
    justify-content: center;
    align-items: center;
}

.loading-content {
    background-color: white;
    padding: 30px;
    border-radius: 8px;
    text-align: center;
    max-width: 500px;
    width: 90%;
}

.progress {
    height: 20px;
    margin: 20px 0;
}

/* Spinner styles for submit buttons */
.btn .spinner-border {
    margin-right: 8px;
}
"@ | Out-File -FilePath "core/static/css/loading.css" -Encoding utf8

@"
/* Table container styles */
.table-container {
    position: relative;
    overflow: auto;
    max-height: calc(100vh - 300px); /* Adjust this value as needed */
}

.table-fixed-header {
    position: sticky;
    top: 0;
    z-index: 10;
    background-color: white;
}

.table-fixed-header th {
    position: sticky;
    top: 0;
    background-color: #f8f9fa; /* Match your table header color */
    z-index: 20;
}

/* Add a shadow to the fixed header */
.table-fixed-header::after {
    content: '';
    position: absolute;
    left: 0;
    right: 0;
    bottom: -5px;
    height: 5px;
    background: linear-gradient(to bottom, rgba(0,0,0,0.1), transparent);
}

/* Styles for fixed columns */
.table-fixed-column {
    position: sticky;
    right: 0;
    background-color: white;
    z-index: 5;
}

.table-fixed-column::before {
    content: '';
    position: absolute;
    top: 0;
    left: -5px;
    width: 5px;
    height: 100%;
    background: linear-gradient(to right, transparent, rgba(0,0,0,0.1));
}

/* Adjust the z-index for header cells to stay above fixed column */
.table-fixed-header th:last-child {
    z-index: 30;
}

/* Ensure the fixed column stays visible when scrolling */
.table-container {
    overflow: auto;
}

/* Table hover effects */
.table-hover tbody tr:hover {
    background-color: rgba(11, 0, 162, 0.05);
}
"@ | Out-File -FilePath "core/static/css/freeze.css" -Encoding utf8

# Create loading.js
@"
document.addEventListener('DOMContentLoaded', function() {
    // Get all forms that should show loading
    const forms = document.querySelectorAll('form');
    
    forms.forEach(form => {
        form.addEventListener('submit', function(e) {
            // Only show loading for forms that aren't the search form
            if (!form.classList.contains('no-loading')) {
                // Show loading overlay
                const loadingOverlay = document.getElementById('loadingOverlay');
                if (loadingOverlay) {
                    loadingOverlay.style.display = 'flex';
                }
                
                // Optional: Disable submit button to prevent double submission
                const submitButton = form.querySelector('button[type="submit"]');
                if (submitButton) {
                    submitButton.disabled = true;
                    submitButton.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Procesando...';
                }
            }
        });
    });
});
"@ | Out-File -FilePath "core/static/js/loading.js" -Encoding utf8

# Create persons template
@" 
{% extends "master.html" %}
{% load static %}

{% block title %}A R P A{% endblock %}
{% block navbar_title %}A R P A{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/finance/" class="btn btn-custom-primary" title="BienesyRentas">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="/cards/" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="/conflicts/" class="btn btn-custom-primary" title="Conflictos">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="/alerts/" class="btn btn-custom-primary" title="Alertas">
        {% if alerts_count > 0 %}
            <span class="badge bg-danger">{{ alerts_count }}</span>
        {% endif %}
        {% if alerts_count == 0 %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="/persons/import/" class="btn btn-custom-primary" title="Importar">
        {% if import_count > 0 %}
            <span class="badge bg-danger">{{ import_count }}</span>
        {% endif %}
        {% if import_count == 0 %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-upload"></i>
    </a>
    <a href="?{% for key, value in request.GET.items %}{{ key }}={{ value }}&{% endfor %}export=excel" class="btn btn-custom-primary btn-my-green" title="Exportar">
        <i class="fas fa-file-excel"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesión">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<!-- Search Form -->
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <!-- General Search -->
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar persona..." 
                       value="{{ request.GET.q }}">
            </div>
            
            <!-- Status Filter -->
            <div class="col-md-2">
                <select name="status" class="form-select form-select-lg">
                    <option value="">Estado</option>
                    <option value="Activo" {% if request.GET.status == 'Activo' %}selected{% endif %}>Activo</option>
                    <option value="Retirado" {% if request.GET.status == 'Retirado' %}selected{% endif %}>Retirado</option>
                </select>
            </div>
            
            <!-- Cargo Filter -->
            <div class="col-md-2">
                <select name="cargo" class="form-select form-select-lg">
                    <option value="">Cargo</option>
                    {% for cargo in cargos %}
                        <option value="{{ cargo }}" {% if request.GET.cargo == cargo %}selected{% endif %}>{{ cargo }}</option>
                    {% endfor %}
                </select>
            </div>
            
            <!-- Compania Filter -->
            <div class="col-md-2">
                <select name="compania" class="form-select form-select-lg">
                    <option value="">Compania</option>
                    {% for compania in companias %}
                        <option value="{{ compania }}" {% if request.GET.compania == compania %}selected{% endif %}>{{ compania }}</option>
                    {% endfor %}
                </select>
            </div>
            
            <!-- Submit Buttons -->
            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<!-- Persons Table -->
<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=revisar&sort_direction={% if current_order == 'revisar' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Revisar
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cedula&sort_direction={% if current_order == 'cedula' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                ID
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=nombre_completo&sort_direction={% if current_order == 'nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Nombre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cargo&sort_direction={% if current_order == 'cargo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=correo&sort_direction={% if current_order == 'correo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Correo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=compania&sort_direction={% if current_order == 'compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Compania
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=estado&sort_direction={% if current_order == 'estado' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Estado
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">Comentarios</th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
                        <tr {% if person.revisar %}class="table-warning"{% endif %}>
                            <td>
                                <a href="/admin/core/person/{{ person.cedula }}/change/" style="text-decoration: none;" title="{% if person.revisar %}Marcado para revisar{% else %}No marcado{% endif %}">
                                    <i class="fas fa-{% if person.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}" style="padding-left: 20px;"></i>
                                </a>
                            </td>
                            <td>{{ person.cedula }}</td>
                            <td>{{ person.nombre_completo }}</td>
                            <td>{{ person.cargo }}</td>
                            <td>{{ person.correo }}</td>
                            <td>{{ person.compania }}</td>
                            <td>
                                <span class="badge bg-{% if person.estado == 'Activo' %}success{% else %}danger{% endif %}">
                                    {{ person.estado }}
                                </span>
                            </td>
                            <td>{{ person.comments|truncatechars:30|default:"" }}</td>
                            <td class="table-fixed-column">
                                <a href="/persons/details/{{ person.cedula }}/" 
                                   class="btn btn-custom-primary btn-sm"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                    {% empty %}
                        <tr>
                            <td colspan="9" class="text-center py-4">
                                {% if request.GET.q or request.GET.status or request.GET.cargo or request.GET.compania %}
                                    Sin registros que coincidan con los filtros.
                                {% else %}
                                    Sin registros
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        <!-- Pagination -->
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/persons.html" -Encoding utf8 -Force

# Create alerts template
@" 
{% extends "master.html" %}

{% block title %}Alertas{% endblock %}
{% block navbar_title %}Alertas{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/persons/" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="/finance/" class="btn btn-custom-primary">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="/cards/" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="/conflicts/" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="/persons/import/" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-upload"></i>
    </a>
    <a href="?{% for key, value in request.GET.items %}{{ key }}={{ value }}&{% endfor %}export=excel" class="btn btn-custom-primary btn-my-green" title="Exportar">
        <i class="fas fa-file-excel"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesión">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<!-- Search Form -->
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <!-- General Search -->
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar persona..." 
                       value="{{ request.GET.q }}">
            </div>
            
            <!-- Status Filter -->
            <div class="col-md-2">
                <select name="status" class="form-select form-select-lg">
                    <option value="">Estado</option>
                    <option value="Activo" {% if request.GET.status == 'Activo' %}selected{% endif %}>Activo</option>
                    <option value="Retirado" {% if request.GET.status == 'Retirado' %}selected{% endif %}>Retirado</option>
                </select>
            </div>
            
            <!-- Cargo Filter -->
            <div class="col-md-2">
                <select name="cargo" class="form-select form-select-lg">
                    <option value="">Cargo</option>
                    {% for cargo in cargos %}
                        <option value="{{ cargo }}" {% if request.GET.cargo == cargo %}selected{% endif %}>{{ cargo }}</option>
                    {% endfor %}
                </select>
            </div>
            
            <!-- Compania Filter -->
            <div class="col-md-2">
                <select name="compania" class="form-select form-select-lg">
                    <option value="">Compania</option>
                    {% for compania in companias %}
                        <option value="{{ compania }}" {% if request.GET.compania == compania %}selected{% endif %}>{{ compania }}</option>
                    {% endfor %}
                </select>
            </div>
            
            <!-- Submit Buttons -->
            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<!-- Alerts Table -->
<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=revisar&sort_direction={% if current_order == 'revisar' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Revisar
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cedula&sort_direction={% if current_order == 'cedula' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                ID
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=nombre_completo&sort_direction={% if current_order == 'nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Nombre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cargo&sort_direction={% if current_order == 'cargo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=correo&sort_direction={% if current_order == 'correo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Correo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=compania&sort_direction={% if current_order == 'compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Compania
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=estado&sort_direction={% if current_order == 'estado' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Estado
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">Comentarios</th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
                        {% if person.revisar or person.comments %}
                        <tr class="{% if person.revisar %}table-warning{% endif %}">
                            <td>
                                {% if person.revisar %}
                                <a href="/admin/core/person/{{ person.cedula }}/change/" style="text-decoration: none;" title="Marcado para revisar">
                                    <i class="fas fa-check-square text-warning" style="padding-left: 20px;"></i>
                                </a>
                                {% else %}
                                <i class="fas fa-square text-secondary" style="padding-left: 20px;"></i>
                                {% endif %}
                            </td>
                            <td>{{ person.cedula }}</td>
                            <td>{{ person.nombre_completo }}</td>
                            <td>{{ person.cargo }}</td>
                            <td>{{ person.correo }}</td>
                            <td>{{ person.compania }}</td>
                            <td>
                                <span class="badge bg-{% if person.estado == 'Activo' %}success{% else %}danger{% endif %}">
                                    {{ person.estado }}
                                </span>
                            </td>
                            <td>
                                {% if person.comments %}
                                    <span class="badge bg-danger">Comentario</span>
                                    {{ person.comments|truncatechars:30 }}
                                {% else %}
                                    -
                                {% endif %}
                            </td>
                            <td class="table-fixed-column">
                                <a href="/persons/details/{{ person.cedula }}/" 
                                   class="btn btn-custom-primary btn-sm"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                        {% endif %}
                    {% empty %}
                        <tr>
                            <td colspan="9" class="text-center py-4">
                                {% if request.GET.q or request.GET.status or request.GET.cargo or request.GET.compania %}
                                    Sin registros que coincidan con los filtros.
                                {% else %}
                                    No hay alertas (ningún registro marcado para revisar o con comentarios)
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        <!-- Pagination -->
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/alerts.html" -Encoding utf8 -Force

# Create import template
@"
{% extends "master.html" %}

{% block title %}Importar desde Excel{% endblock %}
{% block navbar_title %}Importar Datos{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/persons/" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="/finance/" class="btn btn-custom-primary">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="/cards/" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="/conflicts/" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="/alerts/" class="btn btn-custom-primary">
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <form method="post" action="{% url 'process_persons' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary">
            <i class="fas fa-database"></i>
        </button>
    </form>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesiÃ³n">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
{% load static %}
<div class="loading-overlay" id="loadingOverlay">
    <div class="loading-content">
        <h4>Procesando datos...</h4>
        <div class="progress">
            <div class="progress-bar progress-bar-striped progress-bar-animated" 
                 role="progressbar" 
                 style="width: 100%"></div>
        </div>
        <p>Por favor espere, esto puede tomar unos segundos.</p>
    </div>
</div>

<!-- Add loading CSS -->
<link rel="stylesheet" href="{% static 'css/loading.css' %}">

<!-- Add loading JS -->
<script src="{% static 'js/loading.js' %}"></script>

<div class="row">
    <!-- First row with 3 cards -->
    <div class="col-md-4 mb-4">
        <div class="card h-100">
            <div class="card-body">
                <form method="post" enctype="multipart/form-data" action="{% url 'import_period_excel' %}">
                    {% csrf_token %}
                    <div class="mb-3">
                        <input type="file" class="form-control" id="period_excel_file" name="period_excel_file" required>
                        <div class="form-text">El archivo Excel de Periodos debe incluir las columnas: Id, Activo, FechaFinDeclaracion, FechaInicioDeclaracion, Ano declaracion</div>
                    </div>
                    <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar Periodos</button>
                </form>
            </div>
            {% for message in messages %}
                {% if 'import_period_excel' in message.tags %}
                <div class="card-footer">
                    <div class="alert alert-{{ message.tags }} alert-dismissible fade show mb-0">      
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                </div>
                {% endif %}
            {% endfor %}
        </div>
    </div>

    <div class="col-md-4 mb-4">
        <div class="card h-100">
            <div class="card-body">
                <form method="post" enctype="multipart/form-data" action="{% url 'import_persons' %}"> 
                    {% csrf_token %}
                    <div class="mb-3">
                        <input type="file" class="form-control" id="excel_file" name="excel_file" required>
                        <div class="form-text">El archivo Excel de Personas debe incluir las columnas: Id, NOMBRE COMPLETO, CARGO, Cedula, Correo, Compania, Estado</div>
                    </div>
                    <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar Personas</button>
                </form>
            </div>
            {% for message in messages %}
                {% if 'import_persons' in message.tags %}
                <div class="card-footer">
                    <div class="alert alert-{{ message.tags }} alert-dismissible fade show mb-0">      
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                </div>
                {% endif %}
            {% endfor %}
        </div>
    </div>

    <div class="col-md-4 mb-4">
        <div class="card h-100">
            <div class="card-body">
                <form method="post" enctype="multipart/form-data" action="{% url 'import_conflict_excel' %}">
                    {% csrf_token %}
                    <div class="mb-3">
                        <input type="file" class="form-control" id="conflict_excel_file" name="conflict_excel_file" required>
                        <div class="form-text">'ID', 'Cedula', 'Nombre', 'Compania', 'Cargo', 'Email', 'Fecha de Inicio', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11' </div>
                    </div>
                    <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar Conflictos</button>
                </form>
            </div>
            {% for message in messages %}
                {% if 'import_conflict_excel' in message.tags %}
                <div class="card-footer">
                    <div class="alert alert-{{ message.tags }} alert-dismissible fade show mb-0">      
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                </div>
                {% endif %}
            {% endfor %}
        </div>
    </div>
</div>

<div class="row">
    <!-- Left column with 3 import forms in a single card -->
    <div class="col-md-4">
        <div class="card h-100">
            <div class="card-body d-flex flex-column">
                <!-- Bienes y Rentas -->
                <div class="mb-4 flex-grow-1">
                    <form method="post" enctype="multipart/form-data" action="{% url 'import_protected_excel' %}">
                        {% csrf_token %}
                        <div class="mb-3">
                            <input type="file" class="form-control" id="protected_excel_file" name="protected_excel_file" required>
                            <div class="form-text">El archivo Excel de Bienes y Rentas debe incluir las columnas: </div>
                            <div class="mb-3">
                                <input type="password" class="form-control" id="excel_password" name="excel_password">
                                <div class="form-text">Ingrese la contrasena</div>
                            </div>
                        </div>
                        <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar Bienes y Rentas</button>
                    </form>
                </div>

                <!-- Visa -->
                <div class="flex-grow-1">
                    <form method="post" enctype="multipart/form-data" action="{% url 'import_visa_pdfs' %}">
                        {% csrf_token %}
                        <div class="mb-3">
                            <input type="file" class="form-control" id="visa_pdf_files" name="visa_pdf_files" multiple webkitdirectory directory required>
                            <div class="form-text">Seleccione la carpeta con los PDFs de VISA</div>
                        </div>
                        <!-- Add password input field -->
                        <div class="mb-3">
                            <input type="password" class="form-control" id="visa_pdf_password" name="visa_pdf_password" placeholder="Clave">
                            <div class="form-text">Ingrese la contraseña si los PDFs están protegidos</div>
                        </div>
                        <button type="submit" class="btn btn-custom-primary btn-lg text-start">Procesar VISA</button>
                    </form>
                </div>
            </div>
            
            <!-- Messages for all three forms -->
            {% for message in messages %}
                {% if 'import_protected_excel' in message.tags or 'import_visa_pdfs' in message.tags %}
                <div class="card-footer">
                    <div class="alert alert-{{ message.tags }} alert-dismissible fade show mb-0">
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                </div>
                {% endif %}
            {% endfor %}
        </div>
    </div>

    <!-- Right column with analysis results -->
    <div class="col-md-8">
        <div class="card h-100">
            <div class="card-header bg-light">
                <h5 class="mb-0">Resultados del Analisis</h5>
            </div>
            <div class="card-body">
                {% if analysis_results %}
                <div class="table-responsive">
                    <table class="table table-sm">
                        <thead>
                            <tr>
                                <th>Archivo Generado</th>
                                <th>Registros</th>
                                <th>Estado</th>
                                <th>Ultima Actualizacion</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for result in analysis_results %}
                            <tr>
                                <td>{{ result.filename }}</td>
                                <td>{{ result.records|default:"-" }}</td>
                                <td>
                                    <span class="badge bg-{% if result.status == 'success' %}success{% elif result.status == 'error' %}danger{% else %}secondary{% endif %}">
                                        {% if result.status == 'success' %}
                                            Exitoso
                                        {% elif result.status == 'pending' %}
                                            Pendiente
                                        {% elif result.status == 'error' %}
                                            Error
                                        {% else %}
                                            {{ result.status|capfirst }}
                                        {% endif %}
                                    </span>
                                    {% if result.status == 'error' and result.error %}
                                    <small class="text-muted d-block">{{ result.error }}</small>   
                                    {% endif %}
                                </td>
                                <td>
                                    {% if result.last_updated %}
                                    {{ result.last_updated|date:"d/m/Y H:i" }}
                                    {% else %}
                                    -
                                    {% endif %}
                                </td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
                {% else %}
                <div class="text-center py-4">
                    <i class="fas fa-info-circle fa-3x text-muted mb-3"></i>
                    <p class="text-muted">No hay resultados de analisis disponibles</p>
                </div>
                {% endif %}
            </div>
            <div class="card-footer">
                <small class="text-muted">Los archivos se procesan en: core/src/</small>
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/import.html" -Encoding utf8

# Create details template
@"
{% extends "master.html" %}
{% load humanize %}

{% block title %}Detalles - {{ myperson.nombre_completo }}{% endblock %}
{% block navbar_title %}{{ myperson.nombre_completo }}{% endblock %}

{% block navbar_buttons %}
<a href="/admin/core/person/{{ myperson.cedula }}/change/" class="btn btn-outline-dark" title="Admin">
    <i class="fas fa-wrench"></i>
</a>
<a href="/" class="btn btn-custom-primary"><i class="fas fa-arrow-right"></i></a>
{% endblock %}

{% block content %}
<div class="row">
    <div class="col-md-6 mb-4"> {# Column for Informacion Personal - half width #}
        <div class="card h-100"> {# Added h-100 for equal height #}
            <div class="card-header bg-light">
                <h5 class="mb-0">Informacion Personal</h5>
            </div>
            <div class="card-body">
                <table class="table">
                    <tr>
                        <th>ID:</th>
                        <td>{{ myperson.cedula }}</td>
                    </tr>
                    <tr>
                        <th>Nombre:</th>
                        <td>{{ myperson.nombre_completo }}</td>
                    </tr>
                    <tr>
                        <th>Cargo:</th>
                        <td>{{ myperson.cargo }}</td>
                    </tr>
                    <tr>
                        <th>Correo:</th>
                        <td>{{ myperson.correo }}</td>
                    </tr>
                    <tr>
                        <th>Compania:</th>
                        <td>{{ myperson.compania }}</td>
                    </tr>
                    <tr>
                        <th>Estado:</th>
                        <td>
                            <span class="badge bg-{% if myperson.estado == 'Activo' %}success{% else %}danger{% endif %}">
                                {{ myperson.estado }}
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <th>Por revisar:</th>
                        <td>
                            {% if myperson.revisar %}
                                <span class="badge bg-warning text-dark">Si</span>
                            {% else %}
                                <span class="badge bg-secondary">No</span>
                            {% endif %}
                        </td>
                    </tr>
                    <tr>
                        <th>Comentarios:</th>
                        <td>{{ myperson.comments|linebreaks }}</td>
                    </tr>
                </table>
            </div>
        </div>
    </div>

    <div class="col-md-6 mb-4"> {# Column for Conflictos Declarados - half width #}
        <div class="card h-100"> {# Added h-100 for equal height #}
            <div class="card-header bg-light">
                <h5 class="mb-0">Conflictos Declarados</h5>
            </div>
            <div class="card-body p-0">
                {% if conflicts %}
                {% for conflict in conflicts %}
                <div class="table-responsive">
                    <table class="table table-striped table-hover mb-0">
                        <tbody>
                            <tr>
                                <th scope="row">Fecha de Inicio</th>
                                <td>{{ conflict.fecha_inicio|date:"d/m/Y"|default:"-" }}</td>
                            </tr>
                            <tr>
                                <th scope="row">Accionista de algun proveedor del grupo</th>
                                <td class="text-center">{% if conflict.q1 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Familiar de algun accionista, proveedor o empleado</th>
                                <td class="text-center">{% if conflict.q2 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Accionista de alguna companiÂ­a del grupo</th>
                                <td class="text-center">{% if conflict.q3 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Actividades extralaborales</th>
                                <td class="text-center">{% if conflict.q4 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Negocios o bienes con empleados del grupo</th>
                                <td class="text-center">{% if conflict.q5 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Participacion en juntas o consejos directivos</th>
                                <td class="text-center">{% if conflict.q6 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Potencial conflicto diferente a los anteriores</th>
                                <td class="text-center">{% if conflict.q7 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Consciente del codigo de conducta empresarial</th>
                                <td class="text-center">{% if conflict.q8 %}<i class="fas fa-check text-success"></i>{% else %}<i class="fas fa-times text-danger"></i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Veracidad de la informacion consignada</th>
                                <td class="text-center">{% if conflict.q9 %}<i class="fas fa-check text-success"></i>{% else %}<i class="fas fa-times text-danger"></i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Familiar de algun funcionario publico</th>
                                <td class="text-center">{% if conflict.q10 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Relacion con el sector publico o funcionario publico</th>
                                <td class="text-center">{% if conflict.q11 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <hr> {% endfor %}
                {% else %}
                <p class="text-center py-4">No hay conflictos declarados disponibles</p>
                {% endif %}
            </div>
        </div>
    </div>
</div>

<div class="row"> {# New row for Reportes Financieros #}
    <div class="col-md-12 mb-4"> {# Full width column for Reportes Financieros #}
        <div class="card">
            <div class="card-header bg-light d-flex justify-content-between align-items-center">
                <h5 class="mb-0">Reportes Financieros</h5>
                <div>
                    <span class="badge bg-primary">{{ financial_reports.count }} periodos</span>
                </div>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive table-container">
                    <table class="table table-striped table-hover mb-0">
                        <thead class="table-fixed-header">
                            <tr>
                                <th>Ano</th>
                                <th scope="col">Variaciones</th>
                                <th>Activos</th>
                                <th>Pasivos</th>
                                <th>Ingresos</th>
                                <th>Patrimonio</th>
                                <th>Banco</th>
                                <th>Bienes</th>
                                <th>Inversiones</th>
                                <th>Apalancamiento</th>
                                <th>Endeudamiento</th>
                                <th>Indice</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for report in financial_reports %}
                            <tr>
                                <td>{{ report.ano_declaracion|floatformat:"0"|default:"-" }}</td>
                                <th>Relativa</th>
                                <td>{{ report.activos_var_rel|default:"-" }}</td>
                                <td>{{ report.pasivos_var_rel|default:"-" }}</td>
                                <td>{{ report.ingresos_var_rel|default:"-" }}</td>
                                <td>{{ report.patrimonio_var_rel|default:"-" }}</td>
                                <td>{{ report.banco_saldo_var_rel|default:"-" }}</td>
                                <td>{{ report.bienes_var_rel|default:"-" }}</td>
                                <td>{{ report.inversiones_var_rel|default:"-" }}</td>
                                <td>{{ report.apalancamiento_var_rel|default:"-" }}</td>
                                <td>{{ report.endeudamiento_var_rel|default:"-" }}</td>
                                <td>{{ report.aum_pat_subito|default:"-" }}</td>
                            </tr>
                            <tr>
                                <th></th>
                                <th scope="col">Absoluta</th>
                                <td>{{ report.activos_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.pasivos_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.ingresos_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.patrimonio_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.banco_saldo_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.bienes_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.inversiones_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.apalancamiento_var_abs|default:"-" }}</td>
                                <td>{{ report.endeudamiento_var_abs|default:"-" }}</td>
                                <td>{{ report.capital_var_abs|intcomma|default:"-" }}</td>
                            </tr>
                            <tr>
                                <td></td>
                                <th scope="col">Total</th>
                                <td>&#36;{{ report.activos|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.pasivos|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.ingresos|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.patrimonio|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.banco_saldo|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.bienes|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.inversiones|floatformat:2|intcomma|default:"-" }}</td>
                                <td>{{ report.apalancamiento|floatformat:2|default:"-" }}</td>
                                <td>{{ report.endeudamiento|floatformat:2|default:"-" }}</td>
                                <td>&#36;{{ report.capital|floatformat:2|intcomma|default:"-" }}</td>
                            </tr>
                            <tr>
                                <th></th>
                                <th scope="col">Cant.</th>
                                <td></td>
                                <td>{{ report.cant_deudas|default:"-" }}</td>
                                <td>{{ report.cant_ingresos|default:"-" }}</td>
                                <td></td>
                                <td>C{{ report.cant_cuentas|default:"-" }} B{{ report.cant_bancos|default:"-" }}</td>
                                <td>{{ report.cant_bienes|default:"-" }}</td>
                                <td>{{ report.cant_inversiones|default:"-" }}</td>
                                <td></td>
                                <td></td>
                                <td></td>
                            </tr>
                                
                            </tr>
                            {% empty %}
                            <tr>
                                <td colspan="8" class="text-center py-4">
                                    No hay reportes financieros disponibles
                                </td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/details.html" -Encoding utf8

# Create finances template
@"
{% extends "master.html" %}

{% block title %}Bienes y Rentas{% endblock %}
{% block navbar_title %}Bienes y Rentas{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/persons/" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="/cards/" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="/conflicts/" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="/alerts/" class="btn btn-custom-primary">
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="/persons/import/" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-upload"></i>
    </a>
    <a href="?{% for key, value in request.GET.items %}{{ key }}={{ value }}&{% endfor %}export=excel" class="btn btn-custom-primary btn-my-green" title="Exportar">
        <i class="fas fa-file-excel"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesión">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<!-- Search Form -->
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <!-- General Search -->
            <div class="col-md-3">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar persona..." 
                       value="{{ request.GET.q }}">
            </div>
            
            <!-- Column Selector -->
            <div class="col-md-3">
                <select name="column" class="form-select form-select-lg">
                    <option value="">Selecciona Columna</option>
                    <option value="ano_declaracion" {% if request.GET.column == 'ano_declaracion' %}selected{% endif %}>Ano Declaracion</option>
                    <option value="aum_pat_subito" {% if request.GET.column == 'aum_pat_subito' %}selected{% endif %}>Aum. Pat. Subito</option>
                    <option value="activos_var_rel" {% if request.GET.column == 'activos_var_rel' %}selected{% endif %}>Activos Var. Rel.</option>
                    <option value="pasivos_var_rel" {% if request.GET.column == 'pasivos_var_rel' %}selected{% endif %}>Pasivos Var. Rel.</option>
                    <option value="patrimonio_var_rel" {% if request.GET.column == 'patrimonio_var_rel' %}selected{% endif %}>Patrimonio Var. Rel.</option>
                    <option value="apalancamiento_var_rel" {% if request.GET.column == 'apalancamiento_var_rel' %}selected{% endif %}>Apalancamiento Var. Rel.</option>
                    <option value="endeudamiento_var_rel" {% if request.GET.column == 'endeudamiento_var_rel' %}selected{% endif %}>Endeudamiento Var. Rel.</option>
                    <option value="banco_saldo_var_rel" {% if request.GET.column == 'banco_saldo_var_rel' %}selected{% endif %}>BancoSaldo Var. Rel.</option>
                    <option value="bienes_var_rel" {% if request.GET.column == 'bienes_var_rel' %}selected{% endif %}>Bienes Var. Rel.</option>
                    <option value="inversiones_var_rel" {% if request.GET.column == 'inversiones_var_rel' %}selected{% endif %}>Inversiones Var. Rel.</option>
                </select>
            </div>
            
            <!-- Operator Selector -->
            <div class="col-md-2">
                <select name="operator" class="form-select form-select-lg">
                    <option value="">Selecciona operador</option>
                    <option value=">" {% if request.GET.operator == '>' %}selected{% endif %}>Mayor que</option>
                    <option value="<" {% if request.GET.operator == '<' %}selected{% endif %}>Menor que</option>
                    <option value="=" {% if request.GET.operator == '=' %}selected{% endif %}>Igual a</option>
                    <option value=">=" {% if request.GET.operator == '>=' %}selected{% endif %}>Mayor o igual</option>
                    <option value="<=" {% if request.GET.operator == '<=' %}selected{% endif %}>Menor o igual</option>
                    <option value="between" {% if request.GET.operator == 'between' %}selected{% endif %}>Entre</option>
                    <option value="contains" {% if request.GET.operator == 'contains' %}selected{% endif %}>Contiene</option>
                </select>
            </div>
            
            <!-- Value Input -->
            <div class="col-md-2">
                <input type="text" 
                       name="value" 
                       class="form-control form-control-lg" 
                       placeholder="Valor" 
                       value="{{ request.GET.value }}">
            </div>
            
            <!-- Submit Buttons -->
            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<!-- Persons Table -->
<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: red; color: white;">-50%</th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: red; color: white;">-50%</th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: red; color: white;">-50%</th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th class="table-fixed-column"></th>
                    </tr>
                    <tr>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: #228B22; color: white;">-30%</th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: #228B22; color: white;">-30%</th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: #228B22; color: white;">-30%</th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th class="table-fixed-column"></th>
                    </tr>
                    <tr>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th >Medio</th>
                        <th >"<="</th>
                        <th style="background-color: #228B22; "></th>
                        <th ></th>
                        <th style="background-color: #228B22;  color: white;">1.5</th>
                        <th ></th>
                        <th style="background-color: #228B22; color: white;">30%</th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: #228B22; color: white;">30%</th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: #228B22; color: white;">30%</th>
                        <th ></th>
                        <th style="background-color: #228B22; color: white;">4</th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: #228B22; color: white;">>4</th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th class="table-fixed-column"></th>
                    </tr>
                    <tr>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th >Alto</th>
                        <th >"<"</th>
                        <th style="background-color: red;"></th>
                        <th ></th>
                        <th style="background-color: red; color: white;">2</th>
                        <th ></th>
                        <th style="background-color: red; color: white;">50%</th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: red; color: white;">50%</th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: red; color: white;">50%</th>
                        <th ></th>
                        <th style="background-color: red; color: white;">6</th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th style="background-color: red; color: white;">>6</th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th ></th>
                        <th class="table-fixed-column"></th>
                    </tr>
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=revisar&sort_direction={% if current_order == 'revisar' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Revisar
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=nombre_completo&sort_direction={% if current_order == 'nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Nombre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=compania&sort_direction={% if current_order == 'compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Compania
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cargo&sort_direction={% if current_order == 'cargo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargo
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">Comentarios</th>
                        <th style="color: rgb(0, 0, 0);">Periodo</th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__ano_declaracion&sort_direction={% if current_order == 'financial_reports__ano_declaracion' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ano
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__aum_pat_subito&sort_direction={% if current_order == 'financial_reports__aum_pat_subito' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Aum. Pat. Subito
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__patrimonio&sort_direction={% if current_order == 'financial_reports__patrimonio' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Patrimonio
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__patrimonio_var_rel&sort_direction={% if current_order == 'financial_reports__patrimonio_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Patrimonio Var. Rel. % 
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__patrimonio_var_abs&sort_direction={% if current_order == 'financial_reports__patrimonio_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Patrimonio Var. Abs. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__activos&sort_direction={% if current_order == 'financial_reports__activos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Activos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__activos_var_rel&sort_direction={% if current_order == 'financial_reports__activos_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Activos Var. Rel. %
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__activos_var_abs&sort_direction={% if current_order == 'financial_reports__activos_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Activos Var. Abs. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__pasivos&sort_direction={% if current_order == 'financial_reports__pasivos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Pasivos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__pasivos_var_rel&sort_direction={% if current_order == 'financial_reports__pasivos_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Pasivos Var. Rel. 
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__pasivos_var_abs&sort_direction={% if current_order == 'financial_reports__pasivos_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Pasivos Var. Abs. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_deudas&sort_direction={% if current_order == 'financial_reports__cant_deudas' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Deudas
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__ingresos&sort_direction={% if current_order == 'financial_reports__ingresos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ingresos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__ingresos_var_rel&sort_direction={% if current_order == 'financial_reports__ingresos_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ingresos Var. Rel. %
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__ingresos_var_abs&sort_direction={% if current_order == 'financial_reports__ingresos_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ingresos Var. Abs. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_ingresos&sort_direction={% if current_order == 'financial_reports__cant_ingresos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Ingresos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__banco_saldo&sort_direction={% if current_order == 'financial_reports__banco_saldo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bancos Saldo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__banco_saldo_var_rel&sort_direction={% if current_order == 'financial_reports__banco_saldo_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bancos Var. %
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__banco_saldo_var_abs&sort_direction={% if current_order == 'financial_reports__banco_saldo_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bancos Var. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_cuentas&sort_direction={% if current_order == 'financial_reports__cant_cuentas' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Cuentas
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_bancos&sort_direction={% if current_order == 'financial_reports__cant_bancos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Bancos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__bienes_valor&sort_direction={% if current_order == 'financial_reports__bienes_valor' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bienes Valor
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__bienes_var_rel&sort_direction={% if current_order == 'financial_reports__bienes_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bienes Var. %
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__bienes_var_abs&sort_direction={% if current_order == 'financial_reports__bienes_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bienes Var. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_bienes&sort_direction={% if current_order == 'financial_reports__cant_bienes' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Bienes
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__inversiones_valor&sort_direction={% if current_order == 'financial_reports__inversiones_valor' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Inversiones Valor
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__inversiones_var_rel&sort_direction={% if current_order == 'financial_reports__inversiones_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Inversiones Var. %
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__inversiones_var_abs&sort_direction={% if current_order == 'financial_reports__inversiones_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Inversiones Var. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_inversiones&sort_direction={% if current_order == 'financial_reports__cant_inversiones' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Inversiones
                            </a>
                        </th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
                        {% for report in person.financial_reports.all %}
                        <tr {% if person.revisar %}class="table-warning"{% endif %}>
                            <td>
                                <a href="/admin/core/person/{{ person.cedula }}/change/" style="text-decoration: none;" title="{% if person.revisar %}Marcado para revisar{% else %}No marcado{% endif %}">
                                    <i class="fas fa-{% if person.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}" style="padding-left: 20px;"></i>
                                </a>
                            </td>
                            <td>{{ person.nombre_completo }}</td>
                            <td>{{ person.compania }}</td>
                            <td>{{ person.cargo }}</td>
                            <td>{{ person.comments|truncatechars:30|default:"" }}</td>
                            <td>{{ report.fkIdPeriodo|floatformat:"0"|default:"-" }}</td>
                            <td>{{ report.ano_declaracion|floatformat:"0"|default:"-" }}</td>
                            <td>{{ report.aum_pat_subito|default:"-" }}</td>
                            <td>{{ report.patrimonio|default:"-" }}</td>
                            <td>{{ report.patrimonio_var_rel|default:"-" }}</td>
                            <td>{{ report.patrimonio_var_abs|default:"-" }}</td>
                            <td>{{ report.activos|default:"-" }}</td>
                            <td>{{ report.activos_var_rel|default:"-" }}</td>
                            <td>{{ report.activos_var_abs|default:"-" }}</td>
                            <td>{{ report.pasivos|default:"-" }}</td>
                            <td>{{ report.pasivos_var_rel|default:"-" }}</td>
                            <td>{{ report.pasivos_var_abs|default:"-" }}</td>
                            <td>{{ report.cant_deudas|default:"-" }}</td>
                            <td>{{ report.ingresos|default:"-" }}</td>
                            <td>{{ report.ingresos_var_rel|default:"-" }}</td>
                            <td>{{ report.ingresos_var_abs|default:"-" }}</td>
                            <td>{{ report.cant_ingresos|default:"-" }}</td>
                            <td>{{ report.banco_saldo|default:"-" }}</td>
                            <td>{{ report.banco_saldo_var_rel|default:"-" }}</td>
                            <td>{{ report.banco_saldo_var_abs|default:"-" }}</td>
                            <td>{{ report.cant_cuentas|default:"-" }}</td>
                            <td>{{ report.cant_bancos|default:"-" }}</td>
                            <td>{{ report.bienes|default:"-" }}</td>
                            <td>{{ report.bienes_var_rel|default:"-" }}</td>
                            <td>{{ report.bienes_var_abs|default:"-" }}</td>
                            <td>{{ report.cant_bienes|default:"-" }}</td>
                            <td>{{ report.inversiones|default:"-" }}</td>
                            <td>{{ report.inversiones_var_rel|default:"-" }}</td>
                            <td>{{ report.inversiones_var_abs|default:"-" }}</td>
                            <td>{{ report.cant_inversiones|default:"-" }}</td>
                            <td class="table-fixed-column">
                                <a href="/persons/details/{{ person.cedula }}/" 
                                   class="btn btn-custom-primary btn-sm"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                        {% empty %}
                        <tr>
                            <td colspan="14">{{ person.nombre_completo }} - No hay reportes financieros</td>
                        </tr>
                        {% endfor %}
                    {% empty %}
                        <tr>
                            <td colspan="14" class="text-center py-4">
                                {% if request.GET.q or request.GET.column or request.GET.operator or request.GET.value %}
                                    Sin registros que coincidan con los filtros.
                                {% else %}
                                    Sin registros
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        <!-- Pagination -->
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/finances.html" -Encoding utf8

# Create conflicts template
@"
{% extends "master.html" %}

{% block title %}Conflictos de Interes{% endblock %}
{% block navbar_title %}Conflictos de Interes{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/persons/" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="/finance/" class="btn btn-custom-primary">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="/cards/" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="/alerts/" class="btn btn-custom-primary">
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="/persons/import/" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-upload"></i>
    </a>
    <a href="?{% for key, value in request.GET.items %}{{ key }}={{ value }}&{% endfor %}export=excel" class="btn btn-custom-primary btn-my-green" title="Exportar">
        <i class="fas fa-file-excel"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesión">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<!-- Search Form -->
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center no-loading">
            <!-- General Search -->
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar..." 
                       value="{{ request.GET.q }}">
            </div>

            <!-- Compania Filter -->
            <div class="col-md-2">
                <select name="compania" class="form-select form-select-lg">
                    <option value="">Compania</option>
                    {% for compania in companias %}
                        <option value="{{ compania }}" {% if request.GET.compania == compania %}selected{% endif %}>{{ compania }}</option>
                    {% endfor %}
                </select>
            </div>
            
            <!-- Column Selector -->
            <div class="col-md-2">
                <select name="column" class="form-select form-select-lg">
                    <option value="">Selecciona Pregunta</option>
                    <option value="q1" {% if request.GET.column == 'q1' %}selected{% endif %}>Accionista de proveedor</option>
                    <option value="q2" {% if request.GET.column == 'q2' %}selected{% endif %}>Familiar de accionista/empleado</option>
                    <option value="q3" {% if request.GET.column == 'q3' %}selected{% endif %}>Accionista del grupo</option>
                    <option value="q4" {% if request.GET.column == 'q4' %}selected{% endif %}>Actividades extralaborales</option>
                    <option value="q5" {% if request.GET.column == 'q5' %}selected{% endif %}>Negocios con empleados</option>
                    <option value="q6" {% if request.GET.column == 'q6' %}selected{% endif %}>Participacion en juntas</option>
                    <option value="q7" {% if request.GET.column == 'q7' %}selected{% endif %}>Otro conflicto</option>
                    <option value="q8" {% if request.GET.column == 'q8' %}selected{% endif %}>Conoce codigo de conducta</option>
                    <option value="q9" {% if request.GET.column == 'q9' %}selected{% endif %}>Veracidad de informacion</option>
                    <option value="q10" {% if request.GET.column == 'q10' %}selected{% endif %}>Familiar de funcionario</option>
                    <option value="q11" {% if request.GET.column == 'q11' %}selected{% endif %}>Relacion con sector publico</option>
                </select>
            </div>
            
            <!-- Answer Selector -->
            <div class="col-md-2">
                <select name="answer" class="form-select form-select-lg">
                    <option value="">Selecciona Respuesta</option>
                    <option value="yes" {% if request.GET.answer == 'yes' %}selected{% endif %}>Si</option>
                    <option value="no" {% if request.GET.answer == 'no' %}selected{% endif %}>No</option>
                </select>
            </div>
            
            <!-- Submit Buttons -->
            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<!-- Conflicts Table -->
<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=revisar&sort_direction={% if current_order == 'revisar' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Revisar
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=nombre_completo&sort_direction={% if current_order == 'nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Nombre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=compania&sort_direction={% if current_order == 'compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Compania
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q1&sort_direction={% if current_order == 'conflicts__q1' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Accionista de proveedor
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q2&sort_direction={% if current_order == 'conflicts__q2' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Familiar de accionista/empleado
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q3&sort_direction={% if current_order == 'conflicts__q3' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Accionista del grupo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q4&sort_direction={% if current_order == 'conflicts__q4' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Actividades extralaborales
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q5&sort_direction={% if current_order == 'conflicts__q5' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Negocios con empleados
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q6&sort_direction={% if current_order == 'conflicts__q6' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Participacion en juntas
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q7&sort_direction={% if current_order == 'conflicts__q7' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Otro conflicto
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q8&sort_direction={% if current_order == 'conflicts__q8' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Conoce codigo de conducta
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q9&sort_direction={% if current_order == 'conflicts__q9' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Veracidad de informacion
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q10&sort_direction={% if current_order == 'conflicts__q10' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Familiar de funcionario
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=conflicts__q11&sort_direction={% if current_order == 'conflicts__q11' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Relacion con sector publico
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">Comentarios</th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
                        {% for conflict in person.conflicts.all %}
                        <tr {% if person.revisar %}class="table-warning"{% endif %}>
                            <td>
                                <a href="/admin/core/person/{{ person.cedula }}/change/" style="text-decoration: none;" title="{% if person.revisar %}Marcado para revisar{% else %}No marcado{% endif %}">
                                    <i class="fas fa-{% if person.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}" style="padding-left: 20px;"></i>
                                </a>
                            </td>
                            <td>{{ person.nombre_completo }}</td>
                            <td>{{ person.compania }}</td>
                            <td class="text-center">{% if conflict.q1 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q2 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q3 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q4 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q5 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q6 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q7 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q8 %}<i class="fas fa-check text-success"></i>{% else %}<i class="fas fa-times text-danger"></i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q9 %}<i class="fas fa-check text-success"></i>{% else %}<i class="fas fa-times text-danger"></i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q10 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q11 %}<i class="fas fa-check text-danger"></i>{% else %}<i class="fas fa-times text-success"></i>{% endif %}</td>
                            <td>{{ person.comments|truncatechars:30|default:"" }}</td>
                            <td class="table-fixed-column">
                                <a href="/persons/details/{{ person.cedula }}/" 
                                   class="btn btn-custom-primary btn-sm"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                        {% empty %}
                        <tr>
                            <td colspan="14">{{ person.nombre_completo }} - No hay conflictos declarados</td>
                        </tr>
                        {% endfor %}
                    {% empty %}
                        <tr>
                            <td colspan="14" class="text-center py-4">
                                {% if request.GET.q or request.GET.status or request.GET.cargo or request.GET.compania %}
                                    Sin registros que coincidan con los filtros.
                                {% else %}
                                    Sin registros
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        <!-- Pagination -->
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/conflicts.html" -Encoding utf8

# Create tarjetas template
@"
{% extends "master.html" %}

{% block title %}Tarjetas de Crédito{% endblock %}
{% block navbar_title %}Tarjetas de Crédito{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/persons/" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="/finance/" class="btn btn-custom-primary">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="/conflicts/" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="/alerts/" class="btn btn-custom-primary">
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="/persons/import/" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-upload"></i>
    </a>
    <a href="?{% for key, value in request.GET.items %}{{ key }}={{ value }}&{% endfor %}export=excel" class="btn btn-custom-primary btn-my-green" title="Exportar">
        <i class="fas fa-file-excel"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesión">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<!-- Search Form -->
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <!-- General Search 
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar tarjetahabiente..." 
                       value="{{ request.GET.q }}">
            </div> -->
            
            <!-- Card Type Filter -->
            <div class="col-md-3">
                <select name="card_type" class="form-select form-select-lg">
                    <option value="">Tipo</option>
                    <option value="MC" {% if request.GET.card_type == 'MC' %}selected{% endif %}>Mastercard</option>
                    <option value="VI" {% if request.GET.card_type == 'VI' %}selected{% endif %}>Visa</option>
                </select>
            </div>
            
            <!-- Date Range -->
            <div class="col-md-3">
                <input type="date" 
                       name="date_from" 
                       class="form-control form-control-lg" 
                       value="{{ request.GET.date_from }}">
            </div>
            <div class="col-md-3">
                <input type="date" 
                       name="date_to" 
                       class="form-control form-control-lg" 
                       value="{{ request.GET.date_to }}">
            </div>
            
            <!-- Submit Buttons -->
            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<!-- Cards Table -->
<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__nombre_completo&sort_direction={% if current_order == 'person__nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Tarjetahabiente
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=card_type&sort_direction={% if current_order == 'card_type' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Tipo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=card_number&sort_direction={% if current_order == 'card_number' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Número
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=transaction_date&sort_direction={% if current_order == 'transaction_date' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Fecha
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">Descripción</th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=original_value&sort_direction={% if current_order == 'original_value' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Valor Original
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=exchange_rate&sort_direction={% if current_order == 'exchange_rate' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Tasa
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=charges&sort_direction={% if current_order == 'charges' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=balance&sort_direction={% if current_order == 'balance' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Saldo
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">Cuotas</th>
                        <th style="color: rgb(0, 0, 0);">Archivo</th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <!-- In your table body -->
                <tbody>
                    {% for card in page_obj.object_list %}  <!-- Changed from cards to page_obj.object_list -->
                    <tr {% if card.person.revisar %}class="table-warning"{% endif %}>
                        <td>{{ card.person.nombre_completo }}</td>
                        <td>{{ card.get_card_type_display }}</td>
                        <td>**** **** **** {{ card.card_number|slice:"-4:" }}</td>
                        <td>{{ card.transaction_date|date:"d/m/Y" }}</td>
                        <td>{{ card.description|truncatechars:30 }}</td>
                        <td>`$`{{ card.original_value|floatformat:2 }}</td>
                        <td>{{ card.exchange_rate|default_if_none:"-"|floatformat:4 }}</td>
                        <td>`$`{{ card.charges|default_if_none:"-"|floatformat:2 }}</td>
                        <td>`$`{{ card.balance|default_if_none:"-"|floatformat:2 }}</td>
                        <td>{{ card.installments|default:"-" }}</td>
                        <td>{{ card.source_file|truncatechars:15 }}</td>
                        <td class="table-fixed-column">
                            <a href="/persons/details/{{ card.person.cedula }}/" 
                            class="btn btn-custom-primary btn-sm"
                            title="Ver detalles">
                                <i class="bi bi-person-vcard-fill"></i>
                            </a>
                        </td>
                    </tr>
                    {% empty %}
                        <tr>
                            <td colspan="12" class="text-center py-4">
                                {% if request.GET.q or request.GET.card_type or request.GET.date_from or request.GET.date_to %}
                                    No se encontraron transacciones con los filtros aplicados.
                                {% else %}
                                    No hay transacciones de tarjetas registradas.
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        <!-- Pagination -->
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/cards.html" -Encoding utf8

# Create login template
@"
{% extends "master.html" %}

{% block title %}Acceder{% endblock %}
{% block navbar_title %}Acceder{% endblock %}

{% block navbar_buttons %}
{% endblock %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-6">
        <div class="card border-0 shadow">
            <div class="card-body p-5">
                <div style="align-items: center; text-align: center;"> 
                        <a href="/" style="text-decoration: none;" >
                            <div class="logoIN" style="margin: 20px auto;"></div>
                        </a>
                    {% if form.errors %}
                    <div class="alert alert-danger">
                        Tu nombre de usuario y clave no coinciden. Por favor intenta de nuevo.
                    </div>
                    {% endif %}

                    {% if next %}
                        {% if user.is_authenticated %}
                        <div class="alert alert-warning">
                            Tu cuenta no tiene acceso a esta pagina. Para continuar,
                            por favor ingresa con una cuenta que tenga acceso.
                        </div>
                        {% else %}
                        <div class="alert alert-info">
                            Por favor accede con tu clave para ver esta pagina.
                        </div>
                        {% endif %}
                    {% endif %}

                    <form method="post" action="{% url 'login' %}">
                        {% csrf_token %}

                        <div class="mb-3">
                            <input type="text" name="username" class="form-control form-control-lg" id="id_username" placeholder="Usuario" required>
                        </div>

                        <div class="mb-4">
                            <input type="password" name="password" class="form-control form-control-lg" id="id_password" placeholder="Clave" required>
                        </div>

                        <div class="d-flex align-items-center justify-content-between">
                            <button type="submit" class="btn btn-custom-primary btn-lg">
                                <i class="fas fa-sign-in-alt"style="color: green;"></i>
                            </button>
                            <div>
                                <a href="{% url 'register' %}" class="btn btn-custom-primary" title="Registrarse">  
                                    <i class="fas fa-user-plus fa-lg"></i>
                                </a>
                                <a href="{% url 'password_reset' %}" class="btn btn-custom-primary" title="Recupera tu acceso">
                                    <i class="fas fa-key fa-lg" style="color: orange;"></i>
                                </a>
                            </div>
                        </div>

                        <input type="hidden" name="next" value="{{ next }}">
                    </form>
                </div> 
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/registration/login.html" -Encoding utf8

# Create login template
@"
{% extends "master.html" %}

{% block title %}Registro{% endblock %}
{% block navbar_title %}Registro{% endblock %}

{% block navbar_buttons %}
{% endblock %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-6">
        <div class="card border-0 shadow">
            <div class="card-body p-5">
                <div style="align-items: center; text-align: center;"> 
                        <a href="/" style="text-decoration: none;" >
                            <div class="logoIN" style="margin: 20px auto;"></div>
                        </a>
                    {% if messages %}
                        {% for message in messages %}
                            <div class="alert alert-{% if message.tags == 'error' %}danger{% else %}{{ message.tags }}{% endif %}">
                                {{ message }}
                            </div>
                        {% endfor %}
                    {% endif %}

                    <form method="post" action="{% url 'register' %}">
                        {% csrf_token %}
                        <div class="mb-3">
                            <input type="text" name="username" class="form-control form-control-lg" id="username" placeholder="Usuario" required>
                        </div>

                        <div class="mb-3">
                            <input type="email" name="email" class="form-control form-control-lg" id="email" placeholder="Correo" required>
                        </div>

                        <div class="mb-3">
                            <input type="password" name="password1" class="form-control form-control-lg" id="password1" placeholder="Clave" required>
                        </div>

                        <div class="mb-3">
                            <input type="password" name="password2" class="form-control form-control-lg" id="password2" placeholder="Repite tu clave" required>
                        </div>

                        <div class="d-flex align-items-center justify-content-between">
                            <button type="submit" class="btn btn-custom-primary btn-lg">
                                <i class="fas fa-user-plus fa-lg" style="color: green;"></i>
                            </button>
                            <div>
                                <a href="{% url 'login' %}" class="btn btn-custom-primary" title="Recupera tu acceso">
                                    <i class="fas fa-sign-in-alt" style="color: rgb(0, 0, 255);"></i>
                                </a>
                            </div>
                        </div>

                        <input type="hidden" name="next" value="{{ next }}">
                    </form>
                </div> 
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/registration/register.html" -Encoding utf8

    # Update settings.py
    $settingsContent = Get-Content -Path ".\arpa\settings.py" -Raw
    $settingsContent = $settingsContent -replace "INSTALLED_APPS = \[", "INSTALLED_APPS = [
    'core.apps.CoreConfig',
    'django.contrib.humanize',"
    $settingsContent = $settingsContent -replace "from pathlib import Path", "from pathlib import Path
import os"
    $settingsContent | Set-Content -Path ".\arpa\settings.py"

    # Add static files configuration
    Add-Content -Path ".\arpa\settings.py" -Value @"

# Static files (CSS, JavaScript, Images)
STATIC_URL = 'static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'
STATICFILES_DIRS = [
    BASE_DIR / "core/static",
]

MEDIA_URL = 'media/'
MEDIA_ROOT = BASE_DIR / 'media'

# Custom admin skin
ADMIN_SITE_HEADER = "A R P A"
ADMIN_SITE_TITLE = "ARPA Admin Portal"
ADMIN_INDEX_TITLE = "Bienvenido a A R P A"

LOGIN_REDIRECT_URL = '/'  # Where to redirect after login
LOGOUT_REDIRECT_URL = '/accounts/login/'  # Where to redirect after logout
"@

    # Run migrations
    python manage.py makemigrations core
    python manage.py migrate

    # Create superuser
    #python manage.py createsuperuser

    python manage.py collectstatic --noinput

    # Start the server
    Write-Host "🚀 Starting Django development server..." -ForegroundColor $GREEN
    python manage.py runserver

}

migratoDjango