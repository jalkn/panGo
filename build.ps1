function migratoDjango {
    param (
        [string]$ExcelFilePath = $null
    )

    $YELLOW = [ConsoleColor]::Yellow
    $GREEN = [ConsoleColor]::Green

    Write-Host "🚀 Creating Django Project with Excel Import Functionality" -ForegroundColor $YELLOW

    # Install required Python packages
    python -m pip install django whitenoise django-bootstrap-v5 openpyxl pandas

    # Create Django project
    django-admin startproject arpa
    cd arpa

    # Create apps
    python manage.py startapp persons
    python manage.py startapp bienes
    python manage.py startapp transacciones

    # Create models.py for persons app
    @"
from django.db import models

class Person(models.Model):
    ESTADO_CHOICES = [
        ('Activo', 'Activo'),
        ('Retirado', 'Retirado'),
    ]
    
    id = models.IntegerField(primary_key=True)
    nombre_completo = models.CharField(max_length=255, verbose_name="Nombre Completo")
    cargo = models.CharField(max_length=255, verbose_name="Cargo")
    cedula = models.CharField(max_length=20, verbose_name="Cedula")
    correo = models.EmailField(max_length=255, verbose_name="Correo")
    compania = models.CharField(max_length=255, verbose_name="Compania")
    estado = models.CharField(max_length=20, choices=ESTADO_CHOICES, default='Activo', verbose_name="Estado")

    def __str__(self):
        return f"{self.id} - {self.nombre_completo}"

    class Meta:
        verbose_name = "Persona"
        verbose_name_plural = "Personas"
"@ | Out-File -FilePath "persons/models.py" -Encoding utf8

    # Create models.py for bienes app
# Create models.py for bienes app with proper encoding
@"
from django.db import models
from persons.models import Person

class BienesyRentas(models.Model):
    idBien = models.AutoField(primary_key=True, verbose_name="ID Bien")
    Cedula = models.CharField(max_length=20)
    Usuario = models.CharField(max_length=255)
    Nombre = models.CharField(max_length=255)
    Compania = models.CharField(max_length=255, verbose_name="Compañía")
    Cargo = models.CharField(max_length=255)
    fkIdPeriodo = models.ForeignKey('Periodo', on_delete=models.CASCADE, verbose_name="Periodo")
    Ano_Declaracion = models.IntegerField(verbose_name="Año Declaración")
    Ano_Creacion = models.IntegerField(verbose_name="Año Creación")
    Activos = models.DecimalField(max_digits=20, decimal_places=2)
    Cant_Bienes = models.IntegerField(verbose_name="Cantidad Bienes")
    Cant_Bancos = models.IntegerField(verbose_name="Cantidad Bancos")
    Cant_Cuentas = models.IntegerField(verbose_name="Cantidad Cuentas")
    Cant_Inversiones = models.IntegerField(verbose_name="Cantidad Inversiones")
    Pasivos = models.DecimalField(max_digits=20, decimal_places=2)
    Cant_Deudas = models.IntegerField(verbose_name="Cantidad Deudas")
    Patrimonio = models.DecimalField(max_digits=20, decimal_places=2)
    Apalancamiento = models.DecimalField(max_digits=20, decimal_places=2)
    Endeudamiento = models.DecimalField(max_digits=20, decimal_places=2)
    Activos_Var_Abs = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Activos Variación Absoluta")
    Activos_Var_Rel = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Activos Variación Relativa")
    Pasivos_Var_Abs = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Pasivos Variación Absoluta")
    Pasivos_Var_Rel = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Pasivos Variación Relativa")
    Patrimonio_Var_Abs = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Patrimonio Variación Absoluta")
    Patrimonio_Var_Rel = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Patrimonio Variación Relativa")
    Apalancamiento_Var_Abs = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Apalancamiento Variación Absoluta")
    Apalancamiento_Var_Rel = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Apalancamiento Variación Relativa")
    Endeudamiento_Var_Abs = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Endeudamiento Variación Absoluta")
    Endeudamiento_Var_Rel = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Endeudamiento Variación Relativa")
    BancoSaldo = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Saldo Bancario")
    Bienes = models.DecimalField(max_digits=20, decimal_places=2)
    Inversiones = models.DecimalField(max_digits=20, decimal_places=2)
    BancoSaldo_Var_Abs = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Saldo Bancario Var. Abs.")
    BancoSaldo_Var_Rel = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Saldo Bancario Var. Rel.")
    Bienes_Var_Abs = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Bienes Variación Absoluta")
    Bienes_Var_Rel = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Bienes Variación Relativa")
    Inversiones_Var_Abs = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Inversiones Variación Absoluta")
    Inversiones_Var_Rel = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Inversiones Variación Relativa")
    Ingresos = models.DecimalField(max_digits=20, decimal_places=2)
    Cant_Ingresos = models.IntegerField(verbose_name="Cantidad Ingresos")
    Ingresos_Var_Abs = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Ingresos Variación Absoluta")
    Ingresos_Var_Rel = models.DecimalField(max_digits=20, decimal_places=2, verbose_name="Ingresos Variación Relativa")

    def __str__(self):
        return f"{self.idBien} - {self.Nombre}"

    class Meta:
        verbose_name = "Bienes y Rentas"
        verbose_name_plural = "Bienes y Rentas"
"@ | Out-File -FilePath "bienes/models.py" -Encoding utf8 -Force

# Create the Periodo model that fkIdPeriodo references
@"
from django.db import models

class Periodo(models.Model):
    nombre = models.CharField(max_length=100)
    fecha_inicio = models.DateField()
    fecha_fin = models.DateField()

    def __str__(self):
        return self.nombre
"@ | Out-File -FilePath "bienes/models.py" -Encoding utf8 -Append

    # Create models.py for transacciones app
    @"
from django.db import models
from persons.models import Person

class TransaccionesTarjetas(models.Model):
    Archivo = models.CharField(max_length=255)
    Tarjetahabiente = models.CharField(max_length=255)
    Número_de_Tarjeta = models.CharField(max_length=20)
    Moneda = models.CharField(max_length=10)
    Tipo_de_Cambio = models.DecimalField(max_digits=15, decimal_places=2, null=True, blank=True)
    Número_de_Autorización = models.CharField(max_length=50)
    Fecha_de_Transacción = models.DateField()
    Descripción = models.TextField()
    Valor_Original = models.DecimalField(max_digits=15, decimal_places=2)
    Tasa_Pactada = models.DecimalField(max_digits=15, decimal_places=2)
    Tasa_EA_Facturada = models.DecimalField(max_digits=15, decimal_places=2)
    Cargos_y_Abonos = models.DecimalField(max_digits=15, decimal_places=2)
    Saldo_a_Diferir = models.DecimalField(max_digits=15, decimal_places=2)
    Cuotas = models.CharField(max_length=20)
    Página = models.IntegerField()
    persona = models.ForeignKey(Person, on_delete=models.CASCADE, related_name='transacciones_mastercard')

    def __str__(self):
        return f"Transacción MC - {self.Tarjetahabiente} - {self.Fecha_de_Transacción}"

    class Meta:
        verbose_name = "Transacción Tarjeta MC"
        verbose_name_plural = "Transacciones Tarjetas MC"

class TransaccionesVisa(models.Model):
    Archivo = models.CharField(max_length=255)
    Tarjetahabiente = models.CharField(max_length=255)
    Número_de_Tarjeta = models.CharField(max_length=20)
    Número_de_Autorización = models.CharField(max_length=50)
    Fecha_de_Transacción = models.DateField()
    Descripción = models.TextField()
    Valor_Original = models.DecimalField(max_digits=15, decimal_places=2)
    Tasa_Pactada = models.DecimalField(max_digits=15, decimal_places=2)
    Tasa_EA_Facturada = models.DecimalField(max_digits=15, decimal_places=2)
    Cargos_y_Abonos = models.DecimalField(max_digits=15, decimal_places=2)
    Saldo_a_Diferir = models.DecimalField(max_digits=15, decimal_places=2)
    Cuotas = models.CharField(max_length=20)
    Página = models.IntegerField()
    persona = models.ForeignKey(Person, on_delete=models.CASCADE, related_name='transacciones_visa')

    def __str__(self):
        return f"Transacción Visa - {self.Tarjetahabiente} - {self.Fecha_de_Transacción}"

    class Meta:
        verbose_name = "Transacción Tarjeta Visa"
        verbose_name_plural = "Transacciones Tarjetas Visa"
"@ | Out-File -FilePath "transacciones/models.py" -Encoding utf8

    # Create views.py for bienes app
    @"
from django.shortcuts import render, redirect
from django.contrib import messages
from .models import BienesyRentas
import pandas as pd
from persons.models import Person

def list_bienes(request):
    bienes = BienesyRentas.objects.all()
    return render(request, 'bienes/list.html', {'bienes': bienes})

def import_bienes(request):
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        try:
            df = pd.read_excel(excel_file)
            
            for _, row in df.iterrows():
                # Find or create related Person
                person, _ = Person.objects.get_or_create(cedula=row['Cedula'], defaults={
                    'nombre_completo': row['Nombre'],
                    'cargo': row['Cargo'],
                    'compania': row['Compania']
                })
                
                BienesyRentas.objects.update_or_create(
                    Cedula=row['Cedula'],
                    fkIdPeriodo=row['fkIdPeriodo'],
                    defaults={
                        'Usuario': row.get('Usuario', ''),
                        'Nombre': row['Nombre'],
                        'Compañía': row['Compañía'],
                        'Cargo': row['Cargo'],
                        'Año_Declaración': row['Año Declaración'],
                        'Año_Creación': row['Año Creación'],
                        'Activos': row['Activos'],
                        'Cant_Bienes': row['Cant_Bienes'],
                        'Cant_Bancos': row['Cant_Bancos'],
                        'Cant_Cuentas': row['Cant_Cuentas'],
                        'Cant_Inversiones': row['Cant_Inversiones'],
                        'Pasivos': row['Pasivos'],
                        'Cant_Deudas': row['Cant_Deudas'],
                        'Patrimonio': row['Patrimonio'],
                        'Apalancamiento': row['Apalancamiento'],
                        'Endeudamiento': row['Endeudamiento'],
                        'Activos_Var_Abs': row['Activos Var. Abs.'],
                        'Activos_Var_Rel': row['Activos Var. Rel.'],
                        'Pasivos_Var_Abs': row['Pasivos Var. Abs.'],
                        'Pasivos_Var_Rel': row['Pasivos Var. Rel.'],
                        'Patrimonio_Var_Abs': row['Patrimonio Var. Abs.'],
                        'Patrimonio_Var_Rel': row['Patrimonio Var. Rel.'],
                        'Apalancamiento_Var_Abs': row['Apalancamiento Var. Abs.'],
                        'Apalancamiento_Var_Rel': row['Apalancamiento Var. Rel.'],
                        'Endeudamiento_Var_Abs': row['Endeudamiento Var. Abs.'],
                        'Endeudamiento_Var_Rel': row['Endeudamiento Var. Rel.'],
                        'BancoSaldo': row['BancoSaldo'],
                        'Bienes': row['Bienes'],
                        'Inversiones': row['Inversiones'],
                        'BancoSaldo_Var_Abs': row['BancoSaldo Var. Abs.'],
                        'BancoSaldo_Var_Rel': row['BancoSaldo Var. Rel.'],
                        'Bienes_Var_Abs': row['Bienes Var. Abs.'],
                        'Bienes_Var_Rel': row['Bienes Var. Rel.'],
                        'Inversiones_Var_Abs': row['Inversiones Var. Abs.'],
                        'Inversiones_Var_Rel': row['Inversiones Var. Rel.'],
                        'Ingresos': row['Ingresos'],
                        'Cant_Ingresos': row['Cant_Ingresos'],
                        'Ingresos_Var_Abs': row['Ingresos Var. Abs.'],
                        'Ingresos_Var_Rel': row['Ingresos Var. Rel.'],
                        'Compania': row['Compania'],
                        'persona': person
                    }
                )
            
            messages.success(request, f'Successfully imported/updated {len(df)} records!')
        except Exception as e:
            messages.error(request, f'Error importing data: {str(e)}')
        
        return redirect('list_bienes')
    
    return render(request, 'bienes/import.html')
"@ | Out-File -FilePath "bienes/views.py" -Encoding utf8

    # Create views.py for transacciones app
    @"
from django.shortcuts import render, redirect
from django.contrib import messages
from .models import TransaccionesTarjetas, TransaccionesVisa
import pandas as pd
from persons.models import Person

def list_transacciones(request):
    transacciones_mc = TransaccionesTarjetas.objects.all()
    transacciones_visa = TransaccionesVisa.objects.all()
    return render(request, 'transacciones/list.html', {
        'transacciones_mc': transacciones_mc,
        'transacciones_visa': transacciones_visa
    })

def import_mastercard(request):
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        try:
            df = pd.read_excel(excel_file)
            
            for _, row in df.iterrows():
                # Find or create related Person
                person, _ = Person.objects.get_or_create(
                    nombre_completo=row['Tarjetahabiente'],
                    defaults={'cedula': ''}
                )
                
                TransaccionesTarjetas.objects.create(
                    Archivo=row['Archivo'],
                    Tarjetahabiente=row['Tarjetahabiente'],
                    Número_de_Tarjeta=row['Número de Tarjeta'],
                    Moneda=row['Moneda'],
                    Tipo_de_Cambio=row['Tipo de Cambio'],
                    Número_de_Autorización=row['Número de Autorización'],
                    Fecha_de_Transacción=row['Fecha de Transacción'],
                    Descripción=row['Descripción'],
                    Valor_Original=row['Valor Original'],
                    Tasa_Pactada=row['Tasa Pactada'],
                    Tasa_EA_Facturada=row['Tasa EA Facturada'],
                    Cargos_y_Abonos=row['Cargos y Abonos'],
                    Saldo_a_Diferir=row['Saldo a Diferir'],
                    Cuotas=row['Cuotas'],
                    Página=row['Página'],
                    persona=person
                )
            
            messages.success(request, f'Successfully imported {len(df)} Mastercard transactions!')
        except Exception as e:
            messages.error(request, f'Error importing Mastercard data: {str(e)}')
        
        return redirect('list_transacciones')
    
    return render(request, 'transacciones/import_mastercard.html')

def import_visa(request):
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        try:
            df = pd.read_excel(excel_file)
            
            for _, row in df.iterrows():
                # Find or create related Person
                person, _ = Person.objects.get_or_create(
                    nombre_completo=row['Tarjetahabiente'],
                    defaults={'cedula': ''}
                )
                
                TransaccionesVisa.objects.create(
                    Archivo=row['Archivo'],
                    Tarjetahabiente=row['Tarjetahabiente'],
                    Número_de_Tarjeta=row['Número de Tarjeta'],
                    Número_de_Autorización=row['Número de Autorización'],
                    Fecha_de_Transacción=row['Fecha de Transacción'],
                    Descripción=row['Descripción'],
                    Valor_Original=row['Valor Original'],
                    Tasa_Pactada=row['Tasa Pactada'],
                    Tasa_EA_Facturada=row['Tasa EA Facturada'],
                    Cargos_y_Abonos=row['Cargos y Abonos'],
                    Saldo_a_Diferir=row['Saldo a Diferir'],
                    Cuotas=row['Cuotas'],
                    Página=row['Página'],
                    persona=person
                )
            
            messages.success(request, f'Successfully imported {len(df)} Visa transactions!')
        except Exception as e:
            messages.error(request, f'Error importing Visa data: {str(e)}')
        
        return redirect('list_transacciones')
    
    return render(request, 'transacciones/import_visa.html')
"@ | Out-File -FilePath "transacciones/views.py" -Encoding utf8

    # Create urls.py for bienes app
    @"
from django.urls import path
from . import views

urlpatterns = [
    path('', views.list_bienes, name='list_bienes'),
    path('import/', views.import_bienes, name='import_bienes'),
]
"@ | Out-File -FilePath "bienes/urls.py" -Encoding utf8

    # Create urls.py for transacciones app
    @"
from django.urls import path
from . import views

urlpatterns = [
    path('', views.list_transacciones, name='list_transacciones'),
    path('import-mastercard/', views.import_mastercard, name='import_mastercard'),
    path('import-visa/', views.import_visa, name='import_visa'),
]
"@ | Out-File -FilePath "transacciones/urls.py" -Encoding utf8

    # Update project urls.py
    @"
from django.contrib import admin
from django.urls import include, path

# Customize default admin interface
admin.site.site_header = 'ARPA Admin Portal'
admin.site.site_title = 'ARPA Admin Portal'

urlpatterns = [
    path('persons/', include('persons.urls')),
    path('bienes/', include('bienes.urls')),
    path('transacciones/', include('transacciones.urls')),
    path('admin/', admin.site.urls),
    path('', include('persons.urls')), 
]
"@ | Out-File -FilePath "arpa/urls.py" -Encoding utf8

    # Create templates for bienes app
    $bienesTemplates = @(
        "bienes/templates/bienes",
        "bienes/templates/bienes/admin"
    )
    foreach ($dir in $bienesTemplates) {
        New-Item -Path $dir -ItemType Directory -Force
    }

    # Create list template for bienes
    @"
{% extends "master.html" %}

{% block title %}
    Bienes y Rentas
{% endblock %}

{% block content %}
    <div class="d-flex justify-content-between mb-4">
        <a href="/" class="btn btn-secondary">Inicio</a>
        <div>
            <a href="import/" class="btn btn-primary">Cargar Datos</a>
        </div>
    </div>
    
    <h1 class="mb-4">Bienes y Rentas</h1>
    
    <div class="table-responsive">
        <table class="table table-striped table-hover">
            <thead class="table-dark">
                <tr>
                    <th>Cédula</th>
                    <th>Nombre</th>
                    <th>Compañía</th>
                    <th>Activos</th>
                    <th>Patrimonio</th>
                    <th>Acciones</th>
                </tr>
            </thead>
            <tbody>
                {% for bien in bienes %}
                    <tr>
                        <td>{{ bien.Cedula }}</td>
                        <td>{{ bien.Nombre }}</td>
                        <td>{{ bien.Compañía }}</td>
                        <td>{{ bien.Activos }}</td>
                        <td>{{ bien.Patrimonio }}</td>
                        <td>
                            <a href="/admin/bienes/bienesyrentas/{{ bien.id }}/change/" class="btn btn-warning btn-sm">Editar</a>
                        </td>
                    </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
{% endblock %}
"@ | Out-File -FilePath "bienes/templates/bienes/list.html" -Encoding utf8

    # Create import template for bienes
    @"
{% extends "master.html" %}

{% block title %}
    Cargar Bienes y Rentas
{% endblock %}

{% block content %}
    <div class="card">
        <div class="card-body">
            <form method="post" enctype="multipart/form-data">
                {% csrf_token %}
                <div class="mb-3">
                    <input type="file" class="form-control" id="excel_file" name="excel_file" required>
                    <div class="form-text">El archivo debe incluir todas las columnas del modelo Bienes y Rentas</div>
                </div>
                <button type="submit" class="btn btn-primary">Cargar</button>
                <a href="/bienes/" class="btn btn-secondary">Cancelar</a>
            </form>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "bienes/templates/bienes/import.html" -Encoding utf8

    # Create templates for transacciones app
    $transaccionesTemplates = @(
        "transacciones/templates/transacciones",
        "transacciones/templates/transacciones/admin"
    )
    foreach ($dir in $transaccionesTemplates) {
        New-Item -Path $dir -ItemType Directory -Force
    }

    # Create list template for transacciones
    @"
{% extends "master.html" %}

{% block title %}
    Transacciones de Tarjetas
{% endblock %}

{% block content %}
    <div class="d-flex justify-content-between mb-4">
        <a href="/" class="btn btn-secondary">Inicio</a>
        <div>
            <a href="import-mastercard/" class="btn btn-primary me-2">Cargar Mastercard</a>
            <a href="import-visa/" class="btn btn-success">Cargar Visa</a>
        </div>
    </div>
    
    <h1 class="mb-4">Transacciones de Tarjetas</h1>
    
    <div class="card mb-4">
        <div class="card-header bg-primary text-white">
            <h2>Mastercard</h2>
        </div>
        <div class="card-body">
            <div class="table-responsive">
                <table class="table table-striped table-hover">
                    <thead>
                        <tr>
                            <th>Tarjetahabiente</th>
                            <th>Fecha</th>
                            <th>Descripción</th>
                            <th>Valor</th>
                            <th>Acciones</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for trans in transacciones_mc %}
                            <tr>
                                <td>{{ trans.Tarjetahabiente }}</td>
                                <td>{{ trans.Fecha_de_Transacción }}</td>
                                <td>{{ trans.Descripción }}</td>
                                <td>{{ trans.Valor_Original }}</td>
                                <td>
                                    <a href="/admin/transacciones/transaccionestarjetas/{{ trans.id }}/change/" class="btn btn-warning btn-sm">Editar</a>
                                </td>
                            </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <div class="card">
        <div class="card-header bg-success text-white">
            <h2>Visa</h2>
        </div>
        <div class="card-body">
            <div class="table-responsive">
                <table class="table table-striped table-hover">
                    <thead>
                        <tr>
                            <th>Tarjetahabiente</th>
                            <th>Fecha</th>
                            <th>Descripción</th>
                            <th>Valor</th>
                            <th>Acciones</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for trans in transacciones_visa %}
                            <tr>
                                <td>{{ trans.Tarjetahabiente }}</td>
                                <td>{{ trans.Fecha_de_Transacción }}</td>
                                <td>{{ trans.Descripción }}</td>
                                <td>{{ trans.Valor_Original }}</td>
                                <td>
                                    <a href="/admin/transacciones/transaccionesvisa/{{ trans.id }}/change/" class="btn btn-warning btn-sm">Editar</a>
                                </td>
                            </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "transacciones/templates/transacciones/list.html" -Encoding utf8

    # Create import template for mastercard
    @"
{% extends "master.html" %}

{% block title %}
    Cargar Transacciones Mastercard
{% endblock %}

{% block content %}
    <div class="card">
        <div class="card-body">
            <form method="post" enctype="multipart/form-data">
                {% csrf_token %}
                <div class="mb-3">
                    <input type="file" class="form-control" id="excel_file" name="excel_file" required>
                    <div class="form-text">El archivo debe incluir las columnas del extracto Mastercard</div>
                </div>
                <button type="submit" class="btn btn-primary">Cargar</button>
                <a href="/transacciones/" class="btn btn-secondary">Cancelar</a>
            </form>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "transacciones/templates/transacciones/import_mastercard.html" -Encoding utf8

    # Create import template for visa
    @"
{% extends "master.html" %}

{% block title %}
    Cargar Transacciones Visa
{% endblock %}

{% block content %}
    <div class="card">
        <div class="card-body">
            <form method="post" enctype="multipart/form-data">
                {% csrf_token %}
                <div class="mb-3">
                    <input type="file" class="form-control" id="excel_file" name="excel_file" required>
                    <div class="form-text">El archivo debe incluir las columnas del extracto Visa</div>
                </div>
                <button type="submit" class="btn btn-primary">Cargar</button>
                <a href="/transacciones/" class="btn btn-secondary">Cancelar</a>
            </form>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "transacciones/templates/transacciones/import_visa.html" -Encoding utf8

    # Update main template to include new menu items
    @"
{% extends "master.html" %}

{% block title %}
    A R P A
{% endblock %}

{% block content %}
    <div class="card">
        <div class="card-body">
            <div class="d-grid gap-3 mt-4">
                <a href="persons/" class="btn btn-primary btn-lg">Personas</a>
                <a href="bienes/" class="btn btn-info btn-lg">Bienes y Rentas</a>
                <a href="transacciones/" class="btn btn-warning btn-lg">Transacciones Tarjetas</a>
                <a href="/admin/" class="btn btn-dark btn-lg">ARPA Admin Portal</a>
            </div>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "persons/templates/main.html" -Encoding utf8

    # Create admin.py for bienes app
    @"
from django.contrib import admin
from .models import BienesyRentas

@admin.register(BienesyRentas)
class BienesyRentasAdmin(admin.ModelAdmin):
    list_display = ('Cedula', 'Nombre', 'Compañía', 'Activos', 'Patrimonio')
    search_fields = ('Cedula', 'Nombre')
    list_filter = ('Compañía',)
    list_per_page = 25
    ordering = ('Nombre',)
    
    fieldsets = (
        (None, {
            'fields': ('Cedula', 'Nombre', 'Compañía', 'Cargo')
        }),
        ('Datos Financieros', {
            'classes': ('collapse',),
            'fields': ('Activos', 'Pasivos', 'Patrimonio', 'Apalancamiento', 'Endeudamiento'),
        }),
    )
"@ | Out-File -FilePath "bienes/admin.py" -Encoding utf8

    # Create admin.py for transacciones app
    @"
from django.contrib import admin
from .models import TransaccionesTarjetas, TransaccionesVisa

@admin.register(TransaccionesTarjetas)
class TransaccionesTarjetasAdmin(admin.ModelAdmin):
    list_display = ('Tarjetahabiente', 'Fecha_de_Transacción', 'Descripción', 'Valor_Original', 'Moneda')
    search_fields = ('Tarjetahabiente', 'Número_de_Tarjeta')
    list_filter = ('Moneda',)
    list_per_page = 25
    ordering = ('-Fecha_de_Transacción',)

@admin.register(TransaccionesVisa)
class TransaccionesVisaAdmin(admin.ModelAdmin):
    list_display = ('Tarjetahabiente', 'Fecha_de_Transacción', 'Descripción', 'Valor_Original')
    search_fields = ('Tarjetahabiente', 'Número_de_Tarjeta')
    list_per_page = 25
    ordering = ('-Fecha_de_Transacción',)
"@ | Out-File -FilePath "transacciones/admin.py" -Encoding utf8

    # Update settings.py
    $settingsContent = Get-Content -Path ".\arpa\settings.py" -Raw
    $settingsContent = $settingsContent -replace "INSTALLED_APPS = \[", "INSTALLED_APPS = [
    'persons.apps.PersonsConfig',
    'bienes.apps.BienesConfig',
    'transacciones.apps.TransaccionesConfig',"
    $settingsContent = $settingsContent -replace "from pathlib import Path", "from pathlib import Path
import os"
    $settingsContent | Set-Content -Path ".\arpa\settings.py"

    # Add static files configuration
    Add-Content -Path ".\arpa\settings.py" -Value @"

# Static files (CSS, JavaScript, Images)
STATIC_URL = 'static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'

MEDIA_URL = 'media/'
MEDIA_ROOT = BASE_DIR / 'media'
"@

    # Run migrations
    python manage.py makemigrations
    python manage.py migrate

    # Create superuser
    python manage.py createsuperuser

    # Import data if Excel file provided
    if ($ExcelFilePath -and (Test-Path $ExcelFilePath)) {
        Write-Host "Importing data from Excel file..." -ForegroundColor $GREEN
        
        $tempScriptPath = "temp_import.py"
        @"
import os
import django
import pandas as pd

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'arpa.settings')
django.setup()

from persons.models import Person

df = pd.read_excel(r'$ExcelFilePath')
for _, row in df.iterrows():
    Person.objects.update_or_create(
        id=row['Id'],
        defaults={
            'nombre_completo': row['NOMBRE COMPLETO'],
            'cargo': row['CARGO'],
            'cedula': row['Cedula'],
            'correo': row['Correo'],
            'compania': row['Compania'],
            'estado': row['Estado']
        }
    )
print(f"Successfully processed {len(df)} records")
"@ | Out-File -FilePath $tempScriptPath -Encoding utf8

        python $tempScriptPath
        Remove-Item -Path $tempScriptPath
    }

    # Start the server
    Write-Host "🚀 Starting Django development server..." -ForegroundColor $GREEN
    python manage.py runserver
}

migratoDjango