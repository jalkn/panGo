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

    # Create persons app
    python manage.py startapp persons

    # Create models.py with all required fields including ID
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

    # Create views.py with import functionality
    @"
from django.http import HttpResponse, HttpResponseRedirect
from django.template import loader
from django.shortcuts import render
from .models import Person
import pandas as pd
from django.contrib import messages

def persons(request):
    mypersons = Person.objects.all()
    template = loader.get_template('persons.html')
    context = {'mypersons': mypersons}
    return HttpResponse(template.render(context, request))
  
def details(request, id):
    myperson = Person.objects.get(id=id)
    template = loader.get_template('details.html')
    context = {'myperson': myperson}
    return HttpResponse(template.render(context, request))
  
def main(request):
    template = loader.get_template('main.html')
    return HttpResponse(template.render())

def import_from_excel(request):
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        try:
            df = pd.read_excel(excel_file)
            
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
            
            messages.success(request, f'Successfully imported/updated {len(df)} records!')
        except Exception as e:
            messages.error(request, f'Error importing data: {str(e)}')
        
        return HttpResponseRedirect('/persons/')
    
    return render(request, 'import_excel.html')
"@ | Out-File -FilePath "persons/views.py" -Encoding utf8

    # Create urls.py for persons app
    @"
from django.urls import path
from . import views

urlpatterns = [
    path('', views.main, name='main'),
    path('persons/', views.persons, name='persons'),
    path('persons/details/<int:id>', views.details, name='details'),
    path('persons/import/', views.import_from_excel, name='import_excel'),
]
"@ | Out-File -FilePath "persons/urls.py" -Encoding utf8

    # Create admin.py with enhanced configuration
    @"
from django.contrib import admin
from .models import Person

def make_active(modeladmin, request, queryset):
    queryset.update(estado='Activo')
make_active.short_description = "Mark selected as Active"

def make_retired(modeladmin, request, queryset):
    queryset.update(estado='Retirado')
make_retired.short_description = "Mark selected as Retired"

class PersonAdmin(admin.ModelAdmin):
    list_display = ("id", "nombre_completo", "cargo", "cedula", "correo", "compania", "estado")
    search_fields = ("nombre_completo", "cedula")
    list_filter = ("estado", "compania")
    list_per_page = 25
    ordering = ('nombre_completo',)
    actions = [make_active, make_retired]
    
    fieldsets = (
        (None, {
            'fields': ('id', 'nombre_completo', 'cargo', 'cedula')
        }),
        ('Advanced options', {
            'classes': ('collapse',),
            'fields': ('correo', 'compania', 'estado'),
        }),
    )
    
admin.site.register(Person, PersonAdmin)
"@ | Out-File -FilePath "persons/admin.py" -Encoding utf8

    # Update project urls.py with proper admin configuration
    @"
from django.contrib import admin
from django.urls import include, path

# Customize default admin interface
admin.site.site_header = 'ARPA Administration'
admin.site.site_title = 'ARPA Admin Portal'
admin.site.index_title = 'Welcome to ARPA Administration'

urlpatterns = [
    path('persons/', include('persons.urls')),
    path('admin/', admin.site.urls),
    path('', include('persons.urls')), 
]
"@ | Out-File -FilePath "arpa/urls.py" -Encoding utf8

    # Create templates directory structure
    $directories = @(
        "persons/templates",
        "persons/templates/admin",
        "persons/templates/admin/persons"
    )
    foreach ($dir in $directories) {
        New-Item -Path $dir -ItemType Directory -Force
    }

    # Create custom admin base template
    @"
{% extends "admin/base.html" %}

{% block title %}{{ title }} | {{ site_title|default:_('ARPA Administration') }}{% endblock %}

{% block branding %}
<h1 id="site-name"><a href="{% url 'admin:index' %}">{{ site_header|default:_('ARPA Administration') }}</a></h1>
{% endblock %}

{% block nav-global %}{% endblock %}
"@ | Out-File -FilePath "persons/templates/admin/base_site.html" -Encoding utf8

    # Create master template
    @"
<!DOCTYPE html>
<html>
<head>
    <title>{% block title %}{% endblock %}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {
            background-color: #f8f9fa;
        }
        .navbar-custom {
            background-color: #0056b3;
        }
        .footer {
            background-color: #343a40;
            color: white;
            padding: 20px 0;
            margin-top: 40px;
        }
        .navbar-title {
            color: white;
            margin-right: auto;
            padding-left: 15px;
            font-size: 1.25rem;
        }
    </style>
</head>
<body>
    <nav class="navbar navbar-expand-lg navbar-dark navbar-custom mb-4">
        <div class="container-fluid">
            <a class="navbar-brand" href="/">ARPA</a>
            <div class="navbar-title">{% block navbar_title %}{% endblock %}</div>
            <div class="navbar-nav">
                {% block navbar_buttons %}{% endblock %}
            </div>
        </div>
    </nav>
    
    <div class="container mt-4">
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

    <!--<footer class="footer">
        <div class="container text-center">
            <p class="mb-0">© 2025 ARPA</p>
        </div>
    </footer>-->
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"@ | Out-File -FilePath "persons/templates/master.html" -Encoding utf8

    # Create main template
    @"
{% extends "master.html" %}

{% block title %}A R P A{% endblock %}
{% block navbar_title %}{% endblock %}

{% block content %}
    <div class="card">
        <!--<div class="card-header bg-primary text-white">
            <h1 class="mb-0">A R P A</h1>
        </div>-->
        <div class="card-body">
            <!--<h3 class="card-title">Person Management System</h3>-->
            <div class="d-grid gap-3 mt-4">
                <a href="persons/" class="btn btn-primary btn-lg">Personas</a>
                <!--<a href="persons/import/" class="btn btn-success btn-lg">Importar Archivo Excel</a>-->
                <a href="bienesyRentas/" class="btn btn-primary btn-lg">Bienes y Rentas</a>
                <a href="conflictos/" class="btn btn-primary btn-lg">Conflictos de Interes</a>
                <a href="/admin/" class="btn btn-dark btn-lg">Admin ARPA</a>
            </div>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "persons/templates/main.html" -Encoding utf8

    # Create all persons template
    @"
{% extends "master.html" %}

{% block title %}Personas{% endblock %}
{% block navbar_title %}Personas{% endblock %}

{% block navbar_buttons %}
    <a href="/persons/import/" class="nav-link btn btn-primary me-2">Importar Datos</a>
    <!--<a href="/admin/persons/person/add/" class="nav-link btn btn-success">Agregar</a>-->
{% endblock %}

{% block content %}
    
    <div class="table-responsive">
        <table class="table table-striped table-hover">
            <thead class="table-dark">
                <tr>
                    <th>ID</th>
                    <th>Nombre Completo</th>
                    <th>Cargo</th>
                    <th>Cedula</th>
                    <th>Correo</th>
                    <th>Compania</th>
                    <th>Estado</th>
                    <th>Actions</th>
                </tr>
            </thead>
            <tbody>
                {% for person in mypersons %}
                    <tr>
                        <td>{{ person.id }}</td>
                        <td>{{ person.nombre_completo }}</td>
                        <td>{{ person.cargo }}</td>
                        <td>{{ person.cedula }}</td>
                        <td>{{ person.correo }}</td>
                        <td>{{ person.compania }}</td>
                        <td>
                            <span class="badge bg-{% if person.estado == 'Activo' %}success{% else %}danger{% endif %}">
                                {{ person.estado }}
                            </span>
                        </td>
                        <td>
                            <a href="details/{{ person.id }}" class="btn btn-info btn-sm">Ver</a>
                            <a href="/admin/persons/person/{{ person.id }}/change/" class="btn btn-warning btn-sm">Editar</a>
                        </td>
                    </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
{% endblock %}
"@ | Out-File -FilePath "persons/templates/persons.html" -Encoding utf8

    # Create import template
    @"
{% extends "master.html" %}

{% block title %}Importar desde Excel{% endblock %}
{% block navbar_title %}Importar Datos{% endblock %}

{% block navbar_buttons %}
    <!--<a href="/persons/import/" class="nav-link btn btn-primary me-2">Importar Datos</a>-->
    <!--<a href="/admin/persons/person/add/" class="nav-link btn btn-success">Agregar</a>-->
{% endblock %}

{% block content %}
    <div class="card">
        <div class="card-body">
            <form method="post" enctype="multipart/form-data">
                {% csrf_token %}
                <div class="mb-3">
                    <input type="file" class="form-control" id="excel_file" name="excel_file" required>
                    <div class="form-text">El archivo debe incluir las columnas: id, NOMBRE COMPLETO, CARGO, Cedula, Correo, Compania, Estado</div>
                </div>
                <button type="submit" class="btn btn-primary">Importar</button>
                <a href="/persons/" class="btn btn-secondary">Cancelar</a>
            </form>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "persons/templates/import_excel.html" -Encoding utf8

    # Create details template
    @"
{% extends "master.html" %}

{% block title %}Detalles - {{ myperson.nombre_completo }}{% endblock %}
{% block navbar_title %}Detalles: {{ myperson.nombre_completo }}{% endblock %}

{% block navbar_buttons %}
    <a href="/persons/import/" class="nav-link btn btn-primary me-2">Importar Datos</a>
    <a href="/admin/persons/person/add/" class="nav-link btn btn-success">Agregar</a>
{% endblock %}

{% block content %}
    <div class="card">
        <div class="card-header bg-info text-white">
            <h1>{{ myperson.nombre_completo }}</h1>
        </div>
        <div class="card-body">
            <table class="table">
                <tr>
                    <th>ID:</th>
                    <td>{{ myperson.id }}</td>
                </tr>
                <tr>
                    <th>Cargo:</th>
                    <td>{{ myperson.cargo }}</td>
                </tr>
                <tr>
                    <th>Cedula:</th>
                    <td>{{ myperson.cedula }}</td>
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
            </table>
            
            <div class="mt-3">
                <a href="/persons/persons/" class="btn btn-primary">Regresar</a>
                <a href="/admin/persons/person/{{ myperson.id }}/change/" class="btn btn-warning ms-2">Editar</a>
            </div>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "persons/templates/details.html" -Encoding utf8

    # Update settings.py
    $settingsContent = Get-Content -Path ".\arpa\settings.py" -Raw
    $settingsContent = $settingsContent -replace "INSTALLED_APPS = \[", "INSTALLED_APPS = [
    'persons.apps.PersonsConfig',"
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

# Custom admin skin
ADMIN_SITE_HEADER = "ARPA Administration"
ADMIN_SITE_TITLE = "ARPA Admin Portal"
ADMIN_INDEX_TITLE = "Welcome to ARPA Administration"
"@

    # Run migrations
    python manage.py makemigrations persons
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