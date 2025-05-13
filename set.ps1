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

    # Create models.py with cedula as primary key
    @"
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

    def __str__(self):
        return f"{self.cedula} - {self.nombre_completo}"

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
  
def details(request, cedula):
    myperson = Person.objects.get(cedula=cedula)
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
                    cedula=row['Cedula'],
                    defaults={
                        'nombre_completo': row['NOMBRE COMPLETO'],
                        'cargo': row['CARGO'],
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
    path('persons/details/<str:cedula>', views.details, name='details'),
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
    list_display = ("cedula", "nombre_completo", "cargo", "correo", "compania", "estado")
    search_fields = ("nombre_completo", "cedula")
    list_filter = ("estado", "compania")
    list_per_page = 25
    ordering = ('nombre_completo',)
    actions = [make_active, make_retired]
    
    fieldsets = (
        (None, {
            'fields': ('cedula', 'nombre_completo', 'cargo')
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
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css" integrity="sha512-DTOQO9RWCH3ppGqcWaEA1BIZOC6xxalwEsw9c2QQeAIftl+Vegovlnee1c9QX4TctnWMn13TZye+giMm8e2LwA==" crossorigin="anonymous" referrerpolicy="no-referrer" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {
            margin: 0;
            background-color: white;
        }
        .footer {
            background-color: #343a40;
            color: white;
            padding: 20px 0;
            margin-top: 40px;
        }
        .navbar-title {
            color: black;
            margin-right: auto;
            padding-left: 15px;
            font-size: 1.25rem;
            cursor: pointer;
        }
        .topnav-container {
            display: flex;
            align-items: center;
            padding-right: 40px;
            padding-left: 40px;
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
            font-size: 1.2rem;
        }

        .full-width-container {
            width: 100%;
            padding-right: 10px;
            padding-left: 10px;
            margin-right: auto;
            margin-left: auto;
        }
        
        .card-full-width {
            width: 100%;
            margin-bottom: 20px;
        }
        
        .table-full-width {
            width: 100% !important;
        }


        .btn-custom-primary {
            background-color: #ffffff;
            border-color: #0b00a2;
            color: #090086;
        }
        
        .btn-custom-primary:hover,
        .btn-custom-primary:focus {
            background-color: #090086;
            border-color: #090086;
            color: white;
        }
</style>
    </style>
</head>
<body>
    <div class="topnav-container">
        
        <a href="/" style="text-decoration: none;">
            <div class="logoIN"></div>
        </a>

        <div class="navbar-title">{% block navbar_title %}{% endblock %}</div>
        
        <div class="navbar-nav">
            {% block navbar_buttons %}{% endblock %}
        </div>
    </div>
    
    <div class="container-fluid mt-4 px-4">
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
            <p class="mb-0">Â© 2025 ARPA</p>
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
{% block navbar_title %}Dashboard{% endblock %}

{% block navbar_buttons %}
    <a href="/admin/" class="btn btn-outline-dark btn-lg text-start"><i class="fas fa-cog me-2"></i>Admin</a>
{% endblock %}

{% block content %}
 
    <div class="card-body">
        <div class="d-flex flex-column gap-3 mt-4">
            <a href="persons/" class="btn btn-custom-primary btn-lg text-start"><i class="fas fa-users me-2"></i>Personas</a>
            <a href="bienesyRentas/" class="btn btn-custom-primary btn-lg text-start"><i class="fas fa-building me-2"></i>Bienes y Rentas</a>
            <a href="conflictos/" class="btn btn-custom-primary btn-lg text-start"><i class="fas fa-balance-scale me-2"></i>Conflictos de Interes</a>
            <a href="alertas/" class="btn btn-outline-danger btn-lg text-start"><i class="fas fa-exclamation-triangle me-2"></i>Alertas</a>
        </div>
    </div>


<!-- Font Awesome CDN link -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/js/all.min.js"></script>
{% endblock %}
"@ | Out-File -FilePath "persons/templates/main.html" -Encoding utf8

    # Create all persons template
    @"
{% extends "master.html" %}

{% block title %}Personas{% endblock %}
{% block navbar_title %}Personas{% endblock %}

{% block navbar_buttons %}
    <a href="/persons/import/" class="btn btn-custom-primary btn-lg text-start"><i class="fas fa-database"></i></a>
{% endblock %}

{% block content %}
<div class="card border-0 shadow">  <!-- Added border-0 shadow for better appearance -->
    <div class="card-body p-0">  <!-- Added p-0 to remove inner padding -->
        <div class="table-responsive">
            <table class="table table-striped table-hover mb-0">  <!-- Added mb-0 -->
                <thead class="table-dark">
                    <tr>
                        <th>ID</th>
                        <th>Nombre Completo</th>
                        <th>Cargo</th>
                        <th>Correo</th>
                        <th>Compania</th>
                        <th>Estado</th>
                        <th>Acciones</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in mypersons %}
                        <tr>
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
                                <a href="details/{{ person.cedula }}" class="btn btn-custom-primary btn-lg text-start"><i class="bi bi-person-vcard-fill"></i></a>
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "persons/templates/persons.html" -Encoding utf8

    # Create import template
    @"
{% extends "master.html" %}

{% block title %}Importar desde Excel{% endblock %}
{% block navbar_title %}Importar Datos{% endblock %}

{% block content %}
    <div class="card">
        <div class="card-body">
            <form method="post" enctype="multipart/form-data">
                {% csrf_token %}
                <div class="mb-3">
                    <input type="file" class="form-control" id="excel_file" name="excel_file" required>
                    <div class="form-text">El archivo Excel debe incluir las columnas: Id, NOMBRE COMPLETO, CARGO, Cedula, Correo, Compania, Estado</div>
                </div>
                <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar</button>
                <a href="/persons/persons/" class="btn btn-custom-primary btn-lg text-start">Cancelar</a>
            </form>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "persons/templates/import_excel.html" -Encoding utf8

    # Create details template
    @"
{% extends "master.html" %}

{% block title %}Detalles - {{ myperson.nombre_completo }}{% endblock %}
{% block navbar_title %}{{ myperson.nombre_completo }}{% endblock %}

{% block content %}
    <div class="card">
        <div class="card-body">
            <table class="table">
                <tr>
                    <th>ID:</th>
                    <td>{{ myperson.cedula }}</td>
                </tr>
                <tr>
                    <th>Nombre Completo:</th>
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
            </table>
            
            <div class="mt-3">
                <a href="/persons/persons/" class="btn btn-custom-primary btn-lg text-start">Regresar</a>
                <a href="/admin/persons/person/{{ myperson.cedula }}/change/" class="btn btn-custom-primary btn-lg text-start">Editar</a>
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
        cedula=row['Cedula'],
        defaults={
            'nombre_completo': row['NOMBRE COMPLETO'],
            'cargo': row['CARGO'],
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