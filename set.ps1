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

    # Create core app (changed from persons)
    python manage.py startapp core

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
"@ | Out-File -FilePath "core/models.py" -Encoding utf8

    # Create views.py with import functionality
    @"
from django.http import HttpResponse, HttpResponseRedirect
from django.template import loader
from django.shortcuts import render
from .models import Person
import pandas as pd
from django.contrib import messages
from django.core.paginator import Paginator
from django.db.models import Q
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def persons(request):
    persons = Person.objects.all().order_by('nombre_completo')
    
    # Get filter parameters from request
    search_query = request.GET.get('q', '')
    status_filter = request.GET.get('status', '')
    nombre_filter = request.GET.get('nombre', '')
    cargo_filter = request.GET.get('cargo', '')
    compania_filter = request.GET.get('compania', '')
    order_by = request.GET.get('order_by', 'nombre_completo')
    
    # Apply filters if they exist
    if search_query:
        persons = persons.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(compania__icontains=search_query) |
            Q(cargo__icontains=search_query))
    
    if status_filter:
        persons = persons.filter(estado=status_filter)
        
    if nombre_filter:
        persons = persons.filter(nombre_completo__icontains=nombre_filter)
        
    if cargo_filter:
        persons = persons.filter(cargo__icontains=cargo_filter)
        
    if compania_filter:
        persons = persons.filter(compania__icontains=compania_filter)
    
    # Apply ordering
    if order_by in ['cedula', 'nombre_completo', 'cargo', 'compania']:
        persons = persons.order_by(order_by)
    
    context = {
        'persons': persons,
    }
    return render(request, 'main.html', context)

def details(request, cedula):
    myperson = Person.objects.get(cedula=cedula)
    return render(request, 'details.html', {'myperson': myperson})
  
def main(request):
    # Get all persons with filters and pagination
    persons = Person.objects.all()
    
    # Get filter parameters from request
    search_query = request.GET.get('q', '')
    status_filter = request.GET.get('status', '')
    nombre_filter = request.GET.get('nombre', '')
    cargo_filter = request.GET.get('cargo', '')
    compania_filter = request.GET.get('compania', '')
    order_by = request.GET.get('order_by', 'nombre_completo')
    
    # Apply filters if they exist
    if search_query:
        persons = persons.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(compania__icontains=search_query) |
            Q(cargo__icontains=search_query))
    
    if status_filter:
        persons = persons.filter(estado=status_filter)
        
    if nombre_filter:
        persons = persons.filter(nombre_completo__icontains=nombre_filter)
        
    if cargo_filter:
        persons = persons.filter(cargo__icontains=cargo_filter)
        
    if compania_filter:
        persons = persons.filter(compania__icontains=compania_filter)
    
    # Apply ordering
    if order_by in ['cedula', 'nombre_completo', 'cargo', 'compania']:
        persons = persons.order_by(order_by)
    
    # Get unique values for dropdown filters
    cargos = Person.objects.values_list('cargo', flat=True).distinct().order_by('cargo')
    companias = Person.objects.values_list('compania', flat=True).distinct().order_by('compania')
    
    # Pagination
    paginator = Paginator(persons, 1000) 
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    return render(request, 'main.html', {
        'page_obj': page_obj,
        'persons': page_obj.object_list,
        'persons_count': persons.count(),
        'cargos': cargos,
        'companias': companias,
        'current_order': order_by,
    })

def import_from_excel(request):
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        try:
            df = pd.read_excel(excel_file)
            
            # Convert column names to match your model
            column_mapping = {
                'Cedula': 'cedula',
                'NOMBRE COMPLETO': 'nombre_completo',
                'CARGO': 'cargo',
                'Correo': 'correo',
                'Compania': 'compania',
                'Estado': 'estado'
            }
            df.rename(columns=column_mapping, inplace=True)
            
            # Handle missing or null values
            df.fillna('', inplace=True)
            
            # Process each row
            for _, row in df.iterrows():
                Person.objects.update_or_create(
                    cedula=row['cedula'],
                    defaults={
                        'nombre_completo': row['nombre_completo'],
                        'cargo': row['cargo'],
                        'correo': row['correo'],
                        'compania': row['compania'],
                        'estado': row['estado'] if row['estado'] in ['Activo', 'Retirado'] else 'Activo'
                    }
                )
            
            messages.success(request, f'Carga exitosa! {len(df)} filas procesadas.')
        except Exception as e:
            messages.error(request, f'Error importing data: {str(e)}')
        
        return HttpResponseRedirect('/')
    
    return render(request, 'import_excel.html')

def export_to_excel(request):
    # Get filtered queryset using the same filters as the main view
    persons = Person.objects.all()
    
    # Apply the same filters as in the main view
    search_query = request.GET.get('q', '')
    status_filter = request.GET.get('status', '')
    nombre_filter = request.GET.get('nombre', '')
    cargo_filter = request.GET.get('cargo', '')
    compania_filter = request.GET.get('compania', '')
    order_by = request.GET.get('order_by', 'nombre_completo')
    
    if search_query:
        persons = persons.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(compania__icontains=search_query) |
            Q(cargo__icontains=search_query))
    
    if status_filter:
        persons = persons.filter(estado=status_filter)
        
    if nombre_filter:
        persons = persons.filter(nombre_completo__icontains=nombre_filter)
        
    if cargo_filter:
        persons = persons.filter(cargo__icontains=cargo_filter)
        
    if compania_filter:
        persons = persons.filter(compania__icontains=compania_filter)
    
    if order_by in ['cedula', 'nombre_completo', 'cargo', 'compania']:
        persons = persons.order_by(order_by)

    # Create a workbook and add a worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Personas"

    # Add headers
    headers = ["Cedula", "Nombre Completo", "Cargo", "Correo", "Compania", "Estado"]
    for col_num, header in enumerate(headers, 1):
        col_letter = get_column_letter(col_num)
        ws[f"{col_letter}1"] = header

    # Add data
    for row_num, person in enumerate(persons, 2):
        ws[f"A{row_num}"] = person.cedula
        ws[f"B{row_num}"] = person.nombre_completo
        ws[f"C{row_num}"] = person.cargo
        ws[f"D{row_num}"] = person.correo
        ws[f"E{row_num}"] = person.compania
        ws[f"F{row_num}"] = person.estado

    # Prepare response
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=personas.xlsx'
    
    # Save workbook to response
    wb.save(response)
    return response
"@ | Out-File -FilePath "core/views.py" -Encoding utf8

    # Create urls.py for core app
    @"
from django.urls import path
from . import views

urlpatterns = [
    path('', views.main, name='main'),
    path('persons/details/<str:cedula>/', views.details, name='details'),
    path('persons/import/', views.import_from_excel, name='import_excel'),
    path('persons/export/', views.export_to_excel, name='export_excel'),  # Add this line
]
"@ | Out-File -FilePath "core/urls.py" -Encoding utf8

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
"@ | Out-File -FilePath "core/admin.py" -Encoding utf8

    # Update project urls.py with proper admin configuration
    @"
from django.contrib import admin
from django.urls import include, path

# Customize default admin interface
admin.site.site_header = 'A R P A'
admin.site.site_title = 'ARPA Admin Portal'
admin.site.index_title = 'Bienvenido a A R P A'

urlpatterns = [
    path('persons/', include('core.urls')),
    path('admin/', admin.site.urls),
    path('', include('core.urls')), 
]
"@ | Out-File -FilePath "arpa/urls.py" -Encoding utf8

    # Create templates directory structure
    $directories = @(
        "core/templates",
        "core/templates/admin",
        "core/templates/admin/core"
    )
    foreach ($dir in $directories) {
        New-Item -Path $dir -ItemType Directory -Force
    }

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
    <style>
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
    </style>
</head>
<body>
    <div class="topnav-container">
        <a href="/" style="text-decoration: none;">
            <div class="logoIN"></div>
        </a>
        <div class="navbar-title">{% block navbar_title %}ARPA{% endblock %}</div>
            <!-- Update the navbar-buttons section in master.html -->
            <div class="navbar-buttons">
                {% block navbar_buttons %}
                <a href="/admin/" class="btn btn-outline-dark" title="Admin">
                    <i class="fas fa-wrench"></i>
                </a>
                <a href="/persons/import/" class="btn btn-custom-primary" title="Importar">
                    <i class="fas fa-database"></i>
                </a>
                <a href="/persons/export/?{{ request.GET.urlencode }}" class="btn btn-custom-primary" title="Exportar a Excel">
                    <i class="fas fa-file-excel"></i>
                </a>
                {% endblock %}
            </div>
    </div>
    
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

    # Create main template
    @"
{% extends "master.html" %}

{% block title %}A R P A{% endblock %}
{% block navbar_title %}A R P A{% endblock %}

{% block navbar_buttons %}
<a href="/admin/" class="btn btn-outline-dark btn-lg text-start" title="Admin Panel">
    <i class="fas fa-wrench"></i>
</a>
<a href="/persons/import/" class="btn btn-custom-primary btn-lg text-start" title="Import Data">
    <i class="fas fa-database"></i>
</a>
<a href="/persons/export/?{{ request.GET.urlencode }}" class="btn btn-custom-primary btn-lg text-start" title="Export to Excel">
    <i class="fas fa-file-excel"></i>
</a>
{% endblock %}

{% block content %}
<!-- Search Form -->
<div class="card mb-4 border-0 shadow">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <!-- General Search -->
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar..." 
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
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1">Filtrar</button>
                <a href="." class="btn btn-outline-secondary btn-lg flex-grow-1">Limpiar</a>
            </div>
        </form>
    </div>
</div>

<!-- Persons Table -->
<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-dark">
                    <tr>
                        <th><a href="?order_by=cedula{% for key, value in request.GET.items %}{% if key != 'order_by' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">ID {% if current_order == 'cedula' %}<i class="fas fa-sort-up"></i>{% endif %}</a></th>
                        <th><a href="?order_by=nombre_completo{% for key, value in request.GET.items %}{% if key != 'order_by' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">Nombre {% if current_order == 'nombre_completo' %}<i class="fas fa-sort-up"></i>{% endif %}</a></th>
                        <th><a href="?order_by=cargo{% for key, value in request.GET.items %}{% if key != 'order_by' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">Cargo {% if current_order == 'cargo' %}<i class="fas fa-sort-up"></i>{% endif %}</a></th>
                        <th>Correo</th>
                        <th><a href="?order_by=compania{% for key, value in request.GET.items %}{% if key != 'order_by' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">Compania {% if current_order == 'compania' %}<i class="fas fa-sort-up"></i>{% endif %}</a></th>
                        <th>Estado</th>
                        <th>Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
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
                                <a href="/persons/details/{{ person.cedula }}/" 
                                   class="btn btn-custom-primary btn-sm"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                    {% empty %}
                        <tr>
                            <td colspan="7" class="text-center py-4">
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
"@ | Out-File -FilePath "core/templates/main.html" -Encoding utf8

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
                    <input type="file" class="form-control" id="excel_file" name="excel_file" required title="Seleccionar archivo">
                    <div class="form-text">El archivo Excel de Personas debe incluir las columnas: Id, NOMBRE COMPLETO, CARGO, Cedula, Correo, Compania, Estado</div>
                </div>
                <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar Personas</button>
                <!--<a href="/" class="btn btn-custom-primary btn-lg text-start">Cancelar</a>-->
            </form>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/import_excel.html" -Encoding utf8

    # Create details template
    @"
{% extends "master.html" %}

{% block title %}Detalles - {{ myperson.nombre_completo }}{% endblock %}
{% block navbar_title %}{{ myperson.nombre_completo }}{% endblock %}

{% block navbar_buttons %}
    <a href="/admin/" class="btn btn-outline-dark" title="Admin">
        <i class="fas fa-wrench"></i>
    </a>
    <a href="/persons/import/" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-database"></i>
    </a>
    <a href="/persons/export/?q={{ myperson.cedula }}" class="btn btn-custom-primary" title="Exportar a Excel">
        <i class="fas fa-file-excel"></i>
    </a>
{% endblock %}

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
                <a href="/" class="btn btn-custom-primary">Regresar</a>
                <a href="/admin/core/person/{{ myperson.cedula }}/change/" class="btn btn-custom-primary">
                    <i class="fas fa-pencil-alt me-2"></i> Editar
                </a>
            </div>
        </div>
    </div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/details.html" -Encoding utf8

    # Update settings.py
    $settingsContent = Get-Content -Path ".\arpa\settings.py" -Raw
    $settingsContent = $settingsContent -replace "INSTALLED_APPS = \[", "INSTALLED_APPS = [
    'core.apps.CoreConfig',"
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
ADMIN_SITE_HEADER = "A R P A"
ADMIN_SITE_TITLE = "ARPA Admin Portal"
ADMIN_INDEX_TITLE = "Bienvenido a A R P A"
"@

    # Run migrations
    python manage.py makemigrations core
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

from core.models import Person

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