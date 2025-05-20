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

    # Create core app
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
    revisar = models.BooleanField(default=False, verbose_name="Revisar")
    comments = models.TextField(blank=True, null=True, verbose_name="Comentarios")

    def __str__(self):
        return f"{self.cedula} - {self.nombre_completo}"

    class Meta:
        verbose_name = "Persona"
        verbose_name_plural = "Personas"
"@ | Out-File -FilePath "core/models.py" -Encoding utf8 -Force

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

def _apply_filters_and_sorting(queryset, request_params):
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

def _get_dropdown_values():
    """
    Helper function to get distinct values for dropdown filters
    """
    return {
        'cargos': Person.objects.values_list('cargo', flat=True).distinct().order_by('cargo'),
        'companias': Person.objects.values_list('compania', flat=True).distinct().order_by('compania'),
        'estados': Person.objects.values_list('estado', flat=True).distinct().order_by('estado'),
    }

def main(request):
    """
    Main view showing the list of persons with filtering and pagination
    """
    persons = Person.objects.all()
    persons = _apply_filters_and_sorting(persons, request.GET)
    
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

def details(request, cedula):
    """
    View showing details for a single person
    """
    myperson = Person.objects.get(cedula=cedula)
    return render(request, 'details.html', {'myperson': myperson})

def import_from_excel(request):
    """
    View for importing data from Excel files
    """
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        try:
            df = pd.read_excel(excel_file)
            
            column_mapping = {
                'Cedula': 'cedula',
                'NOMBRE COMPLETO': 'nombre_completo',
                'CARGO': 'cargo',
                'Correo': 'correo',
                'Compania': 'compania',
                'Estado': 'estado'
            }
            df.rename(columns=column_mapping, inplace=True)
            df.fillna('', inplace=True)
            
            for _, row in df.iterrows():
                Person.objects.update_or_create(
                    cedula=row['cedula'],
                    defaults={
                        'nombre_completo': row['nombre_completo'],
                        'cargo': row['cargo'],
                        'correo': row['correo'],
                        'compania': row['compania'],
                        'estado': row['estado'] if row['estado'] in ['Activo', 'Retirado'] else 'Activo',
                        'revisar': row.get('revisar', False),
                        'comments': row.get('comments', ''),
                    }
                )
            
            messages.success(request, f'Carga exitosa! {len(df)} filas procesadas.')
        except Exception as e:
            messages.error(request, f'Error importing data: {str(e)}')
        
        return HttpResponseRedirect('/')
    
    return render(request, 'import_excel.html')

def export_to_excel(request):
    """
    View for exporting filtered data to Excel
    """
    persons = Person.objects.all()
    persons = _apply_filters_and_sorting(persons, request.GET)

    wb = Workbook()
    ws = wb.active
    ws.title = "Personas"

    headers = ["Cedula", "Nombre Completo", "Cargo", "Correo", "Compania", "Estado", "Revisar", "Comentarios"]
    for col_num, header in enumerate(headers, 1):
        col_letter = get_column_letter(col_num)
        ws[f"{col_letter}1"] = header

    for row_num, person in enumerate(persons, 2):
        ws[f"A{row_num}"] = person.cedula
        ws[f"B{row_num}"] = person.nombre_completo
        ws[f"C{row_num}"] = person.cargo
        ws[f"D{row_num}"] = person.correo
        ws[f"E{row_num}"] = person.compania
        ws[f"F{row_num}"] = person.estado
        ws[f"G{row_num}"] = "Sí" if person.revisar else "No"
        ws[f"H{row_num}"] = person.comments or ""

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=personas.xlsx'
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
    
admin.site.register(Person, PersonAdmin)
"@ | Out-File -FilePath "core/admin.py" -Encoding utf8 -Force

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
        "core/static",
        "core/static/core",
        "core/static/core/css",
        "core/templates",
        "core/templates/admin"
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
    {% load static %}
    <link rel="stylesheet" href="{% static 'core/css/style.css' %}">
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
                <a href="/persons/export/?q={{ myperson.cedula }}" class="btn btn-custom-primary btn-my-green" title="Exportar a Excel">  
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
"@ | Out-File -FilePath "core/static/core/css/style.css" -Encoding utf8

    # Create persons template
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
<a href="/persons/export/?q={{ myperson.cedula }}" class="btn btn-custom-primary btn-my-green" title="Exportar a Excel">  
    <i class="fas fa-file-excel"></i>
</a>
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
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<!-- Persons Table -->
<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive">
            <table class="table table-striped table-hover mb-0">
                <thead>
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
                        <th style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
                        <tr {% if person.revisar %}class="table-warning"{% endif %}>
                            <td>
                                <a href="/admin/core/person/{{ person.cedula }}/change/" style="text-decoration: none;" title="{% if person.revisar %}Marcado para revisión{% else %}No marcado{% endif %}">
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

    # Create import template
    @"
{% extends "master.html" %}

{% block title %}Importar desde Excel{% endblock %}
{% block navbar_title %}Importar Datos{% endblock %}

{% block navbar_buttons %}
<a href="/" class="btn btn-custom-primary"><i class="fas fa-arrow-right"></i></a>
{% endblock %}

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
<a href="/admin/core/person/{{ myperson.cedula }}/change/" class="btn btn-outline-dark" title="Admin">
    <i class="fas fa-wrench"></i>
</a>
<a href="/persons/export/?q={{ myperson.cedula }}" class="btn btn-custom-primary btn-my-green" title="Exportar a Excel">  
    <i class="fas fa-file-excel"></i>
</a>
<a href="/" class="btn btn-custom-primary"><i class="fas fa-arrow-right"></i></a>
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
                <tr>
                    <th>Por revisar:</th>
                    <td>
                        {% if myperson.revisar %}
                            <span class="badge bg-warning text-dark">Sí</span>
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
STATICFILES_DIRS = [
    BASE_DIR / "core/static",
]

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

    python manage.py collectstatic --noinput

    # Start the server
    Write-Host "🚀 Starting Django development server..." -ForegroundColor $GREEN
    python manage.py runserver
}

migratoDjango