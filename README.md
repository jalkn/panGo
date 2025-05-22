# p a n G O

Dajngo framework to analyze a financial historical data.

## 1. Preparation

Execute the main script `set.ps1`. This script installs the dependencies, creates the analysis scripts, and opens the analysis environment in your browser.

```powershell
.\set.ps1
```

## 2. Analysis Execution and Data Visualization

1. Run the script in the terminal:

```
python manage.py runserver
```
2. Click on import data and import your excel file data.

3. To filter the data:

- Use the buttons to add, view, reset, and apply filters.

- Save the filtered results to the downloads folder with the "Save Excel" button.

- By clicking on "details", you can view all data per row and save it to Excel.

## 3. Results

After "Analyze File", the `core/src/` folder will contain the analysis results in Excel files. The resulting structure will be similar to the following:     

```
arpa/                      # Django project root
├── arpa/                  # Project configuration
│   ├── __init__.py
│   ├── asgi.py
│   ├── settings.py        # Modified with core app and static files
│   ├── urls.py            # Configured with core URLs
│   └── wsgi.py
│
├── core/                  # Main app
│   ├── migrations/        
│   ├── static/            # Static files
│   │   └── core/
│   │       └── css/
│   │           └── style.css
│   │
│   ├── templates/         # HTML templates
│   │   ├── admin/
│   │   │   └── base_site.html
│   │   ├── details.html
│   │   ├── import_excel.html
│   │   ├── master.html
│   │   └── persons.html
│   │
│   ├── src/               # Data storage
│   │
│   ├── admin.py           # Custom admin config
│   ├── apps.py
│   ├── cats.py            # Categories analysis
│   ├── conflicts.py       # Conflict data processor
│   ├── idTrends.py        # Trends analysis
│   ├── inTrends.py        # Trends with conflicts
│   ├── models.py          # Person model
│   ├── nets.py            # Net analysis
│   ├── passKey.py         # Excel password handler
│   ├── trends.py          # Trend analysis
│   ├── urls.py            # App URLs
│   └── views.py           # All view functions
│
├── manage.py
└── db.sqlite3        
```     