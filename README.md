# p a n G O

Dajngo framework to analyze a financial historical data.

## 1. Preparation

Execute the main script `run.ps1`. This script installs the dependencies, creates the analysis scripts, and opens the analysis environment in your browser.

```powershell
.\run.ps1
```

## 2. Analysis Execution and Data Visualization

1. Run the script in the terminal:

```
python manage.py runserver
```
2. Load the excel file by clicking on "Load Excel File", enter the password, and generate the table by clicking "Analyze File". The browser will display the data.

3. To filter the data:

- Use the buttons to add, view, reset, and apply filters.

- Save the filtered results to the downloads folder with the "Save Excel" button.

- By clicking on "details", you can view all data per row and save it to Excel.

## 3. Results

After "Analyze File", the `tables/` folder will also contain the analysis results in Excel files, organized in the subfolders `cats/`, `nets/`, and `trends/`. The resulting structure will be similar to the following:     

```
arpa/
├── models/
│   ├── passKey.py
│   ├── server.py
│   ├── cats.py
│   ├── nets.py
│   └── trends.py
├── src/
│   ├── excelFile.xlsx
│   └── data.json
├── tables/
│   ├── cats/
│   ├── nets/
│   └── trends/
├── static/
│   ├── style.css
│   └── script.js
├── favicon.png
├── index.html
├── .gitignore
├── README.md
└── run.ps1