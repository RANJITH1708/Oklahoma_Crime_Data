# Just-Do-It
# Oklahoma Crime Data Analysis Project Report

## Introduction

The Oklahoma Crime Data Analysis Project investigates crime trends across Oklahoma cities from 2014 to 2023. The project integrates crime statistics, geographic coordinates, and labor data to standardize inconsistent datasets, establish a robust MySQL database, impute missing values, and derive actionable insights through correlation analysis, predictive modeling, and interactive visualizations in Power BI. The tools utilized include VBA, Python, SQL, and Power BI, enabling comprehensive data cleaning, organization, analysis, and presentation.

## Data Collection and Preparation

### Data Sources

The crime data was sourced from the FBI Crime Data Explorer, specifically the National Incident-Based Reporting System (NIBRS) Tables, which provide incident-based crime data for Oklahoma. The dataset was initially downloaded as a PDF from [https://cde.ucr.cjis.gov/LATEST/webapp/#](https://cde.ucr.cjis.gov/LATEST/webapp/#) and subsequently converted into Excel format for processing.

### Initial Data Cleaning

Upon conversion to Excel, the dataset exhibited several issues:

- **Merged Columns**: The Excel sheet contained merged columns, which were manually corrected to ensure proper data structure.
- **Missing Data for Universities**: Data for universities (Carl Albert State College, Oklahoma State University, University of Oklahoma, and Sallisaw) was missing. These rows were removed from the dataset, as they lacked sufficient information for analysis.

### Challenges

Further examination revealed inconsistencies across the annual datasets (2014–2023):

- **Missing Columns**:
  - `animal_cruelty`, `identity_theft`, and `hacking_computer_invasion` were absent in 2014 and 2015.
  - `sex_offenses_non_forcible` was missing in 2018, 2019, and from 2020 to 2023.
- **Renamed Columns**:
  - In 2023, the `fondling` offense was renamed to `criminal_sexual_contact`. To maintain consistency across years, it was reverted to `fondling`, despite slight definitional differences: `fondling` refers to intentional, non-consensual touching of intimate areas for sexual gratification, often categorized as a forcible sex offense, whereas `criminal_sexual_contact` is a broader legal term that may include `fondling` but also other forms of non-consensual sexual touching without penetration.

### Standardization Process

To address these discrepancies, a VBA script was developed in Microsoft Excel to compare column names and their order across workbooks, using the 2014 dataset as the reference.

#### VBA Script for Column Comparison

####vba

Sub CompareColumnNamesAndOrder()
    Dim refWorkbook As Workbook, compareWorkbook As Workbook
    Dim refSheet As Worksheet, compareSheet As Worksheet
    Dim refColNames As Object, compareColNames As Object
    Dim refFile As String, compareFiles As Variant
    Dim lastCol As Integer, i As Integer, report As String
    ' Select reference workbook
    refFile = Application.GetOpenFilename("Excel Files (*.xls;*.xlsm), *.xlsx;*.xlsm", , "Select Reference Workbook")
    If refFile = "False" Then Exit Sub
    
  Set refWorkbook = Workbooks.Open(refFile)
    Set refSheet = refWorkbook.Sheets(1)

' Store reference column names and order
    Set refColNames = CreateObject("Scripting.Dictionary")
    lastCol = refSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        refColNames(refSheet.Cells(1, i).Value) = i
    Next i

  ' Select comparison workbooks
    compareFiles = Application.GetOpenFilename("Excel Files (*.xls;*.xlsm), *.xlsx;*.xlsm", , "Select Workbooks to Compare", , True)
    If Not IsArray(compareFiles) Then Exit Sub

  report = "Column Differences Report:" & vbNewLine & "==============================" & vbNewLine

  ' Compare each workbook
    For Each compareFile In compareFiles
        Set compareWorkbook = Workbooks.Open(compareFile)
        Set compareSheet = compareWorkbook.Sheets(1)

  Set compareColNames = CreateObject("Scripting.Dictionary")
        lastCol = compareSheet.Cells(1, Columns.Count).End(xlToLeft).Column
        For i = 1 To lastCol
            compareColNames(compareSheet.Cells(1, i).Value) = i
        Next i

  Dim colDiff As String, orderDiff As String
        colDiff = ""
        orderDiff = ""

  ' Identify missing columns
        For Each Key In compareColNames.Keys
            If Not refColNames.Exists(Key) Then
                colDiff = colDiff & Key & " (Column: " & compareColNames(Key) & "), "
            End If
        Next Key

  ' Check column order
        For Each Key In refColNames.Keys
            If compareColNames.Exists(Key) Then
                If refColNames(Key) <> compareColNames(Key) Then
                    orderDiff = orderDiff & Key & " (Ref: " & refColNames(Key) & " -> Found: " & compareColNames(Key) & "), "
                End If
            End If
        Next Key

  ' Build report
        If colDiff <> "" Or orderDiff <> "" Then
            report = report & "File: " & compareWorkbook.Name & vbNewLine
            If colDiff <> "" Then report = report & "Missing Columns: " & Left(colDiff, Len(colDiff) - 2) & vbNewLine
            If orderDiff <> "" Then report = report & "Order Mismatches: " & Left(orderDiff, Len(orderDiff) - 2) & vbNewLine
            report = report & "------------------------------" & vbNewLine
        End If

  compareWorkbook.Close False
    Next compareFile

  refWorkbook.Close False

  ' Display results
    If report = "Column Differences Report:" & vbNewLine & "==============================" & vbNewLine Then
        MsgBox "All workbooks match the reference!", vbInformation
    Else
        MsgBox report, vbExclamation, "Column Differences Found"
    End If
End Sub


#  Geographic Coordinates
## To facilitate spatial analysis in Power BI, coordinates for Oklahoma cities were retrieved using the LocationIQ API and saved as a CSV file.

## Python Script for Fetching Coordinates
import requests
import csv
import time
from google.colab import files

API_KEY = "******************************"

cities = [
    "Achille", "Ada", "Adair", "Alex", "Allen", "Altus", "Alva", "Amber", "Anadarko", "Antlers",
    # Full list includes 300+ cities (abridged for brevity)
    "Wynnewood", "Wynona", "Yale", "Yukon"
]

def get_coordinates(city):
    url = f"https://us1.locationiq.com/v1/search.php?key={API_KEY}&q={city}, Oklahoma, USA&format=json"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data:
            return data[0]["lat"], data[0]["lon"]
    return None, None

output_file = "/content/oklahoma_city_coordinates.csv"
with open(output_file, "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["City", "Latitude", "Longitude"])
    for city in cities:
        lat, lon = get_coordinates(city)
        writer.writerow([city, lat, lon])
        print(f"{city}: {lat}, {lon}")
        time.sleep(1)

files.download(output_file)

#### Database Design and Setup
A MySQL database named oklahoma_crime_data was created to efficiently store the standardized data.

USE oklahoma_crime_data;

CREATE TABLE cities (
    city_id INT AUTO_INCREMENT PRIMARY KEY,
    city_name VARCHAR(255) NOT NULL UNIQUE,
    latitude DECIMAL(9,6),
    longitude DECIMAL(9,6)
);

CREATE TABLE years (
    year INT PRIMARY KEY
);

INSERT INTO years (year) VALUES
(2014), (2015), (2016), (2017), (2018),
(2019), (2020), (2021), (2022), (2023);

CREATE TABLE crimedata (
    crime_id INT AUTO_INCREMENT PRIMARY KEY,
    city_id INT,
    year INT,
    population INT,
    total_offenses INT,
    crimes_against_persons INT,
    crimes_against_property INT,
    crimes_against_society INT,
    assault_offenses INT,
    aggravated_assault INT,
    simple_assault INT,
    intimidation INT,
    homicide_offenses INT,
    murder_and_nonnegligent_manslaughter INT,
    negligent_manslaughter INT,
    justifiable_homicide INT,
    human_trafficking_offenses INT,
    commercial_sex_acts INT,
    involuntary_servitude INT,
    kidnapping_abduction INT,
    sex_offenses INT,
    rape INT,
    sodomy INT,
    sexual_assault_with_an_object INT,
    fondling INT,
    sex_offenses_non_forcible INT,
    incest INT,
    statutory_rape INT,
    arson INT,
    bribery INT,
    burglary_breaking_entering INT,
    counterfeiting_forgery INT,
    destruction_damage_vandalism_of_property INT,
    embezzlement INT,
    extortion_blackmail INT,
    fraud_offenses INT,
    false_pretenses_swindle_confidence_game INT,
    credit_card_automated_teller_machine_fraud INT,
    impersonation INT,
    welfare_fraud INT,
    wire_fraud INT,
    identity_theft INT,
    hacking_computer_invasion INT,
    larceny_theft_offenses INT,
    pocket_picking INT,
    purse_snatching INT,
    shop_lifting INT,
    theft_from_building INT,
    theft_from_coin_operated_machine_or_device INT,
    theft_from_motor_vehicle INT,
    theft_of_motor_vehicle_parts_or_accessories INT,
    all_other_larceny INT,
    motor_vehicle_theft INT,
    robbery INT,
    stolen_property_offenses INT,
    animal_cruelty INT,
    drug_narcotic_offenses INT,
    drug_narcotic_violations INT,
    drug_equipment_violations INT,
    gambling_offenses INT,
    betting_wagering INT,
    operating_promoting_assisting_gambling INT,
    gambling_equipment_violations INT,
    sports_tampering INT,
    pornography_obscene_material INT,
    prostitution_offenses INT,
    prostitution INT,
    assisting_or_promoting_prostitution INT,
    purchasing_prostitution INT,
    weapon_law_violations INT,
    FOREIGN KEY (city_id) REFERENCES cities(city_id),
    FOREIGN KEY (year) REFERENCES years(year)
);

CREATE TABLE labor_statistics (
    year INT PRIMARY KEY,
    labor_force INT,
    employment INT,
    unemployment INT,
    unemployment_rate DECIMAL(3,1),
    FOREIGN KEY (year) REFERENCES years(year)
);

INSERT INTO labor_statistics (year, labor_force, employment, unemployment, unemployment_rate) VALUES
(2014, 1797769, 1719826, 77943, 4.3),
(2015, 1829046, 1750501, 78545, 4.3),
(2016, 1828055, 1743225, 84830, 4.6),
(2017, 1826272, 1752733, 73539, 4.0),
(2018, 1931256, 1771251, 60005, 3.3),
(2019, 1843331, 1785400, 57931, 3.1),
(2020, 1842829, 1726786, 116043, 6.3),
(2021, 1866370, 1791369, 75001, 4.0),
(2022, 1903508, 1845328, 58180, 3.1),
(2023, 1963240, 1899949, 63291, 3.2);

CREATE USER 'proot'@'%' IDENTIFIED BY 'I********5';
GRANT ALL PRIVILEGES ON oklahoma_crime_data.* TO 'proot'@'%';
FLUSH PRIVILEGES;


##### Data Loading
Data was imported into the database using LOAD DATA INFILE commands, ensuring proper formatting:

LOAD DATA INFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/oklahoma_cities.csv'
INTO TABLE cities
FIELDS TERMINATED BY ','
ENCLOSED BY '"'
LINES TERMINATED BY '\n'
IGNORE 1 LINES
(city_name, latitude, longitude);



#### Remote Access
The database was connected to Google Colab using ngrok for remote access and analysis:

import pymysql

connection = pymysql.connect(
    host="4.tcp.ngrok.io",
    port=16944,
    user="proot",
    password="I*********5",
    database="oklahoma_crime_data"
)
print("Connection successful!")



###### Data Imputation
To address missing data, a new table, crimedata_imputed, was created by imputing values for cities with at least seven years of data.

Methodology
Filtering: Identified cities with seven or more years of data.
Imputation Process:
Replaced zeros with NaN in crime-related columns (excluding population).
Applied linear interpolation across years for each numeric column.
Filled remaining gaps with the city-specific mean for each column, defaulting to 0 if the mean was unavailable.
Python Script for Imputation


import pymysql
import pandas as pd
import numpy as np

connection = pymysql.connect(host="4.tcp.ngrok.io", port=16944, user="proot", password="I**********5", database="oklahoma_crime_data")
cursor = connection.cursor()
cursor.execute("CREATE TABLE crimedata_imputed LIKE crimedata;")

df = pd.read_sql("SELECT * FROM crimedata", connection)
all_years = pd.read_sql("SELECT year FROM years", connection)["year"].tolist()

years_per_city = df.groupby("city_id")["year"].nunique()
cities_with_7_plus_years = years_per_city[years_per_city >= 7].index.tolist()
df_filtered = df[df["city_id"].isin(cities_with_7_plus_years)].copy()

numeric_columns = df_filtered.select_dtypes(include=[np.number]).columns.tolist()
columns_to_impute = [col for col in numeric_columns if col not in ["crime_id", "city_id", "year"]]
crime_columns = [col for col in columns_to_impute if col != "population"]
for col in crime_columns:
    df_filtered[col] = df_filtered[col].replace(0, np.nan)

df_imputed_full = pd.DataFrame()
for city_id in cities_with_7_plus_years:
    city_data = df_filtered[df_filtered["city_id"] == city_id].copy()
    all_years_df = pd.DataFrame({"year": all_years, "city_id": city_id})
    merged_data = pd.merge(all_years_df, city_data, on=["city_id", "year"], how="left")
    merged_data = merged_data.sort_values("year")
    for col in columns_to_impute:
        merged_data[col] = merged_data[col].interpolate(method="linear", limit_direction="both")
        if merged_data[col].isna().any():
            city_mean = merged_data[col].mean()
            merged_data[col] = merged_data[col].fillna(city_mean if not np.isnan(city_mean) else 0)
        merged_data[col] = merged_data[col].fillna(0)
    df_imputed_full = pd.concat([df_imputed_full, merged_data], ignore_index=True)

df_imputed_full = df_imputed_full.drop(columns=["crime_id"])
columns = df_imputed_full.columns.tolist()
insert_query = f"INSERT INTO crimedata_imputed ({', '.join(columns)}) VALUES ({', '.join(['%s'] * len(columns))})"
cursor.executemany(insert_query, df_imputed_full.values.tolist())
connection.commit()
connection.close()


### Data Analysis
Correlation Analysis
Using the crimedata_imputed table, correlations between population and crime variables were calculated:

Strong Correlations:
crimes_against_property: 0.966
larceny_theft_offenses: 0.961
fraud_offenses: 0.950
Weak Correlations:
purchasing_prostitution: 0.040
involuntary_servitude: -0.019
betting_wagering: -0.018
NaN Values: sports_tampering resulted in NaN due to zero variance (all values likely 0), triggering a NumPy warning (RuntimeWarning: invalid value encountered in divide).

### Predictive Modeling
Predictive models were trained to forecast population using crime variety variables, excluding total_offenses, crime_rate, year, latitude, and longitude.

Python Script for Correlation and Predictive Modeling


!pip install pymysql pandas scikit-learn matplotlib seaborn
!pip install sqlalchemy

from sqlalchemy import create_engine
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.linear_model import LinearRegression
from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.svm import SVR
from sklearn.neural_network import MLPRegressor
from sklearn.metrics import r2_score, mean_squared_error
import numpy as np

engine = create_engine(
    f"mysql+pymysql://proot:Ilovesql%402025@4.tcp.ngrok.io:16944/oklahoma_crime_data?connect_timeout=60&read_timeout=60"
)

crime_df = pd.read_sql("SELECT * FROM crimedata_imputed", engine)
crime_df = crime_df.drop(columns=['crime_id'])
print("Missing values:\n", crime_df.isnull().sum())
crime_df = crime_df.fillna(0)

crime_columns = [
    'crimes_against_persons', 'crimes_against_property', 'crimes_against_society',
    'assault_offenses', 'aggravated_assault', 'simple_assault', 'intimidation', 'homicide_offenses',
    'murder_and_nonnegligent_manslaughter', 'negligent_manslaughter', 'justifiable_homicide',
    'human_trafficking_offenses', 'commercial_sex_acts', 'involuntary_servitude', 'kidnapping_abduction',
    'sex_offenses', 'rape', 'sodomy', 'sexual_assault_with_an_object', 'fondling', 'sex_offenses_non_forcible',
    'incest', 'statutory_rape', 'arson', 'bribery', 'burglary_breaking_entering', 'counterfeiting_forgery',
    'destruction_damage_vandalism_of_property', 'embezzlement', 'extortion_blackmail', 'fraud_offenses',
    'false_pretenses_swindle_confidence_game', 'credit_card_automated_teller_machine_fraud', 'impersonation',
    'welfare_fraud', 'wire_fraud', 'identity_theft', 'hacking_computer_invasion', 'larceny_theft_offenses',
    'pocket_picking', 'purse_snatching', 'shop_lifting', 'theft_from_building',
    'theft_from_coin_operated_machine_or_device', 'theft_from_motor_vehicle',
    'theft_of_motor_vehicle_parts_or_accessories', 'all_other_larceny', 'motor_vehicle_theft', 'robbery',
    'stolen_property_offenses', 'animal_cruelty', 'drug_narcotic_offenses', 'drug_narcotic_violations',
    'drug_equipment_violations', 'gambling_offenses', 'betting_wagering', 'operating_promoting_assisting_gambling',
    'gambling_equipment_violations', 'sports_tampering', 'pornography_obscene_material', 'prostitution_offenses',
    'prostitution', 'assisting_or_promoting_prostitution', 'purchasing_prostitution', 'weapon_law_violations'
]

corr_columns = [col for col in crime_columns if col in crime_df.columns]
correlation_matrix = crime_df[corr_columns].corrwith(crime_df['population'])
print("Correlations with population (crime variety variables only):")
print(correlation_matrix.sort_values(ascending=False))

plt.figure(figsize=(12, 10))
sns.heatmap(crime_df[corr_columns].corr(), cmap='coolwarm', vmin=-1, vmax=1, center=0)
plt.title("Correlation Matrix of Crime Variety Variables")
plt.show()

columns_to_drop = ['population', 'total_offenses', 'city_id', 'year', 'latitude', 'longitude']
existing_columns_to_drop = [col for col in columns_to_drop if col in crime_df.columns]
X = crime_df.drop(columns=existing_columns_to_drop)
y = crime_df['population']

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

models = {
    'Linear Regression': LinearRegression(),
    'Decision Tree': DecisionTreeRegressor(random_state=42),
    'Random Forest': RandomForestRegressor(random_state=42),
    'Gradient Boosting': GradientBoostingRegressor(random_state=42),
    'SVR': SVR(),
    'Neural Network': MLPRegressor(random_state=42, max_iter=1000)
}

def evaluate_model(model, X_train, X_test, y_train, y_test):
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    r2 = r2_score(y_test, y_pred)
    mse = mean_squared_error(y_test, y_pred)
    cv_scores = cross_val_score(model, X, y, cv=5, scoring='r2')
    return r2, mse, np.mean(cv_scores), np.std(cv_scores)

results = {}
for name, model in models.items():
    r2, mse, cv_r2_mean, cv_r2_std = evaluate_model(model, X_train, X_test, y_train, y_test)
    results[name] = {
        'R-squared': r2,
        'MSE': mse,
        'Cross-validated R-squared': f"{cv_r2_mean:.2f} (± {cv_r2_std:.2f})"
    }

print("Model Performance on Test Set:")
for name, metrics in results.items():
    print(f"{name}:")
    print(f"  R-squared: {metrics['R-squared']:.4f}")
    print(f"  MSE: {metrics['MSE']:.2f}")
    print(f"  Cross-validated R-squared: {metrics['Cross-validated R-squared']}\n")

champion = max(results, key=lambda k: float(results[k]['Cross-validated R-squared'].split()[0]))
print(f"Champion Model: {champion}")
print(f"  R-squared: {results[champion]['R-squared']:.4f}")
print(f"  MSE: {results[champion]['MSE']:.2f}")
print(f"  Cross-validated R-squared: {results[champion]['Cross-validated R-squared']}")



##### Results
Model Performance:

Neural Network (Champion): R-squared: 0.9772, Cross-validated R-squared: 0.90 ± 0.06, MSE: 2,933,080.30
Random Forest: R-squared: 0.9817, Cross-validated R-squared: 0.86 ± 0.07, MSE: 2,348,983.96
Gradient Boosting: R-squared: 0.9848, Cross-validated R-squared: 0.73 ± 0.24, MSE: 1,953,164.13
Linear Regression: R-squared: 0.9692, Cross-validated R-squared: 0.46 ± 0.83, MSE: 3,958,126.42
Decision Tree: R-squared: 0.8234, Cross-validated R-squared: 0.69 ± 0.19, MSE: 22,700,393.33
SVR: R-squared: -0.0341, Cross-validated R-squared: -0.08 ± 0.06, MSE: 132,886,335.17
The Neural Network was selected as the champion model due to its high and stable cross-validated R-squared.


