import openpyxl
from openpyxl import load_workbook
import unicodedata

#Read datasheets
sample_wb = load_workbook("/Users/noravandevoorde/Downloads/SPHERE/Add_Population_copy.xlsx")
sample_ws = sample_wb.active
census_wb = load_workbook("/Users/noravandevoorde/Downloads/SPHERE/tabela4714.xlsx")
census_ws = census_wb.active

#Print column names in both datasets
print("Columns in sample datasheet:", [cell.value for cell in sample_ws[1] if cell.value is not None])
print("Columns in census dataset:", [cell.value for cell in census_ws[1] if cell.value is not None])

#remove accents
def remove_accents(input_str):
    if input_str is None:
        return None
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return ''.join([c for c in nfkd_form if not unicodedata.combining(c)])

#clean names
def clean(value):
    if value is None:
        return None
    return remove_accents(str(value).lower().strip())

#Get cities in sample and clean names
sample_cities = {}
for excel_row, row in enumerate(sample_ws.iter_rows(min_row=2, values_only=True)):
    sample_cities[excel_row] = {
        "City": clean(row[2]) if row[2] is not None else ""
    }

print(list(sample_cities.items())[:5])

#Get cities and population from census data
cities_and_pop = {}
for census_row, row in enumerate(census_ws.iter_rows(min_row=3, values_only=True)):
    cities_and_pop[census_row] = {
        "City": clean(row[2]) if row[2] is not None else "",
        "Population": clean(row[3]) if row[3] is not None else ""
    }

print(list(cities_and_pop.items())[:5])