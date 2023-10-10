# -- coding: utf-8 --
"""
DiamondStreetStyles - main.py
"""

import xpress as xp
import pandas as pd

DSS = xp.problem("Deinemutter")

# Load the Excel file and read data from the sheets "Material" and "Produkt"
file_path = 'Produktionsplanung.xlsx'  # Replace with your file path
material_df = pd.read_excel(file_path, sheet_name='Material')
produkt_df = pd.read_excel(file_path, sheet_name='Produkt')

# Extracting VPi and KMi from produkt_df
VPi = produkt_df['Verkaufspreis'].values
KMi = produkt_df['Maschinenkosten'].values

# Constructing a dictionary for material costs for quick lookup
material_costs = dict(zip(material_df['Material'], material_df['Kosten / m']))

# Extracting Mij values as explained previously
num_products = len(produkt_df)
num_materials = len(material_df)
Mij = [[0]*num_materials for _ in range(num_products)]
for i in range(num_products):
    for j in range(1, 3):
        material_name_col = f"{j}. Material - Beschreibung"
        material_amt_col = f"{j}. Matrial - Maße in Meter"
        material_name = produkt_df.at[i, material_name_col].strip()
        if material_name in material_costs:
            material_index = list(material_df['Material']).index(material_name)
            Mij[i][material_index] = produkt_df.at[i, material_amt_col]



df1 = pd.read_excel("Produktionsplanung.xlsx", sheet_name=0)
df2 = pd.read_excel("Produktionsplanung.xlsx", sheet_name=1)
# Einlesen der Parameter
print(df1)

# Material M1
maße_essentials_m1 = df1['1. Matrial - Maße in Meter'].values[0]
maße_urban_m1 = df1['1. Matrial - Maße in Meter'].values[1]
maße_cargohose_m1 = df1['1. Matrial - Maße in Meter'].values[2]
maße_sweatshirt_m1 = df1['1. Matrial - Maße in Meter'].values[3]
maße_sweatshorts_m1 = df1['1. Matrial - Maße in Meter'].values[4]
maße_cargoshorts_m1 = df1['1. Matrial - Maße in Meter'].values[5]
maße_sportshirt_m1 = df1['1. Matrial - Maße in Meter'].values[6]
maße_sweatpants_m1 = df1['1. Matrial - Maße in Meter'].values[7]
maße_seidenshorts_m1 = df1['1. Matrial - Maße in Meter'].values[8]
maße_fleece_shirt_m1 = df1['1. Matrial - Maße in Meter'].values[9]
maße_fleece_top_m1 = df1['1. Matrial - Maße in Meter'].values[10]

#Material m2
maße_essentials_m2 = df1['2. Matrial - Maße in Meter'].values[0]
maße_urban_m2 = df1['2. Matrial - Maße in Meter'].values[1]
maße_cargohose_m2 = df1['2. Matrial - Maße in Meter'].values[2]
maße_sweatpants_m2 = df1['2. Matrial - Maße in Meter'].values[7]

vk_essentials = df1['Verkaufspreis'].values[0]
vk_urban = df1['Verkaufspreis'].values[1]
vk_cargohose = df1['Verkaufspreis'].values[2]
vk_sweatshirt = df1['Verkaufspreis'].values[3]
vk_sweatshorts = df1['Verkaufspreis'].values[4]
vk_cargoshorts = df1['Verkaufspreis'].values[5]
vk_sportshirt = df1['Verkaufspreis'].values[6]
vk_sweatpants = df1['Verkaufspreis'].values[7]
vk_seidenshorts = df1['Verkaufspreis'].values[8]
vk_fleece_shirt = df1['Verkaufspreis'].values[9]
vk_fleece_top = df1['Verkaufspreis'].values[10]

# Maschinenkosten
mk_essentials = df1['Maschinenkosten'].values[0]
mk_urban = df1['Maschinenkosten'].values[1]
mk_cargo = df1['Maschinenkosten'].values[2]
mk_sweatshirt = df1['Maschinenkosten'].values[3]
mk_sweatshorts = df1['Maschinenkosten'].values[4]
mk_cargoshorts = df1['Maschinenkosten'].values[5]
mk_sportshirt = df1['Maschinenkosten'].values[6]
mk_sweatpants = df1['Maschinenkosten'].values[7]
mk_seidenshorts = df1['Maschinenkosten'].values[8]
mk_fleece_shirt = df1['Maschinenkosten'].values[9]
mk_fleece_top = df1['Maschinenkosten'].values[10]

limit_baumwolle = df2['Materialbeschränkungen'].values[0]
limit_elastan = df2['Materialbeschränkungen'].values[1]
limit_seide = df2['Materialbeschränkungen'].values[2]
limit_fleece = df2['Materialbeschränkungen'].values[3]
limit_nylon = df2['Materialbeschränkungen'].values[4]
limit_recyceltes_polyester = df2['Materialbeschränkungen'].values[5]
limit_polyester = df2['Materialbeschränkungen'].values[6]

# nparray_df1 = (essentials:[x1, m1, m2, vk, mk],
#               urban:[x2, m1, m2, vk, mk],
#               cargo:[x3, m1, m2, vk, mk],])


# Variables
x1 = xp.var() #  Anzahl der produzierten Trainingshosen “Diamond Essentials”
x2 = xp.var() #  Anzahl der produzierten Trainingshosen “Diamond Urban”
x3 = xp.var() # Anzahl der produzierten Cargohosen
x4 = xp.var() # Anzahl der produzierten Sweatshirts
x5 = xp.var() # Anzahl der produzierten Sweatshorts
x6 = xp.var() # Anzahl der produzierten Cargoshorts
x7 = xp.var() # Anzahl der produzierten Sportshirts im Trikot-Style
x8 = xp.var() # Anzahl der produzierten Sweatpants
x9 = xp.var() # Anzahl der produzierten Seidenshorts
x10 = xp.var() # Anzahl der produzierten Fleece-Shirts
x11 = xp.var() # Anzahl der produzierten Fleece-Tops


DSS.addVariable(x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11)

# Nebenbedingungen

## Materialbeschränkungen

mat_neb1 = maße_urban_m1*x2 + maße_sweatpants_m1*x8 <= limit_baumwolle # Baumwolle
mat_neb2 = maße_essentials_m2*x1 + maße_urban_m2*x2 + maße_cargohose_m2*x3 + maße_sweatpants_m2*x8 <= limit_elastan # Elastan
mat_neb3 = maße_seidenshorts_m1*x9 <= limit_seide # Seide
mat_neb4 = maße_fleece_shirt_m1*x10 + maße_fleece_top_m1*x11 <= limit_fleece # Fleece
mat_neb5 = maße_essentials_m1*x1 + maße_sportshirt_m1*x7 <= limit_nylon  # Nylon
mat_neb6 = maße_cargohose_m1*x3 + maße_cargoshorts_m1*x6 <= limit_recyceltes_polyester # recyceltes Polyester
mat_neb7 = maße_sweatshirt_m1*x4 + maße_sweatshorts_m1*x5 <= limit_polyester # Polyester


## Verschnittregelung
ver_neb1= x11 >= x10
ver_neb2= x5 >= x4

## Maximalprognose
max_neb1 = x3 <= 5000 # Cargohose
max_neb2 = x6 <= 6000 # Cargoshorts
max_neb3 = x9 <= 4000 # Seidenshorts
max_neb4 = x10 <= 12000 # Fleece-Shirt
max_neb5 = x11 <= 15000 # Fleece-Top

## Mindestproduktionsmenge

min_neb1 = x8 >= 0.6 * 7000
min_neb2 = x2 >= 0.6 * 5000
min_neb3 = x1 >= 2800

## Fixkosten
Designer = 940000
Veranstalrungsort = 2250000
FK = Designer + Veranstalrungsort


DSS.addConstraint(mat_neb1, mat_neb2, mat_neb3, mat_neb4, mat_neb5, mat_neb6, mat_neb7, ver_neb1, ver_neb2, max_neb1, max_neb2, max_neb3, max_neb4, max_neb5, min_neb1, min_neb2, min_neb3)


# Zielfunktion

x = [x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11]




objective = sum([(VPi[i] - KMi[i] - sum([Mij[i][j] * material_costs[material] for j, material in enumerate(material_df['Material'])])) * x[i] for i in range(num_products)])

DSS.setObjective(objective, sense=xp.maximize)

DSS.lpoptimize()

print("Lösung:", DSS.getSolution())
print("ZFW:", DSS.getObjVal())
print("Schattenpreise:", DSS.getDual())
print("Schlupf:", DSS.getSlack())
print("RCost:", DSS.getRCost())