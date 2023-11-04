# -- coding: utf-8 --
"""
DiamondStreetStyles - main.py
"""

import xpress as xp
import pandas as pd
import math


#### Helper Classes ####


## Product - {Name, Verkaufspreis, Maschinenkosten, 
#   Maximalprognose, Mindestproduktionsmenge, Materialien}

# Maximalprognose und Mindestproduktionsmenge können undefiniert sein
# Materialien - {name : m/Stück}

class Product:
    
    def __init__(self, name, vk, mk, maxp, minp, materials):
        self.name = name
        self.materials = materials
        self.vk = vk
        self.mk = mk
        self.maxp = maxp
        self.minp = minp
        
## Material - {Name, Kosten / m, Limit}   
class Material:
    
    def __init__(self, name, costs, limit):
        self.name = name
        self.costs = costs
        self.limit = limit

#### End Helper Classes ####

#### Read Excel #### 

## load Excel
DSS = xp.problem("DiamondStreetStyles")

## XPRESS Settings
DSS.setControl('outputlog', 0)

# Laden der Excel File und auslesen von Material, Produkt und Fixkosten
file_path = 'Produktionsplanung.xlsx'  # Replace with your file path
material_df = pd.read_excel(file_path, sheet_name='Material')
produkt_df = pd.read_excel(file_path, sheet_name='Produkt')
fixed_costs_df = pd.read_excel(file_path, sheet_name='Fixkosten')
variables_df = pd.read_excel(file_path, sheet_name='Variablen')


## Create a list with all products

productList = []

## for each product in Produkt Column
for idx in range(len(produkt_df['Produkt'])):
        
        ## read Row with the index of the Product
        productData = produkt_df.iloc[idx]
        
        ## number of materials is dynamic
        # first material name is always in column 5, 
        # materials required is always the next column
        materials = {}
        for matIdx in range(5, len(productData),2):
            materials[productData[matIdx]] = productData[matIdx + 1]
        
        # add Product to product list
        productList.append(Product(productData[0], productData[1], productData[2], productData[3], productData[4], materials) )

    
## create a map / dict with all materials

materialMap = {}

materialNames = material_df['Material']
materialCosts = material_df['Kosten / m']
materialLimits = material_df['Materialbeschränkungen']

for idx in range(len(materialNames)):
    mat = Material(materialNames[idx], materialCosts[idx], materialLimits[idx])
    materialMap[materialNames[idx]] = mat
            

## create a Map / Dict with all Variables
## each variable has the name of the corresponding product
variableMap = {}

for prod in productList:
    variableMap[prod.name] = xp.var(name=prod.name)
    DSS.addVariable(variableMap[prod.name])


    
### Constraints

constraintList = []

## Materialbeschränkungen

for matName in materialMap:   
    tempList = []
    
    for prod in productList:
        if matName in prod.materials:
            tempList.append(prod)
    c_mat = xp.constraint(xp.Sum(tempList[i].materials[matName] * variableMap[tempList[i].name] for i in range(len(tempList))) <= materialMap[matName].limit, name=matName)
        
    DSS.addConstraint(c_mat)
    
    constraintList.append(c_mat)
    
 
## Verschnittregelung
# TO-DO condition in excel

c_ver1 = variableMap['Fleece-Top'] >= variableMap['Fleece-Shirt']
c_ver2 = variableMap['Sweatshorts'] >= variableMap['Sweatshirt']
DSS.addConstraint(c_ver1)
DSS.addConstraint(c_ver2)
constraintList.append(c_ver1)
constraintList.append(c_ver2)

maxConstraintList = [] # wichtig, nicht löschen !!!

## Maximalprognose
for product in productList:
    if product.maxp > 0:
        # Erstelle eine Constraint
        constraint = variableMap[product.name] <= product.maxp
        # Füge die Constraint zum Optimierer hinzu
        DSS.addConstraint(constraint)
        # Füge die Constraint zur Liste hinzu
        constraintList.append(constraint)
        maxConstraintList.append(constraint) # nicht löschen !!!

# Mindestproduktionsmenge
for product in productList:
    if product.minp > 0:
        # Erstelle eine Constraint
        constraint = variableMap[product.name] >= product.minp
        # Füge die Constraint zum Optimierer hinzu
        DSS.addConstraint(constraint)
        # Füge die Constraint zur Liste hinzu
        constraintList.append(constraint)
    
## Übriges Material

remainingMaterial = {}

for matName in materialMap:   
    tempList = []
    
    for prod in productList:
        if matName in prod.materials:
            tempList.append(prod)
    
    remainingMaterial[matName] = materialMap[matName].limit - sum(tempList[i].materials[matName] * variableMap[tempList[i].name] for i in range(len(tempList)))


#### Kosten #### 

## Fixed Costs

totalFixedCosts = sum(fixed_costs_df['Betrag'])

## Material Kosten

totalMaterialCosts = sum(material.limit * material.costs for material in materialMap.values())


## Rücksendekosten ##



returnCosts = sum(remainingMaterial.values()) * variables_df.loc[0, 'Rücksendekosten']

## Rückerstattungspreis

returnMoney = sum(quantity * materialMap[matName].costs for matName, quantity in remainingMaterial.items())

### Total Costs ###

totalCosts = totalFixedCosts + totalMaterialCosts + returnCosts - returnMoney


######## Zielfunktion

objective = sum((product.vk - product.mk) * variableMap[product.name] for product in productList) - totalCosts

DSS.setObjective(objective, sense=xp.maximize)



# ************************************
# LP-OPTIMIERUNG
# ************************************

DSS.lpoptimize()
print("------------------")
print("LP-OPTIMIERUNG")
print("------------------")


solution = DSS.getSolution()
schlupf = DSS.getSlack()
dualwerte = DSS.getDual()
redkosten = DSS.getRCost()
ZFWert = DSS.getObjVal()

optimal_values = {var: DSS.getSolution(var) for var in variableMap.values()}

print("Lösung:", solution)

print("ZFW:", ZFWert)


print("Schlupf:", schlupf)
print("Dualwerte:", dualwerte)
print("Reduzierte Kosten:", redkosten)
print()

## Produktion pro Variable

print('Produktion pro Variable')
print()

gesamtProd = 0

for name, var in variableMap.items():
    print(name + ": " + str(DSS.getSolution(name)))
    gesamtProd += DSS.getSolution(name)
    
print()
print("Gesamt Herstellungsmenge: " + str(gesamtProd))
    

## Zurücksendungen
print()
print('Zurücksendungen')
print()

for matName in materialMap:   
    tempList = []
    
    for prod in productList:
        if matName in prod.materials:
            tempList.append(prod)
    
    print(matName + ': ' +  str(materialMap[matName].limit - sum(tempList[i].materials[matName] * optimal_values[variableMap[tempList[i].name]] for i in range(len(tempList)))))

    
## Deckungsbeitrag

## Deckungsbeitrag pro Produkt
print()
print("Deckungsbeitrag pro Produkt")
print()
for product in productList:
    db_product = (product.vk - product.mk) * optimal_values[variableMap[product.name]]
    print(f"DB - {product.name}: {db_product}")

## Deckungsbeitrag gesamt
dbgesamt = sum((product.vk - product.mk) * optimal_values[variableMap[product.name]] for product in productList) - totalMaterialCosts
print()
print('DB gesamt: ' + str(dbgesamt))

for matName in materialMap:   
    tempList = []
    
    for prod in productList:
        if matName in prod.materials:
            tempList.append(prod)
    
    remainingMaterial[matName] = materialMap[matName].limit - sum(tempList[i].materials[matName] * optimal_values[variableMap[tempList[i].name]] for i in range(len(tempList)))


print(remainingMaterial.values())

print('Return Costs: ' + str(sum(remainingMaterial.values()) * variables_df.loc[0, 'Rücksendekosten']))
print('Return Money: ' + str(sum(quantity * materialMap[matName].costs for matName, quantity in remainingMaterial.items())))
print()



# ************************************
# Sensitivitaetsanalyse
# ************************************

print("------------------")
print("SENSITIVITÄTSANALYSE")
print("------------------")

# Sensitivitätsanalyse für Zielfunktionskoeffizienten
all_variables = list(variableMap.values())

# Prepare empty lists for the lower and upper bounds.
lower_obj = []
upper_obj = []

# Call the objsa function with the correct parameters.
DSS.objsa(all_variables, lower_obj, upper_obj)

# Now lower_obj and upper_obj lists will be populated with the sensitivity ranges for the objective coefficients.
print("\nSensitivity for Objective Function Coefficients:")#
print()


zfkList = []
for prod in productList:
    zfk = 0
    for mat in materialNames:
        if mat in prod.materials:
            x = prod.materials[mat] * 0.1
            zfk += x
            y = prod.materials[mat] * materialMap[mat].costs
            zfk -= y
    
    zfk += prod.vk - prod.mk
    zfkList.append(zfk)
    

print('Steigerung der Maschinenkosten bis ein Impact auf den optimalen Produktionsplan') #(TO-DO BESSER NENNEN)

for var, lo, zfk in zip(all_variables, lower_obj, zfkList):
    print(f"{var.name}: {zfk-lo}")
    


print()
print('So viel mehr Elastan füht zu einer Änderung') #(TO-DO BESSER NENNEN)
print()
lower_rhs, upper_rhs = [], []

idx = list(materialMap).index("Elastan")

DSS.rhssa([idx], lower_rhs, upper_rhs)


print("Untere Grenzen:", lower_rhs)
print("Obere Grenzen:", upper_rhs)

print()

# Schlupfvariablen für jede Nebenbedingung

# Liste der aktiven und inaktiven Nebenbedingungen erstellen
active_constraints = []
inactive_constraints = []

print('Dual Elastan')
print(DSS.getDual(idx))


# Schlupf für jede Nebenbedingung überprüfen
for idx, constr in enumerate(constraintList):
    slack_value = DSS.getSlack(constr)
    if slack_value == 0:
        active_constraints.append(f'NB{idx+1}')  # +1, weil die Zählung der Constraints bei 1 beginnt
    else:
        inactive_constraints.append(f'NB{idx+1}')

# Ausgabe der aktiven und inaktiven Nebenbedingungen
print("Aktive Nebenbedingungen:", active_constraints)
print("Inaktive Nebenbedingungen:", inactive_constraints)

# Erstellen Sie zwei leere Listen für rowstat und colstat
rowstat = [0] * DSS.attributes.rows
colstat = [0] * DSS.attributes.cols

# Rufen Sie getbasis auf, um die Basisstatus zu erhalten
DSS.getbasis(rowstat, colstat)

# Finden Sie die Basis- und Nicht-Basisvariablen
basis_vars = []
nonbasis_vars = []
variable_names = list(variableMap.values())

for i, status in enumerate(colstat):
    if status == 1:
        basis_vars.append(variable_names[i])
    else:
        nonbasis_vars.append(variable_names[i])

# Ausgabe der Basis- und Nicht-Basisvariablen
print("Basisvariablen:", basis_vars)
print("Nicht-Basisvariablen:", nonbasis_vars)




print("------------------")
print("LP-OPTIMIERUNG OHNE RÜCKNAHME PLOYESTER")
print("------------------")



# überschreiben der alten Zielfunktion
#in neuer Zielfunktion werden Rücksendekosten für Polyester addiert und das zurück erstattete Geld für Polyester abgezogen
objective = sum((product.vk - product.mk) * variableMap[product.name] for product in productList) - totalCosts + remainingMaterial['recyceltes Polyester'] * variables_df.loc[0, 'Rücksendekosten'] - remainingMaterial['recyceltes Polyester'] * materialMap['recyceltes Polyester'].costs

DSS.setObjective(objective, sense=xp.maximize)

DSS.lpoptimize()

solution = DSS.getSolution()
ZFWert = DSS.getObjVal()


print("Lösung:", solution)
print("ZFW:", ZFWert)

## Produktion pro Variable

print('Produktion pro Variable')
print()

for name, var in variableMap.items():
    print(name + ": " + str(DSS.getSolution(name)))
    
    
print("------------------")
print("LP-OPTIMIERUNG OUTLET")
print("------------------")


DSS.delConstraint(maxConstraintList)


overstockMap = {p.name: xp.max(0, variableMap[p.name] - p.maxp ) for p in productList}


objective = sum((product.vk - product.mk) * variableMap[product.name] for product in productList) - totalCosts - sum(0.4 * (p.vk - p.mk) * xp.max(variableMap[p.name] - p.maxp, 0) for p in productList)


DSS.lpoptimize()

solution = DSS.getSolution()
ZFWert = DSS.getObjVal()


print("Lösung:", solution)
print("ZFW:", ZFWert)

## Produktion pro Variable

print('Produktion pro Variable, Produktion über MaxPrognose')
print()

for p in productList:
    print(p.name + ": " + str(DSS.getSolution(p.name)) + ", " + str(max(0, DSS.getSolution(p.name) - p.maxp)))
    
    
optimal_values = {var: DSS.getSolution(var) for var in variableMap.values()}

## Zurücksendungen
print()
print('Zurücksendungen')
print()

for matName in materialMap:   
    tempList = []
    
    for prod in productList:
        if matName in prod.materials:
            tempList.append(prod)
    
    print(matName + ': ' +  str(materialMap[matName].limit - sum(tempList[i].materials[matName] * optimal_values[variableMap[tempList[i].name]] for i in range(len(tempList)))))
    
