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
DSS = xp.problem("Deinemutter")

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
        productList.append(Product(productData[0], productData[1], productData[2], productData[3], productData[4], materials))

    
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
    
    minp, maxp = 0,0
    
    if math.isnan(prod.minp):
      minp = 0
    
    else:
        minp = prod.minp
    
    if math.isnan(prod.maxp):
        maxp = 30000
    else:
        maxp = prod.maxp
        
    variableMap[prod.name] = xp.var(name=prod.name, vartype=xp.integer, lb=minp, ub=maxp)
    DSS.addVariable(variableMap[prod.name])


    
### Constraints

## Materialbeschränkungen

for matName in materialMap:   
    tempList = []
    
    for prod in productList:
        if matName in prod.materials:
            tempList.append(prod)
         
    DSS.addConstraint(xp.Sum(tempList[i].materials[matName] * variableMap[tempList[i].name] for i in range(len(tempList))) <= materialMap[matName].limit)

 
## Verschnittregelung
# TO-DO condition in excel

DSS.addConstraint(variableMap['Fleece-Top'] >= variableMap['Fleece-Shirt'])
DSS.addConstraint(variableMap['Sweatshorts'] >= variableMap['Sweatshirt'])

## Maximalprognose

"""
for product in productList:
    if product.maxp > 0:
        DSS.addConstraint(variableMap[product.name] <= product.maxp)
    
 """   

## Mindestproduktionsmenge
"""
for product in productList:
    if product.minp > 0:
        DSS.addConstraint(variableMap[product.name] >= product.minp)
    
    """
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

DSS.mipoptimize()

solution = DSS.getSolution()

optimal_values = {var: DSS.getSolution(var) for var in variableMap.values()}


print("Lösung:", solution)
print()

## Produktion pro Variable

print('Produktion pro Variable')
print()

for name, var in variableMap.items():
    print(name + ": " + str(DSS.getSolution(name)))
    
    
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

dbgesamt = sum((product.vk - product.mk) * optimal_values[variableMap[product.name]] for product in productList) - totalMaterialCosts
print()
print('DB gesamt: ' + str(dbgesamt))
    