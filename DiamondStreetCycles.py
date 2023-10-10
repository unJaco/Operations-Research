# -- coding: utf-8 --
"""
DiamondStreetStyles - main.py
"""

import xpress as xp
import pandas as pd
import math


## Helper Classes

class Product:
    
    def __init__(self, name, vk, mk, maxp, minp, materials):
        self.name = name
        self.materials = materials
        self.vk = vk
        self.mk = mk
        self.maxp = maxp
        self.minp = minp
        
     
class Material:
    
    def __init__(self, name, costs, limit):
        self.name = name
        self.costs = costs
        self.limit = limit


## load Excel

DSS = xp.problem("Deinemutter")

# Load the Excel file and read data from the sheets "Material" and "Produkt"
file_path = 'Produktionsplanung.xlsx'  # Replace with your file path
material_df = pd.read_excel(file_path, sheet_name='Material')
produkt_df = pd.read_excel(file_path, sheet_name='Produkt')
fixed_costs_df = pd.read_excel(file_path, sheet_name='Fixkosten')



## Create a list with all products

productList = []

for idx in range(len(produkt_df['Produkt'])):
        
        productData = produkt_df.iloc[idx]
        materials = {}
        
        
        for matIdx in range(5, len(productData),2):
            materials[productData[matIdx]] = productData[matIdx + 1]
            
        productList.append(Product(productData[0], productData[1], productData[2], productData[3], productData[4], materials))
    
        


## create a Map / Dict with all Variables
## each variable has the name of the corresponding product
variableMap = {}

for prod in productList:
    variableMap[prod.name] = xp.var(name=prod.name)
    DSS.addVariable(variableMap[prod.name])


## create a map / dict with all materials

materialMap = {}

materialNames = material_df['Material']
materialCosts = material_df['Kosten / m']
materialLimits = material_df['Materialbeschränkungen']

for idx in range(len(materialNames)):
    mat = Material(materialNames[idx], materialCosts[idx], materialLimits[idx])
    materialMap[materialNames[idx]] = mat
    
    
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

for product in productList:
    if product.maxp > 0:
        DSS.addConstraint(variableMap[product.name] <= product.maxp)
    
    

## Mindestproduktionsmenge

for product in productList:
    if product.minp > 0:
        DSS.addConstraint(variableMap[product.name] >= product.minp)
    

### Kosten 

## Fixed Costs

totalFixedCosts = sum(fixed_costs_df['Betrag'])

print(totalFixedCosts)

## Material Kosten

totalMaterialCosts = sum(material.limit * material.costs for material in materialMap.values())


##### Total Costs

totalCosts = totalFixedCosts + totalMaterialCosts


######## Zielfunktion


objective = sum((product.vk - product.mk) * variableMap[product.name] for product in productList) - totalCosts

DSS.setObjective(objective, sense=xp.maximize)

DSS.lpoptimize()

print("Lösung:", DSS.getSolution())
print("ZFW:", DSS.getObjVal())
print("Schattenpreise:", DSS.getDual())
print("Schlupf:", DSS.getSlack())
print("RCost:", DSS.getRCost())
    