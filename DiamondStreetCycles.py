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

    
# Create an empty dictionary to store Material objects, with the material name as the key.
materialMap = {}

# Extract lists of material names, costs per meter, and material constraints from a DataFrame.
materialNames = material_df['Material']
materialCosts = material_df['Kosten / m']
materialLimits = material_df['Materialbeschränkungen']

# Iterate through each row of the extracted material data.
for idx in range(len(materialNames)):
    # Create a Material object with the name, cost, and constraints of the material.
    mat = Material(materialNames[idx], materialCosts[idx], materialLimits[idx])
    # Add the Material object to the dictionary, using the material name as the key to the object.
    materialMap[materialNames[idx]] = mat

# Create an empty dictionary to store optimization variables, using product names as keys.
variableMap = {}

# Iterate through a list of product objects.
for prod in productList:
    # Create an optimization variable for each product using the product's name.
    variableMap[prod.name] = xp.var(name=prod.name)
    # Add the created variable to a collection of variables, presumably for an optimization model.
    DSS.addVariable(variableMap[prod.name])



    
### Constraints

# Initialize an empty list to hold all constraints
constraintList = []

## Material Constraints

# Loop over each material name in the materialMap dictionary
for matName in materialMap:   
    # Initialize a temporary list to hold products that use the current material
    tempList = []
    
    # Loop over each product in the productList
    for prod in productList:
        # If the current material is used in the product
        if matName in prod.materials:
            # Add the product to the temporary list
            tempList.append(prod)
    
    # Create a constraint ensuring the sum of the materials used by the products
    # does not exceed the material limit. This is done by summing the product
    # of the material quantity used in each product and the corresponding variable
    # from the variableMap, for all products that use the material.
    c_mat = xp.constraint(
        xp.Sum(tempList[i].materials[matName] * variableMap[tempList[i].name] for i in range(len(tempList))) <= materialMap[matName].limit,
        name=matName
    )
    
    # Add the constraint to the Decision Support System (DSS)
    DSS.addConstraint(c_mat)
    
    # Add the created constraint to the constraint list
    constraintList.append(c_mat)

    
 
## Offcut Control
# TO-DO condition in excel

# Create constraints based on the condition that the quantity of 'Fleece-Top' should be greater than or equal to 'Fleece-Shirt'
c_ver1 = variableMap['Fleece-Top'] >= variableMap['Fleece-Shirt']
# Create constraints based on the condition that the quantity of 'Sweatshorts' should be greater than or equal to 'Sweatshirt'
c_ver2 = variableMap['Sweatshorts'] >= variableMap['Sweatshirt']
# Add the constraints to the Decision Support System (DSS)
DSS.addConstraint(c_ver1)
DSS.addConstraint(c_ver2)
# Append the constraints to the constraint list for later reference
constraintList.append(c_ver1)
constraintList.append(c_ver2)

# Initialize a list for maximum constraints - important, do not delete!!!
maxConstraintList = []

## Maximum Forecast
for product in productList:
    if product.maxp > 0:
        # Create a constraint for maximum production limit
        constraint = variableMap[product.name] <= product.maxp
        # Add the constraint to the optimizer
        DSS.addConstraint(constraint)
        # Add the constraint to the list for tracking
        constraintList.append(constraint)
        # Add the constraint to the maximum constraint list - do not delete!!!
        maxConstraintList.append(constraint)

# Minimum Production Quantity
for product in productList:
    if product.minp > 0:
        # Create a constraint for minimum production quantity
        constraint = variableMap[product.name] >= product.minp
        # Add the constraint to the optimizer
        DSS.addConstraint(constraint)
        # Add the constraint to the constraint list for tracking
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

# Optimize the linear programming model
DSS.lpoptimize()
print("------------------")
print("LP OPTIMIZATION")
print("------------------")

# Retrieve the solution, slack, dual values, and reduced costs from the optimization
solution = DSS.getSolution()
slack = DSS.getSlack()
dual_values = DSS.getDual()
reduced_costs = DSS.getRCost()
objective_value = DSS.getObjVal()

# Generate a dictionary of optimal values for each variable
optimal_values = {var: DSS.getSolution(var) for var in variableMap.values()}

print("Solution:", solution)

# Print the objective function value (OFV)
print("OFV:", objective_value)

# Print slack values
print("Slack:", slack)
# Print dual values
print("Dual values:", dual_values)
# Print reduced costs
print("Reduced Costs:", reduced_costs)
print()

## Production per Variable

# Print production per variable
print('Production per Variable')
print()

# Initialize total production sum
total_production = 0

# Calculate production for each variable and aggregate the total production
for name, var in variableMap.items():
    print(name + ": " + str(DSS.getSolution(name)))
    total_production += DSS.getSolution(name)

# Print total production amount
print()
print("Total Production Amount: " + str(total_production))

## Returns
# Print the section for returns
print()
print('Returns')
print()

# Calculate and print the returns for each material
for matName in materialMap:   
    tempList = []
    
    # Create a list of products containing the material
    for prod in productList:
        if matName in prod.materials:
            tempList.append(prod)
    
    # Calculate and print the material return amount
    print(matName + ': ' + str(materialMap[matName].limit - sum(tempList[i].materials[matName] * optimal_values[variableMap[tempList[i].name]] for i in range(len(tempList)))))

    
## Contribution Margin

# Print the section for contribution margin per product
print()
print("Contribution Margin per Product")
print()

# Calculate and print the contribution margin for each product
for product in productList:
    db_product = (product.vk - product.mk) * optimal_values[variableMap[product.name]]
    print(f"CM - {product.name}: {db_product}")

# Calculate and print the total contribution margin
total_contribution_margin = sum((product.vk - product.mk) * optimal_values[variableMap[product.name]] for product in productList) - totalMaterialCosts
print()
print('Total CM: ' + str(total_contribution_margin))

# Calculate remaining material after returns
for matName in materialMap:   
    tempList = []
    
    # Create a list of products containing the material
    for prod in productList:
        if matName in prod.materials:
            tempList.append(prod)
    
    # Calculate remaining material for each type
    remainingMaterial[matName] = materialMap[matName].limit - sum(tempList[i].materials[matName] * optimal_values[variableMap[tempList[i].name]] for i in range(len(tempList)))

# Print the remaining material values
print(remainingMaterial.values())

# Calculate and print return costs
print('Return Costs: ' + str(sum(remainingMaterial.values()) * variables_df.loc[0, 'Return Costs']))
# Calculate and print return revenue
print('Return Revenue: ' + str(sum(quantity * materialMap[matName].costs for matName, quantity in remainingMaterial.items())))
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

