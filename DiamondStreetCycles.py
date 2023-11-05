# -- coding: utf-8 --
"""
DiamondStreetStyles - main.py
"""

import xpress as xp
import pandas as pd
import xlsxwriter

#excel file which will be created 
workbook = xlsxwriter.Workbook('Output.xlsx')



#### Helper Classes ####


## Product - {Name, Verkaufspreis, Maschinenkosten, 
#   Maximalprognose, Mindestproduktionsmenge, Materialien}

# Maximalprognose und Mindestproduktionsmenge können undefiniert sein
# Materials ist ein Dictionary für jedes Item gilt  {name : m/Stück}

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

# Laden der Excel File und auslesen von Material, Produkt, Fixkosten und Variablen
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
    
    # Add the constraint to the problem
    DSS.addConstraint(c_mat)
    
    # Add the created constraint to the constraint list
    constraintList.append(c_mat)

    
 
## Offcut Control
# TO-DO condition in excel

# Create constraints based on the condition that the quantity of 'Fleece-Top' should be greater than or equal to 'Fleece-Shirt'
c_ver1 = variableMap['Fleece-Top'] >= variableMap['Fleece-Shirt']
# Create constraints based on the condition that the quantity of 'Sweatshorts' should be greater than or equal to 'Sweatshirt'
c_ver2 = variableMap['Sweatshorts'] >= variableMap['Sweatshirt']
# Add the constraints to the problem
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

# Retrieve the optimization solution
solution = DSS.getSolution()
# Get the slack values from the solution
schlupf = DSS.getSlack()
# Get the dual values from the solution
dualwerte = DSS.getDual()
# Get the reduced costs from the solution
redkosten = DSS.getRCost()
# Get the objective function value from the solution
ZFWert = DSS.getObjVal()

# Create a dictionary of optimal values for each variable
optimal_values = {var: DSS.getSolution(var) for var in variableMap.values()}

# Print the solution
print("Solution:", solution)

# Print the objective function value
print("Objective Function Value:", ZFWert)

# Print the slack values
print("Slack:", schlupf)
# Print the dual values
print("Dual Values:", dualwerte)
# Print the reduced costs
print("Reduced Costs:", redkosten)
print()

# Production per variable
print('Production per Variable')
print()

# Initialize the total production counter
gesamtProd = 0


worksheet = workbook.add_worksheet(name="OPRIMALER PRODUKTIONSPLAM")
worksheetP = workbook.add_worksheet(name="KEIN PLOYESTER")
worksheetO = workbook.add_worksheet(name="PRODUKTIONSPLAN OUTLET")

row = 1

# Iterate through each variable and print its production level
for name, var in variableMap.items():
    print(name + ": " + str(DSS.getSolution(name)))
    # Sum up the total production level
    gesamtProd += DSS.getSolution(name)
    worksheet.write(row, 0, name)
    worksheet.write(row, 1, DSS.getSolution(name))
    
    row += 1
    
# Print the total production quantity
print()
print("Total Manufacturing Quantity: " + str(gesamtProd))
row += 1

worksheet.write(row, 0, "Total Manufacturing Quantity: ")
worksheet.write(row, 1, gesamtProd)

row += 2

worksheet.write(row, 0, "Rücksendungen")

row += 1

# Returns
print()
print('Returns')
print()

# For each material, calculate and print the returns
for matName in materialMap:   
    tempList = []
    
    # Collect products that use the current material
    for prod in productList:
        if matName in prod.materials:
            tempList.append(prod)
    
    # Calculate and print the amount of material returned
    r = materialMap[matName].limit - sum(tempList[i].materials[matName] * optimal_values[variableMap[tempList[i].name]] for i in range(len(tempList)))
    print(matName + ': ' +  str(r))
    worksheet.write(row, 0, matName)
    worksheet.write(row, 1, r)
    
    row += 1

row += 1
# Contribution Margin

# Contribution margin per product
print()
print("Contribution Margin per Product")
print()

worksheet.write(row, 0, "Contribution Margin per Product: ")

row += 1
# Calculate and print the contribution margin for each product
for product in productList:
    db_product = (product.vk - product.mk) * optimal_values[variableMap[product.name]]
    worksheet.write(row, 0, product.name)
    worksheet.write(row, 1, db_product)
    row += 1
    print(f"Contribution Margin - {product.name}: {db_product}")

row+= 1


# Total contribution margin
dbgesamt = sum((product.vk - product.mk) * optimal_values[variableMap[product.name]] for product in productList) - totalMaterialCosts
worksheet.write(row, 0, "Contribution Margin total: ")
worksheet.write(row, 1, dbgesamt)


row += 2
print()
print('Total Contribution Margin: ' + str(dbgesamt))

# Calculate the remaining material after production
for matName in materialMap:   
    tempList = []
    
    # Collect products that use the current material
    for prod in productList:
        if matName in prod.materials:
            tempList.append(prod)
    
    # Calculate the remaining material for each material type
    remainingMaterial[matName] = materialMap[matName].limit - sum(tempList[i].materials[matName] * optimal_values[variableMap[tempList[i].name]] for i in range(len(tempList)))

# Print the remaining material quantities
print(remainingMaterial.values())


# Calculate and print the return costs
rc = sum(remainingMaterial.values()) * variables_df.loc[0, 'Rücksendekosten']
print('Return Costs: ' + str(rc))


worksheet.write(row, 0, "Return Costs: ")
worksheet.write(row, 1, rc)

row += 2
  
# Calculate and print the money obtained from returns
rm = sum(quantity * materialMap[matName].costs for matName, quantity in remainingMaterial.items())
print('Return Money: ' + str(rm))
print()


worksheet.write(row, 0, "Return Money: ")
worksheet.write(row, 1, rm)

row += 2


worksheet.write(row, 0, "Profit: ")
worksheet.write(row, 1, ZFWert)

# ************************************
# Sensitivitaetsanalyse
# ************************************

print("------------------")
print("SENSITIVITÄTSANALYSE")
print("------------------")

worksheetS = workbook.add_worksheet(name="INCREASE MACHINE COSTS")

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


# calculate the Zielfunktionskoeffizient
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
    

print('Steigerung der Maschinenkosten bis ein Impact auf den optimalen Produktionsplan')

row = 1

worksheetS.write(row, 0, "Wie viel müssen die Maschinenkosten pro Produkt steigen um den optimalen Plan zu beeinflussen?")

row += 1

for var, lo, zfk in zip(all_variables, lower_obj, zfkList):
    print(f"{var.name}: {zfk-lo}")
    worksheetS.write(row, 0, var.name)
    worksheetS.write(row, 1, zfk-lo)
    
    row += 1
    

worksheetE = workbook.add_worksheet(name="IMPACT OF EXTRA ELASTAN")

row = 1

print()
print('Impact von 1 Elastan auf Gewinn') 
print()
lower_rhs, upper_rhs = [], []

worksheetE.write(row, 0, 'Which Impact has one extra Elastan on the profit?')

row += 1

idx = list(materialMap).index("Elastan")

DSS.rhssa([idx], lower_rhs, upper_rhs)


worksheetE.write(row, 0, 'Increase per Elastan: ')
worksheetE.write(row, 1, DSS.getDual(idx))

print(DSS.getDual(idx))

print("Die trifft zu bis zu einer Menge von: ", upper_rhs)

row += 2

worksheetE.write(row, 0, 'This is valid until a total Elastan of: ')
worksheetE.write(row, 1, upper_rhs[0])

print()

# Schlupfvariablen für jede Nebenbedingung

# Liste der aktiven und inaktiven Nebenbedingungen erstellen
active_constraints = []
inactive_constraints = []


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

row = 1

for name, var in variableMap.items():
    print(name + ": " + str(DSS.getSolution(name)))
    worksheetP.write(row, 0, name)
    worksheetP.write(row, 1, DSS.getSolution(name))
    
    row += 1
    
    
print("------------------")
print("LP-OPTIMIERUNG OUTLET")
print("------------------")

#löschen der Constraints die efür in Produktionsmaximum sorgen
DSS.delConstraint(maxConstraintList)


#herausfinden für welche Produkte das Produktionsmaximum überschritten wurde
overstockMap = {p.name: xp.max(0, variableMap[p.name] - p.maxp ) for p in productList}

#amount über maximum multiplizieren mit gewinn und 0.4 und dann abziehen
objective = sum((product.vk - product.mk) * variableMap[product.name] for product in productList) - totalCosts - sum(0.4 * (p.vk - p.mk) * xp.max(variableMap[p.name] - p.maxp, 0) for p in productList)


DSS.lpoptimize()

solution = DSS.getSolution()
ZFWert = DSS.getObjVal()


print("Lösung:", solution)
print("ZFW:", ZFWert)

## Produktion pro Variable

row = 1

print('Produktion pro Variable, Produktion über MaxPrognose')
print()

worksheetO.write(row, 0, "Name")
worksheetO.write(row, 1, "Amount")
worksheetO.write(row, 2, "Amount over MaxProg")

row += 2
 
for p in productList:
    
    overMax = max(0, DSS.getSolution(p.name) - p.maxp)
    print(p.name + ": " + str(DSS.getSolution(p.name)) + ", " + str(overMax))
    
    worksheetO.write(row, 0, p.name)
    worksheetO.write(row, 1, str(DSS.getSolution(p.name)))
    worksheetO.write(row, 2, str(overMax))
    
    row += 1
    
    
    
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

workbook.close()

