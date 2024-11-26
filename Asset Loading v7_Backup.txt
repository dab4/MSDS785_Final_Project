Asset Loading Optimization
Version 7 - Has GC attributes within the objective function and as a constraint. Added code to combine individual plant optimizations into a single "Grouped" output file.
Version 6 - Adds attribute 2 number of machines from a constraint into the objective function to minimize.
Version 5 - Moves attribute 1 number of machines from a constraint into the objective function to minimize.

Python libraries to load
# EXECUTE FIRST
​
# computational imports
from pyomo.environ import *
import numpy as np
import pandas as pd
import time
from datetime import datetime
Raw Data Loads:
User Inputs: Brand
# Clear brand variable
brand = 'None'
​
# Brands available for analysis:
brands = ["Child Care", "Depend"]
​
# Ask user for brand to optimize
print("Input brand to optimize from the following list:", brands)
brand = input()
​
while brand not in brands:
    print("Input did not correctly match available brands from the following list:", brands)
    brand = input()
    
# Define input file pathway
path = './' + brand + '/Inputs.xlsx'
Input brand to optimize from the following list: ['Child Care', 'Depend']
Child Care
Load Raw Data
### Production rate data ###
​
# Load production rate data
rates_data = pd.read_excel(path, sheet_name = 'Production Rates', index_col=0)
​
# Clean production rate data
rates_data = rates_data.fillna(0)
smallR = 0.000001
rates_data = rates_data.replace(0, smallR)
rates_data = rates_data.round(6)
​
# Sort rows and columns
rates_data.sort_index(inplace=True)
rates_data = rates_data.reindex(sorted(rates_data.columns), axis=1)
### Constraint data ###
​
# Load available hours data
hours_data = pd.read_excel(path, sheet_name = 'Constraints', index_col=1)
constraint_data = hours_data
hours_data = hours_data.drop(columns=['No. of Attr 1s', 'No. of Attr 2s'], axis=1)
​
# Clean available hours data
hours_data = hours_data[hours_data['available_hours'] != 0]
hours_data.dropna(inplace=True)
hours_data.sort_index(inplace=True)
​
# Clean attribute 1 data
attr1_constraint = constraint_data.drop(columns=['available_hours', 'No. of Attr 2s'])
attr1_constraint = attr1_constraint.dropna()
​
# Clean attribute 2 data
attr2_constraint = constraint_data.drop(columns=['available_hours', 'No. of Attr 1s'])
attr2_constraint = attr2_constraint.dropna()
### Demand and Product data ###
​
# Load demand data
demand_data = pd.read_excel(path, sheet_name = 'Demand', index_col=0)
​
# Create segmentation dataframe
seg_data = demand_data['segmentation'].to_frame()
seg_data_original = seg_data
​
# Convert fill rate targets to penalties
seg_list = [0.995, 0.985, 0.970, 0.940]  # Allowable values
seg_data.loc[~seg_data["segmentation"].isin(seg_list), "segmentation"] = 0.970   # Convert non-allowable values to 0.970
seg_data = seg_data.replace({0.995:1000, 0.985:500, 0.970:300, 0.940:1})  # Convert FR target to penalty
seg_data.sort_index(inplace=True)
​
# Create attribute 1 dataframe
attr1_data = demand_data['attribute 1'].to_frame()
attr1_data_original = demand_data['attribute 1'].to_frame()
attr1_data.sort_index(inplace=True)
attr1_data['values'] = 1
attr1_data = attr1_data.pivot(columns='attribute 1', values='values').fillna(0).astype(int)
​
# Create attribute 2 dataframe
attr2_data = demand_data['attribute 2'].to_frame()
attr2_data_original = demand_data['attribute 2'].to_frame()
attr2_data.sort_index(inplace=True)
attr2_data['values'] = 1
attr2_data = attr2_data.pivot(columns='attribute 2', values='values').fillna(0).astype(int)
​
# Clean demand data
demand_data = demand_data.drop(columns=['attribute 1', 'attribute 2', 'segmentation'], axis=1)
demand_data = demand_data.applymap(lambda x: 0 if isinstance(x, str) else x)
demand_data.sort_index(inplace=True)
Construct the Model
User Input: Plants to Optimize
# Clear plant variable
plant = []
​
# Plants available for analysis:
plants = sorted(list(demand_data.columns.values))
​
# Ask user for plants to optimize
for pl in plants:
    answer = 'NaN'
    print("\nWould you like to optimize", pl, "(Input either Y or N)?")
    answer = input()
    if answer.lower() == "y":
        plant.append(pl)
print("\nThe following plants will be optimized indivdually:", plant)
        
# Ask user for utopia
print("\nWould you like to optimize all plants", plants, "together as one (Y/N)?")
utopia = input()
utopia = utopia.lower()
utopia_plants=[]
utopia_plants.append(plants)

Would you like to optimize BIL (Input either Y or N)?
y

Would you like to optimize OGD (Input either Y or N)?
y

Would you like to optimize PAR (Input either Y or N)?
y

The following plants will be optimized indivdually: ['BIL', 'OGD', 'PAR']

Would you like to optimize all plants ['BIL', 'OGD', 'PAR'] together as one (Y/N)?
y
Function: Optimization Model
### Optimization model function ###
​
def lp_function(products, assets, demand, seg_penalty, rates, hours_available, plant, attr1_matrix, attr1_constraint, attr1_list, attr2_matrix, attr2_constraint, attr2_list):
    
    
    # Declaration
    model = ConcreteModel()
​
    
    # Decision Variables
    model.production_volume = Var(products, assets, domain=NonNegativeReals)
    model.attr1_boolean = Var(attr1_list, assets, domain=Boolean)
    model.attr2_boolean = Var(attr2_list, assets, domain=Boolean)
    
​
    # Objective Function  -->  MIN SCHD HRS AND GAP w/ SEG PENALTY & GC PENALTY
    
    GC_penalty = 100
    GC2_penalty = 100
    
    model.schedule_hours = Objective(expr=sum(model.production_volume[pr,a]/rates.loc[pr,a] for a in assets
                                              for pr in products) + 
                                     sum((demand.loc[pr][0] - sum(model.production_volume[pr,a] for a in assets)) * seg_penalty.loc[pr][0] 
                                         for pr in products) +
                                     sum(model.attr1_boolean[attr1, a] for attr1 in attr1_list for a in assets) * GC_penalty +
                                     sum(model.attr2_boolean[attr2, a] for attr2 in attr2_list for a in assets) * GC2_penalty, sense=minimize)
    
​
​
    # Constraints
    model.capability = ConstraintList()
    for a in assets:   
        for pr in products:
            if rates.loc[pr,a] == smallR:
                model.capability.add(model.production_volume[pr, a] == 0)
    #model.capability.pprint()
​
​
    model.capacity = ConstraintList()
    for a in assets:
        model.capacity.add(
            sum(model.production_volume[pr,a] * (1/rates.loc[pr,a]) for pr in products) <= hours_available.loc[a][0])
    #model.capacity.pprint()
​
​
    model.demand = ConstraintList()
    for pr in products:
        model.demand.add(
            sum(model.production_volume[pr,a] for a in assets) <= demand.loc[pr][0])
    #model.demand.pprint()
    
    
    ### GC Constraint - Only produce if attr1_boolean = 1 ###
    
    bigM = 1000000
    
    model.gradeChange = ConstraintList()
    for a in assets:    
        for pr in products:
            for attr1 in attr1_list:
                model.gradeChange.add(
                    model.production_volume[pr,a] * attr1_matrix.loc[pr, attr1] <= bigM * model.attr1_boolean[attr1,a])  
                
    for a in assets:    
        for pr in products:
            for attr2 in attr2_list:
                model.gradeChange.add(
                    model.production_volume[pr,a] * attr2_matrix.loc[pr, attr2] <= bigM * model.attr2_boolean[attr2,a])                 
    
    ### GC Constraint - Limit the number of Grade Changes ###
    
    for a in assets:
        model.gradeChange.add(
            sum(model.attr1_boolean[attr1, a] for attr1 in attr1_list) <= attr1_constraint.loc[a][0])
        
    for a in assets:
        model.gradeChange.add(
            sum(model.attr2_boolean[attr2, a] for attr2 in attr2_list) <= attr2_constraint.loc[a][0])
    #model.gradeChange.pprint()
    
    
    
    ### SOLUTION ###
    results = SolverFactory('glpk').solve(model)
    results.write()
    
    
    
    ### SUMMARIZE RESULTS ###
    
    # Production volume
    prod_volume = pd.DataFrame([[model.production_volume[pr,a]() for a in assets] for pr in products],
                    index = products,
                    columns= assets)
    prod_volume = prod_volume.round(1)
    prod_volume.index.name='product'
    print(prod_volume.head())
    
    # Attribute 1 grade chnage matrix
    try:
        GCmatrix = pd.DataFrame([[model.attr1_boolean[attr1,a]() for attr1 in attr1_list] for a in assets],
                index = assets,
                columns= attr1_list).astype(int)
        print(GCmatrix.head())
    except ValueError:
        GCmatrix=[]
        
    # Attribute 2 grade chnage matrix
    try:
        GC2matrix = pd.DataFrame([[model.attr2_boolean[attr2,a]() for attr2 in attr2_list] for a in assets],
                index = assets,
                columns= attr2_list).astype(int)
        print(GC2matrix.head())
    except ValueError:
        GC2matrix=[]    
    
    # Scheduled hours
    schd_hours_rates = rates[assets]
    schd_hours_rates = schd_hours_rates.loc[products]
    schd_hours = prod_volume / schd_hours_rates
    schd_hours = schd_hours.round(1)
    print(schd_hours.head())
    
    # Production volume vs demand
    prod_volume_bySKU = prod_volume.sum(axis='columns')
    prod_volume_bySKU = prod_volume_bySKU.to_frame(name = "Production")
    prod_volume_bySKU.index.names = ['product']
    prod_volume_bySKU = pd.merge(demand, prod_volume_bySKU, how='outer', on='product')
    prod_volume_bySKU = pd.merge(prod_volume_bySKU, seg_data_original, how='left', on='product')
    prod_volume_bySKU = pd.merge(prod_volume_bySKU, attr1_data_original, how='left', on='product')
    prod_volume_bySKU = pd.merge(prod_volume_bySKU, attr2_data_original, how='left', on='product')
    prod_volume_bySKU.columns.values[0] = "Demand"
    prod_volume_bySKU['Delta'] = prod_volume_bySKU['Demand'] - prod_volume_bySKU['Production']
    prod_volume_bySKU['Demand'] = prod_volume_bySKU['Demand'].round(0)
    prod_volume_bySKU['Production'] = prod_volume_bySKU['Production'].round(0)
    prod_volume_bySKU['Delta'] = prod_volume_bySKU['Delta'].round(0)
    prod_volume_bySKU = prod_volume_bySKU[['segmentation', 'attribute 1', 'attribute 2', 'Demand', 'Production', 'Delta']]
    prod_volume_bySKU = prod_volume_bySKU.sort_index()
    print(prod_volume_bySKU.head())
​
    # Scheduled Hours vs Available Hours
    schd_vs_avail = schd_hours.sum(axis='index')
    schd_vs_avail = schd_vs_avail.to_frame(name = "Schd_Hours")
    schd_vs_avail['Schd_Hours'] = schd_vs_avail['Schd_Hours'].round(0)
    schd_vs_avail.index.names = ['machine']
    schd_vs_avail = pd.merge(hours_available, schd_vs_avail, how='right', on='machine')
    schd_vs_avail.columns.values[0] = "Avail_Hours"
    schd_vs_avail['Avail_Hours'] = schd_vs_avail['Avail_Hours'].round(0)
    schd_vs_avail['Delta'] = schd_vs_avail['Avail_Hours'] - schd_vs_avail['Schd_Hours']
    print(schd_vs_avail.head())
​
    # Production Rates by SKU and Asset
    prod_rates = schd_hours_rates
    prod_rates = prod_rates.fillna(0)
    prod_rates = prod_rates.replace(smallR, 0)
    prod_rates = prod_rates.round(1)
    print(prod_rates.head())
    
    
​
    ### RETURN RESULTS ###
    return prod_volume, schd_hours, prod_volume_bySKU, schd_vs_avail, prod_rates, GCmatrix, GC2matrix
Run the model for individual plants
### SINGLE PLANTS ###
​
if len(plant) > 0:
​
    # Create a list to store solutions
​
    solutions_list_prod_volume = []
    solutions_list_schd_hours = []
    solutions_list_prod_volume_bySKU = []
    solutions_list_schd_vs_avail = []
    solutions_list_prod_rates = []
    solutions_list_GCmatrix = []
    solutions_list_GC2matrix = []
​
​
​
    # Run optimization function for each plant
    for pl in plant:
    
        # Create demand dataset
        demand_function = demand_data[[pl]]
        demand_function = demand_function.loc[~(demand_function==0).all(axis=1)]
    
        # Create product list
        products_function = demand_function.index.to_list()
    
        # Create seg_penalty dataset
        seg_penalty_function = pd.merge(demand_function, seg_data, how='left', on='product')
        seg_penalty_function = seg_penalty_function.drop(pl, axis=1)
    
        # Create available hours dataset
        hours_function = hours_data[hours_data['plant'] == pl]
        hours_function = hours_function.drop('plant', axis=1)
    
        # Create assets list
        assets_function = hours_function.index.tolist()
    
        # Create rates dataset
        rates_function = rates_data[assets_function]
        rates_function = pd.merge(demand_function, rates_function, how='left', on='product')
        rates_function = rates_function.drop(pl, axis=1)
        rates_function = rates_function.fillna(smallR)
        
        # Create attribute 1 & 2 boolean dataset
        attr1_function = pd.merge(demand_function, attr1_data, how='left', on='product').drop(pl, axis=1)
        attr2_function = pd.merge(demand_function, attr2_data, how='left', on='product').drop(pl, axis=1)
        
        # Create attribute 1 & 2 constraint dataset
        attr1_constraint_function = attr1_constraint[attr1_constraint['plant'] == pl].drop('plant', axis=1)
        attr2_constraint_function = attr2_constraint[attr2_constraint['plant'] == pl].drop('plant', axis=1)
        
        # Create attribute 1 & 2 list
        attr1_list = attr1_function.columns.tolist()
        attr2_list = attr2_function.columns.tolist()
    
        # Call function
        PV, SH, VolSKU, SvA, PR, GC, GC2 = lp_function(products_function, assets_function, demand_function, seg_penalty_function, rates_function, hours_function, pl, attr1_function, attr1_constraint_function, attr1_list, attr2_function, attr2_constraint_function, attr2_list)
​
        solutions_list_prod_volume.append(PV)
        solutions_list_schd_hours.append(SH)
        solutions_list_prod_volume_bySKU.append(VolSKU)
        solutions_list_schd_vs_avail.append(SvA)
        solutions_list_prod_rates.append(PR)
        solutions_list_GCmatrix.append(GC)
        solutions_list_GC2matrix.append(GC2)
​
    # Store solutions in a dataframe
    solutions_df = pd.DataFrame(list(zip(plant, solutions_list_prod_volume, solutions_list_schd_hours, solutions_list_prod_volume_bySKU, solutions_list_schd_vs_avail, solutions_list_prod_rates, solutions_list_GCmatrix, solutions_list_GC2matrix)), columns=['plant', 'prod_vol', 'schd_hours', 'prod by SKU', 'schd vs avail', 'rates', 'GCmatrix-Attr1', 'GCmatrix-Attr2'])
    solutions_df.set_index('plant', inplace=True)
# ==========================================================
# = Solver Results                                         =
# ==========================================================
# ----------------------------------------------------------
#   Problem Information
# ----------------------------------------------------------
Problem: 
- Name: unknown
  Lower bound: 2884.38440366519
  Upper bound: 2884.38440366519
  Number of objectives: 1
  Number of constraints: 2562
  Number of variables: 250
  Number of nonzeros: 3378
  Sense: minimize
# ----------------------------------------------------------
#   Solver Information
# ----------------------------------------------------------
Solver: 
- Status: ok
  Termination condition: optimal
  Statistics: 
    Branch and bound: 
      Number of bounded subproblems: 7
      Number of created subproblems: 7
  Error rc: 0
  Time: 0.2158370018005371
# ----------------------------------------------------------
#   Solution Information
# ----------------------------------------------------------
Solution: 
- number of solutions: 0
  number of solutions displayed: 0
         B18     B19  B20
product                  
40531    0.0  2806.3  0.0
40532    0.0  1512.7  0.0
41313    0.0  1674.6  0.0
41314    0.0  1169.3  0.0
43528    0.0   382.0  0.0
     S2  S3  S3.5  S4  S5  S5.5  S6
B18   1   1     0   0   0     0   0
B19   0   0     1   1   0     0   0
B20   0   1     1   0   0     0   0
     DIAPERPNT  SWMPNT  TRNINGPNT  YOUTHPNT
B18          0       0          1         0
B19          0       0          1         1
B20          0       0          1         0
         B18   B19  B20
product                
40531    0.0  90.9  0.0
40532    0.0  48.4  0.0
41313    0.0  59.8  0.0
41314    0.0  39.2  0.0
43528    0.0  11.7  0.0
         segmentation attribute 1 attribute 2  Demand  Production  Delta
product                                                                 
40531           0.985          S4    YOUTHPNT  2806.0      2806.0   -0.0
40532           0.985          S4    YOUTHPNT  1513.0      1513.0    0.0
41313           0.985          S4    YOUTHPNT  1675.0      1675.0   -0.0
41314           0.985          S4    YOUTHPNT  1169.0      1169.0   -0.0
43528           0.985          S4    YOUTHPNT   382.0       382.0    0.0
         Avail_Hours  Schd_Hours  Delta
machine                                
B18            642.0       642.0    0.0
B19            641.0       641.0    0.0
B20            650.0       601.0   49.0
         B18   B19  B20
product                
40531    0.0  30.9  0.0
40532    0.0  31.2  0.0
41313    0.0  28.0  0.0
41314    0.0  29.8  0.0
43528    0.0  32.5  0.0
# ==========================================================
# = Solver Results                                         =
# ==========================================================
# ----------------------------------------------------------
#   Problem Information
# ----------------------------------------------------------
Problem: 
- Name: unknown
  Lower bound: 1906.9302473843
  Upper bound: 1906.9302473843
  Number of objectives: 1
  Number of constraints: 1659
  Number of variables: 163
  Number of nonzeros: 2165
  Sense: minimize
# ----------------------------------------------------------
#   Solver Information
# ----------------------------------------------------------
Solver: 
- Status: ok
  Termination condition: optimal
  Statistics: 
    Branch and bound: 
      Number of bounded subproblems: 1
      Number of created subproblems: 1
  Error rc: 0
  Time: 0.19351482391357422
# ----------------------------------------------------------
#   Solution Information
# ----------------------------------------------------------
Solution: 
- number of solutions: 0
  number of solutions displayed: 0
           U10    U12
product              
45121    325.4    0.0
45122    276.3    0.0
45127      0.0  384.3
45128      0.0  680.5
45491      0.0  116.3
     S2  S3  S3.5  S4  S5  S5.5  S6
U10   1   1     0   0   0     0   0
U12   0   1     1   0   0     0   0
     DIAPERPNT  SWMPNT  TRNINGPNT  YOUTHPNT
U10          0       0          1         0
U12          0       0          1         0
         U10   U12
product           
45121    9.5   0.0
45122    8.2   0.0
45127    0.0   9.6
45128    0.0  17.1
45491    0.0   3.9
         segmentation attribute 1 attribute 2  Demand  Production  Delta
product                                                                 
45121           0.985          S2   TRNINGPNT   325.0       325.0   -0.0
45122           0.985          S2   TRNINGPNT   276.0       276.0    0.0
45127           0.985          S3   TRNINGPNT   384.0       384.0    0.0
45128           0.985          S3   TRNINGPNT   680.0       680.0   -0.0
45491           0.985          S3   TRNINGPNT   116.0       116.0   -0.0
         Avail_Hours  Schd_Hours  Delta
machine                                
U10            640.0       640.0    0.0
U12            639.0       639.0    0.0
          U10   U12
product            
45121    34.2   0.0
45122    33.5   0.0
45127    31.8  40.1
45128    32.4  39.7
45491    26.0  29.8
# ==========================================================
# = Solver Results                                         =
# ==========================================================
# ----------------------------------------------------------
#   Problem Information
# ----------------------------------------------------------
Problem: 
- Name: unknown
  Lower bound: 6797.59588144251
  Upper bound: 6797.59588144251
  Number of objectives: 1
  Number of constraints: 11461
  Number of variables: 1037
  Number of nonzeros: 15216
  Sense: minimize
# ----------------------------------------------------------
#   Solver Information
# ----------------------------------------------------------
Solver: 
- Status: ok
  Termination condition: optimal
  Statistics: 
    Branch and bound: 
      Number of bounded subproblems: 13
      Number of created subproblems: 13
  Error rc: 0
  Time: 0.27489542961120605
# ----------------------------------------------------------
#   Solution Information
# ----------------------------------------------------------
Solution: 
- number of solutions: 0
  number of solutions displayed: 0
         P09  P12  P13     P15    P16  P17    P18
product                                          
16186    0.0  0.0  0.0     0.0    0.0  0.0   98.0
18339    0.0  0.0  0.0     0.0  225.2  0.0    0.0
18342    0.0  0.0  0.0     0.0  405.9  0.0    0.0
18345    0.0  0.0  0.0     0.0    0.0  0.0  259.2
40531    0.0  0.0  0.0  3648.7    0.0  0.0    0.0
     S2  S3  S3.5  S4  S5  S5.5  S6
P09   0   0     1   1   0     0   0
P12   0   1     1   1   0     0   0
P13   0   0     0   0   1     1   1
P15   0   0     0   1   1     0   0
P16   1   1     0   1   0     0   0
     DIAPERPNT  SWMPNT  TRNINGPNT  YOUTHPNT
P09          0       0          1         0
P12          0       0          1         1
P13          0       0          0         1
P15          0       0          0         1
P16          0       1          1         0
         P09  P12  P13    P15   P16  P17   P18
product                                       
16186    0.0  0.0  0.0    0.0   0.0  0.0   3.9
18339    0.0  0.0  0.0    0.0   9.1  0.0   0.0
18342    0.0  0.0  0.0    0.0  16.0  0.0   0.0
18345    0.0  0.0  0.0    0.0   0.0  0.0  13.0
40531    0.0  0.0  0.0  118.1   0.0  0.0   0.0
         segmentation attribute 1 attribute 2  Demand  Production  Delta
product                                                                 
16186           0.970          S5      SWMPNT    98.0        98.0    0.0
18339           0.940          S3      SWMPNT   225.0       225.0    0.0
18342           0.940          S4      SWMPNT   406.0       406.0   -0.0
18345           0.940          S5      SWMPNT   259.0       259.0    0.0
40531           0.985          S4    YOUTHPNT  3649.0      3649.0    0.0
         Avail_Hours  Schd_Hours  Delta
machine                                
P09            645.0       645.0    0.0
P12            649.0       649.0    0.0
P13            538.0       538.0    0.0
P15            648.0       647.0    1.0
P16            555.0       311.0  244.0
         P09   P12  P13   P15   P16  P17   P18
product                                       
16186    0.0   0.0  0.0   0.0   0.0  0.0  25.0
18339    0.0   0.0  0.0   0.0  24.9  0.0   0.0
18342    0.0   0.0  0.0   0.0  25.4  0.0   0.0
18345    0.0   0.0  0.0   0.0   0.0  0.0  20.0
40531    0.0  27.1  0.0  30.9   0.0  0.0   0.0
Write the single plant results to Excel file
if len(plant) > 0:
    
    # Create dataframe to summarize production volumes for all plants
    if len(plant) > 1:
        prod_vol_all_plants = pd.merge(solutions_df.iloc[0][0], solutions_df.iloc[1][0], on='product', how='outer')
        schd_hours_all_plants = pd.merge(solutions_df.iloc[0][1], solutions_df.iloc[1][1], on='product', how='outer')
        Prod_vs_Demand_all_plants = solutions_df.iloc[0][2].append(solutions_df.iloc[1][2])
        Prod_vs_Demand_all_plants = Prod_vs_Demand_all_plants.groupby(['product', 'segmentation', 'attribute 1', 'attribute 2']).sum()
        Prod_vs_Demand_all_plants.reset_index(inplace=True)
        Prod_vs_Demand_all_plants.set_index('product', inplace=True)
        Schd_vs_Avail_all_plants = solutions_df.iloc[0][3].append(solutions_df.iloc[1][3])
        GC_Matrix1_all_plants = solutions_df.iloc[0][5].append(solutions_df.iloc[1][5])
        GC_Matrix2_all_plants = solutions_df.iloc[0][6].append(solutions_df.iloc[1][6])
        
        if len(plant) > 2:
            for pl in range(2,len(plant)):
                prod_vol_all_plants = pd.merge(prod_vol_all_plants, solutions_df.iloc[pl][0], on='product', how='outer')
                #prod_vol_all_plants = prod_vol_all_plants.replace('', 0)
                #prod_vol_all_plants = prod_vol_all_plants.fillna('')
                prod_vol_all_plants.sort_index(inplace=True)
                
                schd_hours_all_plants = pd.merge(schd_hours_all_plants, solutions_df.iloc[pl][1], on='product', how='outer')
                #schd_hours_all_plants = schd_hours_all_plants.replace('', 0)
                #schd_hours_all_plants = schd_hours_all_plants.fillna('')
                schd_hours_all_plants.sort_index(inplace=True)
                
                Prod_vs_Demand_all_plants = Prod_vs_Demand_all_plants.append(solutions_df.iloc[pl][2])
                Prod_vs_Demand_all_plants.reset_index(inplace=True)
                Prod_vs_Demand_all_plants.set_index('product', inplace=True)
                
                Schd_vs_Avail_all_plants = Schd_vs_Avail_all_plants.append(solutions_df.iloc[pl][3])
                
                GC_Matrix1_all_plants = GC_Matrix1_all_plants.append(solutions_df.iloc[pl][5])
                
                GC_Matrix2_all_plants = GC_Matrix2_all_plants.append(solutions_df.iloc[pl][6])
                
            Prod_vs_Demand_all_plants = Prod_vs_Demand_all_plants.groupby(['product', 'segmentation', 'attribute 1', 'attribute 2']).sum()
            Prod_vs_Demand_all_plants = Prod_vs_Demand_all_plants.replace('', 0)
            Prod_vs_Demand_all_plants = Prod_vs_Demand_all_plants.fillna(0)
            
        # Clean up dataframe
        prod_vol_all_plants = prod_vol_all_plants.replace('', 0)
        prod_vol_all_plants = prod_vol_all_plants.fillna(0)
        schd_hours_all_plants = schd_hours_all_plants.replace('', 0)
        schd_hours_all_plants = schd_hours_all_plants.fillna(0)
​
    for pl in plant:
        
        # Create file name
        timestamp = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
        file = './' + brand + '/Results ' + pl + ' ' + timestamp + '.xlsx'
​
        # Write to Excel
        with pd.ExcelWriter(file) as writer:
            solutions_df.loc[pl]['prod_vol'].to_excel(writer, sheet_name="Prod_Volume", index=True)
            if len(plant) > 1:
                prod_vol_all_plants.to_excel(writer, sheet_name="Prod_Volume_All_Plants", index=True)
            solutions_df.loc[pl]['schd_hours'].to_excel(writer, sheet_name="Schd_Hours", index=True)
            solutions_df.loc[pl]['prod by SKU'].to_excel(writer, sheet_name="Prod_vs_Demand", index=True)
            solutions_df.loc[pl]['schd vs avail'].to_excel(writer, sheet_name="Schd_vs_Avail", index=True)
            solutions_df.loc[pl]['rates'].to_excel(writer, sheet_name="Production_Rates", index=True)
            solutions_df.loc[pl]['GCmatrix-Attr1'].to_excel(writer, sheet_name="GC_Matrix-Attr1", index=True)
            solutions_df.loc[pl]['GCmatrix-Attr2'].to_excel(writer, sheet_name="GC_Matrix-Attr2", index=True)
            
    # Write grouped results for individual plant optimizations to Excel
    if len(plant) > 1:
        grouped_plant = ','.join(plant)
        file_grouped = './' + brand + '/Results ' + grouped_plant + '_GROUPED ' + timestamp + '.xlsx'
        with pd.ExcelWriter(file_grouped) as writer:
            prod_vol_all_plants.to_excel(writer, sheet_name="Prod_Volume", index=True)
            schd_hours_all_plants.to_excel(writer, sheet_name="Schd_Hours", index=True)
            Prod_vs_Demand_all_plants.to_excel(writer, sheet_name="Prod_vs_Demand", index=True)
            Schd_vs_Avail_all_plants.to_excel(writer, sheet_name="Schd_vs_Avail", index=True)
            GC_Matrix1_all_plants.to_excel(writer, sheet_name="GC_Matrix-Attr1", index=True)
            GC_Matrix2_all_plants.to_excel(writer, sheet_name="GC_Matrix-Attr2", index=True)
Run the model for multiple plants
### MULTIPLE PLANTS ###
​
if utopia == 'y':
    
    # Create a list to store solutions
​
    utopia_list_plants = []
    utopia_list_prod_volume = []
    utopia_list_schd_hours = []
    utopia_list_prod_volume_bySKU = []
    utopia_list_schd_vs_avail = []
    utopia_list_prod_rates = []
    utopia_list_GCmatrix = []
    utopia_list_GC2matrix = []
    
    
    
    # Run optimization function for each set of plants
    
    for ut in range(0, len(utopia_plants)):
        
        # Create column name
        utopia_col = ','.join(utopia_plants[ut])
        utopia_list_plants.append(utopia_col)
    
        # Create demand dataset
        utopia_demand_function = demand_data[utopia_plants[ut]]
        utopia_demand_function[utopia_col] = utopia_demand_function.sum(axis=1)
        utopia_demand_function = utopia_demand_function[utopia_col]
        utopia_demand_function = utopia_demand_function.to_frame(name = utopia_col)
        utopia_demand_function = utopia_demand_function.loc[~(utopia_demand_function==0).all(axis=1)]
    
        # Create product list
        utopia_products_function = utopia_demand_function.index.to_list()
    
        # Create seg_penalty dataset
        utopia_seg_penalty_function = pd.merge(utopia_demand_function, seg_data, how='left', on='product')
        utopia_seg_penalty_function = utopia_seg_penalty_function.drop(utopia_col, axis=1)
    
        # Create available hours dataset
        utopia_hours_function = hours_data[hours_data['plant'].isin(utopia_plants[ut])]
        utopia_hours_function = utopia_hours_function.drop('plant', axis=1)
    
        # Create assets list
        utopia_assets_function = utopia_hours_function.index.tolist()
    
        # Create rates dataset
        utopia_rates_function = rates_data[utopia_assets_function]
        utopia_rates_function = pd.merge(utopia_demand_function, utopia_rates_function, how='left', on='product')
        utopia_rates_function = utopia_rates_function.drop(utopia_col, axis=1)
        utopia_rates_function = utopia_rates_function.fillna(smallR)
        
        # Create attribute 1 & 2 boolean dataset
        utopia_attr1_function = pd.merge(utopia_demand_function, attr1_data, how='left', on='product').drop(utopia_col, axis=1)
        utopia_attr2_function = pd.merge(utopia_demand_function, attr2_data, how='left', on='product').drop(utopia_col, axis=1)
        
        # Create attribute 1 & 2 constraint dataset
        utopia_attr1_constraint_function = attr1_constraint.drop('plant', axis=1)
        utopia_attr2_constraint_function = attr2_constraint.drop('plant', axis=1)
        
        # Create attribute 1 & 2 list
        utopia_attr1_list = utopia_attr1_function.columns.tolist()
        utopia_attr2_list = utopia_attr2_function.columns.tolist()
    
        # Call function
        uPV, uSH, uVolSKU, uSvA, uPR, uGC, uGC2 = lp_function(utopia_products_function, utopia_assets_function, utopia_demand_function, utopia_seg_penalty_function, utopia_rates_function, utopia_hours_function, utopia_col, utopia_attr1_function, utopia_attr1_constraint_function, utopia_attr1_list, utopia_attr2_function, utopia_attr2_constraint_function, utopia_attr2_list)
​
        utopia_list_prod_volume.append(uPV)
        utopia_list_schd_hours.append(uSH)
        utopia_list_prod_volume_bySKU.append(uVolSKU)
        utopia_list_schd_vs_avail.append(uSvA)
        utopia_list_prod_rates.append(uPR)
        utopia_list_GCmatrix.append(uGC)
        utopia_list_GC2matrix.append(uGC2)
        
    # Store solutions in a dataframe
    utopia_df = pd.DataFrame(list(zip(utopia_list_plants, utopia_list_prod_volume, utopia_list_schd_hours, utopia_list_prod_volume_bySKU, utopia_list_schd_vs_avail, utopia_list_prod_rates, utopia_list_GCmatrix, utopia_list_GC2matrix)), columns=['plant', 'prod_vol', 'schd_hours', 'prod by SKU', 'schd vs avail', 'rates', 'GCmatrix-Attr1', 'GCmatrix-Attr2'])
    utopia_df.set_index('plant', inplace=True)
# ==========================================================
# = Solver Results                                         =
# ==========================================================
# ----------------------------------------------------------
#   Problem Information
# ----------------------------------------------------------
Problem: 
- Name: unknown
  Lower bound: 10899.0791791088
  Upper bound: 10899.0791791088
  Number of objectives: 1
  Number of constraints: 24542
  Number of variables: 2197
  Number of nonzeros: 32722
  Sense: minimize
# ----------------------------------------------------------
#   Solver Information
# ----------------------------------------------------------
Solver: 
- Status: ok
  Termination condition: optimal
  Statistics: 
    Branch and bound: 
      Number of bounded subproblems: 141
      Number of created subproblems: 141
  Error rc: 0
  Time: 0.6738533973693848
# ----------------------------------------------------------
#   Solution Information
# ----------------------------------------------------------
Solution: 
- number of solutions: 0
  number of solutions displayed: 0
         B18     B19  B20  P09  P12  P13  P15    P16  P17    P18  U10  U12
product                                                                   
16186    0.0     0.0  0.0  0.0  0.0  0.0  0.0    0.0  0.0   98.0  0.0  0.0
18339    0.0     0.0  0.0  0.0  0.0  0.0  0.0  225.2  0.0    0.0  0.0  0.0
18342    0.0     0.0  0.0  0.0  0.0  0.0  0.0  405.9  0.0    0.0  0.0  0.0
18345    0.0     0.0  0.0  0.0  0.0  0.0  0.0    0.0  0.0  259.2  0.0  0.0
40531    0.0  6455.0  0.0  0.0  0.0  0.0  0.0    0.0  0.0    0.0  0.0  0.0
     S2  S3  S3.5  S4  S5  S5.5  S6
B18   1   1     0   0   0     0   0
B19   0   0     0   1   0     0   0
B20   0   0     1   0   0     0   0
P09   0   0     1   1   0     0   0
P12   0   1     1   1   0     0   0
     DIAPERPNT  SWMPNT  TRNINGPNT  YOUTHPNT
B18          0       0          1         0
B19          0       0          0         1
B20          0       0          1         0
P09          0       0          1         0
P12          0       0          1         1
         B18    B19  B20  P09  P12  P13  P15   P16  P17   P18  U10  U12
product                                                                
16186    0.0    0.0  0.0  0.0  0.0  0.0  0.0   0.0  0.0   3.9  0.0  0.0
18339    0.0    0.0  0.0  0.0  0.0  0.0  0.0   9.1  0.0   0.0  0.0  0.0
18342    0.0    0.0  0.0  0.0  0.0  0.0  0.0  16.0  0.0   0.0  0.0  0.0
18345    0.0    0.0  0.0  0.0  0.0  0.0  0.0   0.0  0.0  13.0  0.0  0.0
40531    0.0  209.1  0.0  0.0  0.0  0.0  0.0   0.0  0.0   0.0  0.0  0.0
         segmentation attribute 1 attribute 2  Demand  Production  Delta
product                                                                 
16186           0.970          S5      SWMPNT    98.0        98.0    0.0
18339           0.940          S3      SWMPNT   225.0       225.0    0.0
18342           0.940          S4      SWMPNT   406.0       406.0   -0.0
18345           0.940          S5      SWMPNT   259.0       259.0    0.0
40531           0.985          S4    YOUTHPNT  6455.0      6455.0    0.0
         Avail_Hours  Schd_Hours  Delta
machine                                
B18            642.0       642.0    0.0
B19            641.0       642.0   -1.0
B20            650.0       650.0    0.0
P09            645.0       645.0    0.0
P12            649.0       649.0    0.0
         B18   B19  B20  P09   P12  P13   P15   P16  P17   P18  U10  U12
product                                                                 
16186    0.0   0.0  0.0  0.0   0.0  0.0   0.0   0.0  0.0  25.0  0.0  0.0
18339    0.0   0.0  0.0  0.0   0.0  0.0   0.0  24.9  0.0   0.0  0.0  0.0
18342    0.0   0.0  0.0  0.0   0.0  0.0   0.0  25.4  0.0   0.0  0.0  0.0
18345    0.0   0.0  0.0  0.0   0.0  0.0   0.0   0.0  0.0  20.0  0.0  0.0
40531    0.0  30.9  0.0  0.0  27.1  0.0  30.9   0.0  0.0   0.0  0.0  0.0
Write the multiple plant results to Excel file
if utopia == 'y':
    
    for ut in range(0, len(utopia_plants)):
        
        # Create file name
        timestamp = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
        utopia_col = ','.join(utopia_plants[ut])
        file = './' + brand + '/Results ' + utopia_col + '_UTOPIA ' + timestamp + '.xlsx'
​
        # Write to Excel
        with pd.ExcelWriter(file) as writer:
            utopia_df.loc[utopia_list_plants[ut]]['prod_vol'].to_excel(writer, sheet_name="Prod_Volume", index=True)
            utopia_df.loc[utopia_list_plants[ut]]['schd_hours'].to_excel(writer, sheet_name="Schd_Hours", index=True)
            utopia_df.loc[utopia_list_plants[ut]]['prod by SKU'].to_excel(writer, sheet_name="Prod_vs_Demand", index=True)
            utopia_df.loc[utopia_list_plants[ut]]['schd vs avail'].to_excel(writer, sheet_name="Schd_vs_Avail", index=True)
            utopia_df.loc[utopia_list_plants[ut]]['rates'].to_excel(writer, sheet_name="Production_Rates", index=True)
            utopia_df.loc[utopia_list_plants[ut]]['GCmatrix-Attr1'].to_excel(writer, sheet_name="GC_Matrix-Attr1", index=True)
            utopia_df.loc[utopia_list_plants[ut]]['GCmatrix-Attr2'].to_excel(writer, sheet_name="GC_Matrix-Attr2", index=True)
file_loc = brand
print("Asset loading optimization is complete.  Please check", file_loc, "folder for results files.")
Asset loading optimization is complete.  Please check Child Care folder for results files.