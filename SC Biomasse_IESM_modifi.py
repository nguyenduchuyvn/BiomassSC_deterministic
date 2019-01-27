# -*- coding: utf-8 -*-
"""
Created on Thu Feb  9 14:50:15 2017

@author: nguyendu
"""

'Read data from xlsx file and run model'
'En supprimant des contraintes liê aux stocks de produit brut (biomasse) au fournisseur'
'cost fixe de commande founrnisseur'
import matplotlib.pyplot as plt
from gurobipy import *
import os
import xlrd
"how to read data from xlsx to multidict"
book = xlrd.open_workbook(os.path.join("data_S10_P3_R1_Sce1_de1.xls"))

#book = xlrd.open_workbook(os.path.join("C:\\Users\\nguyendu\\Desktop\\IESM_conference","Database_IESM_S100_P02_T12.xlsx"))

pcost=1.06

"Multidict arcs"
sh = book.sheet_by_name("Arcs")
arcs=[]
distance={}
CapV={}
Vcost={}
i=1
while True:
        try:
            a=sh.cell_value(i, 0),sh.cell_value(i,1)
            value=sh.cell_value(i,2)
            value_capV=sh.cell_value(i,4)
            value_Vcost=sh.cell_value(i,5)
            arcs.append(a)
            distance[a]=value
            CapV[a]=value_capV
            Vcost[a]=value_Vcost
            i = i + 1
        except IndexError:
            break
        
arcs = tuplelist(arcs)        
"Multidict supplier"
sh = book.sheet_by_name("supplier")

supplier=[]
costA_I={}
fixed_cost={}
Q_min_contract={}
i=1
while True:
        try:
            a=sh.cell_value(i, 0)
            supplier.append(a)
            costA_I[a]=sh.cell_value(i, 1)
            fixed_cost[a]=sh.cell_value(i, 2)
            Q_min_contract[a]=sh.cell_value(i, 3)
            i=i+1
        except IndexError:
            break 

"Multidict market"
sh = book.sheet_by_name("Market")
market=[]
costA_M={}
i=1
while True:
        try:
            a=sh.cell_value(i, 0)
            market.append(a)
            costA_M[a]=sh.cell_value(i, 1)
            i=i+1
        except IndexError:
            break 

"Multidict pretraitement"
sh = book.sheet_by_name("Pretraitement")
pretraitement=[]
W_initial={}
W_max,W_min={},{}
Omega_max={}
eta_P={}
lamda_P={}
hcost_P={}
tcost_P={}
i=1
while True:
        try:
            a=sh.cell_value(i, 0)
            pretraitement.append(a)
            W_initial[a]=sh.cell_value(i, 1)
            W_min[a]=sh.cell_value(i, 2)
            W_max[a]=sh.cell_value(i, 3)
            Omega_max[a]=sh.cell_value(i, 4)
            eta_P[a]=sh.cell_value(i, 5)
            lamda_P[a]=sh.cell_value(i, 6)
            hcost_P[a]=sh.cell_value(i, 7)
            tcost_P[a]=sh.cell_value(i, 8)
            i=i+1
        except IndexError:
            break 

"Multidict refinery"
sh = book.sheet_by_name("Refinery")
refinery=[]
V_av_initial={}
V_av_max,V_av_min={}, {}
V_ap_initial={}
V_ap_max, V_ap_min={}, {}
eta_B={}
lamda_B={}
tcost_B={}
hcost_B_av={}
hcost_B_ap={}
gamma_max,gamma_min={},{}
i=1
while True:
        try:          
            a=sh.cell_value(i, 0)
            refinery.append(a)
            V_av_initial[a]=sh.cell_value(i, 1)
            V_av_min[a]=sh.cell_value(i, 2)
            V_av_max[a]=sh.cell_value(i, 3)
            V_ap_initial[a]=sh.cell_value(i, 4)
            V_ap_min[a]=sh.cell_value(i, 5)
            V_ap_max[a]=sh.cell_value(i, 6)
            eta_B[a]=sh.cell_value(i, 7)
            lamda_B[a]=sh.cell_value(i,8)
            tcost_B[a]=sh.cell_value(i, 9)
            hcost_B_av[a]=sh.cell_value(i, 10)
            hcost_B_ap[a]=sh.cell_value(i, 11)
            gamma_min[a]=sh.cell_value(i, 12)
            gamma_max[a]=sh.cell_value(i, 13)
            i=i+1
        except IndexError:
            break 
"dict avai_I"
sh = book.sheet_by_name("Avail_scenario_1")
avail_I={}
i=2
j=1
for i in range(2,sh.nrows):     
    for j in range(1,sh.ncols):
        a=sh.cell_value(i, 0), sh.cell_value(1, j)
        avail_I[a]=sh.cell_value(i, j)

"dict avai_M" " period data"
sh = book.sheet_by_name("Avail_market")
avail_M={}
period=[]
i=2
j=1
for j in range(1,sh.ncols):
    b=sh.cell_value(1, j)
    period.append(b)     
    for i in range(2,sh.nrows):
        a=sh.cell_value(i, 0), sh.cell_value(1, j)
        avail_M[a]=sh.cell_value(i, j)

"dict demand"
sh = book.sheet_by_name("Demand")
demand={}
i=2
j=1
for i in range(2,sh.nrows):     
    for j in range(1,sh.ncols):
        a=sh.cell_value(i, 0), sh.cell_value(1, j)
        demand[a]=sh.cell_value(i, j)
        
        
        
# Create optimization model
m = Model('biomassSC')
m.modelSense = GRB.MINIMIZE
#m.params.TimeLimit = 1800
m.params.threads=1
m.params.NodefileStart=0.5
m.params.Presolve =0
#m.params.MIPGap= 5*1e-4
#m.setParam('OutputFlag', False) # turns off solver chatter

# Create variables

# Transport variable
flow = {} #flow entre noeud i et j à la periode t
nbrT = {} #nbr de tour (camions) entre noeud i et j à la periode t
W = {} # Niveau de stock pour le pretraitement p à la period t
V_av = {} # Niveau de stock (avant la transformation en biocarburant) pour la bioraffinerie b à la period t
V_ap = {} # Niveau de stock (apres la transformation en biocarburant)pour la bioraffinerie b à la period t
z = {} # Quantite utilisee lors de la transformation en biocarburant pour la bioraffinerie b à la period t
Y = {}   #Niveau de stock à la fournisseur i
Q_I = {} #Quantite totale à délivrer à partir du fournisseur i à la periode t
Q_P = {} #Quantite de transformation au pretraitement p
contract = {} # contract long terme avec le fournisseur i
shortage = {} 
demandS = {} # La partie satisfaite à la demande pour la bioraffinerie b a la periode t
x_MB = {}
for i in supplier:
    contract[i] = m.addVar(vtype=GRB.BINARY, name='contract %s' % (i))

for t in period:
    for l in market:
        for b in refinery:
            x_MB[l, b, t]= m.addVar(obj=costA_M[l],name='x_%s_%s_%s' % (l, b, t))


for t in period:        
    for i, j in arcs:
        flow[i, j, t] = m.addVar(name='flow_%s_%s_%s' % (i, j, t))
        nbrT[i, j, t] = m.addVar(vtype=GRB.INTEGER,obj=Vcost[i, j], name='nbrT_%s_%s_%s' % (i, j, t) )

    for i in supplier:
        #Y[i, t] = m.addVar(ub=Y_max[i], name='inventory_supplier %s_%s' % (j, t))
        Q_I[i, t] = m.addVar(obj=costA_I[i], name='Quantity_total_supplier %s_%s' % (i, t))

    for j in pretraitement:
        W[j, t] = m.addVar(lb=W_min[j], ub=W_max[j],obj=eta_P[j]*hcost_P[j],  name='inventory_pre %s_%s' % (j, t))
        Q_P[j, t] = m.addVar(ub=Omega_max[j], obj=tcost_P[j], name='Q_P %s_%s' % (j, t))

    for h in refinery:
        V_av[h, t] = m.addVar(lb=V_av_min[h], ub=V_av_max[h], obj= hcost_B_av[h], name='inventory_bio %s_%s' % (h, t))
        V_ap[h, t] = m.addVar(ub=V_ap_max[h], obj= hcost_B_ap[h], name='inventory_bio %s_%s' % (h, t))
        z[h, t] = m.addVar(lb=gamma_min[h],ub=gamma_max[h], obj=eta_B[h]*tcost_B[h], name='Q_prod %s_%s' % (h, t))
        shortage[h,t] = m.addVar(obj=pcost,name='shortage %s_%s' % (h,t))


m.update()


# inventory level constraints
'balance flows constraint'
for i in supplier:
    # Buying constraint (13)
    m.addConstr(quicksum(Q_I[i, t] for t in period) >= Q_min_contract[i]*contract[i])
    
    
for t in period:
    for i in supplier:            
            m.addConstr(Q_I[i, t] == quicksum(flow[i,p,t] for i, p in arcs.select(i,'*')))
            m.addConstr(Q_I[i, t] <= avail_I[i, t] * contract[i]) 

           
'supply market'
for t in period:
    for l in market:
        m.addConstr(quicksum(x_MB[l,b,t] for b in refinery)<=avail_M[l,t])

'Supply avail_I constraints'

 # inventory level constraints (14)

for j in pretraitement:
    m.addConstr(W[j, 1] == W_initial[j]+ eta_P[j]*quicksum(flow[i, j, 1] for i, j in arcs.select('*', j))
                                        - quicksum(flow[j, h, 1] for j, h in arcs.select(j, '*')))
for t in period:
    if t != 1:
        for j in pretraitement:
            m.addConstr(W[j, t] == lamda_P[j]*W[j, t-1]+ eta_P[j]*quicksum(flow[i, j, t] for i, j in arcs.select('*', j))
                                        - quicksum(flow[j, h, t] for j, h in arcs.select(j, '*')))

# Production capacity of pretraitement usine (15)
for t in period:
    for p in pretraitement:
        m.addConstr(sum(flow[i, p, t] for i, p in arcs.select('*', p)) == Q_P[p,t])
      


 # inventory level constraints (19)
for h in refinery:
    m.addConstr(V_av[h, 1] == V_av_initial[h] + sum(flow[j, h, 1] for j, h in arcs.select('*', h)) - z[h,1]
                                + sum(x_MB[m,h,1] for m in market))

for t in period:
    if t != 1:
        for h in refinery:
            m.addConstr(V_av[h, t] == lamda_B[h]*V_av[h, t - 1]
                                            + sum(flow[j, h, t] for j, h in arcs.select('*', h)) - z[h, t]
                                            + sum(x_MB[m,h,t] for m in market))
       
  # inventory level constraints (21)      
for h in refinery:
    m.addConstr(V_ap[h, 1] == V_ap_initial[h]+ eta_B[h]*z[h, 1] - demand[h, 1]+ shortage[h,1])
    
for t in period:
    if t != 1:
        for h in refinery:   
            m.addConstr(V_ap[h, t] == V_ap[h, t - 1] + eta_B[h]*z[h, t] - demand[h, t]+ shortage[h,t])

#demand contraints
#for t in period:
#    for h in refinery:    
#        m.addConstr(shortage[h,t] == demand[h, t] - demandS[h, t])
            
# Transport constraints

for t in period:
    for i, j in arcs:
        m.addConstr(flow[i, j, t] <= nbrT[i, j, t] * CapV[i, j])


# Objective function
#m.setObjective(
#    # Transport cost
#    #sum(sum(sum(nbrT[i, j, t] * distance[i, j] * Vcost[i, j] for i, j in arcs.select('*', j)) for j in  pretraitement) for t in period)
#    sum(nbrT[i, j, t] * Vcost[i, j] for (i, j) in arcs for t in period)
#    # Pretraitement cost
#    + sum(Q_P[p,t]*tcost_P[p]  for p in pretraitement for t in period)
#    # Backorder cost
#    + sum(shortage[h,t]*pcost for t in period for h in refinery)
#    # convert cost
#    + sum(z[h, t] * tcost_B[h] for h in refinery for t in period)
#
#    # holding cost
#    + sum(W[j, t]*hcost_P[j] for j in pretraitement for t in period)
#    + sum(V_av[h, t]*hcost_B_av[h] for h in refinery for t in period)
#    + sum(V_ap[h, t]*hcost_B_ap[h] for h in refinery for t in period)
#    # Buying cost
#    + sum(Q_I[i, t] * costA_I[i] for i in supplier for t in period )
#    + sum(x_MB[l, b, t]*costA_M[l] for l in market for b in refinery for t in period)
#    # fixed cost
#    +quicksum(contract[i]* fixed_cost[i] for i in supplier)
#)
# Compute optimal solution
m.optimize()


#if m.status == GRB.Status.OPTIMAL:
# Print solutions' values
flow = m.getAttr('x', flow)
nbrT = m.getAttr('x', nbrT)
W = m.getAttr('x', W)
V_ap = m.getAttr('x', V_ap)
V_av = m.getAttr('x', V_av)
contract = m.getAttr('x', contract)
Q_I = m.getAttr('x', Q_I)
Q_P = m.getAttr('x', Q_P)
z= m.getAttr('x', z)
x_MB=m.getAttr('x', x_MB)
demandS= m.getAttr('x', demandS)
shortage = m.getAttr('x', shortage)
#inventory=W
#
#for j in pretraitement:
#    for t in period:
#        print('inventory level %s at period %s: %4.2f' % (j, t, W[j, t]))
#print('########### Refinery ################')
#for h in refinery:
#    for t in period:
##        print('inventory level %s AV  at period %s: %4.2f' % (h, t, V_av[h, t]))
##        print('inventory level %s  AP at period %s: %4.2f' % (h, t, V_ap[h, t]))
#        print('Q production %s at %s: %4.2f  ' % (h, t, z[h, t]))
#        print('shortage %s at %s: %s  ' % (h,t,shortage[h,t]))
#print('########### contrat ################')
#for i in supplier:
#    if contract[i]:
#        print('contract was signed with %s' % i)
##        for t in period:
##            if Q_I[i, t]>0:
##                print('Q harvest %s at period %s: %4.2f' % (i, t, Q_I[i, t]))
#
#
##                print('%s -> %s at period %s: %4.2f ' % (i, j, t, flow[i, j, t]))
#
#print('########### Demand statisfaction ################' )
#for t in period:
#    for h in refinery:
#        print('shortage at %s in period %s is: %4.2f ' % (h, t, shortage[h, t]))
#print('########### COST ################')        
#
#print('Pretraitement cost=%4.2f'% sum(Q_P[p,t]*tcost_P[p]*eta_P[p]  for p in pretraitement for t in period))
#
#print('Backorder cost =%4.2f'% sum(shortage[h,t] * pcost for h in refinery for t in period))
#
#print('Buying costnhap=%4.2f' % (sum(flow[i, j, t]*costA_I[i] for t in period for i in supplier for i, j in arcs.select(i,'*')   )
#                + sum(x_MB[l, b, t]*costA_M[l] for l in market for b in refinery for t in period)))
#
##print('Buying cost=%4.2f' % (sum(flow[i, j, t]*costA_I[i] for i in supplier for i, j in arcs.select(i,'*')  for t in period)
#                            #+sum(x_MB[l, b, t]*costA_M[l] for l in market for b in refinery for t in period)))
#
#print('holding cost=%4.2f' % (sum(W[j, t]*hcost_P[j] for j in pretraitement for t in period)
#    + sum(V_av[h, t]*hcost_B_av[h] for h in refinery for t in period)
#    + sum(V_ap[h, t]*hcost_B_ap[h] for h in refinery for t in period)))
#
#print('transport cost=%4.2f' % sum(nbrT[i, j, t]* distance[i, j]* Vcost[i, j] for (i, j) in arcs for t in period))
#
#print('Convert cost=%4.2f'% sum(eta_B[h]*z[h, t] * tcost_B[h] for h in refinery for t in period))
##
#
Pretraitement_cost=sum(Q_P[p,t]*tcost_P[p]*eta_P[p]  for p in pretraitement for t in period)
Backorder_cost =sum(shortage[h,t] * pcost for h in refinery for t in period)
Purchase_cost=sum(flow[i, j, t]*costA_I[i] for t in period for i in supplier for i, j in arcs.select(i,'*'))
#
##   + sum(x_MB[l, b, t]*costA_M[l] for l in market for b in refinery for t in period)
#
#print(Purchase_cost)

holding_cost = (sum(V_av[h, t]*hcost_B_av[h] for h in refinery for t in period)
            +sum(W[p, t]*hcost_P[p] for p in pretraitement for t in period)
            + sum(V_ap[h, t]*hcost_B_ap[h] for h in refinery for t in period))
transport_cost=sum(nbrT[i, j, t]*Vcost[i, j] for (i, j) in arcs for t in period)
Biofuel_production_cost=sum(eta_B[h]*z[h, t] * tcost_B[h] for h in refinery for t in period)



# Pie chart, where the slices will be ordered and plotted counter-clockwise:
labels = 'pretraitement cost', 'backorder cost', 'holding cost','transport cost','purchase cost','Biofuel production cost'
sizes = [Pretraitement_cost, Backorder_cost, holding_cost,transport_cost,Purchase_cost, Biofuel_production_cost]
explode = (0.1, 0,0, 0.2, 0, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')

fig1, ax1 = plt.subplots()
ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
        shadow=True, startangle=90)
ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

plt.show()


