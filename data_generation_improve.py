# -*- coding: utf-8 -*-
"""
Created on Wed Apr 19 13:56:34 2017

@author: nguyendu
"""
from gurobipy import *

import matplotlib.pyplot as plt
import scipy.stats as stats
import numpy as numpy
from tempfile import TemporaryFile
from xlwt import Workbook


#lower, upper, mu, and sigma are four parameters
lower, upper = 2016, 2280
mu_D, sigma_D = 2133, 4138


supplier_range=[10]
nbr_period=12 
harvest_time= [1,2,3]
scenarios_range=[1]
nbr_refinery=1
nbr_pretraitement=3
nbr_market=1
demand_range=[i for i in range(1,11)]
for nbr_scenarios in scenarios_range:
    for nbr_supplier in supplier_range:
        for coeff_demand in demand_range:
            name_file='data_S%s_P%s_R%s_Sce%s_de%s.xls'% (nbr_supplier,nbr_pretraitement,nbr_refinery, nbr_scenarios,coeff_demand)       
            print(name_file)      
        
            scenarios= [i for i in range(1,nbr_scenarios+1)] 
            #nbr_supplier in each zone
            
            #nbr_zone=4
            
            demand,avail_I, area,probability={},{},{},{}
            
            #print(name_file)
            #pcost=16;
            #refinery = list(refinery)
              
            #creat supplier name
            supplier= ['Supplier_%s_%s'% (i,j) for i in range(1,nbr_supplier+1) for j in range(1, nbr_pretraitement+1)]
            refinery= ['Refinery_%s'% i for i in range(1,nbr_refinery+1)]
            pretraitement= ['Pretraitement_%s'% i for i in range(1,nbr_pretraitement+1)]
            market= ['Market_%s'% i for i in range(1,nbr_market+1)]
            
            for s in scenarios:
                probability[s]=1/(len(scenarios))
            
            book = Workbook()
            " Scenarios"
            sheet1 = book.add_sheet('Scenario')
            sheet1.write(0,0,'Scenario')
            for s in range(1,nbr_scenarios + 1):
                sheet1.write(s, 0 , s)
            
            del sheet1
            " Generate demand in different scenarios    "
            
                #instantiate an object X using the above four parameters,
            X = stats.truncnorm((lower - mu_D) / sigma_D, (upper - mu_D) / sigma_D, loc=mu_D, scale=sigma_D)
            #generate 1000 sample data
            samples= X.rvs(nbr_period)
            for t in range(1,nbr_period+1):
                for i in refinery:
                    demand[i,t]= round(samples[t-1])*1000000/12/12/1.5
            
            sheet1 = book.add_sheet('Demand')
            sheet1.write(1,0,'Refinery')
            #sheet1.write(0,1,'Period')
            #print(refinery)
            #print(demand)
            for b in range(nbr_refinery):
                sheet1.write(b+2,0, refinery[b])
            for p in range(1,nbr_period+1):
                sheet1.write(0, p , 'period')
                sheet1.write(1, p , p)
            
            for b in range(nbr_refinery):    
                for t in range(1,nbr_period+1):
                    B=refinery[b]
                    sheet1.write(b+2, t , demand[B,t])
            
            del sheet1
            "Generate the availability of supplier in different scenarios "
            ## Quantité dispobile en fournisseur i à la periode t 
            ## avail_I[i, t, s]  i= supplier; t = period; s =scénario
            
            avail_I,area={}, {}
            #Ơ Generate surface of cultivation 
            for i in supplier:
                area[i]= numpy.random.randint(500,2000)    
                
            for s in scenarios:
                samples_rendement= numpy.random.uniform(3.5,4.5,nbr_period)
                for i in supplier:        
                    for t in range(1,nbr_period+1):
                        if t in harvest_time:
                            avail_I[i,t]=round(samples_rendement[t-1]*area[i])
                        else:
                            avail_I[i,t]=0            
            
                sheet2 = book.add_sheet('Avail_scenario_%s'% s)
                sheet2.write(1,0,'Supplier')
               
                for t in range(1,nbr_period+1):
                    sheet2.write(0, t , 'period')
                    sheet2.write(1, t , t)
                ligne_index=2
                for i in supplier:
                    sheet2.write(ligne_index,0, i)
                    for t in range(1,nbr_period+1):            
                        sheet2.write(ligne_index, t , avail_I[i,t])
                    ligne_index+=1
            
                del sheet2         
            book.save(name_file)        
            "Generate arcs"
            #Generate arcs
            arcs= []
            CapV,Vcost,  unit_cost , distance={},{},{},{}
            #for i in supplier:
            #    for j in pretraitement:
            #        x=i,j
            #        arcs.append(x)
            #
            #
            for j in pretraitement:
                for b in refinery:
                    x=j,b
                    arcs.append(x)
            for i in range(1,nbr_supplier+1) :
                for j in range(1, nbr_pretraitement+1):
                    x= 'Supplier_%s_%s'% (i,j), 'Pretraitement_%s'% j
                    arcs.append(x)      
                    
            
            #arcs= ('Supplier_%s_%s', 'Pretraitement_%s' % (i,j,j) for i in range(1,nbr_supplier+1) for j in range(1, nbr_pretraitement+1))
            
            # generate parameter on arcs
            for (i,j) in arcs:
                distance[i,j]=numpy.random.randint(10,50)
                CapV[i,j]= 25
                unit_cost[i,j]=0.02
                Vcost[i,j]= distance[i,j]*unit_cost[i,j]
            
                    
            sheet2 = book.add_sheet('Arcs')
            sheet2.write(0,0,'Arcs i') 
            sheet2.write(0,1,'Arcs j') 
            sheet2.write(0,2,'Distance')    
            sheet2.write(0,3,'Unite cost') 
            sheet2.write(0,4,'Cap_V') 
            sheet2.write(0,5,'Vcost')
            ligne_index=1
            for (i,j) in arcs:
                sheet2.write(ligne_index,0,i)
                sheet2.write(ligne_index,1,j)
                sheet2.write(ligne_index,2,distance[i,j])
                sheet2.write(ligne_index,3,unit_cost[i,j])
                sheet2.write(ligne_index,4,CapV[i,j])
                sheet2.write(ligne_index,5,Vcost[i,j])
                ligne_index+=1
                
            book.save(name_file)
            del sheet2            
                        
            costA_I,mcost,Q_min_contract={},{},{}  
            "Generate input parameters related to supplier"
            for i in supplier:
                costA_I[i]=35
                mcost[i]=area[i]*5
                Q_min_contract[i]=area[i]*3
            
            sheet2 = book.add_sheet('supplier')
            sheet2.write(0,0,'Supplier')
            sheet2.write(0,1,'costA_I')
            sheet2.write(0,2,'mcost')
            sheet2.write(0,3,'Q_min_contract')
            ligne_index=1
            for i in supplier:
                sheet2.write(ligne_index, 0 , i)
                sheet2.write(ligne_index, 1 , costA_I[i])
                sheet2.write(ligne_index, 2 , mcost[i])
                sheet2.write(ligne_index, 3 ,  Q_min_contract[i])
                ligne_index+=1
            
            del sheet2
            
            
            "Generate input parameters related to market  "
            costA_M,avail_M={},{} 
            for i in market:
#                costA_M[i]=numpy.random.randint(80,100)
                costA_M[i]=145
#                
                ## Quantité dispobile en fournisseur libre m à la periode t     
                for t in range(1,nbr_period+1):
                    avail_M[i,t]= 1000000
                    
            sheet2 = book.add_sheet('Market')
            sheet2.write(0,0,'Market')
            sheet2.write(0,1,'costA_M')        
            ligne_index=1
            for i in market:
                sheet2.write(ligne_index, 0 , i)
                sheet2.write(ligne_index, 1 , costA_M[i])
                ligne_index+=1
                
            sheet2 = book.add_sheet('Avail_market')
            sheet2.write(1,0,'Market')
            for t in range(1,nbr_period+1):
                sheet2.write(0, t , 'period')
                sheet2.write(1, t , t)   
                
            for i in range(0,nbr_market):
                    I=market[i]
                    sheet2.write(i+2,0, I)
                    for t in range(1,nbr_period+1):            
                        sheet2.write(i+2, t , avail_M[I,t])    
             
            del sheet2
            
            "Generate input related to pretraitement"
            ## W_max: Stock maximal pour le pretraitement
            ## W_initial: Stock initial pour le pretraitement
            ## gamma_max_P: Capaacité de transformation pour le pretraitement
            ## eta_P: Taux de transformation pour le pretraitement
            ## lamda_P: Taux de deterioration pour le pretraiement
            ## hcost_P: cost unitaire de stockage pour le pretraiement
            ## tcost_P: cost unitaire de transformation pour le pretraiement
                    
            W_initial,W_min, W_max, gamma_min_P={},{},{},{}
            gamma_max_P,  eta_P, lamda_P, hcost_P, tcost_P={},{},{},{},{}
            for p in pretraitement:
                W_initial[p]=0    
                W_min[p]=0
                W_max[p]=1000000/6   
                gamma_min_P[p]=0
                gamma_max_P[p]=1000000/12
                eta_P[p]=0.80
                lamda_P[p]=0.95
                hcost_P[p]=1.125
                tcost_P[p]=13.39           
                        
            sheet2 = book.add_sheet('Pretraitement')
            sheet2.write(0,0,'Pretraitement')
            sheet2.write(0,1,'W_initial')    
            sheet2.write(0,2,'W_min')    
            sheet2.write(0,3,'W_max')    
            sheet2.write(0,4,'gamma_min_P')    
            sheet2.write(0,5,'gamma_max_P')    
            sheet2.write(0,6,'eta_P')  
            sheet2.write(0,7,'lamda_P')    
            sheet2.write(0,8,'hcost_P')   
            sheet2.write(0,9,'tcost_P')     
            ligne_index=1
            
            for p in pretraitement:
                sheet2.write(ligne_index, 0 ,  p)
                sheet2.write(ligne_index, 1 ,  W_initial[p])
                sheet2.write(ligne_index, 2 ,  W_min[p])
                sheet2.write(ligne_index, 3 ,  W_max[p])
                sheet2.write(ligne_index, 4 ,  gamma_min_P[p])
                sheet2.write(ligne_index, 5 ,  gamma_max_P[p])
                sheet2.write(ligne_index, 6 ,  eta_P[p])
                sheet2.write(ligne_index, 7 ,  lamda_P[p])
                sheet2.write(ligne_index, 8 ,  hcost_P[p])
                sheet2.write(ligne_index, 9 ,  tcost_P[p])
                ligne_index+=1
             
            del sheet2
            "Generate input related to refinery "
            ## V_av_max: Stock maximal (avant la transformation en carburant) pour la bioraffinerie
            ## V_ap_max: Stock maximal (apres la transformation en carburant) pour la bioraffinerie
            ## eta_B: Taux de transformation pour la bioraffinerie
            ## lamda_B: Taux de deterioration pour la bioraffinerie
            ## hcost_B: cost unitaire de stockage pour la bioraffinerie
            ## tcost_B: cout unitaire de transformation en biocarubrant pour la bioraffinerie
            ## gamma_max_B: Capacite de transformation en biocarburant pour la bioraffinerie
            
                
            V_av_initial, V_av_min, V_av_max, V_ap_initial,V_ap_min, V_ap_max={},{},{},{},{},{}
            eta_B, lamda_B, tcost_B, hcost_B_av, hcost_B_ap, gamma_min_B, gamma_max_B={},{},{},{},{},{},{}
            for b in refinery:
                 V_av_initial[b]=0  
                 V_av_min[b]=0
                 V_av_max[b]=1000000/6
                 V_ap_initial[b]=0
                 V_ap_min[b]=0
                 V_ap_max[b]=100*1000000/0.264/12
                 eta_B[b]=313
                 lamda_B[b]=0.99
                 tcost_B[b]=0.2
                 hcost_B_av[b]=0.9
                 hcost_B_ap[b]=0.2272
                 gamma_min_B[b]=0
                 gamma_max_B[b]=1000000/12
            
            
            sheet2 = book.add_sheet('Refinery')
            sheet2.write(0,0,'Refinery')
            sheet2.write(0,1,'V_av_initial')    
            sheet2.write(0,2,'V_av_min')    
            sheet2.write(0,3,'V_av_max')    
            sheet2.write(0,4,'V_ap_initial')    
            sheet2.write(0,5,'V_ap_min')    
            sheet2.write(0,6,'V_ap_max')  
            sheet2.write(0,7,'eta_B')    
            sheet2.write(0,8,'lamda_B')   
            sheet2.write(0,9,'tcost_B') 
            sheet2.write(0,10,'hcost_B_av')   
            sheet2.write(0,11,'hcost_B_ap')   
            sheet2.write(0,12,'gamma_min_B')   
            sheet2.write(0,13,'gamma_max_B')       
            ligne_index=1
            for b in refinery:
                sheet2.write(ligne_index, 0 ,  b)
                sheet2.write(ligne_index, 1 ,  V_av_initial[b])
                sheet2.write(ligne_index, 2 ,  V_av_min[b])
                sheet2.write(ligne_index, 3 ,  V_av_max[b])
                sheet2.write(ligne_index, 4 ,  V_ap_initial[b])
                sheet2.write(ligne_index, 5 ,  V_ap_min[b])
                sheet2.write(ligne_index, 6 ,  V_ap_max[b])
                sheet2.write(ligne_index, 7 ,  eta_B[b])
                sheet2.write(ligne_index, 8 ,  lamda_B[b])
                sheet2.write(ligne_index, 9 ,  tcost_B[b])
                sheet2.write(ligne_index, 10 ,  hcost_B_av[b])
                sheet2.write(ligne_index, 11 ,  hcost_B_ap[b])
                sheet2.write(ligne_index, 12 ,  gamma_min_B[b])
                sheet2.write(ligne_index, 13 ,  gamma_max_B[b])
                ligne_index+=1
            
            
                    
            book.save(name_file)
            #book.save(TemporaryFile())
            
