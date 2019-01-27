# -*- coding: utf-8 -*-
"""
Created on Wed Oct 25 15:43:08 2017

@author: nguyendu
"""

from matplotlib import pyplot as plt
from matplotlib import font_manager as fm

#fig = plt.figure(1, figsize=(6,6))
#ax = fig.add_axes([1, 1, 1, 1])
#plt.title('Raining Hogs and Dogs')
#
#labels = 'Frogs', 'Hogs', 'Dogs', 'Logs'
#fracs = [15,30,45, 10]
#
#patches, texts, autotexts = ax.pie(fracs, labels=labels, autopct='%1.1f%%',startangle=90)
#
#proptease = fm.FontProperties()
#proptease.set_size('xx-large')
#plt.setp(autotexts, fontproperties=proptease)
#plt.setp(texts, fontproperties=proptease)
#
#plt.show()
##
#
#
#
fig = plt.figure(1, figsize=(5,5))
ax = fig.add_axes([1, 1, 1.0, 1])
labels = 'Feedstock\n procurement','Transportation \n cost','Shortage \n cost','Biofuel \n production \n cost', 'Contract cost', 'Pretreatement cost', 'Storage cost'
sizes = [supplier_cost+ market_cost ,transport_cost,shortage_cost,biofuel_cost, contract_cost, pretreatement_cost, storage_cost]
colors = ['gold', 'lightgreen', 'lightcoral', 'lightskyblue','blue','white','brown']
#explode = (0.1, 0.1, 0.1, 0.1)  # explode 1st slice
 
# Plot
patches, texts, autotexts = ax.pie(sizes, labels=labels,colors=colors, autopct='%1.1f%%', shadow=True, startangle=150)
 #♠plt.axis('equal')
proptease = fm.FontProperties()
proptease.set_size('xx-large')
plt.setp(autotexts, fontproperties=proptease)
plt.setp(texts, fontproperties=proptease)
#plt.title('Total system cost', bbox={'facecolor':'0.4', 'pad':1})
#plt.title('Min Life expectancy across London Regions', fontsize=12)
#plt.text(0.5, 1.5, 'Total system cost',
#         horizontalalignment='right',
#         fontsize=20)
plt.show()