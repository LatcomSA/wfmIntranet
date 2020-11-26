# -*- coding: utf-8 -*-
"""
Created on Mon Oct 19 10:24:28 2020

@author: jasson
"""

import pandas as pd
import datetime
from scipy.stats import poisson
import math
import numpy as np
import openpyxl 
import networkx as nx
from math import nan
import openpyxl 


exporte = pd.read_csv('exportehorario.csv', sep=',', header=0, index_col=None).to_numpy()
agentes = pd.read_excel('agentes.xlsx', sheet_name='Sheet1',header=None).to_numpy()

fechas = []
ubicacion = []
for x in range(exporte.size):    
    try:
        fechas.append(datetime.datetime.strptime(exporte[x,0].strip(';'), "%m/%d/%Y"))
        ubicacion.append(x)
    except:
        pass 

del ubicacion[0]
separacion = np.split(exporte,ubicacion) 


class sched:
    def __init__(self,fecha):
        self.fecha = fecha
        self.graph = nx.Graph()

sched_agent = {x:sched(fechas[x]) for x in range(len(fechas))}

total_names = []
count = 0    
for z in separacion:
    info_day = np.delete(z, 0, axis=0)
    for y in info_day:
        start_name = y[0].find(',')
        name1 = y[0][0:start_name]
        name2 = y[0][start_name+2:]
        name = name2 + ' ' + name1
        sched_agent.get(count).graph.add_node(name)
        if count == 0:
           total_names.append(name) 
        try:
            start = y[4].find('Immediate')
            sep = y[4][start+10:].split('-')
            hour_init = sep[0].split(' ')[1].split(':')                                      
            hour_stop = sep[1].split(' ')[1].split(':')
            if sep[0].split(' ')[2]  == 'AM':
                init_del= datetime.timedelta(hours=int(hour_init[0]),minutes=int(hour_init[1]))
                minus = init_del - datetime.timedelta(hours=0,minutes=15)
                dec,ent = math.modf(minus.seconds/(60*60))
                init = datetime.time(int(hour_init[0]),int(hour_init[1]))
                preturno =  datetime.time(int(ent),int(dec*60))
            else:
                init_del = datetime.timedelta(hours=int(hour_init[0])+12,minutes=int(hour_init[1]))
                minus = init_del - datetime.timedelta(hours=0,minutes=15)
                dec,ent = math.modf(minus.seconds/(60*60))
                init = datetime.time(int(hour_init[0]),int(hour_init[1]))
                preturno =  datetime.time(int(ent),int(dec*60))               
            if sep[1].split(' ')[2]  == 'AM' or int(hour_stop[0]) == 12:
                stop = datetime.time(int(hour_stop[0]),int(hour_stop[1]))
            else:
                stop = datetime.time(int(hour_stop[0])+12,int(hour_stop[1])) 
                
            sched_agent.get(count).graph.add_edge(name,"{}_{}".format('Preturno',name), time = preturno)
            sched_agent.get(count).graph.add_edge(name,"{}_{}".format('Init_turno',name), time = init)
            sched_agent.get(count).graph.add_edge(name,"{}_{}".format('Stop_turno',name), time = stop)
            
            try:
                start_breakI = y[5].find('"Break"')
                end_breakI = y[5].find(';',start_breakI)
                sep_breakI = y[5][start_breakI+8:end_breakI].split('-')
                hour_init_breakI = sep_breakI[0].split(' ')[1].split(':')
                if sep_breakI[0].split(' ')[2]  == 'AM' or int(hour_init_breakI[0]) == 12:
                   init_breakI = datetime.time(int(hour_init_breakI[0]),int(hour_init_breakI[1]))
                else:
                   init_breakI = datetime.time(int(hour_init_breakI[0])+12,int(hour_init_breakI[1]))
                sched_agent.get(count).graph.add_edge(name,"{}_{}".format('BreakI',name), time = init_breakI)
            except:
                sched_agent.get(count).graph.add_edge(name,"{}_{}".format('BreakI',name), time = nan )
            
            try: 
                start_breakII = y[5][end_breakI:].find('"Break"')
                end_breakII = y[5][end_breakI:].find(';',start_breakII)
                sep_breakII = y[5][end_breakI:][start_breakII+8:end_breakII].split('-')
                hour_init_breakII = sep_breakII[0].split(' ')[1].split(':')
                if sep_breakII[0].split(' ')[2]  == 'AM' or int(hour_init_breakII[0]) == 12:
                   init_breakII = datetime.time(int(hour_init_breakII[0]),int(hour_init_breakII[1])) 
                else:
                   init_breakII = datetime.time(int(hour_init_breakII[0])+12,int(hour_init_breakII[1]))
                sched_agent.get(count).graph.add_edge(name,"{}_{}".format('BreakII',name), time = init_breakII)                  
            except:
                sched_agent.get(count).graph.add_edge(name,"{}_{}".format('BreakII',name), time = nan)
            
            try:
                start_lunch = y[5].find('Lunch')
                end_lunch = y[5].find(';',start_lunch)
                sep_lunch = y[5][start_lunch+7:end_lunch].split('-')
                hour_init_lunch = sep_lunch[0].split(' ')[1].split(':')
                if sep_lunch[0].split(' ')[2] == 'AM' or int(hour_init_lunch[0]):
                   init_lunch = datetime.time(int(hour_init_lunch[0]),int(hour_init_lunch[1])) 
                else:
                   init_lunch = datetime.time(int(hour_init_lunch[0])+12,int(hour_init_lunch[1])) 
                sched_agent.get(count).graph.add_edge(name,"{}_{}".format('Lunch',name), time = init_lunch)  
            except:    
                sched_agent.get(count).graph.add_edge(name,"{}_{}".format('Lunch',name), time = nan)
        except:
            sched_agent.get(count).graph.add_edge(name,"{}_{}".format('Preturno',name), time = nan)
            sched_agent.get(count).graph.add_edge(name,"{}_{}".format('Init_turno',name), time = nan)
            sched_agent.get(count).graph.add_edge(name,"{}_{}".format('Stop_turno',name), time = nan)                           
    count += 1   

total_auxs = ['Preturno','Init_turno','Stop_turno','BreakI','BreakII','Lunch']




up_massive = []
for names in total_names:
    for agent in sched_agent.values():
          try:
              turn_in = agent.graph[names]["Init_turno_{}".format(names)]['time'].strftime('%H:%M')
              turn_out = agent.graph[names]["Stop_turno_{}".format(names)]['time'].strftime('%H:%M')
              turn_total = "{} - {}".format(turn_in,turn_out)
              break_row = np.zeros([1,7],dtype='O')

              count_name = 0
              find = True
              for info_name in agentes[:,0]:
                  if names in info_name:
                     break_row[0,0] = info_name
                     break_row[0,1] = agentes[count_name,1]
                     find = False
                     break
                  count_name += 1    
                  
              if find:
                  break_row[0,0] = names
                  break_row[0,1] = nan 

              break_row[0,2] = agent.fecha.strftime('%d/%m/%Y')
              break_row[0,3] = agent.fecha.strftime('%d/%m/%Y') 
              break_row[0,4] = turn_total
              break_row[0,5] = 'Break'
              break_row[0,6] = agent.graph[names]["BreakI_{}".format(names)]['time'].strftime('%H:%M')
              up_massive.append(break_row)                                      
          except:
              pass


for names in total_names:
    for agent in sched_agent.values():
          try:
              turn_in = agent.graph[names]["Init_turno_{}".format(names)]['time'].strftime('%H:%M')
              turn_out = agent.graph[names]["Stop_turno_{}".format(names)]['time'].strftime('%H:%M')
              turn_total = "{} - {}".format(turn_in,turn_out)             
              breakII_row = np.zeros([1,7],dtype='O')
              
              count_name = 0
              find = True
              for info_name in agentes[:,0]:
                  if names in info_name:
                     breakII_row[0,0] = info_name
                     breakII_row[0,1] = agentes[count_name,1]
                     find = False
                     break
                  count_name += 1    
                  
              if find:
                  breakII_row[0,0] = names
                  breakII_row[0,1] = nan              
              
              breakII_row[0,2] = agent.fecha.strftime('%d/%m/%Y')
              breakII_row[0,3] = agent.fecha.strftime('%d/%m/%Y') 
              breakII_row[0,4] = turn_total
              breakII_row[0,5] = 'BreakII'
              breakII_row[0,6] = agent.graph[names]["BreakII_{}".format(names)]['time'].strftime('%H:%M')
              up_massive.append(breakII_row)                           
          except:
              pass

for names in total_names:
    for agent in sched_agent.values():
          try:
              turn_in = agent.graph[names]["Init_turno_{}".format(names)]['time'].strftime('%H:%M')
              turn_out = agent.graph[names]["Stop_turno_{}".format(names)]['time'].strftime('%H:%M')
              turn_total = "{} - {}".format(turn_in,turn_out)             
              Preturno_row = np.zeros([1,7],dtype='O')
              
              count_name = 0
              find = True
              for info_name in agentes[:,0]:
                  if names in info_name:
                     Preturno_row[0,0] = info_name
                     Preturno_row[0,1] = agentes[count_name,1]
                     find = False
                     break
                  count_name += 1    
                  
              if find:
                  Preturno_row[0,0] = names
                  Preturno_row[0,1] = nan              
              
              Preturno_row[0,2] = agent.fecha.strftime('%d/%m/%Y')
              Preturno_row[0,3] = agent.fecha.strftime('%d/%m/%Y') 
              Preturno_row[0,4] = turn_total
              Preturno_row[0,5] = 'Preturno'
              Preturno_row[0,6] = agent.graph[names]["Preturno_{}".format(names)]['time'].strftime('%H:%M')
              up_massive.append(Preturno_row)                           
          except:
              pass

for names in total_names:
    for agent in sched_agent.values():
          try:
              turn_in = agent.graph[names]["Init_turno_{}".format(names)]['time'].strftime('%H:%M')
              turn_out = agent.graph[names]["Stop_turno_{}".format(names)]['time'].strftime('%H:%M')
              turn_total = "{} - {}".format(turn_in,turn_out)              
              Lunch_row = np.zeros([1,7],dtype='O')

              count_name = 0
              find = True
              for info_name in agentes[:,0]:
                  if names in info_name:
                     Lunch_row[0,0] = info_name
                     Lunch_row[0,1] = agentes[count_name,1]
                     find = False
                     break
                  count_name += 1    
                  
              if find:
                  Lunch_row[0,0] = names
                  Lunch_row[0,1] = nan
              
              Lunch_row[0,2] = agent.fecha.strftime('%d/%m/%Y')
              Lunch_row[0,3] = agent.fecha.strftime('%d/%m/%Y') 
              Lunch_row[0,4] = turn_total
              Lunch_row[0,5] = 'Almuerzo'
              Lunch_row[0,6] = agent.graph[names]["Lunch_{}".format(names)]['time'].strftime('%H:%M')
              up_massive.append(Lunch_row)                           
          except:
              pass


subida_masiva = openpyxl.Workbook()
sheet_subida = subida_masiva.active
sheet_subida.append(('Nombre','Documento','Fecha Horario', 'Fecha Otro Horario','Horario','Tipo Otro Horario','Horario Otro Horario'))
for massive in up_massive:
    sheet_subida.append(tuple(massive[0]))
subida_masiva.save('./subida_masiva.xlsx')




