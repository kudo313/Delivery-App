import subprocess
import sys


import numpy as np
import copy

import pandas as pd
import PyQt5 
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, Qt, pyqtSignal
from PyQt5.QtWidgets import *
import random


demand = []
segment = 0
total_dm = 0

x = []
y = []
visited = []
vehi_cap = []
cap = 0
num_vehi = 0
distance_mt = []
num_destination = 0
min_dis = 10000000
rs = []
load = []
distance_route = []
def isnan(value):
    try:
        import math
        return math.isnan(float(value))
    except:
        return True
def create_data_model():
    """Stores the data for the problem."""
    x = pd.read_excel('data/data.xlsx',0,header=0, index_col= False)
    sheet_2 = pd.read_excel('data/data.xlsx',sheet_name = 1,header=None, index_col= False)
    data = {}

    title_sheet2 = sheet_2.columns.ravel()
    data['list_cap'] = []
    for t in title_sheet2:
        for k in sheet_2[t].tolist():
            if isnan(k) == False:
                data['list_cap'].append(k)
    data['list_cap'].reverse()

    # print("/n")
    y = x.columns.ravel()
    # print((x[y[0]].tolist()))
    z = []
    # print(x.to_dict(orient = 'record')[1][2])
    # m = x.to_dict(orient = 'record')

    data['wd_names'] = []
    for t in range(len(y)):
        if (y[t].find('Unnamed') == -1):
            if (t > 1):
                z.append(x[y[t]].tolist())
            
            data['wd_names'].append(y[t])
    count_col = 0
    for i in range(len(z)):
        count_col = 0
        while True:
            if not (isnan(z[i][count_col])):
                count_col+=1
            else:
                z[i].remove(z[i][count_col])
            if count_col == len(z[i]) :
                break

    
    for v in z:
        for k in range(len(v)):
            if isnan(v[k]):
                v[k] = 0
    z = np.array(z)
    z = z.T
    z = z.tolist()
    data['distance_matrix'] = z
    data['distance_matrix'] = ( [list( map(int,i) ) for i in data['distance_matrix']] )
    data['demands'] = []
    for p in range(len(data['distance_matrix'])):
        data['demands'].append(0)

    
    data['wd_names_now'] = []
    return data

data1 = create_data_model()
or_data = create_data_model()
rs = []
def modify_data(data,check_run):
    data['wd_names_now'] = []
    goc_appear = False
    data['ori_len'] = len(data['demands'])
    demand1 = copy.deepcopy(data['demands'])
    distance_matrix1 = copy.deepcopy(data['distance_matrix'])
    data['demands'] = [0]
    data['distance_matrix'] = []
    tmp1 =[0]
    data['distance_matrix'].append(tmp1)
    
    index_need_demand= []
    count = 0
    if demand1[0] == 0:
        goc_appear = False
    else:
        goc_appear = True
    if goc_appear == False:
        data['wd_names_now'].append(data['wd_names'][0])

        for p in range(len(demand1)):
            if demand1[p] != 0:
                count+=1
                data['demands'].append(demand1[p])
                data['wd_names_now'].append(data['wd_names'][p])

                data['distance_matrix'][0].append(distance_matrix1[0][p])
                tmp = []
                tmp.append(distance_matrix1[p][0])
                data['distance_matrix'].append(tmp)
                tmp_dic = {'pre' : p, 'now' : count}
                index_need_demand.append(tmp_dic)
        for p in range(1,len(data['distance_matrix'])):
            for t in range(1,len(data['distance_matrix'][0])):
                if t == p:
                    data['distance_matrix'][p].append(0)
                else:
                    for v in index_need_demand:
                        if v['now'] == p:
                            row = v['pre']
                        if v['now'] == t:
                            col = v['pre']
                    data['distance_matrix'][p].append(distance_matrix1[row][col])
    else:
        data['wd_names_now'].append(data['wd_names'][0])

        for p in range(1,len(demand1)):
            if demand1[p] != 0:
                count+=1
                data['demands'].append(demand1[p])
                data['wd_names_now'].append(data['wd_names'][p])

                data['distance_matrix'][0].append(distance_matrix1[0][p])
                tmp = []
                tmp.append(distance_matrix1[p][0])
                data['distance_matrix'].append(tmp)
                tmp_dic = {'pre' : p, 'now' : count}
                index_need_demand.append(tmp_dic)
        for p in range(1,len(data['distance_matrix'])):
            for t in range(1,len(data['distance_matrix'][0])):
                if t == p:
                    data['distance_matrix'][p].append(0)
                else:
                    for v in index_need_demand:
                        if v['now'] == p:
                            row = v['pre']
                        if v['now'] == t:
                            col = v['pre']
                    data['distance_matrix'][p].append(distance_matrix1[row][col])
    data['vehicle_capacities'] = []
    data['num_vehicles'] = 0
    data['depot'] = 0
    total_demand = 0
    if check_run == True:
        data['du_thua'] = copy.deepcopy(data['demands'])
        tmp_divine = 0
        for i in range(len(data['demands'])):
            if data['demands'][i] > data['list_cap'][0]:
                if (data['demands'][i] % data['list_cap'][0] != 0):
                    tmp_divine = data['demands'][i]//data['list_cap'][0]
                    data['demands'][i] = data['demands'][i] - data['list_cap'][0]*tmp_divine
                    data['du_thua'][i] = tmp_divine

                else:
                    tmp_divine = (data['demands'][i]//data['list_cap'][0]) - 1
                    data['demands'][i] = data['list_cap'][0]
                    data['du_thua'][i] = tmp_divine

            else:
                data['du_thua'][i] = 0

    copy_demand = copy.deepcopy(data['demands'])

    for i in data['demands']:
        total_demand += i
    total_cap = 0
    check = True
    posi_cap = 0

    copy_demand.sort(reverse=True)
    nowload = []
    for i in range(len(copy_demand)):
        checkMaintain = False
        for j in range(len(nowload)):
            if nowload[j] + copy_demand[i] <= data['list_cap'][0]:
                nowload[j] += copy_demand[i]
                checkMaintain = True
                break
        if checkMaintain == False :
            nowload.append(copy_demand[i])
    for i in range(len(nowload)):
        data['vehicle_capacities'].append(data['list_cap'][0])

    global demand
    global vehi_cap
    global cap
    global distance_mt
    global num_vehi
    global segment
    global num_destination
    global rs
    global x
    global y
    global load
    global distance_route
    global visited
    global total_dm
    demand = data['demands']
    vehi_cap = data['vehicle_capacities']
    cap = vehi_cap[0]
    num_vehi = len(vehi_cap)
    distance_mt = data['distance_matrix']
    segment = 0
    num_destination = len(demand) - 1
    x = np.zeros(len(demand),dtype=int)
    y = np.zeros(len(demand),dtype=int)
    load = np.zeros(len(demand),dtype=int)
    distance_route = np.zeros(num_vehi+1,dtype=int)
    for i in range(len(demand)):
        visited.append(False)
    for i  in demand:
        total_dm += i
    return data['du_thua']
# [0, 1, 1, 3, 3, 2, 4, 13, 7, 1, 2, 1, 2, 4, 1, 4, 4, 4]
def print_solution(data, presentation):
    """Prints solution on console."""
    result = []
    count_vehi = 0
    temp_vehi = copy.deepcopy(data1['vehicle_capacities'])
    for tmpIndex in range(len(data['wd_names_now'])):
        tmpstr = '{0}: {1}'.format(tmpIndex, data['wd_names_now'][tmpIndex])
        result.append(tmpstr)
    total_distance = 0
    total_load = 0
    vehicle_id = 0
    loaed = []
    new_cap = []
    for i in range(len(presentation)): 
        loaed.append(0)
        for j in range(len(presentation[i])):
            loaed[i] +=demand[presentation[i][j]]
        for j in range(len(data['list_cap'])):
            if j != len(data['list_cap']) - 1:
                if loaed[i] <= data['list_cap'][j] and loaed[i] > data['list_cap'][j+1]:
                    new_cap.append(data['list_cap'][j])
                    break
            else:
                new_cap.append(data['list_cap'][-1])
                break

    for vehicle_id in range(len(presentation)):
        nowload = 0
        route_distance = 0
        tmpstr = 'Route for vehicle {0}:     Capacity: {1} KGS'.format(vehicle_id,new_cap[vehicle_id])
        result.append(tmpstr)
        tmpstr = ''
        tmpstr += '{0} Load({1}) ->'.format(0,0)
        j =0
        for j in range(len(presentation[vehicle_id])):
            nowload += demand[presentation[vehicle_id][j]]
            tmpstr += '{0} Load({1}) ->'.format(presentation[vehicle_id][j],nowload)
            if j == 0:
                route_distance += distance_mt[0][presentation[vehicle_id][j]]
            else:
                route_distance += distance_mt[presentation[vehicle_id][j-1]][presentation[vehicle_id][j]]
        route_distance += distance_mt[presentation[vehicle_id][j]][0]
        total_distance += route_distance
        total_load += nowload
        tmpstr += '{0} Load({1}) '.format(0,0)
        result.append(tmpstr)
        
        tmpstr = 'Distance of the route: {}km'.format(route_distance)
        result.append(tmpstr)
        tmpstr = 'Load of the route: {}'.format(nowload)
        result.append(tmpstr)
        tmpstr = "                                     "
        result.append(tmpstr)
    
    dt = modify_data(data1,False)
    # print(dt)
    # print('((')
    for i in range(len(dt)):
        if dt[i] != 0:
            for v in range(dt[i]):
                vehicle_id+=1
                tmpstr = 'Route for vehicle {0}:  Capacity: {1} KGS'.format(vehicle_id,data['list_cap'][0])
                result.append(tmpstr)
                tmpstr = ' {0} Load({1}) ->  {2} Load({3}) -> {4} Load({5}) '.format(0,0,i,data1['list_cap'][0],0,0)
                result.append(tmpstr)
                rt = int(data1['distance_matrix'][0][i] +   data1['distance_matrix'][i][0])
                tmpstr = 'Distance of the route: {}km'.format(rt)
                result.append(tmpstr)
                tmpstr = 'Load of the route: {}'.format(data1['list_cap'][0])
                result.append(tmpstr)
                total_distance += rt
                tmpstr = '                              '
                result.append(tmpstr)
                total_load += data1['list_cap'][0]
                
    tmp_str1 = 'Total distance of all routes: {}km'.format(total_distance)
    result.append(tmp_str1)
    tmp_str2 = 'Total load of all routes: {}'.format(total_load)
    result.append(tmp_str2)

    return result


def checkX(v,k):
    if v == 0:
        return True
    if visited[v]:
        return False
    if load[k] + demand[v] > cap:
        return False
    return True
#
def checkY(v,k):
    if visited[v]:
        return False
    if load[k] + demand[v] > cap:
        return False
    return True


#
def solution():
    global min_dis
    global rs
    
    # for k in range(1,num_vehi+1):
    #     print("Route {0}: 0-> {1}".format(k,y[k]
    tmp_dis = 0
    for k in range(1,num_vehi+1):
        tmp_dis += distance_route[k]
    if min_dis > tmp_dis:
        min_dis = tmp_dis
        rs = []
        for i in range(num_vehi+1):
            rs.append([])
        for i in range(1,num_vehi+1):
            rs[i].append(y[i])
            tmp = x[y[i]]
            while True:
                if tmp == 0:
                    break
                rs[i].append(tmp)
                tmp = x[tmp]


#

def TryX(v,k):
    global x
    global segment
    global distance_route
    global y
    global visited
    global load
    tmp_load = 0
    for i in range(1,k):
        tmp_load+= load[i]
    # print('load_now:' + str(tmp_load)+' in ' + str(k))
    if (tmp_load + cap*(num_vehi-k+1) < total_dm): 
        return 1
    for i in range(0,num_destination+1):
        if checkX(i,k):
            # print('x: ' +str(k)+ ':'+'dc chon:'+ str(i) + '*' )
            x[v] = i
            distance_route[k] += distance_mt[v][i]
            load[k] += demand[i]
            visited[i] = True
            segment += 1
            if i == 0:
                if k == num_vehi :
                    if segment == num_destination + num_vehi:
                        solution()
                else:
                    TryX(y[k+1],k+1)
            else:
                TryX(i,k)
            # print('x: ' +str(k)+ ':'+'dc chon:'+ str(i) )

            visited[i] = False
            segment -= 1
            distance_route[k] -= distance_mt[v][i]
            load[k] -= demand[i]
            


#
def TryY(k):
    global segment
    global distance_route
    global y
    global visited
    global load
    
    for i in range(y[k-1] + 1,num_destination+1):
        if checkY(i,k):
            # print('y :' + str(k)+ ':'+'dc chon:'+ str(i) )
            y[k] = i
            distance_route[k] += distance_mt[0][i]
            visited[i] = True
            segment += 1
            load[k] += demand[i]
            # print(str(i)+':'+str(visited[i]))
            if (k == num_vehi ):
                TryX(y[1],1)
            else:
                TryY(k+1)
            distance_route[k] -= distance_mt[0][i]
            visited[i] = False
            segment -= 1
            load[k] -= demand[i]

def main():
    """Solve the CVRP problem."""
    modify_data(data1,True)

    # Instantiate the data problem.
    TryY(1)
    rs2 = []
    for i in rs:
        if len(i) != 0:
            rs2.append(i)
    rs1 = print_solution(data1,rs2)
    return rs1

class Ui_MainWindow(object):
    
    def setupUi(self, MainWindow,widget_names):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(770, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(380, 50, 75, 20))
        self.pushButton.setObjectName("pushButton")
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(30, 50, 251, 20))
        self.comboBox.setObjectName("comboBox")
        for i in widget_names:
            self.comboBox.addItem("")
        #

        
        #
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(30, 30, 47, 13))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(290, 30, 47, 13))
        self.label_2.setObjectName("label_2")
        self.spinBox = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox.setGeometry(QtCore.QRect(290, 50, 80, 22))
        self.spinBox.setObjectName("spinBox")
        self.spinBox.setRange(0,100000)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(30, 90, 101, 16))
        self.label_3.setObjectName("label_3")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(90, 120, 181, 20))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_2.textChanged.connect(self.update_display)
        #search bar
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(30, 120, 61, 16))
        self.label_4.setObjectName("label_4")
        self.pushButton2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton2.setGeometry(290,120,80,20)
        # #scroll area
        # self.scroll = QtWidgets.QScrollArea(self.centralwidget)
        # self.scroll.setGeometry(QtCore.QRect(30,150,700,430))
        # self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        # self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        # self.scroll.setWidgetResizable(True)  

        #
        self.controls = QWidget(self.centralwidget) #Controls container Widget
        self.controls.setGeometry(QtCore.QRect(30,150,400,400))
        self.controlsLayOut  = QVBoxLayout() # Controls container Layout
        #List of names, widgets, are stored in a dictionary by these keys
        
        self.widgets = []

        #Iterate the names, creating a new OnOffWidget for
        # each one, adding it to the layout and
        # storing a reference in the self.widgets list
        for name in widget_names:
            item = OnOffWidget(name)
            self.controlsLayOut.addWidget(item)
            self.widgets.append(item)
        spacer = QSpacerItem(1, 1, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.controlsLayOut.addItem(spacer)
        self.controls.setLayout(self.controlsLayOut)

        #Scroll Area Properties
        self.scroll = QScrollArea(self.centralwidget)
        self.scroll.setGeometry(QtCore.QRect(30,150,700,430))
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setWidgetResizable(True)        
        self.scroll.setWidget(self.controls)
        
        
        #
        # #Search bar
        # self.searchbar = QLineEdit()
        # self.searchbar.textChanged.connect(self.update_display)

        # # Adding the completer
        # self.completer = QCompleter(widget_names)
        # self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        # self.searchbar.setCompleter(self.completer)
        

        #Add the items to  VBoxLayout (applied to container widget)
        # which encompasses the whole window
        
        #
        self.pushButton.clicked.connect(self.get_text_in_comboBox)
        self.pushButton2.clicked.connect(self.run)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 770, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow,widget_names)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow,widget_names):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "Add Demand"))
        for i in range(len(widget_names)):
            self.comboBox.setItemText(i, _translate("MainWindow", widget_names[i]))
        
        self.label.setText(_translate("MainWindow", "Tên "))
        self.label_2.setText(_translate("MainWindow", "Yêu cầu"))
        self.label_3.setText(_translate("MainWindow", "Danh sách yêu cầu"))
        self.label_4.setText(_translate("MainWindow", "Tìm kiếm:"))
        self.pushButton2.setText(_translate("MainWindow", "Run"))
    
    def run(self):
        self.w = AnotherWindow()
        self.w.show()
        data1['demands'] = copy.deepcopy(or_data['demands'])
        data1['distance_matrix'] = copy.deepcopy(or_data['distance_matrix'])
        self.update_element()
    def update_element(self):
        for i in range(len(self.widgets)):
            if self.widgets[i].is_on:
                data1['demands'][i] = int(self.widgets[i].lbl1.text())

    def update_display(self, text):
        for widget in self.widgets:
            if text.lower() in widget.name.lower():
                widget.show()
            else:
                widget.hide()
    def get_text_in_comboBox(self):
        text = str(self.comboBox.currentText())
        num = self.spinBox.value()
        count = 0
        for i in self.widgets:
            if text == i.name:
                i.is_on = True
                i.update_button_state()
                i.show()
                data1['demands'][count] = num
                i.update_num(num)
                break   
            count += 1


class OnOffWidget(QWidget):
    def __init__(self,name):
        super(OnOffWidget, self).__init__()
        self.name = name # Name of widget used for searching.
        self.is_on = False # Current state (true=ON, false=OFF)

        self.lbl = QLabel(self.name) #the widget label
        self.lbl1 = QLabel('0') #the widget label1
        self.spnbox = QSpinBox()
        self.spnbox.setRange(0, 100000)
        
        self.btn_on = QPushButton("Sửa") #The ON button
        self.btn_off = QPushButton("Xoá") #the off button
        self.btn_on.clicked.connect(self.update_num_by_spnb) #connect to function On
        self.btn_off.clicked.connect(self.off) #connect to function Off
        self.hbox = QHBoxLayout()  #A horizontal layout  to encapsulate the above
        self.hbox1 = QHBoxLayout() 
        self.hbox2 = QHBoxLayout() 
        self.hbox2.addWidget(self.spnbox) #Add the spin box to the layout
        self.hbox1.addWidget(self.lbl) #Add the label to the layout 
        self.hbox1.addWidget(self.lbl1) #Add the label to the layout 

        self.hbox2.addWidget(self.btn_on)  # Add the ON button to the layout
        self.hbox2.addWidget(self.btn_off)  # Add the OFF button to the layout
        self.hbox.addLayout(self.hbox1)
        self.hbox.addLayout(self.hbox2)
        
        self.setLayout(self.hbox)
        self.update_button_state()
    def update_num_by_spnb(self):
        num = self.spnbox.value()
        for i in range(len(data1['wd_names'])):
            if data1['wd_names'][i] == self.name:
                data1['demands'][i] = num
                break
        self.lbl1.setText(str(num))
    def update_num(self,num):
        self.lbl1.setText(str(num))
    
    def off(self):
        self.is_on = False
        for i in range(len(data1['wd_names'])):
            if data1['wd_names'][i] == self.name:
                data1['demands'][i] = 0
                break
        self.update_button_state()

    def on(self):
        self.is_on = True
        self.update_button_state()
    def update_button_state(self):
        #Update the  appearance of the control buttons(on/off)
        #depending the current state
        if self.is_on == True:
            self.btn_on.setStyleSheet('background-color : #4CA5F0; color : #fff;')
            self.btn_off.setStyleSheet('background-color : none; color: none;')
            self.show()
        else:
            self.btn_off.setStyleSheet('background-color : #D32F2F; color : #fff;')
            self.btn_on.setStyleSheet('background-color : none; color: none;')
            self.hide()
    def show(self):
        # show this widget, and all child widgets
        if self.is_on:
            for w in [self, self.lbl, self.btn_on, self.btn_off, self.spnbox]:
                w.setVisible(True)
    def hide(self):
        # hide this widget, and all childwidgets
        for w in [self, self.lbl, self.btn_on, self.btn_off, self.spnbox]:
            w.setVisible(False)
class AnotherWindow(QMainWindow):
    """
    This "window" is a QWidget. If it has no parent, it
    will appear as a free-floating window as we want.
    """
    def __init__(self, *args, **kwargs):
        super().__init__()

        self.controls = QWidget() #Controls container Widget
        self.controlsLayOut  = QVBoxLayout() # Controls container Layout
        #List of names, widgets, are stored in a dictionary by these keys
        rs = main()
        widget_names = rs
        self.widgets = []

        #Iterate the names, creating a new OnOffWidget for
        # each one, adding it to the layout and
        # storing a reference in the self.widgets list
        for name in widget_names:
            item = QLabel(name)
            self.controlsLayOut.addWidget(item)
            self.widgets.append(item)
        spacer = QSpacerItem(1, 1, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.controlsLayOut.addItem(spacer)
        self.controls.setLayout(self.controlsLayOut)

        #Scroll Area Properties
        self.scroll = QScrollArea()
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setWidgetResizable(True)        
        self.scroll.setWidget(self.controls)

        #Add the items to  VBoxLayout (applied to container widget)
        # which encompasses the whole window

        container = QWidget()
        containerLayOut = QVBoxLayout()
        containerLayOut.addWidget(self.scroll)

        container.setLayout(containerLayOut)
        self.setCentralWidget(container)
        
        self.setGeometry(660, 100, 500, 400)
        self.setWindowTitle('Control Panel')
    

if __name__ == '__main__':

    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow,data1['wd_names'])
    MainWindow.show()

    sys.exit(app.exec_())
