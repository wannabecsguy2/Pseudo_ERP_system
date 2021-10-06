import numpy as np
import pandas as pd
from tkinter import *
import os
import PySimpleGUI as psg

# Daily Entry Dataframe:
daily = pd.read_excel(r'D:\Coding\Random\Inventory Stock Query and Update\Final Files\Excel Files\daily.xlsx', engine = 'openpyxl')
daily['Date Of Entry'] = pd.to_datetime(daily['Date Of Entry'], format = '%d-%m-%Y')
# Columns = Date Of Entry, Order ID, Day/Night, Machine Number, Number Of Shots, Number Of Pieces, Employee 1 Name, Employee 1 Wage, Employee 2 Name, Employee 2 Wage, Raw Material, Material Used

# Employees Dataframe:
employees = pd.read_excel(r'D:\Coding\Random\Inventory Stock Query and Update\Final Files\Excel Files\employees.xlsx', index_col = 'Employee Name', engine = 'openpyxl')
# Columns = Employee Name, Position

# Machines Dataframe:
machines = pd.read_excel(r'D:\Coding\Random\Inventory Stock Query and Update\Final Files\Excel Files\machines.xlsx', index_col = 'Machine Number', engine = 'openpyxl')
# Columns = Machine Number, Machine Name, Rate With 1, Rate With 2

# Raw Materials Inventory Dataframe:
materials = pd.read_excel(r'D:\Coding\Random\Inventory Stock Query and Update\Final Files\Excel Files\materials.xlsx', index_col = 'Material', engine = 'openpyxl')
# Columns = Material, Net Weight, Melting Loss

# Order Status And Updates Dataframe:
orders = pd.read_excel(r'D:\Coding\Random\Inventory Stock Query and Update\Final Files\Excel Files\orders.xlsx', index_col = 'Order ID', engine = 'openpyxl')
orders['Date Of Entry'] = pd.to_datetime(orders['Date Of Entry'])
# Columns = Order ID, Date Of Entry, Product Name, Order Size, Casting, Fettling/Filling, Shot Blasting/Vibro, Drilling/Tapping, CNC, Heat Treatment, Ready For Dispatch, Dispatched, Status

# Product Details Dataframe:
products = pd.read_excel(r'D:\Coding\Random\Inventory Stock Query and Update\Final Files\Excel Files\products.xlsx', index_col = 'Product Name', engine = 'openpyxl')
# Columns = Product Name, Die Number, Company Name, Weight Of Casting, Pieces Per Shot, Raw Material

# Dispatch Details Dataframe:
dispatch = pd.read_excel(r'D:\Coding\Random\Inventory Stock Query and Update\Final Files\Excel Files\dispatch.xlsx', engine = 'openpyxl')
# Columns = Order ID, Date Of Entry, Product Name, Dispatch

def material_inbound(materials = materials):    # Material Inbound Update
    psg.theme('DarkBlue')
    layout = [[psg.Text('Raw Material:')],
              [psg.Listbox(list(materials.index), key = 'name', size = (10, 10))],
              [psg.Text('Weight:'), psg.Input()],
              [psg.Button('Save')]]
    
    win = psg.Window('Raw Material Inbound', layout)
    e, user_input = win.read()
    win.close()
    
    material, weight = user_input['name'], float(user_input[0])
    
    materials['Net Weight'][material] = materials['Net Weight'][material] + 0.91*weight
    materials['Melting Loss'][material] = materials['Melting Loss'][material] + 0.09*weight
    
    return materials

def new_product(products = products, materials = materials):    # New Product Registration
    psg.theme('DarkBlue')
    layout = [[psg.Text('Name Of the Product:'), psg.Input()],
              [psg.Text('Die Number:'), psg.Input()],
              [psg.Text('Company Name:'), psg.Input()],
              [psg.Text('Weight Of Casting:'), psg.Input()],
              [psg.Text('Pieces Per Shot:'), psg.Input()],
              [psg.Text('Raw Material:')],
              [psg.Listbox(list(materials.index), key = 'name', size = (10, 10))],
              [psg.Button('Save')]]
    
    win = psg.Window('New Product Registration', layout)
    e, user_input = win.read()
    win.close()

    products.loc[user_input[0]] = [user_input[1], user_input[2], float(user_input[3]), int(user_input[4]), user_input['name'][0]]

    return products

def new_order(products = products, orders = orders):    # New Order Registration   
    psg.theme('DarkBlue')
    layout = [[psg.Text('Order ID:'), psg.Input()],
              [psg.Text('Date Of Entry (DD-MM-YYYY)'), psg.Input()],
              [psg.Text('Product Name:')],
              [psg.Listbox(list(products.index), key = 'product', size = (15, 5))],
              [psg.Text('Order Size:'), psg.Input()],
              [psg.Button('Save')]]
    
    win = psg.Window('New Order Registration', layout)
    e, user_input = win.read()
    win.close()

    orders.loc[user_input[0]] = [pd.to_datetime(user_input[1], format = '%d-%m-%Y'), user_input['product'][0], int(user_input[2]), int(user_input[2]), 0, 0, 0, 0, 0, 0, 0, 'Incomplete']

    return orders

def production_update(employees = employees, materials = materials, daily = daily, machines = machines, products = products, orders = orders, dispatch = dispatch):    # Updating Production And Other Stages For A Order
    psg.theme('DarkBlue')
    layout = [[psg.Text('Date Of Entry (DD-MM-YYYY):'), psg.Input()],
              [psg.Text('Order ID:')],
              [psg.Listbox(list(orders[orders['Status'] == 'Incomplete'].index), key = 'id', size = (5, 5))],
              [psg.Text('Day Or Night (D/N):'), psg.Input()],
              [psg.Text('Shots Completed:'), psg.Input()],
              [psg.Text('Machine Number:')],
              [psg.Listbox(list(machines.index), key = 'num', size = (3, 3))],
              [psg.Text('Name Of Employee 1:')],
              [psg.Listbox(list(employees.index), key = 'e1', size = (10, 3))],
              [psg.Text('Name Of Employee 2:')],
              [psg.Listbox(list(employees.index), key = 'e2', size = (10, 3))],
              [psg.Text('Fettling/Filling Completed:'), psg.Input()],
              [psg.Text('Shot Blasting/Vibro Completed:'), psg.Input()],
              [psg.Text('Drilling/Tapping Completed:'), psg.Input()],
              [psg.Text('CNC Completed:'), psg.Input()],
              [psg.Text('Heat Treatment Completed:'), psg.Input()],
              [psg.Text('Dispatched:'), psg.Input()],
              [psg.Button('Save')]]
    
    win = psg.Window('Daily Status Update:', layout, size = (1000, 1000))
    e, user_input = win.read()
    win.close()
    print(user_input[1])
    # Updating Daily Entry Dataframe:
    daily.loc[len(daily.index)] = [pd.to_datetime(user_input[0], format = '%d-%m-%Y'), user_input['id'][0], user_input[1], user_input['num'][0], int(user_input[2]), int(user_input[2])*products['Pieces Per Shot'][orders['Product Name'][user_input['id'][0]]], user_input['e1'][0], None, user_input['e2'][0], None, products['Raw Material'][orders['Product Name'][user_input['id'][0]]], int(user_input[2])*(products['Pieces Per Shot'][orders['Product Name'][user_input['id'][0]]])*(products['Weight Of Casting'][orders['Product Name'][user_input['id'][0]]])]
    if user_input['e2'][0] == 'None':
        daily['Employee 1 Wage'][len(daily.index) - 1] = int(user_input[2])*machines['Rate With 1'][user_input['num'][0]]
        daily['Employee 2 Wage'][len(daily.index) - 1] = 0
    else:
        total_wage = int(user_input[2])*machines['Rate With 2'][user_input['num'][0]]
        
        if employees['Position'][user_input['e1'][0]] != employees['Position'][user_input['e2'][0]]:
            if employees['Position'][user_input['e1'][0]] == 'Operator':
                daily['Employee 1 Wage'][len(daily.index) - 1] = total_wage*0.6
                daily['Employee 2 Wage'][len(daily.index) - 1] = total_wage*0.4
            else:
                daily['Employee 2 Wage'][len(daily.index) - 1] = total_wage*0.6
                daily['Employee 1 Wage'][len(daily.index) - 1] = total_wage*0.4
        else:
            daily['Employee 1 Wage'][len(daily.index) - 1] = total_wage*0.5
            daily['Employee 2 Wage'][len(daily.index) - 1] = total_wage*0.5
    
    # Updating Order Status and Updates Dataframe:
    orders['Casting'][user_input['id'][0]] = orders['Casting'][user_input['id'][0]] - int(user_input[2])*products['Pieces Per Shot'][orders['Product Name'][user_input['id'][0]]]
    orders['Fettling/Filling'][user_input['id'][0]] = orders['Fettling/Filling'][user_input['id'][0]] + int(user_input[2])*products['Pieces Per Shot'][orders['Product Name'][user_input['id'][0]]] - int(user_input[3])
    orders['Shot Blasting/Vibro'][user_input['id'][0]] = orders['Shot Blasting/Vibro'][user_input['id'][0]] + int(user_input[3]) - int(user_input[4])
    orders['Drilling/Tapping'][user_input['id'][0]] = orders['Drilling/Tapping'][user_input['id'][0]] + int(user_input[4]) - int(user_input[5])
    orders['CNC'][user_input['id'][0]] = orders['CNC'][user_input['id'][0]] + int(user_input[5]) - int(user_input[6])
    orders['Heat Treatment'][user_input['id'][0]] = orders['Heat Treatment'][user_input['id'][0]] + int(user_input[6]) - int(user_input[7])
    orders['Ready For Dispatch'][user_input['id'][0]] = orders['Ready For Dispatch'][user_input['id'][0]] + int(user_input[7]) - int(user_input[8])
    orders['Dispatched'][user_input['id'][0]] = orders['Dispatched'][user_input['id'][0]] + int(user_input[8])
    
    if orders['Dispatched'][user_input['id'][0]] == orders['Order Size'][user_input['id'][0]]:
        orders['Status'][user_input['id'][0]] == 'Complete'

    # Updating Dispatch Details Dataframe:
    if int(user_input[8]) != 0:
        dispatch[len(dispatch.index)] = [pd.to_datetime(user_input[0], format = '%d-%m-%Y'), user_input['id'][0], orders['Product Name'][user_input['id'][0]], int(user_input[8])]

    # Updating Raw Materials Inventory Dataframe:
    materials['Net Weight'][products['Raw Material'][orders['Product Name'][user_input['id'][0]]]] -= int(user_input[2])*(products['Pieces Per Shot'][orders['Product Name'][user_input['id'][0]]])*(products['Weight Of Casting'][orders['Product Name'][user_input['id'][0]]])
    
    return [orders, daily, materials, dispatch]

def update_order(orders = orders, dispatch = dispatch):    # Updating Stages Other Than Casting:
    psg.theme('DarkBlue')
    layout = [[psg.Text('Order To Edit:')],
              [psg.Listbox(list(orders.index), key = 'id')],
              [psg.Text('Fettling/Filling Completed:'), psg.Input()],
              [psg.Text('Shot Blasting/Vibro Completed:'), psg.Input()],
              [psg.Text('Drilling/Tapping Completed:'), psg.Input()],
              [psg.Text('CNC Completed:'), psg.Input()],
              [psg.Text('Heat Treatment Completed:'), psg.Input()],
              [psg.Text('Dispatched:'), psg.Input()],
              [psg.Text('Date Of Entry:'), psg.Input()],
              [psg.Button('Save')]]
    
    win = psg.Window('Individual Changes In An Order', layout)
    e, user_input = win.read()
    win.close()

    orders['Fettling/Filling'][user_input['id'][0]] = orders['Fettling/Filling'][user_input['id'][0]] - int(user_input[0])
    orders['Shot Blasting/Vibro'][user_input['id'][0]] = orders['Shot Blasting/Vibro'][user_input['id'][0]] - int(user_input[1]) + int(user_input[0])
    orders['Drilling/Tapping'][user_input['id'][0]] = orders['Drilling/Tapping'][user_input['id'][0]] - int(user_input[2]) + int(user_input[1])
    orders['CNC'][user_input['id'][0]] = orders['CNC'][user_input['id'][0]] - int(user_input[3]) + int(user_input[2])
    orders['Heat Treatment'][user_input['id'][0]] = orders['Heat Treatment'][user_input['id'][0]] - int(user_input[4]) + int(user_input[3])
    orders['Ready For Dispatch'][user_input['id'][0]] = orders['Ready For Dispatch'][user_input['id'][0]] + int(user_input[4]) - int(user_input[5])
    orders['Dispatched'][user_input['id'][0]] = orders['Dispatched'][user_input['id'][0]] + int(user_input[5])

    if orders['Dispatched'][user_input['id'][0]] == orders['Order Size'][user_input['id'][0]]:
        orders['Status'][user_input['id'][0]] == 'Complete'   
    if int(user_input[5]) != 0:
        dispatch[len(dispatch.index)] = [pd.to_datetime(user_input[6], format = '%d-%m-%Y'), user_input['id'][0], orders['Product Name'][user_input['id'][0]], int(user_input[5])]
    return [orders, dispatch]

pd.set_option('display.max_columns', None)
pd.options.display.width=None

# Designing Infinite query loop system:
iterate = 1
while iterate > 0:
    actions = ['Raw Material', 'Product', 'Orders', 'Daily', 'Save And Quit']
    psg.theme('DarkBlue')
    layout1 = [[psg.Text('Choose Parameter To Perform Action On:')],
              [psg.Listbox(actions, key = 'action', size = (20, 5))],
              [psg.Button('Save')]]
    win = psg.Window('Action To Perform', layout1, size = (500, 500))
    e, user_input1 = win.read()
    win.close()

    if user_input1['action'][0] == 'Raw Material':
        while iterate > 0:
            layout2 = [[psg.Text('Action To Perform:')],
                       [psg.Listbox(['Inbound', 'Status', 'Quit'], key = 'action', size = (20, 3))],
                       [psg.Button('Save')]]
            win = psg.Window('Further Actions', layout2, size = (500, 500))
            e, user_input2 = win.read()
            win.close()
            if user_input2['action'][0] == 'Inbound':
                materials = material_inbound()
            elif user_input2['action'][0] == 'Status':
                psg.popup_scrolled(materials, size = (40, 10))
            else:
                break
    
    elif user_input1['action'][0] == 'Product':
        while iterate > 0:
            layout2 = [[psg.Listbox(['New Product', 'All Products', 'Quit'], key = 'action', size = (20, 3))],
                       [psg.Button('Save')]]
            win = psg.Window('Futher Actions', layout2, size = (250, 250))
            e, user_input2 = win.read()
            win.close()
            if user_input2['action'][0] == 'New Product':
                products = new_product()
            elif user_input2['action'][0] == 'All Products':
                psg.popup_scrolled(products, size = (150, 50))
            else:
                break
    
    elif user_input1['action'][0] == 'Orders':
        while iterate > 0:
            layout2 = [[psg.Listbox(['New', 'All', 'Complete', 'Incomplete', 'Quit'], key = 'action', size = (20, 3))],
                       [psg.Button('Save')]]
            win = psg.Window('Further Actions', layout2, size = (250, 250))
            e, user_input2 = win.read()
            win.close()
            if user_input2['action'][0] == 'New':
                orders = new_order()
            elif user_input2['action'][0] == 'All':
                psg.popup_scrolled(orders, size = (200, 50))
            elif user_input2['action'][0] == 'Incomplete':
                psg.popup_scrolled(orders[orders['Status'] == 'Incomplete'], size = (200, 50))
            elif user_input2['action'][0] == 'Complete':
                psg.popup_scrolled(orders[orders['Status'] == 'Complete'], size = (200, 50))
            else:
                break
    
    elif user_input1['action'][0] == 'Daily':
        layout2 = [[psg.Listbox(['Update', 'Sort', 'Quit'], key = 'action', size = (20, 3))],
                   [psg.Button('Save')]]
        win = psg.Window('Further Actions', layout2, size = (250, 250))
        e, user_input2 = win.read()
        win.close()
        
        if user_input2['action'][0] == 'Update':
            while iterate > 0:
                layout3 = [[psg.Listbox(['All', 'All Except Casting', 'Quit'], key = 'action', size = (30, 3))],
                           [psg.Button('Save')]]
                win = psg.Window('Further Actions', layout3, size = (200, 100))
                e, user_input3 = win.read()
                win.close()
                if user_input3['action'][0] == 'All':
                    orders, daily, materials, dispatch = production_update()
                elif user_input3['action'][0] == 'All Except Casitng': 
                    orders, dispatch = update_order()
                else:
                    break
        elif user_input2['action'][0] == 'Sort':
            while iterate > 0:
                layout3 = [[psg.Listbox(['Timeline', 'Employee', 'Machine', 'Order ID', 'Quit'], key = 'action', size = (20, 5))],
                           [psg.Button('Save')]]
                win = psg.Window('Further Actions', layout3, size = (200, 200))
                e, user_input3 = win.read()
                win.close()
                if user_input3['action'][0] == 'Timeline':
                    layout4 = [[psg.Text('Enter Start Date (DD-MM-YYYY):'), psg.Input()],
                               [psg.Text('Enter End Date (DD-MM-YYYY):'), psg.Input()],
                               [psg.Button('Save')]]
                    win = psg.Window('Sort By Duration', layout4, size = (200, 50))
                    e, user_input4 = win.read()
                    win.close()
                    daily_req = daily[daily.index <= pd.to_datetime(user_input4[1], format = '%d-%m-%Y')]
                    daily_req = daily_req[daily_req.index >= pd.to_datetime(user_input4[0], format = '%d-%m-%Y')]
                    psg.popup_scrolled(daily_req,size = (150, 30))
                elif user_input3['action'][0] == 'Employee':
                    layout4 = [[psg.Listbox(list(employees.index), key = 'e1', size = (20, 5))],
                               [psg.Button('Save')]]
                    win = psg.Window('Sort By Employee', layout4, size = (200, 200))
                    e, user_input4 = win.read()
                    win.close()
                    mask = (daily['Employee 1 Name'] == user_input4['e1'][0]) | (daily['Employee 2 Name'] == user_input4['e1'][0])
                    daily_req = daily[mask]
                    psg.popup_scrolled(daily_req, size = (150, 30))
                elif user_input3['action'][0] == 'Machine':
                    layout4 = [[psg.Listbox(list(machines.index), key = 'num', size = (20, 5))],
                               [psg.Button('Save')]]
                    win = psg.Window('Sort By Machine Number', layout4, size = (200, 200))
                    e, user_input4 = win.read()
                    win.close()
                    daily_req = daily[daily['Machine Number'] == user_input4['num'][0]]
                    psg.popup_scrolled(daily_req, size = (150, 30))
                elif user_input3['action'][0] == 'Order ID':
                    layout4 = [[psg.Listbox(list(orders.index), key = 'id', size = (20, 5))],
                               [psg.Button('Save')]]
                    win = psg.Window('Sort By Order ID', layout4, size = (200, 200))
                    e, user_input4 = win.read()
                    win.close()
                    daily_req = daily[daily['Order ID'] == user_input4['id'][0]]
                    psg.popup_scrolled(daily_req, size = (150, 30))
                else:
                    break
        else:
            break    
    else:
        daily.to_excel('daily.xlsx', index = False)
        break