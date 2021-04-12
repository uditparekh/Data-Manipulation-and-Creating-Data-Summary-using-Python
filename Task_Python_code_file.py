import pandas as pd
import pyodbc
from spinner import Spinner
import xlsxwriter
import tabulate
import numpy as np
import time
from datetime import datetime



s=Spinner()
s.start()

''' Solution for Question Number 1 '''

Task_B1 = pd.read_excel('TechnicalScreen/Copy of TechnicalScreen.xlsx', sheet_name = 'Task B')
Task_B2 = pd.read_excel('TechnicalScreen/Copy of TechnicalScreen.xlsx', sheet_name = 'Task B Country List')

Product_B = Task_B1.merge(Task_B2, on= 'Account.Number', how= 'left')

Product_B['Year_Date']= Product_B['Date'].dt.strftime('%Y')



''' Solution for Question Number 2 '''

table = pd.pivot_table(Product_B,
                       index=['Product'],
                       columns=['Year_Date'],
                       values=["Quantity"],
                       aggfunc={"Quantity": np.sum},
                       dropna="False",
                       margins='True',
                       margins_name= "x Grand Total",
                       fill_value= 0)


file = pd.crosstab([Product_B.Product],Product_B.Year_Date,values=Product_B.Quantity,aggfunc=sum).apply(lambda x: x/x.sum()).applymap(lambda x: "{:.0f}%".format(100*x))




file['percent_difference'] = 0
for val in range(0, len(file.index)):
    if val > 9:
        file.loc['Prod-' + str(val),'percent_difference'] = str(abs(((int(file.iloc[val]['2021'][:-1]) - int(file.iloc[val]['2020'][:-1])) / int(file.iloc[val]['2020'][:-1]))* 100)) + "%"
    else :
        file.loc['Prod-0' + str(val),'percent_difference'] = str(abs(((int(file.iloc[val]['2021'][:-1]) - int(file.iloc[val]['2020'][:-1])) / int(file.iloc[val]['2020'][:-1]))* 100)) + "%"

''' Solution for Question Number 3 '''

change_new = Product_B[(Product_B["Year_Date"] == '2020')]

table1 = pd.pivot_table(change_new,
                       index=['Country'],
                       columns=['Year_Date'],
                       values=["Quantity"],
                       aggfunc={"Quantity": np.sum},
                       dropna="False",
                       #margins='True',
                       #margins_name= "x Grand Total",
                       fill_value= 0)

df = table1.reindex(table1['Quantity'].sort_values(by='2020', ascending=False).index).head(3)

''' Solution for Question Number 4 '''

Product_B['Defect_Error'] = Product_B['Defect'].str[:8]


table2 = pd.pivot_table(Product_B,
                       index=['Defect_Error', 'Defect' ],
                       #columns=['Year_Date'],
                       values=["Quantity"],
                       aggfunc={"Quantity": np.sum},
                       dropna="False",
                       margins='True',
                       margins_name= "x Grand Total",
                       fill_value= 0)

table2= pd.concat([d.append(d.sum().rename((k, '-Subtotal'))) for k, d in table2.groupby('Defect_Error' )])

table2=table2[:-1]

writer = pd.ExcelWriter('TaskB.xlsx')
Product_B.to_excel(writer, sheet_name='TaskB_Q1', index=False)
file.to_excel(writer, sheet_name='TaskB_Q2', index=True)
df.to_excel(writer, sheet_name='TaskB_Q3', index=True)
table2.to_excel(writer, sheet_name='TaskB_Q4', index=True)
writer.save()
s.stop()