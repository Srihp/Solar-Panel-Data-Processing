import xlsxwriter
import pandas as pd
import numpy as np
import scipy as sc
from sklearn import *
from tkinter import *
from tkinter import messagebox
import tkinter.filedialog
import mysql.connector
from mysql.connector import Error

def CalculateOutput(manufacturerValues,Measured_STC_Rated,path):
    df = pd.read_excel (path, sheet_name='Sheet1')     
    New_data = df.values.tolist()
    Curvename = New_data[0]
    Date = New_data[1]
    Time = New_data[2]
    CurveTracer = New_data[3]
    Tamb = New_data[4]
    Projectno = New_data[5]
    Trefcell1 = New_data[6]
    Tm = New_data[7]
    Irradiance1 = New_data[8]
    Irradiance2 = New_data[9]
    Trefcell2 = New_data[10]
    Cfref1 = New_data[11]
    cfref2 = New_data[12]
    Isc = New_data[13]
    Voc = New_data[14]
    Imp = New_data[15]
    Vmp = New_data[16]
    Pm = New_data[17]
    FF = New_data[18]
    PR = New_data[19]
    SR = New_data[20]

    Temp_coefficent = []
    Corrected_i1 = []      
    for i in range (1,18):
        if(FF[i]== 0):
            Corrected_i1.append(0)
        else:
            c = round(Irradiance1[i] * PR[2]/( PR[2] + PR[3] * (Trefcell1[i] -25)),2)
        Corrected_i1.append(c)

    Corrected_i2 = []      
    for i in range (1,18):
        if(FF[i]== 0):
            Corrected_i2.append(0)
        else:
            c = round(Irradiance2[i] * SR[2]/( SR[2] + SR[3] * (Trefcell2[i] -25)),2)
        Corrected_i2.append(c)
    
    Corrected_i = []
    for i in range (0,17):
        if (PR[1] == "A825P"):
            c = round(Corrected_i1[i] *(200.0 /1000.0),2)
        else:
            c = round(Corrected_i2[i] * (200.0/ 1000.0),2)
        Corrected_i.append(c)
    
    x = []
    for i in range(1,18):
        if(FF[i]== 0):
            x.append(0)
        else:
            c = round(((Tm[i] + 1.5) - 25),2)
        x.append(c)
        
    y = []
    for i in range (0,17):
        c=round(Isc[i+1]/Corrected_i[i]*200,2)
        y.append(c)
    
    Linear_Regression_m = [] 
    Linear_Regression_b = [] 
        
    X = np.array(x)
    Y = np.array(y)
 
    #isc linear reg
    a = []   
    a = sc.stats.linregress(X,Y)

    Linear_Regression_m.append(a[0])
    Linear_Regression_b.append(a[1])


    #voc Linear reg
    Y = np.array(Voc[1:])
    a = sc.stats.linregress(X,Y)

    Linear_Regression_m.append(a[0])
    Linear_Regression_b.append(a[1])


    #imp linear reg
    Imp_200 = []
    for i in range(0,17):
        a = (Imp[i+1]/Corrected_i[i])*200.00
        Imp_200.append(round(a,2))
        
    Y = np.array(Imp_200)
    a = sc.stats.linregress(X,Y)
    Linear_Regression_m.append(a[0])
    Linear_Regression_b.append(a[1]) 

    #vmp linear reg
    Y= np.array(Vmp[1:])
    a = sc.stats.linregress(X,Y)
    Linear_Regression_m.append(a[0])
    Linear_Regression_b.append(a[1]) 

    #ff linear reg

    Y= np.array(FF[1:])
    a = sc.stats.linregress(X,Y)
    Linear_Regression_m.append(a[0])
    Linear_Regression_b.append(a[1])

    #pmax linear reg
    Pmax_200 = []
    for i in range(0,17):
        a = (Pm[i+1]/Corrected_i[i])*200.00
        Pmax_200.append(round(a,2))
    
    Y= np.array(Pmax_200)
    a = sc.stats.linregress(X,Y)
    Linear_Regression_b.append(a[1]) 
    Linear_Regression_m.append(a[0])

    i=0
    while i < 6:
        a = (Linear_Regression_m[i] / Linear_Regression_b[i]) * 100
        Temp_coefficent.append(a)
        i=i+1
   
    Delta_measured = []
    j = 0
    while j < 6:
        a = round (((Linear_Regression_b[j] - Measured_STC_Rated[j]) / Measured_STC_Rated[j]), 2)
        Delta_measured.append(a)
        j = j +1
    
    try:
        connection = mysql.connector.connect(host='localhost',
                                         database='iv',
                                         user='root',
                                         password='root')
        if (connection.is_connected()):
            db_Info = connection.get_server_info()
            print("Connected to MySQL Server version ", db_Info)
            cursor = connection.cursor()
            cursor.execute("select database();")
            record = cursor.fetchone()
            print("You're connected to database: ", record)
            mycursor = connection.cursor()
            mycursor.execute("SELECT * FROM TestData")
            myresult = mycursor.fetchall()
            for x in myresult:
                print(x)

    except Error as e:
        print("Error while connecting to MySQL", e)
    '''finally:
        if (connection.is_connected()):
            cursor.close()
            connection.close()
            print("MySQL connection is closed")'''
    
    mycursor = connection.cursor()
    
    #for i in range(0,7):
    man = manufacturerValues[0]
    man1 = manufacturerValues[1]
    man3 = manufacturerValues[3]
        
    mea = Measured_STC_Rated[0]
    mea1 = Measured_STC_Rated[1]
    mea2= Measured_STC_Rated[2]
    mea3 = Measured_STC_Rated[3]
    mea4 = Measured_STC_Rated[4]
        
    
    
    sql = "INSERT INTO ManufactureData (manufacturer, modeltype, modulearea, Isc, Voc, Imp, Vmp, Pm) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"
    val = (man,man1,man3,mea,mea1,mea2,mea3,mea4 )
    mycursor.execute(sql, val)
    connection.commit()
    
    sql = "INSERT INTO measuredatstc (Isc, Voc, Imp, Vmp, FF, Pm) VALUES (%s, %s, %s, %s, %s, %s)"
    val = (Linear_Regression_b[0],Linear_Regression_b[1],Linear_Regression_b[2],Linear_Regression_b[3],Linear_Regression_b[4],Linear_Regression_b[5] )
    mycursor.execute(sql, val)
    connection.commit()
    
    sql = "INSERT INTO temperaturecoefficient (Isc, Voc, Imp, Vmp, FF, Pm) VALUES (%s, %s, %s, %s, %s, %s)"
    val = (Temp_coefficent[0],Temp_coefficent[1],Temp_coefficent[2],Temp_coefficent[3],Temp_coefficent[4],Temp_coefficent[5] )
    mycursor.execute(sql, val)
    connection.commit()
    
    sql = "INSERT INTO  deltaMeasured (Isc, Voc, Imp, Vmp, FF, Pm) VALUES (%s, %s, %s, %s, %s, %s)"
    val = (Delta_measured[0],Delta_measured[1],Delta_measured[2],Delta_measured[3],Delta_measured[4],Delta_measured[5] )
    mycursor.execute(sql, val)
    connection.commit()
    

    print("")
    print("Linear_regression")
    print("")
    print(Linear_Regression_m)
    print("")
    print("Measured @STC")
    print(Linear_Regression_b)

    print("")
    print("Delta_measured")    
    print("")       
    print(Delta_measured) 
    
    print("")
    print("Temp_coefficent") 
    print("")   
    print(Temp_coefficent)    
    
    parameters = ['Isc','Voc','Imp','Vmp','FF','Pm']
    df = pd.DataFrame({'':parameters,'Measured @STC':Linear_Regression_b,'Delta_measured':Delta_measured,'Temp_coefficent':Temp_coefficent})

    #with xlsxwriter.Workbook('test.xlsx') as workbook:
     #   worksheet = workbook.add_worksheet()
     
    writer = pd.ExcelWriter('IV_Output.xlsx', engine='xlsxwriter')
    df.to_excel(writer,sheet_name='IV',index=False)
    writer.save()
    #newwin = Toplevel(window)
    #display = Label(newwin, text=STCvalues)
    #display.pack()    
    #display2 = Label(newwin, text="Output2")
    #display2.pack()    
    
def submitCallBack():
    inputValues = []
    manufacturerValues = []
    Manufacturer_val = txt_Manufacturer.get()
    Model_val = txt_Model.get()
    Project_val = txt_Project.get()
    Module_val = txt_Module.get()
    Isc_val = float(txt_Isc.get())
    Voc_val = float(txt_Voc.get())
    Imp_val = float(txt_Imp.get())
    Vmp_val = float(txt_Vmp.get())
    Pm_val = float(txt_Pm.get())
    FF_val = float(txt_FF.get())
    Path_val = txt_Path.get()
    inputValues.extend([Isc_val,Voc_val,Imp_val,Vmp_val,Pm_val,FF_val])
    manufacturerValues.extend([Manufacturer_val,Model_val,Project_val,Module_val])
    CalculateOutput(manufacturerValues,inputValues,Path_val)
    messagebox.showinfo( "IV Test", "Output file generated")
    
def print_path():
    path = tkinter.filedialog.askopenfilename(
        parent=window, initialdir='C:/ProgramData',
        title='Choose file',
        filetypes=[('Excel', '.xlsx')]
        )
    txt_Path.insert(0,path)
    
def helloCallBack1():
    messagebox.showinfo( "Hello Python", "file not saved")
    
    
window = Tk()
LBL_ClientInfo = Label(window, text="  Client Information and rated values")
LBL_ClientInfo.grid(column=3, row=0 )

window.title("IV_TEST")
lbl_Manufacturer = Label(window, text="Manufacturer")
lbl_Manufacturer.grid(column=1, row=1)
txt_Manufacturer = Entry(window,width=10)
txt_Manufacturer.grid(column=2, row=1)

lbl_Model = Label(window, text="Model/Type")
lbl_Model.grid(column=1, row=3)
txt_Model = Entry(window,width=10)
txt_Model.grid(column=2, row=3)

lbl_Project = Label(window, text="Project no.")
lbl_Project.grid(column=1, row=5)
txt_Project = Entry(window,width=10)
txt_Project.grid(column=2, row=5)
 
lbl_Module = Label(window, text="Module Area")
lbl_Module.grid(column=1, row=7)
txt_Module = Entry(window,width=10)
txt_Module.grid(column=2, row=7)

 
lbl_Isc = Label(window, text="Isc")
lbl_Isc.grid(column=3, row=1)
txt_Isc = Entry(window,width=10)
txt_Isc.grid(column=4, row=1)

lbl_Voc = Label(window, text="Voc")
lbl_Voc.grid(column=3, row=3)
txt_Voc = Entry(window,width=10)
txt_Voc.grid(column=4, row=3)

lbl_Imp = Label(window, text="Imp")
lbl_Imp.grid(column=3, row=5)
txt_Imp = Entry(window,width=10)
txt_Imp.grid(column=4, row=5)

lbl_Vmp = Label(window, text="vmp")
lbl_Vmp.grid(column=3, row=7)
txt_Vmp = Entry(window,width=10)
txt_Vmp.grid(column=4, row=7)

lbl_Pm = Label(window, text="pm")
lbl_Pm.grid(column=3, row=9)
txt_Pm = Entry(window,width=10)
txt_Pm.grid(column=4, row=9)

lbl_FF = Label(window, text="FF")
lbl_FF.grid(column=3, row=10)
txt_FF = Entry(window,width=10)
txt_FF.grid(column=4, row=10)

txt_Path = Entry(window,width=20)
txt_Path.grid(column=2, row=12)
Upload = tkinter.Button(window, text='Upload a File', command=print_path)
Upload.grid(column=3, row = 12)

btn_Submit = Button(window, text="Submit", bg = "grey", command = submitCallBack)
btn_Submit.grid(column=2, row = 14)

b = Button(window, text="Cancel", bg = "grey", command = helloCallBack1)
b.grid(column=3, row = 14)
window.mainloop()
