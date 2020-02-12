# Simple absent program

# Import library
import pandas as pd
import PySimpleGUI as sg

# Read data
layout = [ [sg.Text('Enter file name'), sg.Input()],      
           [sg.OK()] ]        

window = sg.Window('Simple absent application', layout)
event, values = window.read()
value = values[0]
df = pd.read_excel(value)
window.close()

# Set default value
df["HADIR"] = 0
df["SAKIT"] = 0
df["IZIN"] = 0
df["TANPA KETERANGAN"] = 0

# Data updating function
def updateData(df) :
    layout = [      
            [sg.Text('Silahkan masukkan data nim(3 angka terakhir), keterangan(HADIR, SAKIT, IZIN, TANPA KETERANGAN)')],           
            [sg.Text('NIM', size=(15, 1)), sg.InputText('')],  
            [sg.Text('Keterangan', size=(15, 1)), sg.InputText('')],          
            [sg.Submit(), sg.Cancel()]      
            ]      
    window = sg.Window('Simple absent application').Layout(layout)         
    button, values = window.Read()
    if (button == "Submit") :
        nim = int(values[0])
        data = df.loc[df["NIM"] == nim]
        data[str(values[1])] += 1
        df.loc[df["NIM"] == nim]= data
    window.close()
    return(df)

def addData(df) :
    layout = [      
            [sg.Text('Silahkan masukkan data nama, nim, angkatan')],      
            [sg.Text('Nama', size=(15, 1)), sg.InputText('')],      
            [sg.Text('NIM(3 angka terakhir)', size=(15, 1)), sg.InputText('')],      
            [sg.Text('Angkatan', size=(15, 1)), sg.InputText('')],      
            [sg.Submit(), sg.Cancel()]      
            ]      
    window = sg.Window('Simple absent application').Layout(layout)         
    button, values = window.Read()
    if(button == "Submit") : 
        data = pd.DataFrame({"ANGKATAN" : [int(values[2])],
                            "NIM" : [int(values[1])],
                            "NAMA" : [str(values[0])],
                            "HADIR" : [1],
                            "SAKIT" : [0],
                            "IZIN" : [0],
                            "TANPA KETERANGAN" : [0]})
        df = df.append(data,ignore_index = True)
    window.close()
    return(df)

# Main Program
button = " "
while(button != "Quit") :
    layout = [      
        [sg.Text('Selamat datang di aplikasi absensi')],       
        [sg.Button("Daftar pertama"), sg.Button("Isi kehadiran"), sg.Quit()]      
        ]      
    window = sg.Window('Simple absent application').Layout(layout)         
    button, values = window.Read()
    if(button == "Daftar pertama") :
        df = addData(df)
    elif(button == "Isi kehadiran") :
        df = updateData(df)

layout = [ [sg.Text('Enter output file name'), sg.Input()],      
           [sg.OK()] ]        

window = sg.Window('Simple absent application', layout)
event, values = window.read()
value = values[0]
writer = pd. ExcelWriter("out.xlsx")
df.to_excel(writer , "Sheet1")
writer.save()
