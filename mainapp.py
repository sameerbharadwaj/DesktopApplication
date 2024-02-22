import asyncio
import time
from bleak import BleakScanner, BleakClient
import tkinter as tk
import pandas as pd
import os

import openpyxl

global address
# IO_DATA_CHAR_UUID = "08b332a8-f4f6-4222-b645-60073ac6823f"      #change the uuid to BLE
IO_DATA_CHAR_UUID = "beb5483e-36e1-4688-b7f5-ea07361b26a8"


def printcommunication():
    global mylable2,btn,root,continue_btn,addr
    mylable2.config(text="connecting to Bluetooth...")
    btn.config(text="wait")
    root.update()
    try:
        # addr=''

        async def run():
            global addr
            devices = await BleakScanner.discover()

            print(devices)
            for d in devices:
                if ("MyESP32" in d.name):     #change to test device name
                    print(d.address)
                    addr = d.address
                    retstr="connected to MyESP32(" + addr+")"
                    mylable2.config(text=retstr)
                    btn.config(text="Start Test")
                    continue_btn.config(state=tk.NORMAL, bg="green")
                    root.update()
                    break
                else:
                    continue_btn.config(state=tk.DISABLED,bg='grey')
                    btn.config(text="Start Test")
                    mylable2.config(text="Turn ON Bluetooth of your PC/device")
                    root.update()


            print(addr)
        root.update()

        loop = asyncio.get_event_loop()
        loop.run_until_complete(run())


    # except UnboundLocalError:
    #     btn.config(text="Start Test")
    #     mylable2.config(text="Turn ON Bluetooth of your PC/device")
    except Exception:
        btn.config(text="Start Test")
        mylable2.config(text="Turn ON Bluetooth of your PC/device")
        # print("Turn Bluetooth of your PC/device")


def blereq(x):
    global addr
    print("addressssss: ",addr)
    try:
        async def run():
            global addr,result
            async with BleakClient(str(addr)) as client:
                var=[x]
                print(var)
                await client.write_gatt_char(IO_DATA_CHAR_UUID, var)
                result=await client.read_gatt_char(IO_DATA_CHAR_UUID)
                # result=list(result)
                print("inloop:",(result[0]))

        loop = asyncio.get_event_loop()
        loop.run_until_complete(run())
        return (result)
    except:
        return blereq(x)

def ozonereadings():
    t = time.localtime()
    current_time = time.strftime("%H:%M:%S", t)
    print(current_time)
    print(type(current_time))

    stri='Time(HH:MM:SS)   Temperature("\N{DEGREE SIGN}C")   OzoneValue (ppb)\n'
    global temp,root1,label6
    label6.config(text=stri)
    root1.update()
    dat = [['Time', 'Temperature\N{DEGREE SIGN}C','OzoneValues(ppb)']]
    d = pd.DataFrame(data=dat)
    d.to_excel('sheet1.xlsx', index=False)
    dat = list()
    for i in range(5):
        t = time.localtime()
        current_time = time.strftime("%H:%M:%S", t)
        oze = blereq(50)
        oze = int.from_bytes(oze, byteorder='little')
        print("i3:", oze)
        temperaturee=int.from_bytes(temp, byteorder='little')
        print("temperature:", int.from_bytes(temp, byteorder='little'))
        stri+=current_time+"                      "+str(temperaturee)+"                     "+str(oze)+"\n"
        x = [current_time,temperaturee,oze]
        dat.append(x)
        label6.config(text=stri)
        root1.update()
        d1 = pd.DataFrame(data=dat)
        d1 = d.append(d1)
        # print(d1)
        d1.to_excel('sheet1.xlsx', index=False)



def modereadings():
    global label5 ,root1

    oze = blereq(67)  # mode of operation
    stri="Mode Of operation: "+oze.decode()
    label5.config(text=stri)
    root1.update()
    print("mode:", oze.decode())
    ozonereadings()
def genexcel():
    os.startfile('sheet1.xlsx')

def page2():
    global root1,label3,addr,temp,label5,label6
    # print("page2,", addr)
    root1 = tk.Tk()
    # root1=tk.Frame(root1)
    root.destroy()
    root1.state("zoomed")
    label3=tk.Label(root1,text="Waiting for test to start...",font=("Calibri", 20))
    label3.place(x=0,y=0)
    # label3.grid(row=0,column=0)
    label4=tk.Label(root1,text="",font=("Calibri", 20))
    label4.place(x=0,y=35)
    # label4.grid(row=1,column=0)
    label5=tk.Label(root1,text="",font=("Calibri", 20))
    label5.place(x=0, y=70)
    # label5.grid(row=3,column=0)
    label6=tk.Label(root1,text="",font=("Calibri", 20))
    label6.place(x=80, y=110)
    # label6.grid(row=5,column=10)
    exit_btn=tk.Button(root1,text="Generate Excel Sheet",width=20,command=genexcel, state=tk.DISABLED, bg="grey")
    exit_btn.place(x=1200, y=700)

    root1.update()

    while 1:
        time.sleep(5)
        req=blereq(48)    #replace start test code in place of x

        # print(req[0],"waithing for test to start")
        if req[0]==1:
            break
    if req[0] == 1:
        # print("changing to text started")
        label3.config(text="Test started")
        root1.update()
    label4.config(text="stablizing temperature")
    root1.update()
    while 1:
        temp=blereq(49)
        # label4.config(text="stablizing temperature")
        # root1.update()
        if((int.from_bytes(temp,byteorder='little'))==25):
            label4.config(text="Temperature Stablised to 25 Degree")
            root1.update()
            modereadings()

            break
            # oze = blereq(67)            #mode of operation
            # print("mode:", oze)
            # oze=blereq(50)
            # oze=int.from_bytes(oze,byteorder='little')
            # print("i3:" ,oze)
    exit_btn.config(bg="green",state=tk.NORMAL)
    root1.update()
        # oze=blereq(52)
        # print("req43" ,oze)



def page1():
    global mylable1,mylable2,btn,continue_btn,addr
    root.state('zoomed')
    mylable1=tk.Label(root,text='Ready to start test',font=("Arial", 25))
    mylable1.pack(pady=100)
    btn=tk.Button(root,text= "Start Test",command=printcommunication,width=70, font=("Arial", 15), bg='green')
    btn.pack(pady=100,ipady=10)
    mylable2=tk.Label(root,text="",font=('Aerial',20))
    mylable2.pack()
    continue_btn = tk.Button(root, text="Continue",command=page2, width=20, state=tk.DISABLED, bg="grey")
    continue_btn.pack(side='bottom', padx=100, pady=100, ipady=10)


if __name__=='__main__':
    root=tk.Tk()
    page1()
    tk.mainloop()
