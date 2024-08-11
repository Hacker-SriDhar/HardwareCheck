import SystemRebuilt 
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from ttkbootstrap.constants import *
import ttkbootstrap as tb
from customtkinter import *
import time
import threading
import pythoncom
from generate_report import generate
from functools import partial
from tkinter import filedialog,messagebox
from PIL import Image,ImageTk


directory_path=None
software_gather=[]
background=True
progressbar_load3=None

def progress_bar():
    progressbar_load2=tb.Progressbar(master=main_frame,bootstyle="danger striped",maximum=100,mode="determinate",length="200",value=0)
    progressbar_load2.pack(pady=100)
    progressbar_load2.start(10)
    app.after(1000,lambda:progressbar_load2.destroy())

def progress_bar2():
    global progressbar_load3
    progressbar_load3=tb.Progressbar(master=main_frame,bootstyle="danger striped",maximum=100,mode="determinate",length="200",value=0)
    progressbar_load3.pack(pady=100)
    progressbar_load3.start(10)


def cpu_info():
    del_page()
    get_cpu=SystemRebuilt.get_cpu_info()
    progress_bar()
    cpu_info=["Name","Manufacturer","NumberOfCores","MaxClockSpeed","NumberOfLogicalProcessors","ProcessorId","Architecture","Caption","DeviceID","Family","Revision","Stepping","L2CacheSize","L3CacheSize","LoadPercentage"]
    cpu_frame=tk.Frame(main_frame)
    lb=tk.Label(cpu_frame,text="CPU Information",fg="black",bg="white",font=('Impact',15))
    cpu_frame.pack(pady=20)
    
    def run_after():
        time.sleep(1)
        lb.pack()
        
        cpu_table=ttk.Treeview(cpu_frame,columns=("Name","Manufacturer","Cores","LogicalProcessors","ProcessorId","Architecture","Caption","Stepping"),show='headings')
        cpu_table.heading("Name",text="Name")
        cpu_table.column("Name",width=300,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table.heading("Manufacturer",text="Manufacturer")
        cpu_table.column("Manufacturer",width=150,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table.heading("Cores",text="Cores")
        cpu_table.column("Cores",width=100,stretch=tk.YES,anchor=tk.CENTER)       
        cpu_table.heading("LogicalProcessors",text="LogicalProcessors")
        cpu_table.column("LogicalProcessors",width=150,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table.heading("ProcessorId",text="ProcessorId")
        cpu_table.column("ProcessorId",width=150,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table.heading("Architecture",text="Architecture")
        cpu_table.column("Architecture",width=150,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table.heading("Caption",text="Caption")
        cpu_table.column("Caption",width=150,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table.heading("Stepping",text="Stepping")
        cpu_table.column("Stepping",width=150,stretch=tk.YES,anchor=tk.CENTER)
        
        cpu_table.pack(expand=True,fill='both')
        
        cpu_table2=ttk.Treeview(cpu_frame,columns=("DeviceID","MaxClockSpeed","Family","Revision","L2CacheSize","L3CacheSize","LoadPercentage"),show='headings')
        cpu_table2.heading("DeviceID",text="DeviceID")
        cpu_table2.column("DeviceID",width=150,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table2.heading("MaxClockSpeed",text="MaxClockSpeed")
        cpu_table2.column("MaxClockSpeed",width=150,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table2.heading("Family",text="Family")
        cpu_table2.column("Family",width=200,stretch=tk.YES,anchor=tk.CENTER) 
        cpu_table2.heading("Revision",text="Revision")
        cpu_table2.column("Revision",width=150,stretch=tk.YES,anchor=tk.CENTER)       
        cpu_table2.heading("L2CacheSize",text="L2CacheSize")
        cpu_table2.column("L2CacheSize",width=150,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table2.heading("L3CacheSize",text="L3CacheSize")
        cpu_table2.column("L3CacheSize",width=150,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table2.heading("LoadPercentage",text="LoadPercentage")
        cpu_table2.column("LoadPercentage",width=150,stretch=tk.YES,anchor=tk.CENTER)
        cpu_table2.pack(expand=True,fill='both',pady=10)
      
        cpu_info1=[]
        cpu_info2=[]
        for info in get_cpu:
            part1=info[:8]
            part2=info[8:]

            cpu_info1.append(part1)
            cpu_info2.append(part2)
        
        for info1 in cpu_info1:
            cpu_table.insert("",tk.END,values=info)
            
        for info2 in cpu_info2:
            cpu_table2.insert("",tk.END,values=info2)
            
        
    
    
    delay_thread=threading.Thread(target=run_after)
    delay_thread.start()


def ram_info():
    ram_gather=SystemRebuilt.get_Ram_info()
    del_page()
    ram_frame=tk.Frame(main_frame,borderwidth=3,relief="solid")
    lb=tk.Label(ram_frame,text="RAM Information",fg="white",font=('Impact',15))
    ram_frame.pack(pady=20)
    progress_bar()
    def run_after():
        time.sleep(1)
        lb.pack()
        ram_table=ttk.Treeview(ram_frame,columns=("Caption","Manufacturer","PartNumber","SerialNumber","Speed","Capacity"),show='headings')
        ram_table.heading("Caption",text="Caption")
        ram_table.column("Caption",width=150,stretch=tk.YES,anchor=tk.CENTER)
        ram_table.heading("Manufacturer",text="Manufacturer")
        ram_table.column("Manufacturer",width=150,stretch=tk.YES,anchor=tk.CENTER)
        ram_table.heading("PartNumber",text="PartNumber")
        ram_table.column("PartNumber",width=200,stretch=tk.YES,anchor=tk.CENTER) 
        ram_table.heading("SerialNumber",text="SerialNumber")
        ram_table.column("SerialNumber",width=150,stretch=tk.YES,anchor=tk.CENTER)       
        ram_table.heading("Speed",text="Speed")
        ram_table.column("Speed",width=150,stretch=tk.YES,anchor=tk.CENTER)
        ram_table.heading("Capacity",text="Capacity")
        ram_table.column("Capacity",width=150,stretch=tk.YES,anchor=tk.CENTER)

        ram_table.pack(expand=True,fill='both')
        
        
        
        
        for rams in ram_gather:
            
            ram_table.insert("",tk.END,values=rams)
               
        

    delay_thread=threading.Thread(target=run_after)
    delay_thread.start()

def display_info():
    del_page()
    display_gather=SystemRebuilt.get_monitor_info()
    display_frame=tk.Frame(main_frame,borderwidth=3,relief="solid")
    lb=tk.Label(display_frame,text="Monitor Information",fg="white",font=('Impact',15))
    display_frame.pack(pady=20)
    progress_bar()
    def run_after():
        time.sleep(1)
        lb.pack()
        
        monitor_table=ttk.Treeview(display_frame,columns=("DeviceID","MonitorType","Name","PNPDeviceID","Status","StatusInfo"),show='headings')
        monitor_table.heading("DeviceID",text="DeviceID")
        monitor_table.column("DeviceID",width=150,stretch=tk.YES,anchor=tk.CENTER)
        monitor_table.heading("MonitorType",text="MonitorType")
        monitor_table.column("MonitorType",width=150,stretch=tk.YES,anchor=tk.CENTER)
        monitor_table.heading("Name",text="Name")
        monitor_table.column("Name",width=150,stretch=tk.YES,anchor=tk.CENTER)
        monitor_table.heading("PNPDeviceID",text="PNPDeviceID")
        monitor_table.column("PNPDeviceID",width=350,stretch=tk.YES,anchor=tk.CENTER)        
        monitor_table.heading("Status",text="Status")
        monitor_table.column("Status",width=150,stretch=tk.YES,anchor=tk.CENTER)
        monitor_table.heading("StatusInfo",text="StatusInfo")
        monitor_table.column("StatusInfo",width=150,stretch=tk.YES,anchor=tk.CENTER)
        monitor_table.pack(expand=True,fill='both')

        
        
        
       
        for display in display_gather:
            
            monitor_table.insert("",tk.END,values=display)

        
    delay_thread=threading.Thread(target=run_after)
    delay_thread.start()

def harddisk_info():
    del_page()
    storage_gather=SystemRebuilt.get_drive_info()
    disk_frame=tk.Frame(main_frame)
    lb=tk.Label(disk_frame,text="Harddisk Information",fg="white",font=('Impact',15))
    disk_frame.pack(pady=20)
    progress_bar()
    def run_after():
        time.sleep(1)
        lb.pack()
        disk_table=ttk.Treeview(disk_frame,columns=("DeviceID","Model","SerialNumber","MediaType","InterfaceType","FirmwareRevision","Status","StatusInfo"),show='headings')
        disk_table.heading("DeviceID",text="DeviceID")
        disk_table.column("DeviceID",width=150,stretch=tk.YES,anchor=tk.CENTER)
        disk_table.heading("Model",text="Model")
        disk_table.column("Model",width=250,stretch=tk.YES,anchor=tk.CENTER)
        disk_table.heading("SerialNumber",text="SerialNumber")
        disk_table.column("SerialNumber",width=250,stretch=tk.YES,anchor=tk.CENTER) 
        disk_table.heading("MediaType",text="MediaType")
        disk_table.column("MediaType",width=180,stretch=tk.YES,anchor=tk.CENTER)       
        disk_table.heading("InterfaceType",text="InterfaceType")
        disk_table.column("InterfaceType",width=150,stretch=tk.YES,anchor=tk.CENTER)
        disk_table.heading("FirmwareRevision",text="FirmwareRevision")
        disk_table.column("FirmwareRevision",width=150,stretch=tk.YES,anchor=tk.CENTER)
        disk_table.heading("Status",text="Status")
        disk_table.column("Status",width=150,stretch=tk.YES,anchor=tk.CENTER)
        disk_table.heading("StatusInfo",text="StatusInfo")
        disk_table.column("StatusInfo",width=150,stretch=tk.YES,anchor=tk.CENTER)



        disk_table.pack(expand=True,fill='both')
        
        
        
        
        for storage in storage_gather:
            
            disk_table.insert("",tk.END,values=storage)
               
    
    delay_thread=threading.Thread(target=run_after)
    delay_thread.start()

def partition_info():
    del_page()
    part_gather=SystemRebuilt.get_disk_partition_info()
    part_frame=tk.Frame(main_frame,borderwidth=3,relief="solid")
    lb=tk.Label(part_frame,text="PARTITION Information",fg="white",font=('Impact',15))
    part_frame.pack(pady=20)
    progress_bar()
    def run_after():
        time.sleep(1)
        lb.pack()

        part_table=ttk.Treeview(part_frame,columns=("Device","Mountpoint","Fstype","Total Storage","Used Storage","Free Space"),show='headings')
        part_table.heading("Device",text="Device")
        part_table.column("Device",width=150,stretch=tk.YES,anchor=tk.CENTER)
        part_table.heading("Mountpoint",text="Mountpoint")
        part_table.column("Mountpoint",width=150,stretch=tk.YES,anchor=tk.CENTER)
        part_table.heading("Fstype",text="Fstype")
        part_table.column("Fstype",width=150,stretch=tk.YES,anchor=tk.CENTER)
        part_table.heading("Total Storage",text="Total Storage")
        part_table.column("Total Storage",width=150,stretch=tk.YES,anchor=tk.CENTER)        
        part_table.heading("Used Storage",text="Used Storage")
        part_table.column("Used Storage",width=150,stretch=tk.YES,anchor=tk.CENTER)
        part_table.heading("Free Space",text="Free Space")
        part_table.column("Free Space",width=150,stretch=tk.YES,anchor=tk.CENTER)
        part_table.pack(expand=True,fill='both')

        
       
        for part in part_gather:

            part_table.insert("",tk.END,values=part)
    delay_thread=threading.Thread(target=run_after)
    delay_thread.start()

def software_info():
    global progressbar_load3
    del_page()
    progress_bar2()
    software_frame=tk.Frame(main_frame)
    lb=tk.Label(software_frame,text="Software Information",fg="white",font=('Impact',15))
    software_frame.pack(pady=20)
    def run_after():
        while(background!=False):
            time.sleep(1)
        progressbar_load3.stop()
        app.after(100,lambda:progressbar_load3.destroy())
        lb.pack()
        software_table=ttk.Treeview(software_frame,columns=("Installed Software","Installed Date"),show='headings')
        software_table.heading("Installed Software",text="Installed Software")
        software_table.column("Installed Software",width=550)
        software_table.heading("Installed Date",text="Installed Date")
        software_table.column("Installed Date",width=250)
        software_table.pack(expand=True,fill='both')
    
        for software in software_gather:
            software_table.insert(parent='',index=0,values=software)

    delay_thread=threading.Thread(target=run_after)
    delay_thread.start()

def battery_info():
    del_page()
    progress_bar()
    battery_gather=SystemRebuilt.battery_info()
    battery_frame=tk.Frame(main_frame,borderwidth=3,relief="solid")
    lb=tk.Label(battery_frame,text="Battery Information",fg="white",font=('Impact',15))
    battery_frame.pack(pady=20)
    def run_after():
        time.sleep(1)
        lb.pack()
        battery_table=ttk.Treeview(battery_frame,columns=("Informations","Gathered"),show='headings')
        battery_table.heading("Informations",text="Informations")
        battery_table.column("Informations",width=400)
        battery_table.heading("Gathered",text="Gathered")
        battery_table.column("Gathered",width=350)
        battery_table.pack(expand=True,fill='both')
        for storage in battery_gather:
            battery_table.insert("",tk.END,values=storage)
    delay_thread=threading.Thread(target=run_after)
    delay_thread.start()


def del_page():
    for frame in main_frame.winfo_children():
        frame.destroy()






def generate_report():
    
    del_page()
    report_frame=tk.Frame(main_frame)
    lb=tk.Label(report_frame,text="Report Generator",fg="white",font=('Impact',15))
    report_frame.pack(pady=20)
    progress_bar()
    
    


   
    def run_after():
        def let_gen():
            if((cpu.get(),ram.get(),storage.get(),part.get(),display.get(),software.get())==(0,0,0,0,0,0)):
                messagebox.showinfo("Error",f"Please select the informations to be added")
                
            elif(directory_path==None):
                messagebox.showinfo("Error",f"Please Chose the directory")
            else:
                progressbar_load.pack()
                progressbar_load.start()
                n=False
                n=generate(cpu.get(),ram.get(),storage.get(),part.get(),display.get(),software.get(),directory_path)
                while n==False:
                    progressbar_load.stop()

                




        
        time.sleep(1)
        lb.pack()
        cpu=IntVar()
        ram=IntVar()
        storage=IntVar()
        part=IntVar()
        display=IntVar()
        software=IntVar()
        def choose_directory():
            global directory_path
            directory_path = filedialog.askdirectory()
            if directory_path:
                file_path=directory_path+"/"+"report.pdf"
                if os.path.isfile(file_path):
                    messagebox.showinfo("Message","OLD REPORT DELETED")
                    os.remove(file_path)
                    dir_path.configure(text=f"Chosen path:{directory_path}",fg_color="red",text_color="white")
                else:

                    dir_path.configure(text=f"Chosen path:{directory_path}",fg_color="red",text_color="white")
        
            else: 
                messagebox.showinfo("No file selected")
        
        lable=tk.Label(report_frame,text="Choose The Options Which you wan to Add to the report",font=('courier',10))
        lable.pack()
        c1= tb.Checkbutton(report_frame,text="CPU info",variable=cpu,bootstyle="square-toggle")
        c2=tb.Checkbutton(report_frame,text="RAM info",variable=ram,bootstyle="square-toggle")
        c3=tb.Checkbutton(report_frame,text="Storage info",variable=storage,bootstyle="square-toggle")
        c4=tb.Checkbutton(report_frame,text="Partition info",variable=part,bootstyle="square-toggle")
        c5=tb.Checkbutton(report_frame,text="Monitor/Display info",variable=display,bootstyle="square-toggle")
        c6=tb.Checkbutton(report_frame,text="Installed Software",variable=software,bootstyle="square-toggle")
        

        c1.pack(anchor='w',padx=5,pady=5)
        c2.pack(anchor='w',padx=5,pady=5)
        c3.pack(anchor='w',padx=5,pady=5)
        c4.pack(anchor='w',padx=5,pady=5)
        c5.pack(anchor='w',padx=5,pady=5)
        c6.pack(anchor='w',padx=5,pady=5)


        '''options=["Pdf","Txt"]'''
        top_frame=tk.Frame(report_frame)
        top_frame.pack(padx=10,pady=10,anchor='w')

        text=tk.Label(top_frame,text="Choose Report Path",font=('courier',10))
        text.pack(anchor='w')
    
        choose_button = CTkButton(master=top_frame, text="Choose Directory", command=choose_directory)
        choose_button.pack(side=tk.LEFT, pady=10)
        dir_path=CTkLabel(top_frame,text="chosen directory:None")
        dir_path.pack(side=tk.LEFT,padx=10)
        button_frame=tk.Frame(report_frame)
        button_frame.pack(pady=(0,10))
        generate_button=CTkButton(button_frame,text="Generate Report",width=200,font=('courier',15),command=let_gen)
        generate_button.pack()

        progressbar_load=CTkProgressBar(master=report_frame,orientation="horizontal")
        progressbar_load.set(0)





    delay_thread=threading.Thread(target=run_after)
    delay_thread.start()








app=tb.Window(themename="superhero")
app.geometry('1650x550')
app.title('HardwareCheck v1.0.0')
app.iconbitmap("Images/icon.ico")

title=tk.Label(app,text="HardwareCheck v1.0.0",font=("Impact",20))
title.pack()

cpuimage=Image.open("Images/cpu.png").resize((30,30))
cpuimage_tk=ImageTk.PhotoImage(cpuimage)

storageimage=Image.open("Images/storage.png").resize((30,30))
storageimage_tk=ImageTk.PhotoImage(storageimage)

displayimage=Image.open("Images/display.png").resize((30,30))
displayimage_tk=ImageTk.PhotoImage(displayimage)

exeimage=Image.open("Images/exe.png").resize((30,30))
exeimage_tk=ImageTk.PhotoImage(exeimage)

reportimage=Image.open("Images/report.png").resize((30,30))
reportimage_tk=ImageTk.PhotoImage(reportimage)


partimage=Image.open("Images/part.png").resize((30,30))
partimage_tk=ImageTk.PhotoImage(partimage)

ramimage=Image.open("Images/ram.png").resize((30,30))
ramimage_tk=ImageTk.PhotoImage(ramimage)


batteryimage=Image.open("Images/battery.png").resize((30,30))
batteryimage_tk=ImageTk.PhotoImage(batteryimage)

bannerimage=Image.open("Images/banner.png").resize((300,300))
bannerimage_tk=ImageTk.PhotoImage(bannerimage)

option_frame=tk.Frame(app,highlightthickness=2)
option_frame.pack(side=tk.LEFT)
option_frame.pack_propagate(False)
option_frame.configure(width=260,height=500)


main_frame=tk.Frame(app,highlightbackground="white",highlightthickness=2)
main_frame.pack(side=tk.LEFT)
main_frame.pack_propagate(False)
main_frame.configure(width=1600,height=500)


des=tk.Label(main_frame,text="The PYTHON GUI application is to Retrive Information about the Laptop Spec for Window Operating System  ",font=("impact",15))
des.pack(padx=50,pady=50)



desimg=tk.Label(main_frame,image=bannerimage_tk)
desimg.pack(padx=50,pady=50)




cpuinfo_button=tb.Button(option_frame,text="CPU ",width=20,bootstyle="solid nfo",command=cpu_info,image=cpuimage_tk,compound='left')
cpuinfo_button.place(x=20,y=20)

ram_button=tb.Button(option_frame,text="Ram ",width=20,bootstyle="solid nfo",command=ram_info,image=ramimage_tk,compound='left')
ram_button.place(x=20,y=80)

display_button=tb.Button(option_frame,text="Monitor ",width=20,bootstyle="solid nfo",command=display_info,image=displayimage_tk,compound='left')
display_button.place(x=20,y=140)

storage_button=tb.Button(option_frame,text="Storage ",width=20,bootstyle="solid nfo",command=harddisk_info,image=storageimage_tk,compound='left')
storage_button.place(x=20,y=200)

part_button=tb.Button(option_frame,text="Partition ",width=20,bootstyle="solid nfo",command=partition_info,image=partimage_tk,compound='left')
part_button.place(x=20,y=260)

installed_button=tb.Button(option_frame,text="Software ",width=20,bootstyle="solid nfo",command=software_info,image=exeimage_tk,compound='left')
installed_button.place(x=20,y=320)

battery_button=tb.Button(option_frame,text="Battery ",width=20,bootstyle="solid nfo",command=battery_info,image=batteryimage_tk,compound='left')
battery_button.place(x=20,y=380)

report_button=tb.Button(option_frame,text="Generate Report",width=20,bootstyle="solid nfo",command=generate_report,image=reportimage_tk,compound='left')
report_button.place(x=20,y=440)











stop_thread=threading.Event()


def get_software():
    pythoncom.CoInitialize()
    try:
        global software_gather
        software_gather=SystemRebuilt.installed_software()
       

    finally:
        pythoncom.CoUninitialize()
        global background
        background=False
        
    



def on_closing():
    if messagebox.askokcancel("Quiet","Do you want to Quit?"):
        stop_thread.set()
        background_thread.join()
        exit()
        app.quit()







background_thread=threading.Thread(target=get_software)  
background_thread.start()  
app.protocol("WM_DELETE_WINDOW",on_closing)



app.mainloop()

    

background_thread.join()   


    