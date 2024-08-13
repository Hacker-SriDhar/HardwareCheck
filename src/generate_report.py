import json
from fpdf import FPDF
import SystemRebuilt



def generate_pad_report():
    pdf=FPDF()
    pdf.add_page()
    pdf.set_font("Courier",size=12)

    cpu_get=SystemRebuilt.get_cpu_info()
    cpu_info=["Name","Manufacturer","NumberOfCores","NumberOfLogicalProcessors","ProcessorId","Architecture","Caption","Stepping","DeviceID","MaxClockSpeed","Family","Revision","L2CacheSize","L3CacheSize","LoadPercentage"]
    pdf.cell(200,10,"___________________________________________________________________",0,1,"C")
    pdf.cell(200,10,txt="CPU Information",ln=True,align="c")
    pdf.ln(10)
    
    for infos in cpu_get:
        for title,info in zip(cpu_info,infos):
            pdf.cell(200,10,txt=f"{title}:{info}")
            pdf.ln(10)
    pdf.cell(200,10,"___________________________________________________________________",0,1,"C")

    pdf.output("Informations/report.pdf")  

def generate(cpu,ram,storage,part,display,software,path,softwares_get):
    pdf=FPDF()
    pdf.add_page()
    pdf.set_font("Courier",size=12)
    if cpu==1:
        cpu_get=SystemRebuilt.get_cpu_info()
        cpu_info=["Name","Manufacturer","NumberOfCores","NumberOfLogicalProcessors","ProcessorId","Architecture","Caption","Stepping","DeviceID","MaxClockSpeed","Family","Revision","L2CacheSize","L3CacheSize","LoadPercentage"]
        pdf.set_font("Times",size=12)
        pdf.cell(200,10,txt="CPU Information",ln=True,align="c")
        pdf.ln(10)
        pdf.set_font("Courier",size=8)
        for infos in cpu_get:
            for title,info in zip(cpu_info,infos):
                pdf.cell(200,10,txt=f"{title}:{info}")
     
                pdf.ln(10)
        pdf.cell(200,10,"___________________________________________________________________",0,1,"C")
    
    if ram==1:
        ram_get=SystemRebuilt.get_Ram_info()
        ram_info=["Caption","Manufacturer","PartNumber","SerialNumber","Speed","Capacity"]
        
        pdf.ln(20)
        pdf.set_font("Times",size=12)
        pdf.cell(200,10,txt="RAM Information",ln=True,align="c")
        pdf.ln(10)
        pdf.set_font("Courier",size=8)
        for infos in ram_get:
            pdf.ln(10)
            for title,info in zip(ram_info,infos):
                pdf.cell(200,10,txt=f"{title}:{info}")
                pdf.ln(10)
        pdf.cell(200,10,"___________________________________________________________________",0,1,"C")

    if storage==1:
        storage_get=SystemRebuilt.get_drive_info()
        storage_info=["DeviceID","Model","SerialNumber","MediaType","InterfaceType","FirmwareRevision","Status","StatusInfo"]
        
        pdf.ln(20)
        pdf.set_font("Times",size=12)
        pdf.cell(200,10,txt="STORAGE Information",ln=True,align="c")
        pdf.ln(10)
        pdf.set_font("Courier",size=8)
        for infos in storage_get:
            pdf.ln(10)
            for title,info in zip(storage_info,infos):
                pdf.cell(200,10,txt=f"{title}:{info}")
                pdf.ln(10)
        pdf.cell(200,10,"___________________________________________________________________",0,1,"C")
    
    if part==1:
        part_get=SystemRebuilt.get_disk_partition_info()
        part_info=["device","mountpoint","fstype","TotalSpace","Used Space","Free Space"]
        
        pdf.ln(20)
        pdf.set_font("Times",size=12)
        pdf.cell(200,10,txt="Partition Information",ln=True,align="c")
        pdf.ln(10)
        pdf.set_font("Courier",size=8)
        for infos in part_get:
            pdf.ln(10)
            for title,info in zip(part_info,infos):
                pdf.cell(200,10,txt=f"{title}:{info}")
                pdf.ln(10)
        pdf.cell(200,10,"___________________________________________________________________",0,1,"C")
    
    if display==1:
        display_get=SystemRebuilt.get_monitor_info()
        display_info=["DeviceID","MonitorType","Name","PNPDeviceID","Status","StatusInfo"]
        
        pdf.ln(20)
        pdf.set_font("Times",size=12)
        pdf.cell(200,10,txt="Display Information",ln=True,align="c")
        pdf.ln(10)
        pdf.set_font("Courier",size=8)
        for infos in display_get:
            pdf.ln(10)
            for title,info in zip(display_info,infos):
                pdf.cell(200,10,txt=f"{title}:{info}")
                pdf.ln(10)
        pdf.cell(200,10,"___________________________________________________________________",0,1,"C")

    if software==1:
        software_get=softwares_get
        
        
        pdf.ln(20)
        pdf.set_font("Times",size=12)
        pdf.cell(200,10,txt="Software Information",ln=True,align="c")
        pdf.ln(10)
        pdf.set_font("Courier",size=8)
        for i in range(len(software_get)):
            pdf.cell(200,10,txt=f"{software_get[i][0]}\t:\t{software_get[i][1]}")
            pdf.ln(10)
        pdf.cell(200,10,"___________________________________________________________________",0,1,"C")

    org_path=path+'/'+"report.pdf"
    pdf.output(org_path)
    return True
    