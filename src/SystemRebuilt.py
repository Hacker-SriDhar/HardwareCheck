import psutil
import win32com.client
import subprocess
import psutil
import wmi
import winreg
from datetime import datetime


def list_installed_software():
    software_list = []
    try:
        result=subprocess.run(['wmic','product','get','Name'],capture_output=True,text=True)
        if result.returncode == 0:
            output=result.stdout
            lines=output.splitlines()[1:]
            software_list=[line.strip() for line in lines if line.strip()]
    except Exception as e:
        print(f"error occured:{e}")
    return software_list


def installed_software():
    c=wmi.WMI()
    software_list=[]

    for software in c.Win32_Product():
        name=software.Name
        install_date=software.InstallDate

        if install_date:
            try:
                install_date=datetime.strptime(install_date,"%Y%m%d").date()
            except ValueError:
                install_date="N/A"
        else:
            install_date='N/A'

        software_list.append((name,install_date))

    return software_list

def get_installed_software(registry_path):
    software_list=[]
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,registry_path) as key:
            index=0
            while True:
                try:
                    subkey_name=winreg.EnumKey(key,index)
                    with winreg.OpenKey(key,subkey_name) as subkey:
                        try:
                            display_name=winreg.QueryValueEx(subkey,'DisplayName')[0]
                        except FileNotFoundError:
                            continue
                        try:
                            install_date_str=winreg.QueryValueEx(subkey,'InstallDate')[0]
                            install_date=datetime.strptime(install_date_str,'%Y%m%d').date()
                        except (FileNotFoundError,ValueError):
                            install_date='N/A'
                        
                        print(display_name,install_date)
                        software_list.append((display_name,install_date))
                except OSError:
                    break
                index +=1
    except FileNotFoundError:
        print(f"Registry path {registry_path} not found")

    return software_list
                        




"""
software_list=installed_software()
for software,install_date in software_list:
    print(f"{software}-Install Date:{install_date}")
"""

def get_battery_info():
    try:
        result=subprocess.run(['wmic','path','Win32_Battery','get','/format:list'],capture_output=True,text=True)
        print(result.stdout)
    except Exception as e:
        print(f"Error Regarding {e}")





def get_monitor_info():
    wmi=win32com.client.GetObject("winmgmts://")
    monitor_info=[]
    for monitor in wmi.InstancesOf("Win32_DesktopMonitor"):
        monitor_info.append([monitor.DeviceID,
                             monitor.MonitorType,
                             monitor.Name,
                             monitor.PNPDeviceID,
                             monitor.Status,
                             monitor.StatusInfo
                           
                           
                           ])

    return monitor_info


def get_drive_info():
    wmi=win32com.client.GetObject("winmgmts://")
    drive_info=[]
    for drive in wmi.InstancesOf("Win32_DiskDrive"):
        drive_info.append([drive.DeviceID,
                           drive.Model,
                           drive.SerialNumber,
                           drive.MediaType,
                           drive.InterfaceType,
                           drive.FirmwareRevision,
                           drive.Status,
                           drive.StatusInfo
                           
                           
                           ])

    return drive_info


def get_cpu_info():
    wmi=win32com.client.GetObject("winmgmts://")
    cpu_info=[]
    for cpu in wmi.InstancesOf("Win32_Processor"):
        cpu_info.append([
        cpu.Name,
        cpu.Manufacturer,
        cpu.NumberOfCores,
        cpu.NumberOfLogicalProcessors,
        cpu.ProcessorId,
        cpu.Architecture,
        cpu.Caption,
        cpu.Stepping,
        cpu.DeviceID,
        cpu.MaxClockSpeed,
        cpu.Family,
        cpu.Revision,
        cpu.L2CacheSize,
        cpu.L3CacheSize,
        cpu.LoadPercentage,
        ])
        
                          

    return cpu_info  





def get_Ram_info():
    wmi=win32com.client.GetObject("winmgmts://")
    ram_info=[]
    for ram in wmi.InstancesOf("Win32_PhysicalMemory"):
        ram_info.append([  ram.Caption,
                           ram.Manufacturer,
                           ram.PartNumber,
                           ram.SerialNumber,
                           ram.Speed,
                           int(ram.Capacity)/(1024**3)
                            
                           ])

    return ram_info




def get_ram_info():
    virtual_memory=psutil.virtual_memory()
    return(virtual_memory.total/(1024**3))



def get_disk_partition_info():
    partitions=psutil.disk_partitions()
    disk_info=[]
    for partition in partitions:
        usage=psutil.disk_usage(partition.mountpoint)
        disk_info.append([partition.device,
                          partition.mountpoint,
                          partition.fstype,
                          int(usage.total/(1024**3)),
                          int(usage.used/(1024**3)),
                          int(usage.free/(1024**3))])
    return disk_info


















"""
disk_infoes=get_disk_info()
for disk_info in disk_infoes:
    print("Device:",disk_info[0])
    print("Total Disk Size:",disk_info[1])
    print("used Disk Size:",disk_info[2])
    print("Free Disk Size:",disk_info[3])
    print("\n")





drive_info=get_monitor_info()

for drive in drive_info:
    print(drive)




"""

def battery_info():
    try:
        res=subprocess.run(['wmic','path','Win32_Battery','get','/format:list'],stdout=subprocess.PIPE,text=True)
        output=res.stdout

        lines=output.strip().split('\n')
        battery_info={}
        info=[]
        for line in lines:
            if '=' in line:
                key,value=line.split('=',1)
                org=value.strip()
                if org=="":
                    
                    org=" N/A"
                    battery_info[key.strip()]=org
                    info.append([key,org])
                else:
                    battery_info[key.strip()]=org
                    info.append([key,value])

        #print(info)
        #for key,value in battery_info.items():
            #print(f"{key}:{value}")
        return info
    except Exception as e:
        return e




battery_info()

















