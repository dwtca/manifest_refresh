import os
from shutil import rmtree
import subprocess
from socket import gethostname
from time import sleep
from xml.etree import ElementTree as ET
import tkinter as tk
from tkinter import messagebox

startupinfo = subprocess.STARTUPINFO()
startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

folder_path = r"C:\ProgramData\Zoll Data Systems\eDistribution\DownloadTemp"
xml_file_path = r"C:\ProgramData\ZOLL data systems\eDistribution\eDistribution.Manifest.xml"
services = [
    "ZOLL Data RescueNet eDistribution Client",
    "ZDSMachineLicensingService",
    "ZOLL Data RescueNet Mercury Messaging Service"
]
exe_path = r"C:\Program Files (x86)\ZOLL Data Systems\Common\Bin\ZollData.MSCS.eDistribution.ClientNotifier.exe"
icon_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "settings.ico")

def start_task():
    text.delete('1.0', tk.END)
    
    text.insert(tk.END, "Deleting Temporary eDistribution folders... ")
    root.update_idletasks()
    for folder_name in os.listdir(folder_path):
        if folder_name != "Uncompressed" and folder_name.isdigit():
            rmtree(os.path.join(folder_path, folder_name))
    text.insert(tk.END, "Done!\n")

    text.insert(tk.END, "Removing files list from Manifest... ")
    root.update_idletasks()
    tree = ET.parse(xml_file_path)
    xml_root = tree.getroot()
    for files in xml_root.findall('Files'):
        files.clear()
    tree.write(xml_file_path)
    text.insert(tk.END, "Done!\n")

    text.insert(tk.END, "Stopping Zoll Services... ")
    root.update_idletasks()
    for service in services:
        subprocess.Popen(["net", "stop", service], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, startupinfo=startupinfo)
        sleep(1)
    text.insert(tk.END, "Done!\n")

    messagebox.showinfo("eDistribution", f"Please remove and re-add {gethostname()} in System Group.")

    text.insert(tk.END, "Restarting Zoll Services... ")
    root.update_idletasks()
    for service in reversed(services):
        subprocess.Popen(["net", "start", service], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, startupinfo=startupinfo)
        sleep(2)
    text.insert(tk.END, "Done!\n")

    while True:
        text.insert(tk.END, "Waiting for eDistrubition... ")
        result = subprocess.check_output(f'sc query "{service}"', shell=True).decode()
        if "RUNNING" in result:
            text.insert(tk.END, "Done!\n")
            root.update_idletasks()
            subprocess.Popen(exe_path, shell=True)
            text.insert(tk.END, "Launching eDistribution Client Notifier")
            break
        else:
            sleep(1)  # Wait for 1 second before checking again

    messagebox.showinfo("Success", "Manifest and Temporary Files cleared.")
    root.destroy()

root = tk.Tk()
root.title("Zoll Manifest Refresh")
root.iconbitmap(icon_path)
button = tk.Button(root, text="Start", command=start_task, padx=50, pady=20)
button.pack(pady=20)
text = tk.Text(root, width=50, height=10)
text.pack()
root.mainloop()
