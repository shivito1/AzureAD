import subprocess
from os import environ
import sys
from tkinter import messagebox, Tk
import pyautogui
from time import sleep
import os, winshell
from win32com.client import Dispatch
import shutil
import os
result1 = False
result2 = False

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def creaetshortc(path,name):
    path1 = os.path.join(path, name+".lnk")
    target = path
    wDir = path
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path1)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wDir
    shortcut.save()

def print_file(file_path):
    file_path = resource_path(file_path)
    print(file_path)

print_file('connect+.png')

if "AzureAdJoined : YES" in str(subprocess.check_output("dsregcmd /status")):
    print("joind")
    result1 = messagebox.askokcancel("Azure AD","This computer is already joined to Azure AD")
elif "AzureAdJoined : NO" in str(subprocess.check_output("dsregcmd /status")):
    print("Not Joined")
    result2 = messagebox.askyesno("Azure AD", "This computer has not joined Azure AD. Join now?")

    desktop_src_dir = environ['USERPROFILE']+"\\"+"Desktop"
    desktop_tar_dir = r"C:\second_user_folder"+"\\"+"Desktop_folder"
    desktop_names = os.listdir(desktop_src_dir)
    os.makedirs(desktop_tar_dir, mode=511, exist_ok=True)

    documents_src_dir = environ['USERPROFILE']+"\\"+"Documents"
    documents_tar_dir = r"C:\second_user_folder"+"\\"+"Documents_folder"
    documents_names = os.listdir(documents_src_dir)
    os.makedirs(documents_tar_dir, mode=511, exist_ok=True)

    downloads_src_dir = environ['USERPROFILE']+"\\"+"Downloads"
    downloads_tar_dir = r"C:\second_user_folder"+"\\"+"Downloads_folder"
    downloads_names = os.listdir(downloads_src_dir)
    os.makedirs(downloads_tar_dir, mode=511, exist_ok=True)

    print(desktop_names)
    #Desktop copy
    print(desktop_names)
    for file_name in desktop_names:
        shutil.move(os.path.join(desktop_src_dir, file_name), desktop_tar_dir)

    #Documents copy
    print(documents_names)
    for file_name in documents_names:
        shutil.move(os.path.join(documents_src_dir, file_name), documents_tar_dir)

    #Downloads copy
    print(downloads_names)
    for file_name in downloads_names:
        shutil.move(os.path.join(downloads_src_dir, file_name), downloads_tar_dir)

    creaetshortc(desktop_tar_dir,"Desktop")
    creaetshortc(documents_tar_dir, "Documents")
    creaetshortc(downloads_tar_dir, "Downloads")

if result2:
    os.system("start ms-settings:workplace")
    start = None
    maxtime = 0
    bk = ""
    while start == None:
        if maxtime >= 30:
            bk = "break"
            break
        start = pyautogui.locateCenterOnScreen(resource_path("data_files\\connect+.png"), confidence="0.7",
                                               grayscale=True)  # If the file is not a png file it will not work
        if start == None:
            start = pyautogui.locateCenterOnScreen(resource_path("data_files\\connectlight.png"), confidence="0.7",
                                                   grayscale=True)  # If the file is not a png file it will not work
        print(start)
        if start != None:
            pyautogui.click(start)  # Moves the mouse to the coordinates of the image
        sleep(1)
        maxtime = maxtime + 1
    maxtime = 0
    start = None

while start == None:
    if maxtime >= 30:
        bk = "break"
        break
    start = pyautogui.locateCenterOnScreen(resource_path("data_files\\AzureAD.png"), confidence="0.7", grayscale=True)
    print(start)
    if start != None:
        pyautogui.click(start)  # Moves the mouse to the coordinates of the image

    sleep(1)
    maxtime = maxtime + 1

sleep(1)
pyautogui.typewrite("email@address.com")
pyautogui.press('enter')
sleep(1)
pyautogui.typewrite("testing")
pyautogui.press('enter')

if bk == "break":
    print("Broken")
    seit = messagebox.askyesno("Access work or school", 'Sorry Auto detect failed. Do you see "Access work or school"')
    if seit:
        messagebox.askyesno('Azure AD', 'Please click the big plus symbol and at the bottom of the pop up window click "Join this device to Azure Active Directory" \n After that sign in with your hills email and password.')
    elif seit == False:
        messagebox.askokcancel("Call John","My cell: 513-814-8882")
#else:
    #messagebox.askokcancel("Let's get you signed in", "On the sign in screen use your Hills email and password.")