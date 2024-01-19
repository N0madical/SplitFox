import os, winreg, winshell, random
from win32com.client import Dispatch

profName = "default"

os.system(f"mklink /j \"C:\\Program Files\\Mozilla Firefox - {profName}\" \"C:\\Program Files\\Mozilla Firefox\"")

base = winreg.HKEY_LOCAL_MACHINE

try:
    location = winreg.OpenKeyEx(base, r"SOFTWARE\\Mozilla\\Firefox\\TaskBarIDs")
    value = winreg.QueryValueEx(location, f"C:\\Program Files\\Mozilla Firefox - {profName}")
    print("Found it!", value)
except:
    try:
        location = winreg.OpenKey(base, r"SOFTWARE\\Mozilla\\Firefox\\TaskBarIDs", 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(location, f"C:\\Program Files\\Mozilla Firefox - {profName}", 1, winreg.REG_SZ, ("308046B0" + str(random.randrange(10000000, 99999999))))
    except:
        location = winreg.OpenKeyEx(base, r"SOFTWARE\\Mozilla\\Firefox\\TaskBarIDs")
        value = winreg.QueryValueEx(location, f"C:\\Program Files\\Mozilla Firefox - {profName}")
        print("Wrote Value to Registry:", value)

if location:
    winreg.CloseKey(location)

taskbar = os.getenv('APPDATA') + "\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar\\"
path = os.path.join(taskbar, f"{profName}.lnk")
shell = Dispatch('WScript.Shell')
shortcut = shell.CreateShortCut(path)
shortcut.Targetpath = rf"C:\Program Files\Mozilla Firefox - {profName}\firefox.exe"
shortcut.Arguments = f"-p {profName}"
shortcut.WorkingDirectory = rf"C:\Program Files\Mozilla Firefox - {profName}"
shortcut.IconLocation = rf"C:\Program Files\Mozilla Firefox - {profName}\firefox.exe"
shortcut.save()