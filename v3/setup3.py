import sys
import os
from cx_Freeze import setup, Executable
from PIL import Image

# Desktop = os.path.expanduser("~/Desktop")
exe = Executable(
    script=r"gui_test.py",
    base="Win32GUI",
    targetName="WhatsAppSender.exe",
    shortcutName="WhatsAppSender",
    shortcutDir="DesktopFolder",
    icon="whatsapp.ico",
    )

exe_list = list()
exe_list.append(exe)
# backend = Executable(
#     script=r"backend_driver.py",
#     base="Console",
#     targetName="backend.exe",
#     # icon="important/logiqids.png",
#     )
# TARGETDIR = 
# icon_table = ("whatsapp.ico")

# shortcut_table = [
#     ("DesktopShortcut",        # Shortcut
#      "DesktopFolder",          # Directory_
#      "WhatsAppSender",         # Name
#      "TARGETDIR",              # Component_
#      "[TARGETDIR]WhatsAppSender.exe",        # Target
#      None,                     # Arguments
#      None,                     # Description
#      None,                     # Hotkey
#      None,                     # Icon
#      None,                      # IconIndex
#      None,                     # ShowCmd
#      'TARGETDIR'               # WkDir
#      )
#     ]

# msi_data = {"Shortcut": shortcut_table}

# bdist_msi_options = {'data': msi_data}


build_exe_options = {'include_files': ['important', 'whatsapp.exe'], }
# options = {"build_exe": build_exe_options, ""}
options = {
          # 'bdist_msi': bdist_msi_options,
          'build_exe': build_exe_options,
          }

setup(
    name = "WhatsAppSender",
    version = "0.1",
    description = "Send whatsapp messages",
    options = options,
    executables = exe_list
    )