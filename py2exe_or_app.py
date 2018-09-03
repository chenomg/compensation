import os
cmd1 = "pyinstaller -F -w --icon=icon.ico Main.py --hidden-import=PyQt5.sip"
os.system(cmd1)
print('Converted Processing done...')
