import os
cmd1 = "pyinstaller -F -w Main.py --hidden-import=PyQt5.sip"
os.system(cmd1)
print('Converted Processing done...')
