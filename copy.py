import os, shutil
shutil.copy('out/build/x64-release/xls2csv.dll', 'xls2csv.pyd')
shutil.copy('xls2csv.pyd', r'C:\Samo\Trunk\data\buildgdd\xls2csv.pyd')