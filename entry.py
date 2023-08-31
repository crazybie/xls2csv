import sys, os, xls2csv
import builtins
from io import StringIO


csv_files = xls2csv.load_xls_dir(r'C:\Samo\Trunk\data\GameDatas\datas', '')


open_old = open
def open_new(name, mode='r', *args, **kwargs):        
    if mode == 'r':
        base = os.path.basename(name)        
        if base.endswith('.csv'):
            d = csv_files[base]
            #print('--hit', name)
            return StringIO(d)        
    return open_old(name, mode, *args, **kwargs)    
builtins.open = open_new


listdir_old = os.listdir
def listdir_new(p):
    if p == 'dummy':
        return csv_files.keys()
    return listdir_old(p)
    
os.listdir = listdir_new


sys.argv.append('dummy')
sys.argv.append(r'C:\Samo\Trunk\data\buildgdd\csvs')
exec(open(r'C:\Samo\Trunk\data\buildgdd\csvalid.py').read())