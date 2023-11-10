#!/usr/bin/python3
#
import os
import shutil
import sys

def delist_dir(directory):
    for filename in os.listdir(directory):
        if os.path.isdir(directory + filename):
            delist_dir(directory + filename + '\\')
      #      print directory + filename + '\\'
            os.rmdir(directory + filename + '\\')
        else:
            if filename == 'epsg':
                continue
            os.remove(directory + filename)

# \dist\siren will contain sirenm and most files
dirs = ['app_flexiplot', 'app_getera5', 'app_getmap', 'app_getmerra2', 'app_indexweather',
        'app_makegrid', 'app_makeweatherfiles', 'app_powermatch', 'app_powerplot',
        'app_siren', 'app_sirenupd', 'app_updateswis']
for di in dirs:
    print('Including:', di)
    fils = sorted(os.listdir('dist\\' + di))
    for fil in fils:
        if os.path.isdir('dist\\' + di + '\\' + fil):
            print('Directory encountered:', fil)
            try:
                shutil.copytree('dist\\' + di + '\\' + fil, 'dist\\siren\\' + fil)
            except:
                pass
        else:
            try:
                shutil.copy2('dist\\' + di + '\\' + fil, 'dist\\siren')
            except:
                pass

del_dirs = [] #'tcl\\tzdata', 'tcl\\encoding', 'mpl-data\\fonts', 'pytz\\zoneinfo',
           # 'mpl_toolkits\\basemap\\data']
for di in del_dirs:
    if os.path.exists('dist\\siren\\' + di):
        print('Removing:', di)
        src_dir = 'dist\\siren\\' + di + '\\'
        delist_dir(src_dir)
