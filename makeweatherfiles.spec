# -*- mode: python -*-
#
# This is a pyinstaller spec file to build a version of makeweatherfiles.py for Windows
#
import os

def list_dir(directory):
    for filename in os.listdir(directory):
        if os.path.isdir(directory + filename):
            list_dir(directory + filename + '\\')
      #      print directory + filename + '\\'
            os.rmdir(directory + filename + '\\')
        else:
            if filename == 'epsg':
                continue
            os.remove(directory + filename)

this_dir = os.getcwd()
block_cipher = None
a = Analysis(['makeweatherfiles.py'],
             pathex=[this_dir],
             binaries=None,
             datas=None,
             hiddenimports=['FileDialog', 'netCDF4_utils', 'Tkinter'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='makeweatherfiles',
          debug=False,
          strip=False,
          upx=True,
          console=True,
          icon=this_dir + os.sep + 'sen_icon32.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas + [('COPYING.txt', 'GNU Affero General Public License Windows.txt', 'DATA'),
                          ('makeweatherfiles.py', 'makeweatherfiles.py', 'DATA'),
                          ('makeweatherfiles.html', 'makeweatherfiles.html', 'DATA')],
               strip=False,
               upx=True,
               name='makeweatherfiles')

del_dirs = ['tcl\\tzdata', 'tcl\\encoding']
for di in del_dirs:
    print 'Removing:', di
    src_dir = 'dist\\makeweatherfiles\\' + di + '\\'
    list_dir(src_dir)