#!/usr/bin/python3
#
#  Copyright (C) 2020-2025 Sustainable Energy Now Inc., Angus King
#
#  getmodels.py - This file is part of SIREN.
#
#  SIREN is free software: you can redistribute it and/or modify
#  it under the terms of the GNU Affero General Public License as
#  published by the Free Software Foundation, either version 3 of
#  the License, or (at your option) any later version.
#
#  SIREN is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU Affero General Public License for more details.
#
#  You should have received a copy of the GNU Affero General
#  Public License along with SIREN.  If not, see
#  <http://www.gnu.org/licenses/>.
#

import configparser
import os
from PyQt5 import QtGui, QtWidgets
from shutil import copy
import sys

def commonprefix(args, chr='/'):
    arg2 = []
    for arg in args:
        arg2.append(arg)
        if arg[-1] != chr:
            arg2[-1] += chr
    return os.path.commonprefix(arg2).rpartition(chr)[0]

def getModelFile(*args):
    def set_models_locn(ini_file):
        app = QtWidgets.QApplication.instance()
        if app is None:
            app = QtWidgets.QApplication(sys.argv)
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            fldr_div = '\\'
        else:
            fldr_div = '/'
        siren_dir = '.' + fldr_div
        newdir = QtWidgets.QFileDialog.getExistingDirectory(None,
                 'SIREN. Choose location for SIREN Model (Preferences) files',
                 '', QtWidgets.QFileDialog.ShowDirsOnly)
        if newdir != '':
            mydir = os.path.abspath(__file__)
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                newdir = newdir.replace('/', fldr_div)
            mydir = mydir[:mydir.rfind(fldr_div)]
            that_len = len(commonprefix([newdir, mydir], chr=fldr_div))
            if that_len > 0:
                bits = mydir[that_len:].split(fldr_div)
                if that_len < len(mydir): # go up the tree
                    pfx = ('..' + fldr_div) * (len(bits) - 1)
                else:
                    pfx = ''
                siren_dir = pfx + newdir[that_len + 1:]
                if siren_dir == '':
                    siren_dir = '.'
                if siren_dir[-1] != fldr_div:
                    siren_dir += fldr_div
            else:
                siren_dir = newdir + fldr_div
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                pfx = os.getenv('LOCALAPPDATA') + fldr_div + 'siren' + fldr_div
                if not os.path.exists(pfx):
                    os.mkdir(pfx)
            else:
                pfx = os.path.expanduser('~') + '/.siren/'
                if not os.path.exists(pfx):
                    os.mkdir(pfx)
            mf = open(pfx + 'siren_models_location.txt', 'a')
            mf.write(siren_dir + '\n')
            mf.close()
            updir = mydir[:mydir.rfind(fldr_div) + 1]
            # copy getfiles.ini
            if not os.path.exists(siren_dir + 'getfiles.ini') \
              and os.path.exists(mydir + fldr_div + 'getfiles.ini'): # no getfiles.ini
                copy(mydir + fldr_div +  'getfiles.ini', siren_dir + 'getfiles.ini')
            if os.path.exists(mydir + fldr_div + 'SIREN.ini') \
              and not os.path.exists(siren_dir + 'SIREN.ini'): # has the sample
                have_sample = False
                if os.path.exists(updir + 'siren_sample'):
                    have_sample = True
                    msi = ''
                elif sys.platform == 'win32' or sys.platform == 'cygwin': # maybe in ProgramData
                    up1more = updir[:updir[:-1].rfind(fldr_div) + 1]
                    up2more = up1more[:up1more[:-1].rfind(fldr_div) + 1]
                    if os.path.exists(f'{up2more}\\ProgramData\\SIREN\\siren_sample') or \
                      os.path.exists('C:\\ProgramData\\SIREN\\siren_sample'):
                        have_sample = True
                        msi = 'msi'
                if have_sample:
                    msgbox = QtWidgets.QMessageBox()
                    msgbox.setWindowTitle('SIREN - Copy File')
                    msgbox.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
                    msgbox.setText('SIREN sample_files folder found')
                    msgbox.setInformativeText('Do you want to copy the sample SIREN Model (Y)?')
                    msgbox.setDetailedText('It seems you have the siren sample files. ' + \
                        "If you reply 'Y'es the sample preferences file, SIREN.ini, will " + \
                        'be copied to '+ siren_dir + '.')
                    msgbox.setIcon(QtWidgets.QMessageBox.Question)
                    msgbox.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
                    reply = msgbox.exec_()
                    if reply == QtWidgets.QMessageBox.Yes:
                        if os.path.exists(siren_dir + 'SIREN.ini'):
                            if os.path.exists(siren_dir + 'SIREN.ini~'):
                                os.remove(siren_dir + 'SIREN.ini~')
                            os.rename(siren_dir + 'SIREN.ini', siren_dir + 'SIREN.ini~')
                        copy(f'{mydir}{fldr_div}SIREN{msi}.ini', siren_dir + 'SIREN.ini')
        return siren_dir + ini_file

    def get_models_dirs(args):
        models_dirs = []
        models_location = os.getcwd() + '/siren_models_location.txt'
        if not os.path.exists(models_location):
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                models_location = os.getenv('LOCALAPPDATA') + '\\siren\\siren_models_location.txt'
            else:
                models_location = os.path.expanduser('~') + '/.siren/siren_models_location.txt'
        if os.path.exists(models_location):
            mf = open(models_location, 'r')
            maybe_dirs = mf.readlines()
            mf.close()
            models_dirs = []
            for maybe_dir in maybe_dirs:
                maybe_dir = maybe_dir.strip('\n')
                if os.path.exists(maybe_dir):
                    models_dirs.append(maybe_dir)
            if len(models_dirs) > 1:
                dups = []
                for i in range(len(models_dirs)):
                    for j in range(i + 1, len(models_dirs)):
                        if os.path.samefile(models_dirs[i], models_dirs[j]):
                            if j not in dups:
                                dups.append(j)
                dups.sort(reverse=True)
                for dup in dups:
                    del models_dirs[dup]
        return models_location, models_dirs

    def get_common(config, common_files, model_file):
        config.read(model_file)
        try:
            common = config.get('Base', 'common').split(',')
            for cfil in common:
                if cfil[-4:] != '.ini':
                    cfil += '.ini'
                if os.path.exists(cfil):
                    common_file = cfil
                if cfil.find('/') > 0 and os.path.exists(os.getcwd() + '/' + cfil):
                    common_file = os.getcwd() + '/' + cfil
                else:
                    common_file = models_dir.strip('\n') + cfil
                if os.path.exists(common_file):
                    if common_file not in common_files:
                        common_files.append(common_file)
                        common_files = get_common(config, common_files, common_file)
        except:
            pass
        return common_files

    models_location, models_dirs = get_models_dirs(args)
    if len(args) > 0:
        if isinstance(args[0], list):
            ini_file = args[0]
        else:
            ini_file = args[0].split(',')
        common_files = []
        model_files = []
        config = configparser.RawConfigParser()
        for models_dir in models_dirs:
            for fil in ini_file:
                if fil.find('/') > 0: # and fil[:len(models_dir.strip('\n'))] == models_dir.strip('\n'):
                    model_file = fil
                else:
                    model_file = models_dir.strip('\n') + fil
                if os.path.exists(model_file):
                    if model_file not in model_files:
                        model_files.append(model_file)
                    common_files = get_common(config, common_files, model_file)
        if len(model_files) > 0:
            for cfil in common_files:
                if cfil not in model_files:
                    model_files.insert(0, cfil)
 #           if len(model_files) > 1:
  #              for f in range(len(model_files) - 1, -2, -1):
   #                 if os.path.samefile(model_files[f - 1], model_files[f]):
    #                    del model_files[f]
            return model_files
        return ini_file # default to just return filename
    if len(models_dirs) == 0:
        app = QtWidgets.QApplication.instance()
        if app is None:
            app = QtWidgets.QApplication(sys.argv)
        msgbox = QtWidgets.QMessageBox()
        msgbox.setWindowTitle("SIREN - Can't find Models")
        msgbox.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        msgbox.setText('SIREN Models folder missing')
        msgbox.setInformativeText('Do you want to reset the Models location (Y)?')
        msgbox.setDetailedText("Can't find locations in\n'" + models_location + "'.\n" + \
           "If you reply 'Y'es you can choose a new location for the Models.")
        msgbox.setIcon(QtWidgets.QMessageBox.Question)
        msgbox.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        reply = msgbox.exec_()
        if reply == QtWidgets.QMessageBox.Yes:
            good_dirs = [set_models_locn('')]
        else:
            sys.exit(8)
        return good_dirs
    else:
        return models_dirs
