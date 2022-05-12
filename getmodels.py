#!/usr/bin/python3
#
#  Copyright (C) 2020-2022 Sustainable Energy Now Inc., Angus King
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
                siren_dir = newdir
            mf = open('siren_models_location.txt', 'w')
            mf.write(siren_dir)
            mf.close()
            updir = mydir[:mydir.rfind(fldr_div) + 1]
            if os.path.exists(updir + 'siren_sample') \
              and os.path.exists(mydir + fldr_div + 'SIREN.ini') \
              and not os.path.exists(siren_dir + 'SIREN.ini'): # has the sample
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
                    copy(mydir + fldr_div +  'SIREN.ini', siren_dir + 'SIREN.ini')
        return siren_dir + ini_file

    siren_dir = ''
    if len(args) > 0:
        ini_file = args[0]
        if os.path.exists('siren_models_location.txt'):
            mf = open('siren_models_location.txt', 'r')
            siren_dir = mf.readline()
            mf.close()
            siren_dir = siren_dir.strip('\n')
            siren_dir = siren_dir + ini_file
            return siren_dir
        else:
            return ini_file
    ini_file = ''
    if os.path.exists('siren_models_location.txt'):
        mf = open('siren_models_location.txt', 'r')
        siren_dir = mf.readline()
        mf.close()
        siren_dir = siren_dir.strip('\n')
        if not(os.path.exists(siren_dir)):
            app = QtWidgets.QApplication.instance()
            if app is None:
                app = QtWidgets.QApplication(sys.argv)
            msgbox = QtWidgets.QMessageBox()
            msgbox.setWindowTitle("SIREN - Can't find Models")
            msgbox.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
            msgbox.setText('SIREN Models folder missing')
            msgbox.setInformativeText('Do you want to reset the Models location (Y)?')
            msgbox.setDetailedText("Can't find " + siren_dir + '. ' + \
                "If you reply 'Y'es you can choose a new location for the Models.")
            msgbox.setIcon(QtWidgets.QMessageBox.Question)
            msgbox.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            reply = msgbox.exec_()
            if reply == QtWidgets.QMessageBox.Yes:
                siren_dir = set_models_locn(ini_file)
            else:
                sys.exit(8)
        siren_dir = siren_dir + ini_file
    else:
        siren_dir = set_models_locn(ini_file)
    return siren_dir
