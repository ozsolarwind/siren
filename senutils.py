#!/usr/bin/python3
#
#  Copyright (C) 2015-2022 Sustainable Energy Now Inc., Angus King
#
#  senutils.py - This file is part of SIREN.
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
try:
    import pwd
except:
    pass
import sys
from PyQt5 import QtCore, QtWidgets


class ClickableQLabel(QtWidgets.QLabel):
    clicked = QtCore.pyqtSignal()

    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtWidgets.QApplication.widgetAt(event.globalPos()).setFocus()
        self.clicked.emit()

#
# replace parent string in filenames
def getParents(aparents):
    parents = []
    for key, value in aparents:
        for key2, value2 in aparents:
            if key2 == key:
                continue
            value = value.replace(key2, value2)
        for key2, value2 in parents:
            if key2 == key:
                continue
            value = value.replace(key2, value2)
        parents.append((key, value))
    return parents

#
# return current userid
def getUser():
    if sys.platform == 'win32' or sys.platform == 'cygwin':   # windows
        return os.environ.get("USERNAME")
    elif sys.platform == 'darwin':   # osx64
        return pwd.getpwuid(os.geteuid()).pw_name
    elif sys.platform == 'linux' or sys.platform == 'linux2':   # linux
        return pwd.getpwuid(os.geteuid()).pw_name
    else:
        return os.environ.get("USERNAME")

#
# clean up tech names
def techClean(tech, full=False):
    cleantech = tech.replace('_', ' ').title()
    cleantech = cleantech.replace('Bm', 'BM')
    cleantech = cleantech.replace('Ccgt', 'CCGT')
    cleantech = cleantech.replace('Ccg', 'CCG')
    cleantech = cleantech.replace('Cst', 'CST')
    cleantech = cleantech.replace('Lng', 'LNG')
    cleantech = cleantech.replace('Ocgt', 'OCGT')
    cleantech = cleantech.replace('Ocg', 'OCG')
    cleantech = cleantech.replace('Phs', 'PHS')
    cleantech = cleantech.replace('Pv', 'PV')
    if full:
        alll = [['Cf', 'CF'], ['hi ', 'HI '], ['Lcoe', 'LCOE'], ['Mw', 'MW'],
                ['ni ', 'NI '], ['Npv', 'NPV'], ['Re', 'RE'], ['Tco2E', 'tCO2e']]
        for each in alll:
            cleantech = cleantech.replace(each[0], each[1])
    return cleantech
