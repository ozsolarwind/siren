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

import math
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
                ['ni ', 'NI '], ['Npv', 'NPV'], ['Re', 'RE'], ['Tco2E', 'tCO2e'],
                ['REference', 'Reference']]
        for each in alll:
            cleantech = cleantech.replace(each[0], each[1])
    return cleantech

#
# add another windspeed height
def extrapolateWind(wind_file, tgt_height, law='logarithmic', replace=False):
    if not os.path.exists(wind_file):
        if replace:
            return False
        else:
            return None
    if wind_file[-4:] != '.srw':
        if replace:
            return False
        else:
            return None
    tf = open(wind_file, 'r')
    lines = tf.readlines()
    tf.close()
    fst_row = 5
    units = lines[3].rstrip(',\n').split(',')
    hghts = lines[4].rstrip(',\n').split(',')
    col = -1
    heights_ms = []
    heights_dirn = []
    for j in range(len(units)):
        if units[j] == 'm/s':
             heights_ms.append([int(hghts[j]), j])
             if heights_ms[-1][0] == tgt_height:
                 if replace:
                     return False
                 else:
                     return None
        elif units[j] == 'degrees':
             heights_dirn.append([int(hghts[j]), j])
    lines[2] = lines[2].rstrip(',\n') + ',Direction,Speed\n'
    lines[3] = lines[3].rstrip(',\n') + ',degrees,m/s\n'
    lines[4] = lines[4].rstrip(',\n') + ',' + str(tgt_height) + ',' + str(tgt_height) + '\n'
    heights_ms.sort(key=lambda x: x[0], reverse=True)
    heights_dirn.sort(key=lambda x: x[0], reverse=True)
    height = float(heights_ms[0][0])
    col = heights_ms[0][1]
    height0 = float(heights_ms[1][0])
    if height0 == height:
        if replace:
            return False
        else:
            return None
    col0 = heights_ms[1][1]
    cold = heights_dirn[0][1]
    for i in range(fst_row, len(lines)):
        bits = lines[i].rstrip(',\n').split(',')
        speed = float(bits[col])
        speed0 = float(bits[col0])
        if speed0 >= speed:
            alpha = 1. / 7. # one-seventh power law
        else:
            alpha = (math.log(speed)-math.log(speed0))/(math.log(height)-math.log(height0))
        z0 = math.exp(((pow(height0, alpha) * math.log(height)) - pow(height, alpha) * math.log(height0)) \
                      / ( pow(height0, alpha) - pow(height, alpha)))
        if z0 < 1e-308:
            z0 = 0.03
        if law.lower()[0] == 'l': # law == 'logarithmic'
            speedz = math.log(tgt_height / z0) / math.log(height0 / z0) * speed0
            lines[i] = lines[i].strip() + ',' + bits[cold] + ',' + str(round(speedz, 4)) + '\n'
        else: # law == 'hellmann'
            speeda = pow(tgt_height / height0, alpha) * speed0
            lines[i] = lines[i].strip() + ',' + bits[cold] + ',' + str(round(speeda, 4)) + '\n'
    if replace:
        if os.path.exists(wind_file + '~'):
            os.remove(wind_file + '~')
        os.rename(wind_file, wind_file + '~')
        nf = open(wind_file, 'w')
        for line in lines:
            nf.write(line)
        nf.close()
        return True
    else:
        return lines
        array = [] # this doesn't work yet
        for i in range(4):
            bits = lines[i].split(',')
            array.append(bits)
        for i in range(fst_row, len(lines)):
            bits = lines[i].split(',')
            for j in range(len(bits)):
                bits[j] = float(bits[j])
            array.append(bits)
        return array

# split a string
def strSplit(string, char=',', dropquote=True):
    last = 0
    splits = []
    inQuote = None
    for i, letter in enumerate(string):
        if inQuote:
            if (letter == inQuote):
                inQuote = None
                if dropquote:
                    splits.append(string[last:i])
                    last = i + 1
                    continue
        elif (letter == '"' or letter == "'"):
            inQuote = letter
            if dropquote:
                last += 1
        elif letter == char:
            if last != i:
                splits.append(string[last:i])
            last = i + 1
    if last < len(string):
        splits.append(string[last:])
    return splits
