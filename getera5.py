#!/usr/bin/python3
#
#  Copyright (C) 2023 Sustainable Energy Now Inc., Angus King
#
#  getera5.py - This file is part of SIREN.
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

import cdsapi
import datetime
import os
import subprocess
import sys
from netCDF4 import Dataset
import configparser   # decode .ini file
from PyQt5 import QtCore, QtGui, QtWidgets

from credits import fileVersion
import displayobject
from editini import SaveIni
from getmodels import getModelFile
from senutils import ClickableQLabel, getUser
import worldwindow

def spawn(who, log):
    pid = ''
    cwd = os.getcwd()
    stdoutf = log
    stdout = open(stdoutf, 'wb')
    if sys.argv[0][-4:] == '.exe': # avoid need for console window?
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    else:
        startupinfo = None
    try:
        if type(who) is list:
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                if os.path.exists(who[0] + '.py'):
                    who[0] += '.py'
                    shell = True
                elif os.path.exists(who[0] + '.exe'):
                    who[0] += '.exe'
                    shell = False
                else:
                    return pid, who[0] + ' not found'
                who_str = ''
                for wh in who:
                    who_str += '"' + wh + '" '
                pid = subprocess.Popen(who_str, shell=shell, cwd=cwd, startupinfo=startupinfo, stderr=subprocess.STDOUT,
                                       stdout=stdout).pid
            else:
                who[0] += '.py'
                who.insert(0, 'python3')
                pid = subprocess.Popen(who, cwd=cwd, startupinfo=startupinfo, stderr=subprocess.STDOUT,
                               stdout=stdout).pid
        else:
            # may not work as it ain;t been tested
            pid = subprocess.Popen([who], cwd=cwd, startupinfo=startupinfo,
                                   stderr=subprocess.STDOUT, stdout=stdout).pid
    except:
        return pid, sys.exc_info()[0]
        pass
    return pid, None

def the_period(period, way='-'):
    year = int(period[:4])
    mth = int(period[-2:])
    if way in ['-', '<']:
        if mth == 1:
            return '{:0>4d}{}'.format(year - 1, '12')
        else:
            return '{:0>4d}{:0>2d}'.format(year, mth - 1)
    elif way in ['+', '>']:
        if mth == 12:
            return '{:0>4d}{}'.format(year + 1, '01')
        else:
            return '{:0>4d}{:0>2d}'.format(year, int(mth + 1))
    return period

def checkFiles(tgt_dir, ini_file=None):
    def get_range(top, prefix, suffix):
        chk_folders.append(top)
        chk_src_files.append([])
        ndx = len(chk_src_files) - 1
        fils = sorted(os.listdir(top))
        for fil in fils:
            if os.path.isdir(top + '/' + fil):
                get_range(top + '/' + fil, prefix, suffix)
            elif fil[-len(suffix):] == suffix:
                if prefix == '' or fil[:len(prefix)] == prefix:
                    chk_src_files[ndx].append(fil)
        del fils

    msg_text = ''
    config = configparser.RawConfigParser()
    config.read(ini_file)
    try:
        filename = config.get('getera5', 'filename')
    except:
        filename = '.nc'
    if filename.find('$') > 0:
        bits = filename.split('$')
        file_pfx = bits[0]
        file_sfx = bits[-1]
    else:
        file_pfx = ''
        file_sfx = filename[filename.rfind('.'):]
    chk_folders = []
    chk_src_files = []
    get_range(tgt_dir, file_pfx, file_sfx)
    log = None
    fst_period = ''
    lst_period = ''
    gap_periods = []
    chk_days = []
    for i in range(len(chk_folders)):
        if len(chk_src_files[i]) == 0:
            continue
        for src_file in chk_src_files[i]:
            if file_pfx == '':
                period = src_file[:-len(file_sfx)]
            else:
                period = src_file[len(file_pfx):-len(file_sfx)]
            if len(period) == 4:
                if fst_period == '':
                    fst_period = period + '01'
                elif int(period) != int(lst_period[:4]) + 1:
                    gap_periods.append(str(int(lst_period[:4]) + 1))
                lst_period = period + '12'
            elif len(period) == 6:
                chk_period = the_period(period)
                if chk_period != lst_period:
                    gap_periods.append(the_period(lst_period, '+'))
                lst_period = period
            else:
                chk_days.append(period)
            log = fileInfo(chk_folders[i] + '/' + src_file)
            if not log.ok:
                msg_text = log.log
                return [log.log]
        break
    if log is None:
        msg_text = 'Check ERA5 incomplete.'
        return [msg_text]
    lats = log.latitudes[-1]
    latn = log.latitudes[0]
    lonw = log.longitudes[0]
    lone = log.longitudes[-1]
    grd1 = abs(log.latitudes[0] - log.latitudes[1])
    grd2 = abs(log.longitudes[0] - log.longitudes[1])
    msg_text = 'Boundaries and period set for ERA5 data'
    if fst_period != '':
        reqd = ''
        if lone >= 15.:
            chk_day = '{:0>4d}1231'.format(int(fst_period[:4]) - 1)
            if chk_day not in chk_days:
                reqd = '; {}{}{}'.format(file_pfx, chk_day, file_sfx)
        if lonw <= -15.0:
            chk_day = '{:0>4d}0101'.format(int(lst_period[:4]) + 1)
            if chk_day not in chk_days:
                if reqd != '':
                    reqd += ' & {}{}{}'.format(file_pfx, chk_day, file_sfx)
                else:
                    reqd = '; {}{}{}'.format(file_pfx, chk_day, file_sfx)
        if reqd != '':
            reqd += ' needed'
        msg_text += ' (' + fst_period
        if lst_period == fst_period:
            msg_text += ' exists'
        else:
            msg_text += ' to ' + lst_period + ' exist'
            if len(gap_periods) > 0:
                msg_text += ' with ' + str(len(gap_periods)) + ' gaps'
                print(gap_periods)
        msg_text += reqd + ')'
    return [msg_text, latn, lats, lonw, lone, grd1, grd2, the_period(lst_period, '+')]

def retrieve_era5(ini_file, lat1, lat2, lon1, lon2, grd1, grd2, year, tgt_dir, launch=False):
    config = configparser.RawConfigParser()
    config.read(ini_file)
    tgt_file = config.get('getera5', 'filename').replace('$year$', year)
    tgt_log = tgt_file[:tgt_file.rfind('.')] + '.log'
    if launch:
        parmstr = 'getera5,ini={},lat1={:.2f},lat2={:.2f},lon1={:.2f},lon2={:.2f},grd1={:.2f},grd2={:.2f},year={},tgt_dir={}'.format(ini_file,
                  lat1, lat2, lon1, lon2, grd1, grd2, year, tgt_dir)
        parms = parmstr.split(',')
        spawn(parms, tgt_dir + '/' + tgt_log)
        return 'Request for ' + tgt_file + ' launched.'
    era5_dict = {'product_type': 'reanalysis',
                 'format': 'netcdf'}
    variables = []
    var_list = config.get('getera5', 'variables').split(',')
    for var in var_list:
        variables.append(config.get('getera5', 'var_' + var.strip()))
    era5_dict['variable'] = variables
    if len(year) > 6:
        era5_dict['year'] = year[:4]
        era5_dict['month'] = year[4 : -2]
        era5_dict['day'] = year[-2:]
    else:
        days = []
        for d in range(1, 32):
            days.append('{:0>2d}'.format(d))
        era5_dict['day'] = days
        if len(year) > 4:
            era5_dict['year'] = year[:4]
            era5_dict['month'] = year[-2:]
        else:
            era5_dict['year'] = year
            mths = []
            for m in range(1, 13):
                mths.append('{:0>2d}'.format(m))
            era5_dict['month'] = mths
    times = []
    for h in range(24):
        times.append('{:0>2d}:00'.format(h))
    era5_dict['time'] = times
    era5_dict['area'] = [lat1, lon1, lat2, lon2]
    era5_dict['grid'] = [grd1, grd2]
    c = cdsapi.Client(verify=True)
    c.retrieve('reanalysis-era5-single-levels',
               era5_dict,
               tgt_dir + '/' + tgt_file)


class fileInfo:
    def __init__(self, inp_file):
        self.ok = False
        self.log = ''
        if not os.path.exists(inp_file):
            self.log = 'File not found: ' + inp_file
            return
        try:
            cdf_file = Dataset(inp_file, 'r')
        except:
            self.log = 'Error reading: ' + inp_file
            return
        i = inp_file.rfind('/')
        self.file = inp_file[i + 1:]
        try:
            self.format = cdf_file.Format
        except:
            self.format = '?'
        self.dimensions = ''
        keys = list(cdf_file.dimensions.keys())
        values = list(cdf_file.dimensions.values())
        if type(cdf_file.dimensions) is dict:
            for i in range(len(keys)):
                self.dimensions += keys[i] + ': ' + str(values[i]) + ', '
        else:
            for i in range(len(keys)):
                bits = str(values[i]).strip().split()
                self.dimensions += keys[i] + ': ' + bits[-1] + ', '
        self.variables = ''
        for key in iter(sorted(cdf_file.variables.keys())):
            self.variables += key + ', '
        self.latitudes = []
        try:
            latitude = cdf_file.variables['lat'][:]
        except:
            latitude = cdf_file.variables['latitude'][:]
        lati = latitude[:]
        for val in lati:
            self.latitudes.append(val)
        self.longitudes = []
        try:
            longitude = cdf_file.variables['lon'][:]
        except:
            longitude = cdf_file.variables['longitude'][:]
        longi = longitude[:]
        for val in longi:
            self.longitudes.append(val)
        cdf_file.close()
        self.ok = True


class subwindow(QtWidgets.QDialog):
    def __init__(self, parent = None):
        super(subwindow, self).__init__(parent)

    def closeEvent(self, event):
        event.accept()

    def exit(self):
        self.close()


class getERA5(QtWidgets.QDialog):
    procStart = QtCore.pyqtSignal(str)
    statusmsg = QtCore.pyqtSignal()

    def get_config(self):
        self.config = configparser.RawConfigParser()
        self.config_file = self.ini_file
        self.config.read(self.config_file)
        try:
            self.help = self.config.get('Files', 'help')
        except:
            self.help = 'help.html'
        self.restorewindows = False
        try:
            rw = self.config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
        self.zoom = 0.8
        try:
            self.zoom = float(self.config.get('View', 'zoom_rate'))
            if self.zoom > 1:
                self.zoom = 1 / self.zoom
            if self.zoom > 0.95:
                self.zoom = 0.95
            elif self.zoom < 0.75:
                self.zoom = 0.75
        except:
            pass
        self.wait_days = 93
        try:
            self.wait_days = int(self.config.get('getera5', 'wait_days'))
        except:
            pass
        self.retrieve_year = False
        try:
            yr = self.config.get('getera5', 'retrieve_year')
            if yr.lower() in ['true', 'yes', 'on']:
                self.retrieve_year = True
        except:
            pass

    def __init__(self, help='help.html', ini_file='getfiles.ini', parent=None):
        super(getERA5, self).__init__(parent)
        self.help = help
        self.ini_file = getModelFile(ini_file)
        self.get_config()
        self.ignore = False
        self.worldwindow = None
        ok = False
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            try:
                cdsapirc = '~\\.cdsapirc'.replace('~', os.environ['HOME'])
            except:
                cdsapirc = ''
        else:
            cdsapirc = '~/.cdsapirc'.replace('~', os.path.expanduser('~'))
        if cdsapirc != '' and os.path.exists(cdsapirc):
            ok = True
        if not ok:
            cdsapircmsg = '.cdsapirc file missing.\n(' + cdsapirc + ')\nDo you want create one?'
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                cdsapircmsg += '\nYou will need to reinvoke getera5.'
            reply = QtWidgets.QMessageBox.question(self, 'SIREN - getera5 - No .cdsapirc file',
                    cdsapircmsg,
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.Yes:
                self.create_cdsapirc()
        self.northSpin = QtWidgets.QDoubleSpinBox()
        self.northSpin.setDecimals(3)
        self.northSpin.setSingleStep(.5)
        self.northSpin.setRange(-90, 90)
        self.northSpin.setObjectName('north')
        self.westSpin = QtWidgets.QDoubleSpinBox()
        self.westSpin.setDecimals(3)
        self.westSpin.setSingleStep(.5)
        self.westSpin.setRange(-180, 180)
        self.westSpin.setObjectName('west')
        self.southSpin = QtWidgets.QDoubleSpinBox()
        self.southSpin.setDecimals(3)
        self.southSpin.setSingleStep(.5)
        self.southSpin.setRange(-90, 90)
        self.southSpin.setObjectName('south')
        self.eastSpin = QtWidgets.QDoubleSpinBox()
        self.eastSpin.setDecimals(3)
        self.eastSpin.setSingleStep(.5)
        self.eastSpin.setRange(-180, 180)
        self.eastSpin.setObjectName('east')
        self.latSpin = QtWidgets.QDoubleSpinBox()
        self.latSpin.setDecimals(3)
        self.latSpin.setSingleStep(.5)
        self.latSpin.setRange(-90, 90)
        self.latSpin.setObjectName('lat')
        self.lonSpin = QtWidgets.QDoubleSpinBox()
        self.lonSpin.setDecimals(3)
        self.lonSpin.setSingleStep(.5)
        self.lonSpin.setRange(-180, 180)
        self.lonSpin.setObjectName('lon')
        self.latwSpin = QtWidgets.QDoubleSpinBox()
        self.latwSpin.setDecimals(3)
        self.latwSpin.setSingleStep(.5)
        self.latwSpin.setRange(0, 180)
        self.latwSpin.setObjectName('latw')
        self.lonwSpin = QtWidgets.QDoubleSpinBox()
        self.lonwSpin.setDecimals(3)
        self.lonwSpin.setSingleStep(.5)
        self.lonwSpin.setRange(0, 360)
        self.lonwSpin.setObjectName('lonw')
        if len(sys.argv) > 1:
            his_config_file = sys.argv[1]
            his_config = configparser.RawConfigParser()
            his_config.read(his_config_file)
            try:
                mapp = his_config.get('Map', 'map_choice')
            except:
                mapp = ''
            try:
                 upper_left = his_config.get('Map', 'upper_left' + mapp).split(',')
                 self.northSpin.setValue(float(upper_left[0].strip()))
                 self.westSpin.setValue(float(upper_left[1].strip()))
                 lower_right = his_config.get('Map', 'lower_right' + mapp).split(',')
                 self.southSpin.setValue(float(lower_right[0].strip()))
                 self.eastSpin.setValue(float(lower_right[1].strip()))
            except:
                 try:
                     lower_left = his_config.get('Map', 'lower_left' + mapp).split(',')
                     upper_right = his_config.get('Map', 'upper_right' + mapp).split(',')
                     self.northSpin.setValue(float(upper_right[0].strip()))
                     self.westSpin.setValue(float(lower_left[1].strip()))
                     self.southSpin.setValue(float(lower_left[0].strip()))
                     self.eastSpin.setValue(float(upper_right[1].strip()))
                 except:
                     pass
        self.northSpin.valueChanged.connect(self.showArea)
        self.westSpin.valueChanged.connect(self.showArea)
        self.southSpin.valueChanged.connect(self.showArea)
        self.eastSpin.valueChanged.connect(self.showArea)
        self.latSpin.valueChanged.connect(self.showArea)
        self.lonSpin.valueChanged.connect(self.showArea)
        self.latwSpin.valueChanged.connect(self.showArea)
        self.lonwSpin.valueChanged.connect(self.showArea)
        self.grid = QtWidgets.QGridLayout()
        self.grid.addWidget(QtWidgets.QLabel('Area of Interest:'), 0, 0)
        area = QtWidgets.QPushButton('Choose area via Map', self)
        self.grid.addWidget(area, 0, 1, 1, 2)
        area.clicked.connect(self.areaClicked)
        self.grid.addWidget(QtWidgets.QLabel('Upper left:'), 1, 0, 1, 1)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(QtWidgets.QLabel('North'), 2, 0)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.northSpin, 2, 1)
        self.grid.addWidget(QtWidgets.QLabel('West'), 3, 0)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.westSpin, 3, 1)
        self.grid.addWidget(QtWidgets.QLabel('Lower right:'), 1, 2)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(QtWidgets.QLabel('South'), 2, 2)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.southSpin, 2, 3)
        self.grid.addWidget(QtWidgets.QLabel('East'), 3, 2)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.eastSpin, 3, 3)
        self.grid.addWidget(QtWidgets.QLabel('Centre:'), 1, 4)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(QtWidgets.QLabel('Lat.'), 2, 4)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.latSpin, 2, 5)
        self.grid.addWidget(QtWidgets.QLabel('Lon.'), 3, 4)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.lonSpin, 3, 5)
        self.grid.addWidget(QtWidgets.QLabel('Degrees'), 1, 6)
        self.grid.addWidget(self.latwSpin, 2, 6)
        self.grid.addWidget(self.lonwSpin, 3, 6)
        self.grid.addWidget(QtWidgets.QLabel('Grid:'), 4, 0)
        try:
            era5_grid = self.config.get('getera5', 'grid').split(',')
        except:
            era5_grid = ['0.25', '0.25']
        self.era5grid = []
        eragrid = QtWidgets.QHBoxLayout()
        for i in range(2):
            self.era5grid.append(QtWidgets.QDoubleSpinBox())
            self.era5grid[-1].setValue(float(era5_grid[i].strip()))
            self.era5grid[-1].setDecimals(2)
            self.era5grid[-1].setSingleStep(.25)
            self.era5grid[-1].setRange(0.25, 3.0)
            self.era5grid[-1].valueChanged.connect(self.showArea)
            eragrid.addWidget(self.era5grid[-1])
            if i == 0:
                eragrid.addWidget(QtWidgets.QLabel(' x'))
        rw = 4
        eragrid.addWidget(QtWidgets.QLabel('(size of weather "cells")'))
        self.grid.addLayout(eragrid, rw, 1, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Approx. area:'), rw, 0)
        self.approx_area = QtWidgets.QLabel('')
        self.grid.addWidget(self.approx_area, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('ERA5 dimensions:'), rw, 2)
        self.merra_cells = QtWidgets.QLabel('')
        self.grid.addWidget(self.merra_cells, rw, 3, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Start month:'), rw, 0)
        self.strt_date = QtWidgets.QDateEdit(self)
        self.strt_date.setDate(QtCore.QDate.currentDate().addDays(-self.wait_days))
        self.strt_date.setCalendarPopup(False)
        self.strt_date.setMinimumDate(QtCore.QDate(1940, 1, 1))
        self.strt_date.setMaximumDate(QtCore.QDate.currentDate().addDays(-self.wait_days))
        self.strt_date.setDisplayFormat('yyyy-MM')
        self.grid.addWidget(self.strt_date, rw, 1)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('End month:'), rw, 0)
        self.end_date = QtWidgets.QDateEdit(self)
        self.end_date.setDate(QtCore.QDate.currentDate().addDays(-self.wait_days))
        self.end_date.setCalendarPopup(False)
        self.end_date.setMinimumDate(QtCore.QDate(1940,1,1))
        self.end_date.setMaximumDate(QtCore.QDate.currentDate().addDays(-self.wait_days))
        self.end_date.setDisplayFormat('yyyy-MM')
        self.grid.addWidget(self.end_date, rw, 1)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Variables:'), rw, 0)
        self.grid.addWidget(QtWidgets.QLabel(self.config.get('getera5', 'variables')), rw, 1, 1, 5)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Target Folder:'), rw, 0)
        cur_dir = os.getcwd()
        self.tgt_dir = ClickableQLabel()
        try:
            self.tgt_dir.setText(self.config.get('Files', 'era5_files').replace('$USER$', getUser()))
        except:
            self.tgt_dir.setText(cur_dir)
        self.tgt_dir.setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
        self.tgt_dir.clicked.connect(self.dirChanged)
        self.grid.addWidget(self.tgt_dir, rw, 1, 1, 8)
        rw += 2
        self.log = QtWidgets.QLabel('')
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.log.setPalette(msg_palette)
        self.grid.addWidget(self.log, rw, 1, 2, 4)
        sw = self.northSpin.minimumSizeHint().width()
        dw = self.strt_date.minimumSizeHint().width()
        if sw > dw: # fix for wide QDoubleSpinBox width in Windows
            self.northSpin.setMinimumWidth(self.strt_date.minimumSizeHint().width())
            self.westSpin.setMinimumWidth(self.strt_date.minimumSizeHint().width())
            self.southSpin.setMinimumWidth(self.strt_date.minimumSizeHint().width())
            self.eastSpin.setMinimumWidth(self.strt_date.minimumSizeHint().width())
            self.latSpin.setMinimumWidth(self.strt_date.minimumSizeHint().width())
            self.lonSpin.setMinimumWidth(self.strt_date.minimumSizeHint().width())
            self.latwSpin.setMinimumWidth(self.strt_date.minimumSizeHint().width())
            self.lonwSpin.setMinimumWidth(self.strt_date.minimumSizeHint().width())
        buttongrid = QtWidgets.QGridLayout()
        quit = QtWidgets.QPushButton('Done', self)
        buttongrid.addWidget(quit, 0, 0)
        quit.clicked.connect(self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('x'), self, self.quitClicked)
        getdata = QtWidgets.QPushButton('Get Data', self)
        buttongrid.addWidget(getdata, 0, 1)
        getdata.clicked.connect(self.getClicked)
        check = QtWidgets.QPushButton('Check Data', self)
        buttongrid.addWidget(check, 0, 2)
        check.clicked.connect(self.checkClicked)
        help = QtWidgets.QPushButton('Help', self)
        buttongrid.addWidget(help, 0, 5)
        help.clicked.connect(self.helpClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        frame2 = QtWidgets.QFrame()
        frame2.setLayout(buttongrid)
        self.layout.addWidget(frame2)
        self.setWindowTitle('SIREN - getera5 (' + fileVersion() + ') - Get ERA5 data')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.resize(int(self.sizeHint().width() * 1.5), int(self.sizeHint().height() * 1.07))
        self.updated = False
        self.show()

    def dirChanged(self):
        curdir = self.tgt_dir.text()
        newdir = QtWidgets.QFileDialog.getExistingDirectory(self, 'Choose Target Folder',
                 curdir, QtWidgets.QFileDialog.ShowDirsOnly)
        if newdir != '':
            self.tgt_dir.setText(newdir)
            self.updated = True

    def helpClicked(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.help,
                 title='Help for getting ERA5 data (' + fileVersion() + ')', section='getera5')
        dialog.exec_()

    def exit(self):
        self.close()

    def closeEvent(self, event):
        try:
            self.mySubwindow.close()
        except:
            pass
        event.accept()

    @QtCore.pyqtSlot(list, str)
    def maparea(self, rectangle, approx_area=None):
        if type(rectangle) is str:
            if rectangle == 'goodbye':
                 self.worldwindow = None
                 return
        elif type(rectangle[0]) is str:
            if rectangle[0] == 'goodbye':
                 self.worldwindow = None
                 return
        self.ignore = True
        self.northSpin.setValue(rectangle[0].y())
        self.westSpin.setValue(rectangle[0].x())
        self.southSpin.setValue(rectangle[1].y())
        self.eastSpin.setValue(rectangle[1].x())
        self.ignore = False
        self.approx_area.setText(approx_area)
        merra_dims = worldwindow.merra_cells(self.northSpin.value(), self.westSpin.value(),
                     self.southSpin.value(), self.eastSpin.value(),
                     self.era5grid[0].value(), self.era5grid[1].value())
        self.merra_cells.setText(merra_dims)

    def areaClicked(self):
        if self.worldwindow is None:
            scene = worldwindow.WorldScene()
            self.worldwindow = worldwindow.WorldWindow(self, scene, era5=True)
            self.worldwindow.view.tellarea.connect(self.maparea)
            self.worldwindow.show()
            self.showArea('init')

    def showArea(self, event):
        if self.ignore:
            return
        if self.sender().objectName() == 'north' or self.sender().objectName() == 'south':
            self.ignore = True
            if self.southSpin.value() > self.northSpin.value():
                y = self.northSpin.value()
                self.northSpin.setValue(self.southSpin.value())
                self.southSpin.setValue(y)
            self.latwSpin.setValue(self.northSpin.value() - self.southSpin.value())
            self.latSpin.setValue(self.northSpin.value() - (self.northSpin.value() - self.southSpin.value()) / 2.)
            self.ignore = False
        elif self.sender().objectName() == 'east' or self.sender().objectName() == 'west':
            self.ignore = True
            if self.eastSpin.value() < self.westSpin.value():
                x = self.westSpin.value()
                self.westSpin.setValue(self.eastSpin.value())
                self.eastSpin.setValue(x)
            self.lonwSpin.setValue(self.eastSpin.value() - self.westSpin.value())
            self.lonwSpin.setValue(self.eastSpin.value() - (self.eastSpin.value() - self.westSpin.value()) / 2.)
            self.ignore = False
        elif self.sender().objectName() == 'lat':
            self.ignore = True
            if self.latwSpin.value() == 0:
                self.latwSpin.setValue(0.25)
            self.northSpin.setValue(self.latSpin.value() + self.latwSpin.value() / 2)
            self.southSpin.setValue(self.latSpin.value() - self.latwSpin.value() / 2)
            self.ignore = False
        elif self.sender().objectName() == 'lon':
            self.ignore = True
            if self.lonwSpin.value() == 0:
                self.lonwSpin.setValue(0.25)
            self.eastSpin.setValue(self.lonSpin.value() + self.lonwSpin.value() / 2)
            self.westSpin.setValue(self.lonSpin.value() - self.lonwSpin.value() / 2)
            self.ignore = False
        elif self.sender().objectName() == 'latw':
            self.ignore = True
            self.northSpin.setValue(self.latSpin.value() + self.latwSpin.value() / 2)
            self.southSpin.setValue(self.latSpin.value() - self.latwSpin.value() / 2)
            self.ignore = False
        elif self.sender().objectName() == 'lonw':
            self.ignore = True
            self.eastSpin.setValue(self.lonSpin.value() + self.lonwSpin.value() / 2)
            self.westSpin.setValue(self.lonSpin.value() - self.lonwSpin.value() / 2)
            self.ignore = False
        merra_dims = worldwindow.merra_cells(self.northSpin.value(), self.westSpin.value(),
                     self.southSpin.value(), self.eastSpin.value(),
                     self.era5grid[0].value(), self.era5grid[1].value())
        self.merra_cells.setText(merra_dims)
        if self.worldwindow is None:
            return
        approx_area = self.worldwindow.view.drawRect([QtCore.QPointF(self.westSpin.value(), self.northSpin.value()),
                                   QtCore.QPointF(self.eastSpin.value(), self.southSpin.value())])
        self.approx_area.setText(approx_area)
        if event != 'init':
            self.worldwindow.view.statusmsg.emit(approx_area + ' ' + merra_dims)

    def quitClicked(self):
        if self.updated:
            updates = {}
            lines = []
            lines.append('era5_files=' + self.tgt_dir.text().replace(getUser(), '$USER$'))
            updates['Files'] = lines
            SaveIni(updates, ini_file=self.config_file)
        self.close()

    def checkClicked(self):
        self.log.setText('')
        check = checkFiles(self.tgt_dir.text(), ini_file=self.ini_file)
        if len(check) > 1: # ok
            self.ignore = True
            self.northSpin.setValue(check[1])
            self.southSpin.setValue(check[2])
            self.latwSpin.setValue(self.northSpin.value() - self.southSpin.value())
            self.latSpin.setValue(self.northSpin.value() - (self.northSpin.value() - self.southSpin.value()) / 2.)
            self.westSpin.setValue(check[3])
            self.eastSpin.setValue(check[4])
            self.lonwSpin.setValue(self.eastSpin.value() - self.westSpin.value())
            self.lonSpin.setValue(self.eastSpin.value() - (self.eastSpin.value() - self.westSpin.value()) / 2.)
            self.era5grid[0].setValue(check[5])
            self.era5grid[1].setValue(check[6])
            self.strt_date.setDate(QtCore.QDate(int(check[7][:4]), int(check[7][-2:]), 1))
            self.end_date.setDate(QtCore.QDate(int(check[7][:4]), 12, 31))
            self.ignore = False
            merra_dims = worldwindow.merra_cells(self.northSpin.value(), self.westSpin.value(),
                         self.southSpin.value(), self.eastSpin.value(),
                         self.era5grid[0].value(), self.era5grid[1].value())
            self.merra_cells.setText(merra_dims)
        self.log.setText(check[0])
        if self.worldwindow is None:
            return
        approx_area = self.worldwindow.view.drawRect([QtCore.QPointF(self.westSpin.value(), self.northSpin.value()),
                                   QtCore.QPointF(self.eastSpin.value(), self.southSpin.value())])
        self.approx_area.setText(approx_area)
        self.worldwindow.view.statusmsg.emit(approx_area + ' ' + merra_dims)

    def getClicked(self):
        def get_file(year):
            wgot = retrieve_era5(self.ini_file,
                                 self.northSpin.value(), self.southSpin.value(),
                                 self.westSpin.value(), self.eastSpin.value(),
                                 self.era5grid[0].value(), self.era5grid[1].value(),
                                 year, self.tgt_dir.text(), True)
            self.log.setText(wgot)
        try:
            filename = self.config.get('getera5', 'filename')
        except:
            filename = '.nc'
        if filename.find('$') > 0:
            bits = filename.split('$')
            file_pfx = bits[0]
            file_sfx = bits[-1]
        else:
            file_pfx = ''
            file_sfx = filename[filename.rfind('.'):]
        strt_period = '{:0>4d}{:0>2d}'.format(int(self.strt_date.date().year()),
                                              int(self.strt_date.date().month()))
        stop_period = '{:0>4d}{:0>2d}'.format(int(self.end_date.date().year()),
                                              int(self.end_date.date().month()))
        if self.eastSpin.value() >= 15. and self.strt_date.date().month() == 1:
            chk_period = the_period(strt_period, '-')
            lst_fil = '{}{}{}'.format(file_pfx, chk_period[:4], file_sfx)
            if not os.path.exists(self.tgt_dir.text() + '/' + lst_fil):
                lst_fil = '{}{}{}'.format(file_pfx, chk_period, file_sfx)
                if not os.path.exists(self.tgt_dir.text() + '/' + lst_fil):
                    fst_fil = '{}{}31{}'.format(file_pfx, chk_period, file_sfx)
                    if not os.path.exists(self.tgt_dir.text() + '/' + fst_fil):
                        get_file(chk_period + '31')
        if self.westSpin.value() <= -15. and self.end_date.date().month() == 12:
            chk_period = the_period(stop_period, '+')
            nxt_fil = '{}{}{}'.format(file_pfx, chk_period[:4], file_sfx)
            if not os.path.exists(self.tgt_dir.text() + '/' + nxt_fil):
                nxt_fil = '{}{}{}'.format(file_pfx, chk_period, file_sfx)
                if not os.path.exists(self.tgt_dir.text() + '/' + nxt_fil):
                    nxt_fil = '{}{}01{}'.format(file_pfx, chk_period, file_sfx)
                    if not os.path.exists(self.tgt_dir.text() + '/' + nxt_fil):
                        get_file(chk_period + '01')
        if self.retrieve_year and strt_period[:4] == stop_period[:4] and \
            strt_period[-2:] == '01' and stop_period[-2:] == '12':
            get_file(strt_period[:4])
            return
        nxt_period = strt_period
        while nxt_period <= stop_period:
            the_file = '{}{}{}'.format(file_pfx, nxt_period, file_sfx)
            if not os.path.exists(self.tgt_dir.text() + '/' + the_file):
                get_file(nxt_period)
            nxt_period = the_period(nxt_period, '+')
        return

    def create_cdsapirc(self):
        self.mySubwindow = subwindow()
        grid = QtWidgets.QGridLayout()
        grid.addWidget(QtWidgets.QLabel('Enter details for URS Registration'), 0, 0, 1, 2)
        r = 1
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            try:
                self.home_dir = os.environ['HOME']
            except:
                self.home_dir = os.getcwd()
                grid.addWidget(QtWidgets.QLabel('HOME directory:'), 1, 0)
                self.home = ClickableQLabel()
                self.home.setText(self.home_dir)
                self.home.setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
                self.home.clicked.connect(self.ursdirChanged)
                grid.addWidget(self.home, 1, 1, 1, 3)
                r = 2
        grid.addWidget(QtWidgets.QLabel('API UID:'), r, 0)
        self.api_uid = QtWidgets.QLineEdit('')
        grid.addWidget(self.api_uid, r, 1)
        grid.addWidget(QtWidgets.QLabel('API Key:'), r + 1, 0)
        self.api_key = QtWidgets.QLineEdit('')
        grid.addWidget(self.api_key, r + 1, 1)
        cdsapirc_button = QtWidgets.QPushButton('Create .cdsapirc file', self.mySubwindow)
        cdsapirc_button.clicked.connect(self.createcdsapirc)
        grid.addWidget(cdsapirc_button, r + 2, 0)
        self.mySubwindow.setLayout(grid)
        self.mySubwindow.exec_()

    def ursdirChanged(self):
        curdir = self.home.text()
        newdir = QtWidgets.QFileDialog.getExistingDirectory(self, 'Choose .cdsapirc Folder',
                 curdir, QtWidgets.QFileDialog.ShowDirsOnly)
        if newdir != '':
            self.home.setText(newdir)
            self.home_dir = newdir

    def createcdsapirc(self):
        config = configparser.RawConfigParser()
        config.read(ini_file)
        api_url = config.get('getera5', 'api_url')
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            env_var = 'HOME'
            os.system('SETX {0} "{1}"'.format(env_var, self.home_dir))
            cdsapirc = self.home_dir + '\\.cdsapirc'
        else:
            cdsapirc = '~/.cdsapirc'.replace('~', os.path.expanduser('~'))
        cdsapirc_string = 'url: %s\nkey: %s:%s' % (api_url, self.api_uid.text(), self.api_key.text())
        fou = open(cdsapirc, 'wb')
        fou.write(cdsapirc_string.encode())
        fou.close()
        self.mySubwindow.close()
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            self.close()


if '__main__' == __name__:
    batch = False
    check = False
    ini_file = 'getfiles.ini'
    if len(sys.argv) > 2:  # arguments
        batch = True
    elif len(sys.argv) == 2:
        if sys.argv[1][-4:] == '.ini':
            ini_file = sys.argv[1]
        else:
            batch = True
    if batch:
        ini_parms = ['ini', 'config', 'configuration']
        lat1 = 0.
        lat1_parms = ['lat1', 'toplat', 'northlat', 'north']
        lat2 = 0.
        lat2_parms = ['lat2', 'bottomlat', 'botlat', 'southlat', 'south']
        lon1 = 0.
        lon1_parms = ['lon1', 'leftlon', 'westlon', 'west']
        lon2 = 0.
        lon2_parms = ['lon2', 'rightlon', 'eastlon', 'east']
        grd1 = 0.
        grd1_parms = ['grd1', 'grid1', 'latgrid']
        grd2 = 0.
        grd2_parms = ['grd2', 'grid1', 'longrid']
        year = ''
        year_parms = ['date', 'year']
        tgt_dir = os.getcwd()
        tgt_parms = ['tgt_dir', 'dir', 'folder', 'target']
        errors = []
        for i in range(1, len(sys.argv)):
            argv = sys.argv[i].split('=')
            if len(argv) > 1:
                if argv[0] in ini_parms:
                    ini_file = argv[1]
                elif argv[0] in lat1_parms:
                    try:
                        lat1 = float(argv[1])
                        if lat1 < -85. or lat1 > 85.:
                            errors.append(sys.argv[i])
                    except:
                        errors.append(sys.argv[i])
                elif argv[0] in lat2_parms:
                    try:
                        lat2 = float(argv[1])
                        if lat2 < -85. or lat2 > 85.:
                            errors.append(sys.argv[i])
                    except:
                        errors.append(sys.argv[i])
                elif argv[0] in lon1_parms:
                    try:
                        lon1 = float(argv[1])
                        if lon1 < -180. or lon1 > 180.:
                            errors.append(sys.argv[i])
                    except:
                        errors.append(sys.argv[i])
                elif argv[0] in lon2_parms:
                    try:
                        lon2 = float(argv[1])
                        if lon2 < -180. or lon2 > 180.:
                            errors.append(sys.argv[i])
                    except:
                        errors.append(sys.argv[i])
                elif argv[0] in grd1_parms:
                    try:
                        grd1 = float(argv[1])
                        if grd1 < 0.25 or grd1 > 2.:
                            errors.append(sys.argv[i])
                    except:
                        errors.append(sys.argv[i])
                elif argv[0] in grd2_parms:
                    try:
                        grd2 = float(argv[1])
                        if grd2 < 0.25 or grd2 > 2.:
                            errors.append(sys.argv[i])
                    except:
                        errors.append(sys.argv[i])
                elif argv[0] in year_parms:
                    year = argv[1]
                elif argv[0] in tgt_parms:
                    tgt_dir = argv[1]
                    if not os.path.exists(tgt_dir):
                        errors.append(sys.argv[i])
                elif argv[0] == 'check':
                    if argv[1].lower() in ['true', 'yes', 'on']:
                        check = True
                    elif argv[1].lower() in ['false', 'no', 'off']:
                        check = False
                    else:
                        errors.append(sys.argv[i])
            else:
                if sys.argv[i][-4:] == '.ini':
                    ini_file = sys.argv[i]
                elif sys.argv[i] == 'check':
                    check = True
                else:
                    errors.append(sys.argv[i])
        if len(errors) > 0:
            print('Errors in parameters:', errors)
            sys.exit(4)
        if check:
            check = checkFiles(tgt_dir, ini_file=ini_file)
            if len(check) > 1: # ok
                lat1 = check[1]
                lat2 = check[2]
                lon1 = check[3]
                lon2 = check[4]
                grd1 = check[5]
                grd2 = check[6]
                year = check[7]
            else:
                print(check[0])
                sys.exit(4)
        wgot = retrieve_era5(ini_file, lat1, lat2, lon1, lon2, grd1, grd2, year, tgt_dir)
        sys.exit()
    else:
        app = QtWidgets.QApplication(sys.argv)
        ex = getERA5(ini_file=ini_file)
        app.exec_()
        app.deleteLater()
        sys.exit()
