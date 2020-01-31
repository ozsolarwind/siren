#!/usr/bin/python3
#
#  Copyright (C) 2017-2020 Sustainable Energy Now Inc., Angus King
#
#  getmerra2.py - This file is part of SIREN.
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

import datetime
import gzip
import os
import subprocess
import sys
from netCDF4 import Dataset
import configparser   # decode .ini file
from PyQt4 import QtCore, QtGui

import displayobject
import worldwindow
from credits import fileVersion

def spawn(who, cwd, log):
    stdoutf = cwd + '/' + log
    stdout = open(stdoutf, 'wb')
    who = who.split(' ')
  #  for i in range(len(who)):
   #     who[i] = who[i].replace('~', os.path.expanduser('~'))
    try:
        if type(who) is list:
            pid = subprocess.Popen(who, cwd=cwd, stderr=subprocess.STDOUT, stdout=stdout).pid
        else:
            pid = subprocess.Popen([who], cwd=cwd, stderr=subprocess.STDOUT, stdout=stdout).pid
    except:
        pass
    return

def checkFiles(chk_key, tgt_dir, ini_file=None, collection=None):
    def get_range(top):
        chk_folders.append(top)
        ndx = len(chk_folders) - 1
        chk_src_key.append([])
        chk_src_dte.append([])
        fils = sorted(os.listdir(top))
        for fil in fils:
            if os.path.isdir(top + '/' + fil):
                if fil.isdigit() and len(fil) == 4:
                    get_range(top + '/' + fil)
            elif fil.find('MERRA') >= 0:
                if fil[-6:] == '.gz.nc': # ignore unzipped
                    continue
                ndy = 0 #value[1]
                j = fil.find(chk_collection)
                l = len(chk_collection)
                if j > 0:
                    fil_key = fil[:j + l + 1] + fil[j + l + 9:]
                    if fil_key not in chk_src_key[-1]:
                        chk_src_key[-1].append(fil_key)
                        chk_src_dte[-1].append([])
                        chk_src_dte[-1][-1].append(fil[j + l + 1: j + l + 9])
                    else:
                        k = chk_src_key[-1].index(fil_key)
                        chk_src_dte[-1][k].append(fil[j + l + 1: j + l + 9])
        del fils

    if collection is None:
        config = configparser.RawConfigParser()
        config_file = ini_file
        config.read(config_file)
        chk_collection = config.get('getmerra2', chk_key + '_collection')
    else:
        chk_collection = collection
    msg_text = ''
    chk_folders = []
    chk_src_key = []
    chk_src_dte = []
    top = tgt_dir
    get_range(top)
    first_dte = 0
    not_contiguous = False
    log1 = None
    log2 = None
    for i in range(len(chk_folders)):
        for j in range(len(chk_src_key[i])):
            if len(chk_src_key[i][j]) > 0:
                break
        else:
            continue
        for j in range(len(chk_src_key[i])):
            dtes = sorted(chk_src_dte[i][j])
            try:
                dte1 = datetime.datetime.strptime(dtes[0], '%Y%m%d')
                if first_dte == 0:
                    first_dte = dtes[0]
            except:
                continue
            try:
                dte2 = datetime.datetime.strptime(dtes[-1], '%Y%m%d')
            except:
                continue
            l = chk_src_key[i][j].index(chk_collection)
            file1 = chk_src_key[i][j][: l + len(chk_collection) + 1] + dtes[0] + \
                    chk_src_key[i][j][l + len(chk_collection) + 1:]
            log1 = fileInfo(chk_folders[i] + '/' + file1)
            if not log1.ok:
                msg_text = log1.log
                return [log1.log]
            file2 = chk_src_key[i][j][: l + len(chk_collection) + 1] + dtes[-1] + \
                    chk_src_key[i][j][l + len(chk_collection) + 1:]
            log2 = fileInfo(chk_folders[i] + '/' + file2)
            if not log2.ok:
                msg_text = log2.log
                return [log2.log]
            dte_msg = '.\nCurrent day range %s to %s' % (str(first_dte), str(dtes[-1]))
            if len(dtes) != (dte2-dte1).days + 1:
                print('(128)', len(dtes), (dte2-dte1).days + 1, dte2, dte1)
                not_contiguous = True
                print('getmerra2: File template ' + chk_src_key[i][j])
                years = {}
                for dte in dtes:
                    y = dte[:4]
                    m = int(dte[4:6])
                    try:
                        years[y][m] = years[y][m] + 1
                    except:
                        years[y] = []
                        for l in range(13):
                            years[y].append(0)
                        years[y][m] = years[y][m] + 1
                for key, value in iter(sorted(years.items())):
                    print('getmerra2: Days per month', key, value[1:])
    if log1 is None or log2 is None:
        msg_text = 'Check ' + chk_key.title() + ' incomplete.'
        return [msg_text]
    lats = latn = lonw = lone = 0.
    if log1.latitudes == log2.latitudes:
        lats = log1.latitudes[0]
        latn = log1.latitudes[-1]
        lat_rnge = latn - lats
        if lat_rnge < 1:
            if len(log1.latitudes) == 2:
                lats = lats - .005
            latn = lats + 1
    if log1.longitudes == log2.longitudes:
        lonw = log1.longitudes[0]
        lone = log1.longitudes[-1]
        lon_rnge = lone - lonw
        if lon_rnge < 1.25:
            if len(log1.longitudes) == 2:
                lonw = lonw - .005
            lone = lonw + 1.25
    dte = dte2 + datetime.timedelta(days=1)
    if not_contiguous:
        dte_msg += ' but days not contiguous'
    msg_text = 'Boundaries and start date set for ' + chk_key.title() + dte_msg
    return [msg_text, dte, dte2, latn, lats, lonw, lone]

def invokeWget(ini_file, coll, date1, date2, lat1, lat2, lon1, lon2, tgt_dir, spawn_wget):
    config = configparser.RawConfigParser()
    config.read(ini_file)
    if coll == 'solar':
        ignor = 'wind'
    else:
        ignor = 'solar'
    variables = []
    try:
        variables = config.items('getmerra2')
        wget = config.get('getmerra2', 'wget')
    except:
        return 'Error accessing', ini_file, 'variables'
    working_vars = []
    for prop, value in variables:
        valu = value
        if prop != 'url_prefix':
            valu = valu.replace('/', '%2F')
            valu = valu.replace(',', '%2C')
        if prop[:len(ignor)] == ignor:
            continue
        elif prop[:len(coll)] == coll:
            working_vars.append(('$' + prop[len(coll) + 1:] + '$', valu))
        else:
            working_vars.append(('$' + prop + '$', valu))
    working_vars.append(('$lat1$', str(lat1)))
    working_vars.append(('$lon1$', str(lon1)))
    working_vars.append(('$lat2$', str(lat2)))
    working_vars.append(('$lon2$', str(lon2)))
    wget_base = wget[:]
    while '$' in wget:
        for key, value in working_vars:
            wget = wget.replace(key, value)
        if wget == wget_base:
            break
        wget_base = wget
    a_date = datetime.datetime.now().strftime('%Y-%m-%d_%H%M')
    wget_file = 'wget_' + coll + '_' + a_date + '.txt'
    wf = open(tgt_dir + '/' + wget_file, 'w')
    days = (date2 - date1).days + 1
    while date1 <= date2:
        date_vars = []
        date_vars.append(('$year$','{0:04d}'.format(date1.year)))
        date_vars.append(('$month$', '{0:02d}'.format(date1.month)))
        date_vars.append(('$day$', '{0:02d}'.format(date1.day)))
        wget = wget_base[:]
        for key, value in date_vars:
            wget = wget.replace(key, value)
        wf.write(wget + '\n')
        date1 = date1 + datetime.timedelta(days=1)
    wf.close()
    curdir = os.getcwd()
    os.chdir(tgt_dir)
    wget_cmd = config.get('getmerra2', 'wget_cmd')
    os.chdir(curdir)
    log_file = wget_file[:-3] + 'log'
    cwd = tgt_dir
    bat_file = wget_file[:-3] + 'bat'
    bf = open(tgt_dir + '/' + bat_file, 'w')
    who = wget_cmd.split(' ')
    for i in range(len(who)):
        if who[i] == '-i': # input file
            who[i] = who[i] + ' ' + wget_file
        elif who[i] == '-o' or who[i] == '-a':
            who[i] = who[i] + ' ' + log_file
        who[i] = who[i].replace('~', os.path.expanduser('~'))
    bat_cmd = ' '.join(who)
    bf.write(bat_cmd)
    bf.close()
    if sys.platform == 'linux' or sys.platform == 'linux2':
        os.chmod(tgt_dir + '/' + bat_file, 0o777)
    if spawn_wget:
        spawn(bat_cmd, cwd, log_file)
        return 'wget being launched (logging to: ' + log_file +')'
    return bat_file + ', ' + wget_file + ' (' + str(days) + ' days) created.'


class fileInfo:
    def __init__(self, inp_file):
        self.ok = False
        self.log = ''
        if inp_file[-3:] == '.gz':
            out_file = inp_file + '.nc'
            if os.path.exists(out_file):
                pass
            else:
                fin = gzip.open(inp_file, 'rb')
                file_content = fin.read()
                fin.close()
                fou = open(out_file, 'wb')
                fou.write(file_content)
                fou.close()
            inp_file = out_file
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
            self.format = str(cdf_file.Format)
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
                bits = str(values[i]).strip().split(' ')
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


class ClickableQLabel(QtGui.QLabel):
    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtGui.QApplication.widgetAt(event.globalPos()).setFocus()
        self.emit(QtCore.SIGNAL('clicked()'))


class subwindow(QtGui.QDialog):
    def __init__(self, parent = None):
        super(subwindow, self).__init__(parent)

    def closeEvent(self, event):
        event.accept()

    def exit(self):
        self.close()


class getMERRA2(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

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
        self.wait_days = 42
        try:
            self.wait_days = int(self.config.get('getmerra2', 'wait_days'))
        except:
            pass

    def __init__(self, help='help.html', ini_file='getfiles.ini', parent=None):
        super(getMERRA2, self).__init__(parent)
        self.help = help
        self.ini_file = ini_file
        self.get_config()
        self.ignore = False
        self.worldwindow = None
        ok = False
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            try:
                netrc = '~\\.netrc'.replace('~', os.environ['HOME'])
            except:
                netrc = ''
        else:
            netrc = '~/.netrc'.replace('~', os.path.expanduser('~'))
        if netrc != '' and os.path.exists(netrc):
            ok = True
        if not ok:
            netrcmsg = '.netrc file missing.\n(' + netrc + ')\nDo you want create one?'
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                netrcmsg += '\nYou will need to reinvoke getMERRA2.'
            reply = QtGui.QMessageBox.question(self, 'SIREN - getmerra2 - No .netrc file',
                    netrcmsg,
                    QtGui.QMessageBox.Yes | QtGui.QMessageBox.No, QtGui.QMessageBox.No)
            if reply == QtGui.QMessageBox.Yes:
                self.create_netrc()
        self.northSpin = QtGui.QDoubleSpinBox()
        self.northSpin.setDecimals(3)
        self.northSpin.setSingleStep(.5)
        self.northSpin.setRange(-85.06, 85.06)
        self.northSpin.setObjectName('north')
        self.westSpin = QtGui.QDoubleSpinBox()
        self.westSpin.setDecimals(3)
        self.westSpin.setSingleStep(.5)
        self.westSpin.setRange(-180, 180)
        self.westSpin.setObjectName('west')
        self.southSpin = QtGui.QDoubleSpinBox()
        self.southSpin.setDecimals(3)
        self.southSpin.setSingleStep(.5)
        self.southSpin.setRange(-85.06, 85.06)
        self.southSpin.setObjectName('south')
        self.eastSpin = QtGui.QDoubleSpinBox()
        self.eastSpin.setDecimals(3)
        self.eastSpin.setSingleStep(.5)
        self.eastSpin.setRange(-180, 180)
        self.eastSpin.setObjectName('east')
        self.latSpin = QtGui.QDoubleSpinBox()
        self.latSpin.setDecimals(3)
        self.latSpin.setSingleStep(.5)
        self.latSpin.setRange(-84.56, 84.56)
        self.latSpin.setObjectName('lat')
        self.lonSpin = QtGui.QDoubleSpinBox()
        self.lonSpin.setDecimals(3)
        self.lonSpin.setSingleStep(.5)
        self.lonSpin.setRange(-179.687, 179.687)
        self.lonSpin.setObjectName('lon')
        self.latwSpin = QtGui.QDoubleSpinBox()
        self.latwSpin.setDecimals(3)
        self.latwSpin.setSingleStep(.5)
        self.latwSpin.setRange(0, 170.12)
        self.latwSpin.setObjectName('latw')
        self.lonwSpin = QtGui.QDoubleSpinBox()
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
        self.grid = QtGui.QGridLayout()
        self.grid.addWidget(QtGui.QLabel('Area of Interest:'), 0, 0)
        area = QtGui.QPushButton('Choose area via Map', self)
        self.grid.addWidget(area, 0, 1, 1, 2)
        area.clicked.connect(self.areaClicked)
        self.grid.addWidget(QtGui.QLabel('Upper left:'), 1, 0, 1, 2)
   #     self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignCenter)
        self.grid.addWidget(QtGui.QLabel('North'), 2, 0)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.northSpin, 2, 1)
        self.grid.addWidget(QtGui.QLabel('West'), 3, 0)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.westSpin, 3, 1)
        self.grid.addWidget(QtGui.QLabel('Lower right:'), 1, 2)
    #    self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignCenter)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(QtGui.QLabel('South'), 2, 2)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.southSpin, 2, 3)
        self.grid.addWidget(QtGui.QLabel('East'), 3, 2)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.eastSpin, 3, 3)
        self.grid.addWidget(QtGui.QLabel('Centre:'), 1, 4)
   #     self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignCenter)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(QtGui.QLabel('Lat.'), 2, 4)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.latSpin, 2, 5)
        self.grid.addWidget(QtGui.QLabel('Lon.'), 3, 4)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.lonSpin, 3, 5)
        self.grid.addWidget(QtGui.QLabel('Degrees'), 1, 6)
   #     self.grid.addWidget(QtGui.QLabel('  South'), 2, 2)
        self.grid.addWidget(self.latwSpin, 2, 6)
  #      self.grid.addWidget(QtGui.QLabel('  East'), 3, 2)
        self.grid.addWidget(self.lonwSpin, 3, 6)
        self.grid.addWidget(QtGui.QLabel('Approx. area:'), 4, 0)
        self.approx_area = QtGui.QLabel('')
        self.grid.addWidget(self.approx_area, 4, 1)
        self.grid.addWidget(QtGui.QLabel('MERRA dimensions:'), 4, 2)
        self.merra_cells = QtGui.QLabel('')
        self.grid.addWidget(self.merra_cells, 4, 3)
        self.grid.addWidget(QtGui.QLabel('Start date:'), 5, 0)
        self.strt_date = QtGui.QDateEdit(self)
        self.strt_date.setDate(QtCore.QDate.currentDate().addDays(-self.wait_days))
        self.strt_date.setCalendarPopup(True)
        self.strt_date.setMinimumDate(QtCore.QDate(1980, 1, 1))
        self.strt_date.setMaximumDate(QtCore.QDate.currentDate().addDays(-self.wait_days))
        self.grid.addWidget(self.strt_date, 5, 1)
        self.grid.addWidget(QtGui.QLabel('End date:'), 6, 0)
        self.end_date = QtGui.QDateEdit(self)
        self.end_date.setDate(QtCore.QDate.currentDate().addDays(-self.wait_days))
        self.end_date.setCalendarPopup(True)
        self.end_date.setMinimumDate(QtCore.QDate(1980,1,1))
        self.end_date.setMaximumDate(QtCore.QDate.currentDate())
        self.grid.addWidget(self.end_date, 6, 1)
        self.grid.addWidget(QtGui.QLabel('Copy folder down:'), 7, 0)
        self.checkbox = QtGui.QCheckBox()
        self.checkbox.setCheckState(QtCore.Qt.Checked)
        self.grid.addWidget(self.checkbox, 7, 1)
        self.grid.addWidget(QtGui.QLabel('If checked will copy solar folder changes down'), 7, 2, 1, 3)
        cur_dir = os.getcwd()
        self.dir_labels = ['Solar', 'Wind']
        datasets = []
        self.collections = []
        for typ in self.dir_labels:
            datasets.append(self.config.get('getmerra2', typ.lower() + '_collection') + ' (Parameters: ' + \
                            self.config.get('getmerra2', typ.lower() + '_variables') + ')')
            self.collections.append(self.config.get('getmerra2', typ.lower() + '_collection'))
        self.dirs = [None, None, None]
        for i in range(2):
            self.grid.addWidget(QtGui.QLabel(self.dir_labels[i] + ' files:'), 8 + i * 2, 0)
            self.grid.addWidget(QtGui.QLabel(datasets[i]), 8 + i * 2, 1, 1, 4)
            self.grid.addWidget(QtGui.QLabel('    Target Folder:'), 9 + i * 2, 0)
            self.dirs[i] = ClickableQLabel()
            self.dirs[i].setText(cur_dir)
            self.dirs[i].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
            self.connect(self.dirs[i], QtCore.SIGNAL('clicked()'), self.dirChanged)
            self.grid.addWidget(self.dirs[i], 9 + i * 2, 1, 1, 4)
        self.log = QtGui.QLabel('')
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.log.setPalette(msg_palette)
        self.grid.addWidget(self.log, 12, 1, 2, 4)
        buttongrid = QtGui.QGridLayout()
        quit = QtGui.QPushButton('Quit', self)
        buttongrid.addWidget(quit, 0, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('x'), self, self.quitClicked)
        solar = QtGui.QPushButton('Get Solar', self)
        buttongrid.addWidget(solar, 0, 1)
        solar.clicked.connect(self.getClicked)
        wind = QtGui.QPushButton('Get Wind', self)
        buttongrid.addWidget(wind, 0, 2)
        wind.clicked.connect(self.getClicked)
        check = QtGui.QPushButton('Check Solar', self)
        buttongrid.addWidget(check, 0, 3)
        check.clicked.connect(self.checkClicked)
        check = QtGui.QPushButton('Check Wind', self)
        buttongrid.addWidget(check, 0, 4)
        check.clicked.connect(self.checkClicked)
        help = QtGui.QPushButton('Help', self)
        buttongrid.addWidget(help, 0, 5)
        help.clicked.connect(self.helpClicked)
        QtGui.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        frame2 = QtGui.QFrame()
        frame2.setLayout(buttongrid)
        self.layout.addWidget(frame2)
        self.setWindowTitle('SIREN - getmerra2 (' + fileVersion() + ') - Get MERRA-2 data')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.resize(int(self.sizeHint().width()* 1.4), int(self.sizeHint().height() * 1.07))
        self.show()

    def dirChanged(self):
        for i in range(2):
            if self.dirs[i].hasFocus():
                break
        curdir = str(self.dirs[i].text())
        newdir = str(QtGui.QFileDialog.getExistingDirectory(self, 'Choose ' + self.dir_labels[i] + ' Folder',
                 curdir, QtGui.QFileDialog.ShowDirsOnly))
        if newdir != '':
            self.dirs[i].setText(newdir)
            if self.checkbox.isChecked():
                if i == 0:
                    self.dirs[1].setText(newdir)

    def zoomChanged(self, val):
        self.zoomScale.setText('(' + scale[int(val)] + ')')
        self.zoomScale.adjustSize()

    def fileChanged(self):
        self.filename.setText(QtGui.QFileDialog.getSaveFileName(self, 'Save Image File',
                              self.filename.text(),
                              'Images (*.jpeg *.jpg *.png)'))
        if self.filename.text() != '':
            i = str(self.filename.text()).rfind('.')
            if i < 0:
                self.filename.setText(self.filename.text() + '.png')

    def helpClicked(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.help,
                 title='Help for getting MERRA data (' + fileVersion() + ')', section='getmerra2')
        dialog.exec_()

    def exit(self):
        self.close()

    def closeEvent(self, event):
        try:
            self.mySubwindow.close()
        except:
            pass
        event.accept()

    @QtCore.pyqtSlot()
    def maparea(self, rectangle, approx_area=None):
        if type(rectangle) is str:
            if rectangle == 'goodbye':
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
                     self.southSpin.value(), self.eastSpin.value())
        self.merra_cells.setText(merra_dims)

    def areaClicked(self):
        if self.worldwindow is None:
            scene = worldwindow.WorldScene()
            self.worldwindow = worldwindow.WorldWindow(self, scene)
            self.connect(self.worldwindow.view, QtCore.SIGNAL('tellarea'), self.maparea)
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
                self.latwSpin.setValue(1.)
            self.northSpin.setValue(self.latSpin.value() + self.latwSpin.value() / 2)
            self.southSpin.setValue(self.latSpin.value() - self.latwSpin.value() / 2)
            self.ignore = False
        elif self.sender().objectName() == 'lon':
            self.ignore = True
            if self.lonwSpin.value() == 0:
                self.lonwSpin.setValue(1.5)
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
                     self.southSpin.value(), self.eastSpin.value())
        self.merra_cells.setText(merra_dims)
        if self.worldwindow is None:
            return
        approx_area = self.worldwindow.view.drawRect([QtCore.QPointF(self.westSpin.value(), self.northSpin.value()),
                                   QtCore.QPointF(self.eastSpin.value(), self.southSpin.value())])
        self.approx_area.setText(approx_area)
        if event != 'init':
            self.worldwindow.view.emit(QtCore.SIGNAL('statusmsg'), approx_area + ' ' + merra_dims)

    def quitClicked(self):
        self.close()

    def checkClicked(self):
        def get_range(top):
            self.chk_folders.append(top)
            ndx = len(self.chk_folders) - 1
            self.chk_src_key.append([])
            self.chk_src_dte.append([])
            fils = sorted(os.listdir(top))
            for fil in fils:
                if os.path.isdir(top + '/' + fil):
                    if fil.isdigit() and len(fil) == 4:
                        get_range(top + '/' + fil)
                elif fil.find('MERRA') >= 0:
                    if fil[-6:] == '.gz.nc': # ignore unzipped
                        continue
                    ndy = 0 #value[1]
                    j = fil.find(self.chk_collection)
                    l = len(self.chk_collection)
                    if j > 0:
                        fil_key = fil[:j + l + 1] + fil[j + l + 9:]
                        if fil_key not in self.chk_src_key[-1]:
                            self.chk_src_key[-1].append(fil_key)
                            self.chk_src_dte[-1].append([])
                            self.chk_src_dte[-1][-1].append(fil[j + l + 1: j + l + 9])
                        else:
                            k = self.chk_src_key[-1].index(fil_key)
                            self.chk_src_dte[-1][k].append(fil[j + l + 1: j + l + 9])
            del fils

        if self.sender().text() == 'Check Solar':
            chk_key = 'solar'
            ndx = 0
        elif self.sender().text() == 'Check Wind':
            chk_key = 'wind'
            ndx = 1
        self.log.setText('')
        self.chk_collection = self.collections[ndx]

        check = checkFiles(chk_key, str(self.dirs[ndx].text()), ini_file=self.ini_file, collection=self.chk_collection)
        if len(check) > 1: # ok
            self.ignore = True
            self.southSpin.setValue(check[4])
            self.northSpin.setValue(check[3])
            self.latwSpin.setValue(self.northSpin.value() - self.southSpin.value())
            self.latSpin.setValue(self.northSpin.value() - (self.northSpin.value() - self.southSpin.value()) / 2.)
            self.westSpin.setValue(check[5])
            self.eastSpin.setValue(check[6])
            self.lonwSpin.setValue(self.eastSpin.value() - self.westSpin.value())
            self.lonSpin.setValue(self.eastSpin.value() - (self.eastSpin.value() - self.westSpin.value()) / 2.)
            self.strt_date.setDate(check[1])
            self.end_date.setDate(datetime.datetime.now() - datetime.timedelta(days=self.wait_days))
            if self.end_date.date() < self.strt_date.date():
                self.end_date.setDate(self.strt_date.date())
            self.ignore = False
            merra_dims = worldwindow.merra_cells(self.northSpin.value(), self.westSpin.value(),
                         self.southSpin.value(), self.eastSpin.value())
            self.merra_cells.setText(merra_dims)
        self.log.setText(check[0])
        if self.worldwindow is None:
            return
        approx_area = self.worldwindow.view.drawRect([QtCore.QPointF(self.westSpin.value(), self.northSpin.value()),
                                   QtCore.QPointF(self.eastSpin.value(), self.southSpin.value())])
        self.approx_area.setText(approx_area)
        self.worldwindow.view.emit(QtCore.SIGNAL('statusmsg'), approx_area + ' ' + merra_dims)

    def getClicked(self):
        lat_rnge = self.northSpin.value() - self.southSpin.value()
        lon_rnge = self.eastSpin.value() - self.westSpin.value()
        if lat_rnge < 1 or lon_rnge < 1.25:
            self.log.setText('Area too small. Range must be at least 1 degree Lat x 1.25 Lon')
            return
        if self.sender().text() == 'Get Solar':
            me = 'solar'
        elif self.sender().text() == 'Get Wind':
            me = 'wind'
        date1 = datetime.date(int(self.strt_date.date().year()), int(self.strt_date.date().month()),
                int(self.strt_date.date().day()))
        date2 = datetime.date(int(self.end_date.date().year()), int(self.end_date.date().month()),
                int(self.end_date.date().day()))
        i = self.dir_labels.index(me.title())
        tgt_dir = str(self.dirs[i].text())
        wgot = invokeWget(self.ini_file, me, date1, date2, self.northSpin.value(),
               self.southSpin.value(), self.westSpin.value(), self.eastSpin.value(), tgt_dir, True)
        self.log.setText(wgot)
        return

    def create_netrc(self):
        self.mySubwindow = subwindow()
        grid = QtGui.QGridLayout()
        grid.addWidget(QtGui.QLabel('Enter details for URS Registration'), 0, 0, 1, 2)
        r = 1
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            try:
                self.home_dir = os.environ['HOME']
            except:
                self.home_dir = os.getcwd()
                grid.addWidget(QtGui.QLabel('HOME directory:'), 1, 0)
                self.home = ClickableQLabel()
                self.home.setText(self.home_dir)
                self.home.setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
                self.connect(self.home, QtCore.SIGNAL('clicked()'), self.ursdirChanged)
                grid.addWidget(self.home, 1, 1, 1, 3)
                r = 2
        grid.addWidget(QtGui.QLabel('URS Userid:'), r, 0)
        self.urs_id = QtGui.QLineEdit('')
        grid.addWidget(self.urs_id, r, 1)
        grid.addWidget(QtGui.QLabel('URS Password:'), r + 1, 0)
        self.urs_pwd = QtGui.QLineEdit('')
        grid.addWidget(self.urs_pwd, r + 1, 1)
        netrc_button = QtGui.QPushButton('Create .netrc file', self.mySubwindow)
        self.connect(netrc_button, QtCore.SIGNAL('clicked()'), self.createnetrc)
        grid.addWidget(netrc_button, r + 2, 0)
        self.mySubwindow.setLayout(grid)
        self.mySubwindow.exec_()

    def ursdirChanged(self):
        curdir = str(self.home.text())
        newdir = str(QtGui.QFileDialog.getExistingDirectory(self, 'Choose .netrc Folder',
                 curdir, QtGui.QFileDialog.ShowDirsOnly))
        if newdir != '':
            self.home.setText(newdir)
            self.home_dir = newdir

    def createnetrc(self):
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            env_var = 'HOME'
            os.system('SETX {0} "{1}"'.format(env_var, self.home_dir))
            netrc = self.home_dir + '\\.netrc'
        else:
            netrc = '~/.netrc'.replace('~', os.path.expanduser('~'))
        netrc_string = 'machine urs.earthdata.nasa.gov login %s password %s' % (str(self.urs_id.text()),
                       str(self.urs_pwd.text()))
        fou = open(netrc, 'wb')
        fou.write(netrc_string)
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
        coll = 'solar'
        coll_parms = ['coll', 'collection']
        date1 = ''
        date1_parms = ['date', 'date1', 'strtdate', 'startdate', 'year']
        date2 = ''
        date2_parms = ['date2', 'enddate']
        lat1 = 0.
        lat1_parms = ['lat1', 'toplat', 'northlat', 'north']
        lat2 = 0.
        lat2_parms = ['lat2', 'bottomlat', 'botlat', 'southlat', 'south']
        lon1 = 0.
        lon1_parms = ['lon1', 'leftlon', 'westlon', 'west']
        lon2 = 0.
        lon2_parms = ['lon2', 'rightlon', 'eastlon', 'east']
        tgt_dir = os.getcwd()
        tgt_parms = ['tgt_dir', 'dir', 'folder', 'target']
        spawn_wget = False
        spawn_parms = ['get', 'wget', 'spawn']
        errors = []
        for i in range(1, len(sys.argv)):
            argv = sys.argv[i].split('=')
            if len(argv) > 1:
                if argv[0] in ini_parms:
                    ini_file = argv[1]
                elif argv[0] in coll_parms:
                    coll = argv[1]
                    if coll not in ['wind', 'solar']:
                        errors.append(sys.argv[i])
                elif argv[0] in date1_parms:
                    date1 = argv[1]
                    if argv[0] == 'year':
                        date1 = argv[1] + '-01-01'
                        try:
                            date1 = datetime.datetime.strptime(date1, '%Y-%m-%d')
                            date2 = date1.replace(date1.year + 1)
                            date1 = date1 - datetime.timedelta(days=1)
                        except:
                            errors.append(sys.argv[i])
                    else:
                        try:
                            date1 = datetime.datetime.strptime(date1, '%Y-%m-%d')
                        except:
                            errors.append(sys.argv[i])
                elif argv[0] in date2_parms:
                    date2 = argv[1]
                    try:
                        date2 = datetime.datetime.strptime(date2, '%Y-%m-%d')
                    except:
                        errors.append(sys.argv[i])
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
                elif argv[0] in tgt_parms:
                    tgt_dir = argv[1]
                    if not os.path.exists(tgt_dir):
                        errors.append(sys.argv[i])
                elif argv[0] in spawn_parms:
                    if argv[1].lower() in ['true', 'yes', 'on']:
                        spawn_wget = True
                    elif argv[1].lower() in ['false', 'no', 'off']:
                        spawn_wget = False
                    else:
                        errors.append(sys.argv[i])
                elif argv[0] == 'check':
                    if argv[1].lower() in ['true', 'yes', 'on']:
                        check = True
                    elif argv[1].lower() in ['false', 'no', 'off']:
                        check = False
                    else:
                        errors.append(sys.argv[i])
            else:
                if sys.argv[i] in spawn_parms:
                    spawn_wget = True
                elif sys.argv[i][-4:] == '.ini':
                    ini_file = sys.argv[i]
                elif sys.argv[i] == 'solar' or sys.argv[i] == 'wind':
                    coll = sys.argv[i]
                elif sys.argv[i] == 'check':
                    check = True
                else:
                    errors.append(sys.argv[i])
        if len(errors) > 0:
            print('Errors in parameters:', errors)
            sys.exit(4)
        if check:
            check = checkFiles(coll, tgt_dir, ini_file=ini_file)
            if len(check) > 1: # ok
                date1 = check[1]
                if date2 == '':
                    date2 = datetime.datetime.now() - datetime.timedelta(days=self.wait_days)
                    if date2 < date1:
                        date2 = date1
                lat1 = check[3]
                lat2 = check[4]
                lon1 = check[5]
                lon2 = check[6]
                print(check[0])
            else:
                print(check[0])
                sys.exit(4)
        if abs(lat1 - lat2) < 1 or abs(lon2 - lon1) < 1.25:
            print('Area too small. Range must be at least 1 degree Lat x 1.25 Lon')
            sys.exit(4)
        if date2 < date1:
            print('Date2 (End) less than Date1 (Start)')
            sys.exit(4)
        wgot = invokeWget(ini_file, coll, date1, date2, lat1, lat2, lon1, lon2, tgt_dir, spawn_wget)
        print(wgot)
        sys.exit()
    else:
        app = QtGui.QApplication(sys.argv)
        ex = getMERRA2(ini_file=ini_file)
        app.exec_()
        app.deleteLater()
        sys.exit()
