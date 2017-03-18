#!/usr/bin/python
#
#  Copyright (C) 2017 Sustainable Energy Now Inc., Angus King
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
import os
import subprocess
import sys
if sys.platform == 'win32' or sys.platform == 'cygwin':
    from netCDF4 import Dataset
else:
    from Scientific.IO.NetCDF import *
import ConfigParser   # decode .ini file
from PyQt4 import QtCore, QtGui

from credits import fileVersion
import displayobject
import worldwindow

def spawn(who, cwd, log):
    stdoutf = cwd + '/' + log
    stdout = open(stdoutf, 'wb')
    who = who.split(' ')
    for i in range(len(who)):
        who[i] = who[i].replace('~', os.path.expanduser('~'))
    try:
        if type(who) is list:
            pid = subprocess.Popen(who, cwd=cwd, stderr=subprocess.STDOUT, stdout=stdout).pid
        else:
            pid = subprocess.Popen([who], cwd=cwd, stderr=subprocess.STDOUT, stdout=stdout).pid
    except:
        pass
    return


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
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                cdf_file = Dataset(inp_file, 'r')
            else:
                cdf_file = NetCDFFile(inp_file, 'r')
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
        keys = cdf_file.dimensions.keys()
        values = cdf_file.dimensions.values()
        if type(cdf_file.dimensions) is dict:
            for i in range(len(keys)):
                self.dimensions += keys[i] + ': ' + str(values[i]) + ', '
        else:
            for i in range(len(keys)):
                bits = str(values[i]).strip().split(' ')
                self.dimensions += keys[i] + ': ' + bits[-1] + ', '
        self.variables = ''
        for key in iter(sorted(cdf_file.variables.iterkeys())):
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


class getMERRA2(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def get_config(self):
        self.config = ConfigParser.RawConfigParser()
        self.config_file = 'getfiles.ini'
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

    def __init__(self, help='help.html', parent=None):
        super(getMERRA2, self).__init__(parent)
        self.get_config()
        self.help = help
        self.ignore = False
        self.worldwindow = None
        self.once = True
        self.northSpin = QtGui.QDoubleSpinBox()
        self.northSpin.setDecimals(3)
        self.northSpin.setSingleStep(.5)
        self.northSpin.setRange(-85.06, 85.06)
        self.westSpin = QtGui.QDoubleSpinBox()
        self.westSpin.setDecimals(3)
        self.westSpin.setSingleStep(.5)
        self.westSpin.setRange(-180, 180)
        self.southSpin = QtGui.QDoubleSpinBox()
        self.southSpin.setDecimals(3)
        self.southSpin.setSingleStep(.5)
        self.southSpin.setRange(-85.06, 85.06)
        self.eastSpin = QtGui.QDoubleSpinBox()
        self.eastSpin.setDecimals(3)
        self.eastSpin.setSingleStep(.5)
        self.eastSpin.setRange(-180, 180)
        if len(sys.argv) > 1:
            his_config_file = sys.argv[1]
            his_config = ConfigParser.RawConfigParser()
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
        self.grid = QtGui.QGridLayout()
        self.grid.addWidget(QtGui.QLabel('Area of Interest:'), 0, 0)
        area = QtGui.QPushButton('Choose area via Map', self)
        self.grid.addWidget(area, 0, 1, 1, 2)
        area.clicked.connect(self.areaClicked)
        self.grid.addWidget(QtGui.QLabel('Upper left:'), 1, 0)
        self.grid.addWidget(QtGui.QLabel('  North'), 2, 0)
        self.grid.addWidget(self.northSpin, 2, 1)
        self.grid.addWidget(QtGui.QLabel('  West'), 3, 0)
        self.grid.addWidget(self.westSpin, 3, 1)
        self.grid.addWidget(QtGui.QLabel('Lower right:'), 1, 2)
        self.grid.addWidget(QtGui.QLabel('  South'), 2, 2)
        self.grid.addWidget(self.southSpin, 2, 3)
        self.grid.addWidget(QtGui.QLabel('  East'), 3, 2)
        self.grid.addWidget(self.eastSpin, 3, 3)
        self.grid.addWidget(QtGui.QLabel('Approx. area:'), 4, 0)
        self.approx_area = QtGui.QLabel('')
        self.grid.addWidget(self.approx_area, 4, 1)
        self.grid.addWidget(QtGui.QLabel('MERRA dimensions:'), 4, 2)
        self.merra_cells = QtGui.QLabel('')
        self.grid.addWidget(self.merra_cells, 4, 3)
        self.grid.addWidget(QtGui.QLabel('Start date:'), 5, 0)
        self.strt_date = QtGui.QDateEdit(self)
        self.strt_date.setDate(QtCore.QDate.currentDate().addMonths(-1))
        self.strt_date.setCalendarPopup(True)
        self.strt_date.setMinimumDate(QtCore.QDate(1980, 1, 1))
        self.strt_date.setMaximumDate(QtCore.QDate.currentDate().addMonths(-1))
        self.grid.addWidget(self.strt_date, 5, 1)
        self.grid.addWidget(QtGui.QLabel('End date:'), 6, 0)
        self.end_date = QtGui.QDateEdit(self)
        self.end_date.setDate(QtCore.QDate.currentDate().addMonths(-1))
        self.end_date.setCalendarPopup(True)
        self.end_date.setMinimumDate(QtCore.QDate(1980,1,1))
        self.end_date.setMaximumDate(QtCore.QDate.currentDate().addMonths(-1))
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
            self.dirs[i].setFrameStyle(6)
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
        self.setWindowTitle('SIREN - getmerra2 - Get MERRA-2 data')
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
        if self.sender() == self.southSpin or self.sender() == self.eastSpin:
            if self.southSpin.value() > self.northSpin.value():
                y = self.northSpin.value()
                self.northSpin.setValue(self.southSpin.value())
                self.southSpin.setValue(y)
            if self.eastSpin.value() < self.westSpin.value():
                x = self.westSpin.value()
                self.westSpin.setValue(self.eastSpin.value())
                self.eastSpin.setValue(x)
        if self.worldwindow is None:
            return
        approx_area = self.worldwindow.view.drawRect([QtCore.QPointF(self.westSpin.value(), self.northSpin.value()),
                                   QtCore.QPointF(self.eastSpin.value(), self.southSpin.value())])
        self.approx_area.setText(approx_area)
        merra_dims = worldwindow.merra_cells(self.northSpin.value(), self.westSpin.value(),
                     self.southSpin.value(), self.eastSpin.value())
        self.merra_cells.setText(merra_dims)
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
        self.chk_folders = []
        self.chk_src_key = []
        self.chk_src_dte = []
        top = str(self.dirs[ndx].text())
        get_range(str(self.dirs[ndx].text()))
        first_dte = 0
        not_contiguous = False
        log1 = None
        log2 = None
        for i in range(len(self.chk_folders)):
            for j in range(len(self.chk_src_key[i])):
                if len(self.chk_src_key[i][j]) > 0:
                    break
            else:
                continue
            for j in range(len(self.chk_src_key[i])):
                dtes = sorted(self.chk_src_dte[i][j])
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
                l = self.chk_src_key[i][j].index(self.chk_collection)
                file1 = self.chk_src_key[i][j][: l + len(self.chk_collection) + 1] + dtes[0] + \
                        self.chk_src_key[i][j][l + len(self.chk_collection) + 1:]
                log1 = fileInfo(self.chk_folders[i] + '/' + file1)
                if not log1.ok:
                    self.log.setText(log1.log)
                    return
                file2 = self.chk_src_key[i][j][: l + len(self.chk_collection) + 1] + dtes[-1] + \
                        self.chk_src_key[i][j][l + len(self.chk_collection) + 1:]
                log2 = fileInfo(self.chk_folders[i] + '/' + file2)
                if not log2.ok:
                    self.log.setText(log2.log)
                    return
                dte_msg = '.\nCurrent day range %s to %s' % (str(first_dte), str(dtes[-1]))
                if len(dtes) != (dte2-dte1).days + 1:
                    not_contiguous = True
                    print 'getmerra2: File template ' + self.chk_src_key[i][j]
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
                    for key, value in iter(sorted(years.iteritems())):
                        print 'getmerra2: Days per month', key, value[1:]
        if log1 is None or log2 is None:
            self.log.setText('Check ' + chk_key.title() + ' incomplete.')
            return
        if log1.latitudes == log2.latitudes:
            self.ignore = True
            self.southSpin.setValue(log1.latitudes[0])
            self.ignore = False
            self.northSpin.setValue(log1.latitudes[-1])
        if log1.longitudes == log2.longitudes:
            self.ignore = True
            self.westSpin.setValue(log1.longitudes[0])
            self.ignore = False
            self.eastSpin.setValue(log1.longitudes[-1])
            dte = dte2 + datetime.timedelta(days=1)
            self.strt_date.setDate(dte)
        if not_contiguous:
            dte_msg += ' but days not contiguous'
        self.log.setText('Boundaries and start date set for ' + chk_key.title() + dte_msg)
        del self.chk_collection, self.chk_folders, self.chk_src_key, self.chk_src_dte

    def getClicked(self):
        if self.southSpin.value() == self.northSpin.value() or self.eastSpin.value() == self.westSpin.value():
            self.log.setText('Area too small')
            return
        if self.sender().text() == 'Get Solar':
            me = 'solar'
            ignor = 'wind'
        elif self.sender().text() == 'Get Wind':
            me = 'wind'
            ignor = 'solar'
        variables = []
        try:
            variables = self.config.items('getmerra2')
            wget = self.config.get('getmerra2', 'wget')
        except:
            self.log.setText('Error accessing getfiles.ini variables')
            return
        if self.once: # check .netrc exists
            self.once = False
            ok = False
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                netrc = '~\\.netrc'.replace('~', os.environ['HOME'])
            else:
                netrc = '~/.netrc'.replace('~', os.path.expanduser('~'))
            if os.path.exists(netrc):
                ok = True
            if not ok:
                reply = QtGui.QMessageBox.question(self, 'SIREN - getmerra2 - No .netrc file',
                        '.netrc file missing.\n(' + netrc + ')\nOK to continue?',
                        QtGui.QMessageBox.Yes | QtGui.QMessageBox.No, QtGui.QMessageBox.No)
                if reply == QtGui.QMessageBox.Yes:
                    pass
                else:
                    return
        working_vars = []
        for prop, value in variables:
            valu = value
            if prop != 'url_prefix':
                valu = valu.replace('/', '%2F')
                valu = valu.replace(',', '%2C')
            if prop[:len(ignor)] == ignor:
                continue
            elif prop[:len(me)] == me:
                working_vars.append(('$' + prop[len(me) + 1:] + '$', valu))
            else:
                working_vars.append(('$' + prop + '$', valu))
        working_vars.append(('$lat1$', str(self.northSpin.value())))
        working_vars.append(('$lon1$', str(self.westSpin.value())))
        working_vars.append(('$lat2$', str(self.southSpin.value())))
        working_vars.append(('$lon2$', str(self.eastSpin.value())))
        wget_base = wget[:]
        while '$' in wget:
            for key, value in working_vars:
                wget = wget.replace(key, value)
            if wget == wget_base:
                break
            wget_base = wget
        a_date = QtCore.QDate()
        y = int(self.strt_date.date().year())
        m = int(self.strt_date.date().month())
        d = int(self.strt_date.date().day())
        a_date.setDate(y, m, d)
        i = self.dir_labels.index(me.title())
        wget_file = 'wget_' + me + '_' + \
                    str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), \
                                                  'yyyy-MM-dd_hhmm')) + '.txt'
        wf = open(str(self.dirs[i].text()) + '/' + wget_file, 'w')
        while a_date.__le__(self.end_date.date()):
            date_vars = []
            date_vars.append(('$year$','{0:04d}'.format(a_date.year())))
            date_vars.append(('$month$', '{0:02d}'.format(a_date.month())))
            date_vars.append(('$day$', '{0:02d}'.format(a_date.day())))
            wget = wget_base[:]
            for key, value in date_vars:
                wget = wget.replace(key, value)
            wf.write(wget + '\n')
            a_date = QtCore.QDate(a_date.addDays(1))
        wf.close()
        curdir = os.getcwd()
        os.chdir(str(self.dirs[i].text()))
        wget_cmd = self.config.get('getmerra2', 'wget_cmd') + ' ' + wget_file
        os.chdir(curdir)
        log = wget_file[:-3] + 'log'
        cwd = str(self.dirs[i].text())
        spawn(wget_cmd, cwd, log)
        self.log.setText('wget launched (logging to: ' + log +')')
        bat = wget_file[:-3] + 'bat'
        bf = open(str(self.dirs[i].text()) + '/' + bat, 'w')
        who = wget_cmd.split(' ')
        for i in range(len(who)):
            who[i] = who[i].replace('~', os.path.expanduser('~'))
        bat_cmd = ' '.join(who)
        bf.write(bat_cmd)
        bf.close()


if '__main__' == __name__:
    app = QtGui.QApplication(sys.argv)
    ex = getMERRA2()
    app.exec_()
    app.deleteLater()
    sys.exit()
