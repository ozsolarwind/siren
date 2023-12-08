#!/usr/bin/python3
#
#  Copyright (C) 2022 Sustainable Energy Now Inc., Angus King
#
#  dataview.py - This file is part of SIREN.
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
import sys
import time
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.Qt import QColor
#
import configparser   # decode .ini file
import displayobject
from editini import SaveIni
from getmodels import getModelFile
from senutils import ClickableQLabel, getParents, getUser, strSplit, techClean, WorkBook

def rreplace(s, old, new, occurrence):
    li = s.rsplit(old, occurrence)
    return new.join(li)


class DataView(QtWidgets.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, scene=None):
        super(DataView, self).__init__()
        self.scene = scene
        if not self.scene.show_coord:
            self.scene._coordGroup.setVisible(True)
        self.setup = [False, False]
        self.book = None
        self.history = None
        self.max_files = 10
        self.colours = {}
        opacity = 1.
        numbers = True
        multiple = False
        number_min = 0
        self.number_steps = [80, 60, 40, 20]
        dynamic = True
        plot_centre = True
        plot_rotate = False
        plot_scale = 1.
        grid_centre = True
        self.dataview_items = []
        self.datacells = []
        self.be_open = True
        self.hourly = False
        self.daily = False
        self.exclude = []
        self.technologies = ''
        self.show_maximums = False
        ifiles = {}
        ifile = ''
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            self.config_file = sys.argv[1]
        else:
            self.config_file = getModelFile('SIREN.ini')
        config.read(self.config_file)
        parents = []
        try:
            parents = getParents(config.items('Parents'))
        except:
            pass
        try:
            base_year = config.get('Base', 'year')
        except:
            base_year = '2012'
        try:
            scenario_prefix = config.get('Files', 'scenario_prefix')
        except:
            scenario_prefix = ''
        try:
            self.scenarios = config.get('Files', 'scenarios')
            if scenario_prefix != '' :
                self.scenarios += '/' + scenario_prefix
            for key, value in parents:
                self.scenarios = self.scenarios.replace(key, value)
            self.scenarios = self.scenarios.replace('$USER$', getUser())
            self.scenarios = self.scenarios.replace('$YEAR$', base_year)
            self.scenarios = self.scenarios[: self.scenarios.rfind('/') + 1]
            if self.scenarios[:3] == '../':
                ups = self.scenarios.split('../')
                me = os.getcwd().split(os.sep)
                me = me[: -(len(ups) - 1)]
                me.append(ups[-1])
                self.scenarios = '/'.join(me)
        except:
            self.scenarios = ''
        try:
            self.year = config.get('Base', 'year')
        except:
            self.year = '2021'
        try:
            self.technologies = config.get('Power', 'technologies')
        except:
            pass
        self.map = ''
        try:
            tmap = config.get('Map', 'map_choice')
        except:
            tmap = ''
        self.map_upper_left = [0., 0.]
        self.map_lower_right = [-90., 180.]
        try:
             upper_left = config.get('Map', 'upper_left' + self.map).split(',')
             self.map_upper_left[0] = float(upper_left[0].strip())
             self.map_upper_left[1] = float(upper_left[1].strip())
             lower_right = config.get('Map', 'lower_right' + self.map).split(',')
             self.map_lower_right[0] = float(lower_right[0].strip())
             self.map_lower_right[1] = float(lower_right[1].strip())
        except:
             try:
                 lower_left = config.get('Map', 'lower_left' + self.map).split(',')
                 upper_right = config.get('Map', 'upper_right' + self.map).split(',')
                 self.map_upper_left[0] = float(upper_right[0].strip())
                 self.map_upper_left[1] = float(lower_left[1].strip())
                 self.map_lower_right[0] = float(lower_left[0].strip())
                 self.map_lower_right[1] = float(upper_right[1].strip())
             except:
                 pass
      # get default colours
        try:
            colors = config.items('Colors')
            for item, colour in colors:
                if item in self.technologies or item in self.colours or item == 'station_name':
                    itm = techClean(item)
                    self.colours[itm] = colour
        except:
            pass
        if tmap != '':
            try:
                colors = config.items('Colors' + tmap)
                for item, colour in colors:
                    if item in self.technologies or item in self.colours or item == 'station_name':
                        itm = techClean(item)
                        self.colours[itm] = colour
            except:
                pass
        self.colours['number_colour'] = self.colours['Station Name']
        try:
            self.helpfile = config.get('Files', 'help')
        except:
            self.helpfile = ''
        self.restorewindows = False
        try:
            rw = config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
        try:
            items = config.items('Dataview')
            for key, value in items:
                if key == 'file_history':
                    self.history = value.split(',')
                elif key == 'file_choices':
                    self.max_files = int(value)
                elif key[:4] == 'file':
                    ifiles[key[4:]] = value.replace('$USER$', getUser())
                elif key == 'grid_centre':
                    if value.lower() in ['false', 'no', 'off']:
                        grid_centre = False
                elif key == 'number_multiple':
                    if value.lower() in ['true', 'yes', 'on']:
                        multiple = True
                elif key == 'number_colour':
                    self.colours[key] = value
                elif key == 'number_dynamic':
                    if value.lower() in ['false', 'no', 'off']:
                        dynamic = False
                elif key == 'number_steps':
                    steps = value.split(',')
                    number_steps = []
                    try:
                        for step in steps:
                            number_steps.append(int(step))
                        self.number_steps = number_steps
                    except:
                        pass
                elif key == 'plot_opacity':
                    try:
                        opacity = float(value)
                    except:
                        pass
                elif key == 'plot_centre':
                    if value.lower() in ['false', 'no', 'off']:
                        plot_centre = False
                elif key == 'number_show':
                    if value.lower() in ['false', 'no', 'off']:
                        numbers = False
                elif key == 'plot_rotate':
                    if value.lower() in ['true', 'yes', 'on']:
                        plot_rotate = True
                elif key == 'plot_scale':
                    try:
                        plot_scale = float(value)
                        if plot_scale > 1.:
                            plot_scale = 1.
                    except:
                        pass
                elif key == 'show_maximums':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.show_maximums = True
                elif key == 'number_min':
                    try:
                        number_min = int(value)
                    except:
                        pass
        except:
            pass
        if len(ifiles) > 0:
            if self.history is None:
                self.history = sorted(ifiles.keys(), reverse=True)
            ifile = ifiles[self.history[0]]
        self.grid = QtWidgets.QGridLayout()
        self.updated = False
        self.log = QtWidgets.QLabel('')
        rw = 0
        self.grid.addWidget(QtWidgets.QLabel('Recent Files:'), rw, 0)
        self.files = QtWidgets.QComboBox()
        if ifile != '':
            self.popfileslist(ifile, ifiles)
        self.grid.addWidget(self.files, rw, 1, 1, 5)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Data File:'), rw, 0)
        self.file = ClickableQLabel()
        self.file.setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
        self.file.setText(ifile)
        self.grid.addWidget(self.file, rw, 1, 1, 5)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Sheet:'), rw, 0)
        self.sheet = QtWidgets.QComboBox()
        self.sheet.currentIndexChanged.connect(self.sheetChanged)
        self.grid.addWidget(self.sheet, rw, 1)
        label = QtWidgets.QLabel('exclude cells:')
        label.setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(label, rw, 2)
        self.xsheet = QtWidgets.QComboBox()
        self.xsheet.addItem('n/a')
        self.xsheet.addItem('Exclude')
        self.xsheet.setCurrentIndex(0)
    #    self.xsheet.currentIndexChanged.connect(self.xsheetChanged)
        self.grid.addWidget(self.xsheet, rw, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Data Choices:'), rw, 0)
        col = 1
        techs = self.technologies.split(' ')
        self.columns = []
        for tech in sorted(techs):
            self.columns.append(QtWidgets.QCheckBox(techClean(tech), self))
            self.columns[-1].setCheckState(QtCore.Qt.Checked)
            self.columns[-1].stateChanged.connect(self.somethingChanged)
            self.grid.addWidget(self.columns[-1], rw, col)
            col += 1
            if col > 4:
                rw += 1
                col = 1
        rw += 1
        self.maximums = {}
        if self.show_maximums:
            self.grid.addWidget(QtWidgets.QLabel('Maximums:'), rw, 0)
            self.capture = QtWidgets.QCheckBox('Capture')
            self.capture.setCheckState(QtCore.Qt.Unchecked)
            self.grid.addWidget(self.capture, rw + 1, 0)
            self.usemax = QtWidgets.QCheckBox('Use')
            self.usemax.setCheckState(QtCore.Qt.Unchecked)
            self.grid.addWidget(self.usemax, rw + 2, 0)
            rw_2 = rw
            col = 1
            for tech in sorted(techs):
                self.maximums[techClean(tech)] = [0, QtWidgets.QLineEdit('')]
                self.grid.addWidget(self.maximums[techClean(tech)][1], rw, col)
                col += 1
                if col > 4:
                    rw += 1
                    col = 1
            rw += 1
            if rw <= rw_2:
                rw = rw_2 + 1
        else:
            for tech in sorted(techs):
                self.maximums[techClean(tech)] = [0]
        self.grid.addWidget(QtWidgets.QLabel('Type of Chart:'), rw, 0)
        plots = ['Bar Chart', 'Pie Chart'] #, 'Grid fill']
        self.plottype = QtWidgets.QComboBox()
        for plot in plots:
             self.plottype.addItem(plot)
        self.grid.addWidget(self.plottype, rw, 1) #, 1, 2)
        self.rotate = QtWidgets.QCheckBox('Vertical Bar Chart')
        if plot_rotate:
            self.rotate.setCheckState(QtCore.Qt.Checked)
        else:
            self.rotate.setCheckState(QtCore.Qt.Unchecked)
        self.rotate.stateChanged.connect(self.somethingChanged)
        self.grid.addWidget(self.rotate, rw, 2)
        self.plotcentre = QtWidgets.QCheckBox('Centre in cell')
        if plot_centre:
            self.plotcentre.setCheckState(QtCore.Qt.Checked)
        else:
            self.plotcentre.setCheckState(QtCore.Qt.Unchecked)
        self.plotcentre.stateChanged.connect(self.somethingChanged)
        self.grid.addWidget(self.plotcentre, rw, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Opacity:'), rw, 0)
        self.opacitySpin = QtWidgets.QDoubleSpinBox()
        self.opacitySpin.setRange(0, 1.)
        self.opacitySpin.setDecimals(2)
        self.opacitySpin.setSingleStep(.05)
        self.opacitySpin.setValue(opacity)
        self.opacitySpin.valueChanged[str].connect(self.somethingChanged)
        self.grid.addWidget(self.opacitySpin, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('(of pie slices)'), rw, 2)
        self.grid.addWidget(QtWidgets.QLabel('Scale of chart in cell:'), rw, 3)
        self.scaleSpin = QtWidgets.QDoubleSpinBox()
        self.scaleSpin.setRange(0.1, 1.)
        self.scaleSpin.setDecimals(2)
        self.scaleSpin.setSingleStep(.05)
        self.scaleSpin.setValue(plot_scale)
        self.scaleSpin.valueChanged[str].connect(self.somethingChanged)
        self.grid.addWidget(self.scaleSpin, rw, 4)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Show pct. values:'), rw, 0)
        self.numbers = QtWidgets.QCheckBox('Show')
        if numbers:
            self.numbers.setCheckState(QtCore.Qt.Checked)
        else:
            self.numbers.setCheckState(QtCore.Qt.Unchecked)
        self.numbers.stateChanged.connect(self.somethingChanged)
        self.multiple = QtWidgets.QCheckBox('Each')
        if multiple:
            self.multiple.setCheckState(QtCore.Qt.Checked)
        else:
            self.multiple.setCheckState(QtCore.Qt.Unchecked)
        self.multiple.stateChanged.connect(self.somethingChanged)
        two_in_one = QtWidgets.QHBoxLayout()
        two_in_one.addWidget(self.numbers)
        two_in_one.addWidget(self.multiple)
        self.grid.addLayout(two_in_one, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('(of maximum value)'), rw, 2, 1, 2)
        self.dynamic = QtWidgets.QCheckBox('Dynamic scale')
        if dynamic:
            self.dynamic.setCheckState(QtCore.Qt.Checked)
        else:
            self.dynamic.setCheckState(QtCore.Qt.Unchecked)
        self.dynamic.stateChanged.connect(self.somethingChanged)
        self.grid.addWidget(self.dynamic, rw, 3)
        self.numbtn = QtWidgets.QPushButton('Pct. colour', self)
        self.numbtn.clicked.connect(self.colourChanged)
        self.numbtn.setStyleSheet('QPushButton {background-color: %s; color: %s;}' %
                                  (self.colours['number_colour'], self.colours['number_colour']))
        if self.colours['number_colour'].lower() == '#ffffff':
            self.numbtn.setStyleSheet('QPushButton {color: #808080;}')
        self.grid.addWidget(self.numbtn, rw, 4)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Show pct. >=:'), rw, 0)
        self.number_min = QtWidgets.QSpinBox()
        self.number_min.setRange(0, 100)
        self.number_min.setSingleStep(10)
        self.number_min.setValue(number_min)
        self.grid.addWidget(self.number_min, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('% (of maximum value)'), rw, 2, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Centre on grid:'), rw, 0)
        self.gridcentre = QtWidgets.QCheckBox()
        if grid_centre:
            self.gridcentre.setCheckState(QtCore.Qt.Checked)
        else:
            self.gridcentre.setCheckState(QtCore.Qt.Unchecked)
        self.gridcentre.stateChanged.connect(self.somethingChanged)
        self.grid.addWidget(self.gridcentre, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('(centre charts on weather grid)'), rw, 2, 1, 3)
        if ifile != '':
            self.get_file_config(self.history[0])
        self.files.currentIndexChanged.connect(self.filesChanged)
        self.file.clicked.connect(self.fileChanged)
        self.plottype.currentIndexChanged.connect(self.somethingChanged)
        rw += 1
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.log.setPalette(msg_palette)
        self.grid.addWidget(self.log, rw, 1, 1, 4)
        rw += 1
        quit = QtWidgets.QPushButton('Done', self)
        self.grid.addWidget(quit, rw, 0)
        quit.clicked.connect(self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        doit = QtWidgets.QPushButton('Show', self)
        self.grid.addWidget(doit, rw, 1)
        doit.clicked.connect(self.showClicked)
        hide = QtWidgets.QPushButton('Hide', self)
        self.grid.addWidget(hide, rw, 2)
        hide.clicked.connect(self.hideClicked)
        help = QtWidgets.QPushButton('Help', self)
        self.grid.addWidget(help, rw, 4)
        help.clicked.connect(self.helpClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Renewable DataView Overlay')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            move_right = False
        else:
            move_right = True
        if self.restorewindows:
            try:
                rw = config.get('Windows', 'dataview_size').split(',')
                self.resize(int(rw[0]), int(rw[1]))
                mp = config.get('Windows', 'dataview_pos').split(',')
                self.move(int(mp[0]), int(mp[1]))
                move_right = False
            except:
                pass
        if move_right:
            frameGm = self.frameGeometry()
            screen = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
            trPoint = QtWidgets.QApplication.desktop().availableGeometry(screen).topRight()
            frameGm.moveTopRight(trPoint)
            self.move(frameGm.topRight())
        self.show()

    def colourChanged(self):
        col = QtWidgets.QColorDialog.getColor(QtGui.QColor(''))
        if col.isValid():
            self.colours['number_colour'] = col.name()
            self.numbtn.setStyleSheet('QPushButton {background-color: %s; color: %s;}' % (col.name(), col.name()))
            if self.colours['number_colour'].lower() == '#ffffff':
                self.numbtn.setStyleSheet('QPushButton {color: #808080;}')
            self.somethingChanged()

    def clear_Resource(self):
        try:
            for i in range(len(self.resource_items)):
                self.scene.removeItem(self.resource_items[i])
            del self.resource_items
        except:
            pass

    def get_file_config(self, choice=''):
        ifile = ''
        isheet = ''
        ignore = True
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        columns = []
        self.plottype.setCurrentIndex(0)
        try: # get list of files if any
            items = config.items('Dataview')
            for key, value in items:
                if key == 'columns' + choice:
                    columns = strSplit(value)
                    for col in range(len(self.columns)):
                        if self.columns[col].text() in columns:
                            self.columns[col].setCheckState(QtCore.Qt.Checked)
                        else:
                            self.columns[col].setCheckState(QtCore.Qt.Unchecked)
                elif key == 'file' + choice:
                    ifile = value.replace('$USER$', getUser())
                elif key == 'sheet' + choice:
                    isheet = value
                elif key == 'plot' + choice:
                    self.plottype.setCurrentIndex(self.plottype.findText(value))
        except:
             pass
        if ifile != '':
            self.openFile(ifile)
            self.setSheet(isheet)
        ignore = False

    def popfileslist(self, ifile, ifiles=None):
        self.setup[1] = True
        if ifiles is None:
             ifiles = {}
             for i in range(self.files.count()):
                 ifiles[self.history[i]] = self.files.itemText(i)
        if self.history is None:
            self.history = ['']
            ifiles = {'': ifile}
        else:
            for i in range(len(self.history)):
                if ifile == ifiles[self.history[i]]:
                    self.history.insert(0, self.history.pop(i)) # make this entry first
                    break
            else:
                if len(self.history) >= self.max_files:
                    self.history.insert(0, self.history.pop(-1)) # make last entry first
                else:
                    hist = sorted(self.history)
                    if hist[0] != '':
                        ent = ''
                    else:
                        for i in range(1, len(hist)):
                            if str(i) != hist[i]:
                                ent = str(i)
                                break
                        else:
                            ent = str(i + 1)
                    self.history.insert(0, ent)
                ifiles[self.history[0]] = ifile
     #           self.getData(ifile)
        self.files.clear()
        for i in range(len(self.history)):
            try:
                self.files.addItem(ifiles[self.history[i]])
            except:
                pass
        self.files.setCurrentIndex(0)
        self.setup[1] = False

    def openFile(self, ifile):
        if self.book is not None:
            self.book.close()
            self.book = None
        self.book = WorkBook()
        try:
            if os.path.exists(ifile):
                self.book.open_workbook(ifile)
            else:
                self.book.open_workbook(self.scenarios + ifile)
        except:
            self.log.setText("Can't open file - " + ifile)
            return
        self.log.setText('File opened - ' + ifile)
        self.file.setText(ifile)
        self.exclude = []
        if 'Exclude' in self.book.sheet_names():
            exclude_worksheet = self.book.sheet_by_name('Exclude')
            num_cols = exclude_worksheet.ncols - 1
            curr_col = -1
            latc = -1
            lonc = -1
            reac = -1
            while curr_col < num_cols:
                curr_col += 1
                if exclude_worksheet.cell_value(0, curr_col) == 'Latitude':
                    latc = curr_col
                elif exclude_worksheet.cell_value(0, curr_col) == 'Longitude':
                    lonc = curr_col
                elif exclude_worksheet.cell_value(0, curr_col) == 'Reasons':
                    reac = curr_col
            num_rows = exclude_worksheet.nrows - 1
            curr_row = 0
            while curr_row < num_rows:
                curr_row += 1
                if exclude_worksheet.cell_value(curr_row, reac) != None and exclude_worksheet.cell_value(curr_row, reac) != '':
                    self.exclude.append([exclude_worksheet.cell_value(curr_row, latc), exclude_worksheet.cell_value(curr_row, lonc)])

    def fileChanged(self):
        self.log.setText('')
        if os.path.exists(self.file.text()):
            curfile = self.file.text()
        else:
            curfile = self.scenarios + self.file.text()
        newfile = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file', curfile, 'Excel Files (*.xls*);;CSV Files (*.csv)')[0]
        if newfile == '':
            return
        if newfile != '':
            self.openFile(newfile)
        self.setSheet()
        for col in range(len(self.columns)):
            if self.columns[col].text() in self.dataview_vars.keys():
                self.columns[col].setVisible(True)
                if self.show_maximums:
                    self.maximums[self.columns[col].text()][1].setVisible(True)
            else:
                self.columns[col].setVisible(False)
                if self.show_maximums:
                    self.maximums[self.columns[col].text()][1].setVisible(False)
            isheet = self.sheet.currentText()
            if newfile[: len(self.scenarios)] == self.scenarios:
                self.file.setText(newfile[len(self.scenarios):])
            else:
                self.file.setText(newfile)
            self.setSheet(isheet)
            self.popfileslist(self.file.text())
            self.setup[1] = False
      ##      self.get_file_config()
            self.updated = True

    def filesChanged(self):
        if self.setup[1]:
            return
        self.setup[0] = True
        self.log.setText('')
        self.saveConfig()
        self.get_file_config(self.history[self.files.currentIndex()])
        self.popfileslist(self.files.currentText())
        self.log.setText("File " + self.file.text() + " 'loaded'")
        self.setup[0] = False

    def somethingChanged(self):
        if not self.setup[0]:
            self.updated = True

    def setSheet(self, isheet=''):
        if isheet == '':
            isheet = self.book.sheet_names()[0]
            try:
                if self.book.sheet_names()[0] == 'Exclude':
                    isheet = self.book.sheet_names()[1]
            except:
                pass
        ndx = 0
        self.sheet.clear()
        j = -1
        for sht in self.book.sheet_names():
            j += 1
            self.sheet.addItem(sht)
            if sht == isheet:
                ndx = j
        self.sheet.setCurrentIndex(ndx)
        try:
            self.getData()
        except:
            self.log.setText('Error getting Data from worksheet')

    def sheetChanged(self):
        self.log.setText('')
        isheet = self.sheet.currentText()
        if isheet not in self.book.sheet_names():
            self.log.setText("Can't find sheet - " + isheet)
            return
        self.getData(isheet)
        self.updated = True

    def helpClicked(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.helpfile, title='DataView Help', section='dataview')
        dialog.exec_()

    @QtCore.pyqtSlot()
    def exit(self):
        self.be_open = False
        self.close()

    def closeEvent(self, event):
        self.clear_DataView()
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)   # need to account for border
            lines.append('dataview_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('dataview_size=%s,%s' % (str(self.width()), str(self.height())))
            updates['Windows'] = lines
            SaveIni(updates)
        if not self.scene.show_coord:
            self.scene._coordGroup.setVisible(False)
        if self.be_open:
            self.procStart.emit('goodbye')
        self.hideClicked()
        event.accept()

    def quitClicked(self):
        if self.book is not None:
            self.book.close()
        if self.updated:
            self.saveConfig()
        self.close()

    def showClicked(self):
        if self.book is None:
            self.log.setText('No Data file!')
            return
        self.dataviewGrid()

    def hideClicked(self):
        self.clear_DataView()

    def saveConfig(self):
        updates = {}
        if self.updated:
            config = configparser.RawConfigParser()
            config.read(self.config_file)
            choice = self.history[0]
            save_file = self.file.text().replace(getUser(), '$USER$')
            try:
                self.max_files = int(config.get('Dataview', 'file_choices'))
            except:
                pass
            lines = []
            lines.append('grid_centre=')
            if not self.gridcentre.isChecked():
                lines[-1] += 'False'
            lines.append('number_colour=')
            if self.colours['number_colour'] != self.colours['Station Name']:
                lines[-1] += self.colours['number_colour']
            lines.append('number_dynamic=')
            if not self.dynamic.isChecked():
                lines[-1] += 'False'
            lines.append('number_min=')
            if self.number_min.value() != 0:
                lines[-1] += str(self.number_min.value())
            lines.append('number_multiple=')
            if self.multiple.isChecked():
                lines[-1] += 'True'
            lines.append('number_show=')
            if not self.numbers.isChecked():
                lines[-1] += 'False'
            lines.append('plot_centre=')
            if not self.plotcentre.isChecked():
                lines[-1] += 'False'
            lines.append('plot_opacity=')
            if self.opacitySpin.value() != 1:
                lines[-1] += str(self.opacitySpin.value())
            lines.append('plot_rotate=')
            if self.rotate.isChecked():
                lines[-1] += 'True'
            lines.append('plot_scale=')
            if self.scaleSpin.value() != 1:
                lines[-1] += str(self.scaleSpin.value())
            if len(self.history) > 0:
                line = ''
                for itm in self.history:
                    line += itm + ','
                line = line[:-1]
                lines.append('file_history=' + line)
            lines.append('file' + choice + '=' + self.file.text().replace(getUser(), '$USER$'))
            lines.append('plot' + choice + '=' + self.plottype.currentText())
            lines.append('sheet' + choice + '=')
            if self.sheet.currentIndex() > 0:
                lines[-1] += self.sheet.currentText()
            try:
                cols = 'columns' + choice + '='
                for col in range(len(self.columns)):
                    if self.columns[col].isChecked() and self.columns[col].text() in self.dataview_vars.keys():
                        try:
                            if self.columns[col].text().index(',') >= 0:
                                try:
                                    if self.columns[col].text().index("'") >= 0:
                                        qte = '"'
                                except:
                                    qte = "'"
                        except:
                           qte = ''
                        cols += qte + self.columns[col].text() + qte + ','
                if cols[-1] != '=':
                    cols = cols[:-1]
                lines.append(cols)
            except:
                pass
            updates['Dataview'] = lines
            SaveIni(updates, ini_file=self.config_file)
            self.updated = False

    def clear_DataView(self):
        for i in range(len(self.dataview_items)):
            self.scene.removeItem(self.dataview_items[i])
        self.dataview_items = []
        QtCore.QCoreApplication.processEvents()
        QtCore.QCoreApplication.flush()

    def getData(self, sheet=None):
        def in_map(a_min, a_max, b_min, b_max):
            return (a_min <= b_max) and (b_min <= a_max)

        self.clear_DataView()
        self.dataview_vars = {}
        self.dataview_items = []
        self.datacells = []
        self.dataview_worksheet = self.book.sheet_by_name(self.sheet.currentText())
        num_rows = self.dataview_worksheet.nrows - 1
        num_cols = self.dataview_worksheet.ncols - 1
#       get column names and format of data
        if self.dataview_worksheet.cell_value(0, 0) == 'Name':  # powerclasses summary table
            curr_col = -1
            tmp_cols = {}
            while curr_col < num_cols:
                curr_col += 1
                tmp_cols[self.dataview_worksheet.cell_value(0, curr_col)] = curr_col
            tmp_tech = []
            curr_row = 0
            while curr_row < num_rows:
                curr_row += 1
                if self.dataview_worksheet.cell_value(curr_row, tmp_cols['Name']) is not None \
                  and self.dataview_worksheet.cell_value(curr_row, tmp_cols['Name']) != '':
                    if not self.dataview_worksheet.cell_value(curr_row, tmp_cols['Technology']) in tmp_tech:
                        tmp_tech.append(self.dataview_worksheet.cell_value(curr_row, tmp_cols['Technology']))
            tmp_tech.sort()
            self.dataview_vars['Latitude'] = 0
            self.dataview_vars['Longitude'] = 1
            i = 2
            for tech in tmp_tech:
                self.dataview_vars[tech] = i
                i += 1
            curr_row = 0
            while curr_row < num_rows:
                curr_row += 1
                if self.dataview_worksheet.cell_value(curr_row, tmp_cols['Name']) is not None\
                  and self.dataview_worksheet.cell_value(curr_row, tmp_cols['Name']) != '':
                    for stn in self.scene._stations.stations:
                        if stn.name == self.dataview_worksheet.cell_value(curr_row, tmp_cols['Name']):
                            a_lat = stn.lat
                            a_lon = stn.lon
                            break
                    else:
                        a_lat = 0
                        a_lon = 0
                    if in_map(a_lat - 0.25, a_lat + 0.25, self.scene.map_lower_right[0],
                              self.scene.map_upper_left[0]) \
                      and in_map(a_lon - 0.3333, a_lon + 0.3333,
                                 self.scene.map_upper_left[1], self.scene.map_lower_right[1]):
                        self.datacells.append([stn.lat, stn.lon])
                        for i in tmp_tech:
                            self.datacells[-1].append(0)
                        i = tmp_tech.index(self.dataview_worksheet.cell_value(curr_row, tmp_cols['Technology'])) + 2
                        self.datacells[-1][i] = float(self.dataview_worksheet.cell_value(curr_row, tmp_cols['Generation (MWh)']))
        else:
            curr_col = -1
            while curr_col < num_cols:
                curr_col += 1
                self.dataview_vars[self.dataview_worksheet.cell_value(0, curr_col)] = curr_col #[curr_col, 0, -1]
            curr_row = 0
            while curr_row < num_rows:
                curr_row += 1
                try:
                    a_lat = float(self.dataview_worksheet.cell_value(curr_row, self.dataview_vars['Latitude']))
                except:
                    continue
                a_lon = float(self.dataview_worksheet.cell_value(curr_row, self.dataview_vars['Longitude']))
                if in_map(a_lat - 0.25, a_lat + 0.25, self.scene.map_lower_right[0],
                          self.scene.map_upper_left[0]) \
                  and in_map(a_lon - 0.3333, a_lon + 0.3333,
                          self.scene.map_upper_left[1], self.scene.map_lower_right[1]):
                    try:
                        self.datacells.append([float(self.dataview_worksheet.cell_value(curr_row, self.dataview_vars['Latitude'])),
                                          float(self.dataview_worksheet.cell_value(curr_row, self.dataview_vars['Longitude']))])
                        for key in self.dataview_vars.keys():
                            if key == 'Latitude' or key == 'Longitude':
                                continue
                            self.datacells[-1].append(self.dataview_worksheet.cell_value(curr_row, self.dataview_vars[key]))
                            if self.datacells[-1][-1] is None:
                                self.datacells[-1][-1] = 0
                    except:
                        pass
        for col in range(len(self.columns)):
            if self.columns[col].text() in self.dataview_vars.keys():
                self.columns[col].setVisible(True)
                if self.show_maximums:
                    self.maximums[self.columns[col].text()][1].setVisible(True)
            else:
                self.columns[col].setVisible(False)
                if self.show_maximums:
                    self.maximums[self.columns[col].text()][1].setVisible(False)

    def dataviewGrid(self):
        def dv_setText():
            if pct == 0:
                return
            txt = QtWidgets.QGraphicsSimpleTextItem('{:d}%'.format(pct))
            txt.setPen(pen)
            txt.setBrush(brush)
            div = 4
            if multiple:
                div += .5
            for step in self.number_steps:
                if num_pct >= step:
                    break
                div += .5
            font.setPointSizeF(x_d / div)
            txt.setFont(font)
            txt.setZValue(0)
            if self.plottype.currentText() == 'Bar Chart':
                if self.rotate.isChecked():
                    txt.setRotation(-90)
                    txt.setPos(p_x, p.y() + y_d + font.pointSize() * 1.5 - txt.boundingRect().height())
                else:
                    if multiple:
                        txt.setPos(p.x(), p_y + y_b - txt.boundingRect().height())
                    else:
                        txt.setPos(p.x(), p.y() + y_d / 2 - font.pointSize())
            else:
                x_b = (1 - (len(txt.text()) + 1) / 5.) * x_d
                txt.setPos(p.x() + x_b, p.y() + y_d / 2 - font.pointSize())
            self.dataview_items.append(txt)
            self.scene.addItem(self.dataview_items[-1])

        self.clear_DataView()
        opacity = self.opacitySpin.value()
        plot_scale = self.scaleSpin.value()
        lon_cell = .3125
        cells = []
        if self.xsheet.currentText() == 'Exclude':
            for cell in self.datacells:
                lat = round(round(cell[0] / 0.5, 0) * 0.5, 4)
                lon = round(round(cell[1] / 0.625, 0) * 0.625, 4)
                ignore = False
                for exclude in self.exclude:
                    if exclude[0] == lat and exclude[1] == lon:
                        ignore = True
                        break
                if ignore:
                    continue
                cells.append(cell)
        else:
            for cell in self.datacells:
                cells.append(cell)
        if self.gridcentre.isChecked():
            lat_lons = {}
            for cel in range(len(cells) -1, -1, -1):
                lat = round(round(cells[cel][0] / 0.5, 0) * 0.5, 4)
                lon = round(round(cells[cel][1] / 0.625, 0) * 0.625, 4)
                lat_lon = str(lat) + '_' + str(lon)
                if lat_lon in lat_lons.keys():
                    tc = lat_lons[lat_lon]
                    for cl in range(2, len(cells[cel])):
                        cells[tc][cl] += cells[cel][cl]
                    del cells[cel]
                else:
                    lat_lons[lat_lon] = cel
                    cells[cel][0] = lat
                    cells[cel][1] = lon
        cols = []
        keys = []
        for col in self.columns:
            if col.isVisible() and col.isChecked():
                cols.append(self.dataview_vars[col.text()])
                keys.append(col.text())
                if self.show_maximums:
                    if self.capture.isChecked():
                        self.maximums[keys[-1]][0] = 0
                        self.maximums[keys[-1]][1].setText('')
                else:
                    self.maximums[keys[-1]][0] = 0
        if len(cols) == 0:
            self.log.setText('No data to display')
            return
        if len(cols) > 1 and self.multiple.isChecked():
            multiple = True
        else:
            multiple = False
        hi_valu = 0
        hi_bar = 0
        lo_valu = -1
        if self.show_maximums and self.usemax.isChecked():
            for key in keys:
                hi_bar = max(self.maximums[key][0], hi_bar)
                hi_valu += self.maximums[key][0]
        else:
            for cell in cells:
                valu = 0
                for cl in range(len(cols)):
                    valu += cell[cols[cl]]
                    hi_bar = max(cell[cols[cl]], hi_bar)
                    if self.show_maximums:
                        if self.capture.isChecked():
                            if cell[cols[cl]] > self.maximums[keys[cl]][0]:
                                 self.maximums[keys[cl]][0] = cell[cols[cl]]
                                 self.maximums[keys[cl]][1].setText(str(cell[cols[cl]]))
                    else:
                        if cell[cols[cl]] > self.maximums[keys[cl]][0]:
                             self.maximums[keys[cl]][0] = cell[cols[cl]]
                hi_valu = max(valu, hi_valu)
        for cell in cells:
            valu = 0
            for cl in range(len(cols)):
                valu += cell[cols[cl]]
            if lo_valu < 0:
                lo_valu = valu
            else:
                lo_valu = min(valu, lo_valu)
        color = QtGui.QColor(self.colours['number_colour'])
        pen = QtGui.QPen(color, self.scene.line_width)
        pen.setJoinStyle(QtCore.Qt.RoundJoin)
        pen.setCapStyle(QtCore.Qt.RoundCap)
        brush = QtGui.QBrush(color)
        font = QtGui.QFont()
        for cell in cells:
            if self.number_min.value() > 0:
                c = 2
                totl = 0
                for key in self.dataview_vars.keys():
                    if key == 'Latitude' or key == 'Longitude':
                        continue
                    for col in self.columns:
                        if key == col.text() and col.isChecked():
                            totl += float(cell[c])
                            break
                    c += 1
                if totl * 100 / hi_valu < self.number_min.value():
                    continue
            lat = cell[0]
            lon = cell[1]
            p = self.scene.mapFromLonLat(QtCore.QPointF(lon - lon_cell, lat + .25))
            pe = self.scene.mapFromLonLat(QtCore.QPointF(lon + lon_cell, lat + .25))
            ps = self.scene.mapFromLonLat(QtCore.QPointF(lon - lon_cell, lat - .25))
            if plot_scale != 1.:
                x_di = (pe.x() - p.x()) * (1 - plot_scale) / 2.
                y_di = (ps.y() - p.y()) * (1 - plot_scale) / 2.
                if self.plotcentre.isChecked(): # centre in cell
                    p.setX(p.x() + x_di)
                    pe.setX(pe.x() - x_di)
                    p.setY(p.y() + y_di)
                    ps.setY(ps.y() - y_di)
                else: # top left
                    pe.setX(pe.x() - x_di * 2)
                    ps.setY(ps.y() - y_di * 2)
            x_d = pe.x() - p.x()
            y_d = ps.y() - p.y()
            if self.plottype.currentText() == 'Bar Chart':
                first_x = -1
                if self.rotate.isChecked():
                    if self.plotcentre.isChecked():
                        x_b = x_d / (len(cols) + 1)
                        p_x = p.x() + x_b / 2
                    else:
                        x_b = x_d / len(cols)
                        p_x = p.x()
                else:
                    if self.plotcentre.isChecked():
                        y_b = y_d / (len(cols) + 1)
                        p_y = p.y() + y_b / 2
                    else:
                        y_b = y_d / len(cols)
                        p_y = p.y()
                c = 2
                totl = 0
                ctr = 0
                for key in self.dataview_vars.keys():
                    if key == 'Latitude' or key == 'Longitude':
                        continue
                    for col in self.columns:
                        if key == col.text() and col.isChecked():
                            ctr += 1
                            bar = cell[c] / hi_bar
                            totl += cell[c]
                            if self.rotate.isChecked():
                                rect = QtWidgets.QGraphicsRectItem(int(p_x), int(ps.y() - y_d * bar), int(x_b), int(y_d * bar))
                            else:
                                rect = QtWidgets.QGraphicsRectItem(int(p.x()), int(p_y), int(x_d * bar), int(y_b))
                            rect.setBrush(QtGui.QColor(self.colours[key]))
                            rect.setOpacity(opacity)
                            self.dataview_items.append(rect)
                            self.scene.addItem(self.dataview_items[-1])
                            if self.numbers.isChecked():
                                if first_x < 0:
                                    first_x = p.x() + (x_d / 2) / 2
                                if multiple:
                                 #   pct = round(bar * 100) # change to be pct of this technology
                                    try:
                                        pct = round(cell[c] / self.maximums[key][0] * 100)
                                    except:
                                        pct = 0
                                    num_pct = pct
                                    if self.dynamic.isChecked():
                                        try:
                                            num_pct = round((totl - lo_valu) / (hi_valu - lo_valu) * 100)
                                        except:
                                            pass
                                    dv_setText()
                                elif ctr == len(cols):
                                    pct = round(totl / hi_valu * 100)
                                    num_pct = pct
                                    if self.dynamic.isChecked():
                                        try:
                                            num_pct = round((totl - lo_valu) / (hi_valu - lo_valu) * 100)
                                        except:
                                            pass
                                    if ctr > 1 and self.rotate.isChecked():
                                        p_x = first_x
                                    dv_setText()
                            if self.rotate.isChecked():
                                p_x += x_b
                            else:
                                p_y += y_b
                            break
                    c += 1
            elif self.plottype.currentText() == 'Pie Chart':
                set_angle = -round(5760 - 5760 / 4)
                c = 2
                totl = 0
                ctr = 0
                for key in self.dataview_vars.keys():
                    if key == 'Latitude' or key == 'Longitude':
                        continue
                    for col in self.columns:
                        if key == col.text() and col.isChecked():
                            ctr += 1
                            totl += cell[c]
                            ellipse = QtWidgets.QGraphicsEllipseItem(int(p.x()), int(p.y()), int(x_d), int(y_d))
                            ellipse.setPos(0, 0)
                            ellipse.setStartAngle(set_angle)
                            angle = -round(float(cell[c] / hi_valu) * 5760)
                            set_angle += angle
                            ellipse.setSpanAngle(angle)
                            ellipse.setBrush(QtGui.QColor(self.colours[key]))
                            ellipse.setOpacity(opacity)
                            self.dataview_items.append(ellipse)
                            self.scene.addItem(self.dataview_items[-1])
                            if self.numbers.isChecked():
                                if multiple:
                                    pct = round((set_angle + round(5760 - 5760 / 4)) / -5760. * 100) # change to be pct of this technology
                                    if self.dynamic.isChecked():
                                        num_pct = round((totl - lo_valu) / (hi_valu - lo_valu) * 100)
                                    else:
                                        num_pct = pct
                                    dv_setText()
                                elif ctr == len(cols):
                                    pct = round((set_angle + round(5760 - 5760 / 4)) / -5760. * 100)
                                    if self.dynamic.isChecked():
                                        num_pct = round((totl - lo_valu) / (hi_valu - lo_valu) * 100)
                                    else:
                                        num_pct = pct
                                    dv_setText()
                            break
                    c += 1
            else:
                continue
        QtCore.QCoreApplication.processEvents()
        QtCore.QCoreApplication.flush()
