#
#  Copyright (C) 2019-2023 Sustainable Energy Now Inc., Angus King
#
#  powerplot.py - This file is possibly part of SIREN.
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

import configparser  # decode .ini file
import os
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from math import log10, ceil
import matplotlib
if matplotlib.__version__ > '3.5.1':
    matplotlib.use('Qt5Agg')
else:
    matplotlib.use('TkAgg')
from matplotlib.font_manager import FontProperties
import numpy as np
import matplotlib.pyplot as plt
import displayobject
from colours import Colours
from credits import fileVersion
from displaytable import Table
from editini import EdtDialog, SaveIni
from getmodels import getModelFile
from senutils import ClickableQLabel, getParents, getUser, ListWidget, strSplit, techClean, WorkBook
from zoompan import ZoomPanX

col_letters = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ'
def ss_col(col, base=0):
    if base == 1:
        col -= 1
    c1 = 0
    c2, c3 = divmod(col, 26)
    c3 += 1
    if c2 > 26:
        c1, c2 = divmod(c2, 26)
    return (col_letters[c1] + col_letters[c2] + col_letters[c3]).strip()


class PowerPlot(QtWidgets.QWidget):

    def set_colour(self, item):
        tgt = item.lower()
        alpha = ''
        for alph in self.alpha_word:
            tgt = tgt.replace(alph, '')
        tgt = tgt.replace('  ', ' ')
        tgt = tgt.strip()
        if tgt == item.lower():
            return self.colours[item.lower()]
        else:
            return self.colours[tgt.strip()] + self.alphahex

    def set_hatch(self, item):
        tgt = item.lower()
        hatch = None
        for alph in self.hatch_word:
            tgt = tgt.replace(alph, '')
        tgt = tgt.replace('  ', ' ')
        tgt = tgt.strip()
        if tgt == item.lower():
            return None
        else:
            # pattern = ['-', '+', 'x', '\\', '|', '/', '*', 'o', 'O', '.']
            return '.'

    def replace_words(self, what, src, tgt):
        words = {'m': ['$MTH$', '$MONTH$'],
                 'y': ['$YEAR$'],
                 's': ['$SHEET']}
        tgt_str = src
        if tgt == 'find':
            for wrd in words[what]:
                tgt_num = tgt_str.find(wrd)
                if tgt_num >= 0:
                    return tgt_num
                tgt_num = tgt_str.find(wrd.lower())
                if tgt_num >= 0:
                    return tgt_num
            return -1
        else:
            for wrd in words[what]:
                tgt_str = tgt_str.replace(wrd, tgt)
                tgt_str = tgt_str.replace(wrd.lower(), tgt)
        return tgt_str

    def __init__(self, help='help.html'):
        super(PowerPlot, self).__init__()
        self.help = help
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            self.config_file = sys.argv[1]
        else:
            self.config_file = getModelFile('powerplot.ini')
        config.read(self.config_file)
        parents = []
        self.colours = {}
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
            colours = config.items('Plot Colors')
            for item, colour in colours:
                itm = item.replace('_', ' ')
                self.colours[itm] = colour
        except:
            pass
        mth_labels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        ifile = ''
        isheet = ''
        columns = []
        self.setup = [False, False]
        self.details = True
        self.book = None
        self.error = False
        # create a colour map based on traffic lights
        cvals  = [-2., -1, 2]
        colors = ['red' ,'orange', 'green']
        norm=plt.Normalize(min(cvals), max(cvals))
        tuples = list(zip(map(norm, cvals), colors))
        self.cmap = matplotlib.colors.LinearSegmentedColormap.from_list('', tuples)
        self.cbar = True
        self.cbar2 = True
        self.toprow = None
        self.rows = None
        self.breakdown_row = -1
        self.zone_row = -1
        self.leapyear = False
        iper = '<none>'
        imax = 0
        self.alpha = 0.25
        self.alpha_word = ['surplus', 'charge']
        self.margin_of_error = .0001
        self.constrained_layout = False
        self.target = ''
        self.overlay = '<none>'
        self.palette = True
        self.hatch_word = ['charge']
        self.history = None
        self.max_files = 10
        self.seasons = {}
        self.interval = 24
        self.show_contribution = False
        self.show_correlation = False
        self.select_day = False
        ifiles = {}
        try:
            items = config.items('Powerplot')
            for key, value in items:
                if key == 'alpha':
                    try:
                        self.alpha = float(value)
                    except:
                        pass
                if key == 'alpha_word':
                    self.alpha_word = value.split(',')
                elif key == 'cbar':
                    if value.lower() in ['false', 'no', 'off']:
                        self.cbar = False
                elif key == 'cbar2':
                    if value.lower() in ['false', 'no', 'off']:
                        self.cbar2 = False
                elif key == 'cmap':
                    self.cmap = value
                elif key == 'constrained_layout':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.constrained_layout = True
                elif key == 'file_history':
                    self.history = value.split(',')
                elif key == 'file_choices':
                    self.max_files = int(value)
                elif key[:4] == 'file':
                    ifiles[key[4:]] = value.replace('$USER$', getUser())
                if key == 'hatch_word':
                    self.hatch_word = value.split(',')
                elif key == 'margin_of_error':
                    try:
                        self.margin_of_error = float(value)
                    except:
                        pass
                elif key == 'palette':
                    if value.lower() in ['false', 'no', 'off']:
                        self.palette = False
                elif key == 'select_day':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.select_day = True
        except:
            pass
        self.alphahex = hex(int(self.alpha * 255))[2:]
        try:
            items = config.items('Power')
            for item, values in items:
                if item[:6] == 'season':
                    if item == 'season':
                        continue
                    mths = values.split(',')
                    if mths[0] in self.seasons.keys():
                        per = mths[0] + ' Period'
                        self.seasons[per] = self.seasons[mths[0]]
                    self.seasons[mths[0]] = list(map(lambda x: int(x) - 1, mths[1:]))
                elif item[:6] == 'period':
                    if item == 'period':
                        continue
                    mths = values.split(',')
                    if mths[0] in self.seasons.keys():
                        per = mths[0] + ' Period'
                    else:
                        per = mths[0]
                    self.seasons[per] = list(map(lambda x: int(x) - 1, mths[1:]))
        except:
            pass
        try:
            value = config.get('Powermatch', 'show_contribution')
            if value.lower() in ['true', 'yes', 'on']:
                self.show_contribution = True
        except:
            pass
        try:
            value = config.get('Powermatch', 'show_correlation')
            if value.lower() in ['true', 'yes', 'on']:
                self.show_correlation = True
        except:
            pass
        if len(ifiles) > 0:
            if self.history is None:
                self.history = sorted(ifiles.keys(), reverse=True)
            ifile = ifiles[self.history[0]]
        matplotlib.rcParams['savefig.directory'] = os.getcwd()
        self.grid = QtWidgets.QGridLayout()
        self.updated = False
        self.colours_updated = False
        self.log = QtWidgets.QLabel('')
        self.log.setText('Preferences file: ' + self.config_file)
        rw = 0
        self.grid.addWidget(QtWidgets.QLabel('Recent Files:'), rw, 0)
        self.files = QtWidgets.QComboBox()
        if ifile != '':
            self.popfileslist(ifile, ifiles)
        self.grid.addWidget(self.files, rw, 1, 1, 5)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('File:'), rw, 0)
        self.file = ClickableQLabel()
        self.file.setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
        self.file.setText('')
        self.grid.addWidget(self.file, rw, 1, 1, 5)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Sheet:'), rw, 0)
        self.sheet = QtWidgets.QComboBox()
        self.grid.addWidget(self.sheet, rw, 1, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Period:'), rw, 0)
        self.period = QtWidgets.QComboBox()
        self.period.addItem('<none>')
        self.period.addItem('Year')
        for mth in mth_labels:
            self.period.addItem(mth)
        for per in sorted(self.seasons.keys()):
            self.period.addItem(per)
        self.grid.addWidget(self.period, rw, 1, 1, 1)
        self.cperiod = QtWidgets.QComboBox()
        self.cperiod.addItem('<none>')
        for mth in mth_labels:
            self.cperiod.addItem(mth)
        self.grid.addWidget(self.cperiod, rw, 2, 1, 1)
        self.grid.addWidget(QtWidgets.QLabel('(Diurnal profile for Period or Month(s))'), rw, 3, 1, 2)
        if self.select_day:
            rw += 1
            self.grid.addWidget(QtWidgets.QLabel('Day:'), rw, 0)
            self.aday = QtWidgets.QComboBox()
            self.aday.addItem('')
            for d in range(1, 32):
                self.aday.addItem(str(d))
            self.grid.addWidget(self.aday, rw, 1)
            self.adaylbl = QtWidgets.QLabel('')
            self.grid.addWidget(self.adaylbl, rw, 2, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Target:'), rw, 0)
        self.targets = QtWidgets.QComboBox()
        self.grid.addWidget(self.targets, rw, 1, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(e.g. Load)'), rw, 3, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Overlay:'), rw, 0)
        self.overlays = QtWidgets.QComboBox()
        self.overlays.addItem('<none>')
        self.overlays.addItem('Charge')
        self.overlays.addItem('Underlying Load')
        self.overlays.setCurrentIndex(0)
        self.grid.addWidget(self.overlays, rw, 1, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(e.g. Charge)'), rw, 3, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Header:'), rw, 0)
        self.suptitle = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.suptitle, rw, 1, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Title:'), rw, 0)
        self.title = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.title, rw, 1, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Maximum:'), rw, 0)
        self.maxSpin = QtWidgets.QDoubleSpinBox()
        self.maxSpin.setDecimals(1)
        self.maxSpin.setRange(0, 10000)
        self.maxSpin.setSingleStep(500)
        self.grid.addWidget(self.maxSpin, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('(Handy if you want to produce a series of plots)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Type of Plot:'), rw, 0)
        plots = ['Bar Chart', 'Cumulative', 'Heat Map', 'Line Chart', 'Step Chart']
        self.plottype = QtWidgets.QComboBox()
        for plot in plots:
             self.plottype.addItem(plot)
        self.plottype.setCurrentIndex(2) # default to cumulative
        self.grid.addWidget(self.plottype, rw, 1) #, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Type of plot - stacked except for Line Chart)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Percentage:'), rw, 0)
        self.percentage = QtWidgets.QCheckBox()
        self.percentage.setCheckState(QtCore.Qt.Unchecked)
        self.grid.addWidget(self.percentage, rw, 1) #, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Check for percentage distribution)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Spill label:'), rw, 0)
        self.spill_label = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.spill_label, rw, 1, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Enter suffix if you want spill labels)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Show Grid:'), rw, 0)
        grids = ['Both', 'Horizontal', 'Vertical', 'None']
        self.gridtype = QtWidgets.QComboBox()
        for grid in grids:
             self.gridtype.addItem(grid)
        self.grid.addWidget(self.gridtype, rw, 1) #, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Choose gridlines)'), rw, 3, 1, 3)
        rw += 1
        self.brk_label = QtWidgets.QLabel('Breakdowns:')
        self.grid.addWidget(self.brk_label, rw, 0)
        brk_h = self.spill_label.sizeHint().height() * 2
        self.brk_order = ListWidget(self)
        self.brk_order.setFixedHeight(brk_h)
        self.grid.addWidget(self.brk_order, rw, 1, 1, 2)
        self.brk_ignore = ListWidget(self)
        self.brk_ignore.setFixedHeight(brk_h)
        self.grid.addWidget(self.brk_ignore, rw, 3, 1, 2)
        self.brk_label.setHidden(True)
        self.brk_order.setHidden(True)
        self.brk_ignore.setHidden(True)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Column Order:\n(move to right\nto exclude)'), rw, 0)
        self.order = ListWidget(self)
        self.grid.addWidget(self.order, rw, 1, 1, 2)
        self.ignore = ListWidget(self)
        self.grid.addWidget(self.ignore, rw, 3, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel(' '), rw, 5)
        if ifile != '':
            self.get_file_config(self.history[0])
        self.files.currentIndexChanged.connect(self.filesChanged)
        self.file.clicked.connect(self.fileChanged)
        self.period.currentIndexChanged.connect(self.periodChanged)
        self.cperiod.currentIndexChanged.connect(self.somethingChanged)
        self.files.currentIndexChanged.connect(self.targetChanged)
        self.sheet.currentIndexChanged.connect(self.sheetChanged)
        self.targets.currentIndexChanged.connect(self.targetChanged)
        self.overlays.currentIndexChanged.connect(self.overlayChanged)
        self.suptitle.textChanged.connect(self.somethingChanged)
        self.title.textChanged.connect(self.somethingChanged)
        self.maxSpin.valueChanged.connect(self.somethingChanged)
        self.plottype.currentIndexChanged.connect(self.somethingChanged)
        self.gridtype.currentIndexChanged.connect(self.somethingChanged)
        self.percentage.stateChanged.connect(self.somethingChanged)
        self.spill_label.textChanged.connect(self.somethingChanged)
        self.order.itemSelectionChanged.connect(self.somethingChanged)
        self.brk_order.itemSelectionChanged.connect(self.somethingChanged)
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
        pp = QtWidgets.QPushButton('Plot', self)
        self.grid.addWidget(pp, rw, 1)
        pp.clicked.connect(self.ppClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('p'), self, self.ppClicked)
        cb = QtWidgets.QPushButton('Colours', self)
        self.grid.addWidget(cb, rw, 2)
        cb.clicked.connect(self.editColours)
        if self.show_contribution or self.show_correlation:
            if self.show_contribution:
                co = 'Contribution'
            else:
                co = 'Correlation'
            corr = QtWidgets.QPushButton(co, self)
            self.grid.addWidget(corr, rw, 3)
            corr.clicked.connect(self.corrClicked)
        editini = QtWidgets.QPushButton('Preferences', self)
        self.grid.addWidget(editini, rw, 4)
        editini.clicked.connect(self.editIniFile)
        help = QtWidgets.QPushButton('Help', self)
        self.grid.addWidget(help, rw, 5)
        help.clicked.connect(self.helpClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        if self.select_day:
            self.adaylbl.setText('(Diurnal profile for a day of ' + self.period.currentText() + ')')
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - powerplot (' + fileVersion() + ') - PowerPlot')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.center()
        self.resize(int(self.sizeHint().width() * 1.4), int(self.sizeHint().height() * 1.4))
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        centerPoint = QtWidgets.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def get_file_config(self, choice=''):
        ifile = ''
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        breakdowns = []
        columns = []
        self.period.setCurrentIndex(0)
        self.cperiod.setCurrentIndex(0)
        self.gridtype.setCurrentIndex(0)
        self.plottype.setCurrentIndex(3)
        try: # get list of files if any
            items = config.items('Powerplot')
            for key, value in items:
                if key == 'breakdown' + choice:
                    breakdowns = strSplit(value)
                elif key == 'columns' + choice:
                    columns = strSplit(value)
                elif key == 'cperiod' + choice:
                    i = self.cperiod.findText(value, QtCore.Qt.MatchExactly)
                    if i >= 0 :
                        self.cperiod.setCurrentIndex(i)
                    else:
                        self.cperiod.setCurrentIndex(0)
                elif key == 'file' + choice:
                    ifile = value.replace('$USER$', getUser())
                elif key == 'grid' + choice:
                    self.gridtype.setCurrentIndex(self.gridtype.findText(value))
                elif key == 'overlay' + choice:
                    try:
                        i = self.overlays.findText(value, QtCore.Qt.MatchExactly)
                        if i >= 0:
                            self.overlays.setCurrentIndex(i)
                        else:
                            self.overlays.setCurrentIndex(0)
                        self.overlay = value
                    except:
                        pass
                elif key == 'percentage' + choice:
                    if value.lower() in ['true', 'yes', 'on']:
                        self.percentage.setCheckState(QtCore.Qt.Checked)
                    else:
                        self.percentage.setCheckState(QtCore.Qt.Unchecked)
                elif key == 'period' + choice:
                    i = self.period.findText(value, QtCore.Qt.MatchExactly)
                    if i >= 0 :
                        self.period.setCurrentIndex(i)
                    else:
                        self.period.setCurrentIndex(0)
                elif key == 'plot' + choice:
                    if value == 'Linegraph':
                        value = 'Line Chart'
                    elif value == 'Step Plot':
                        value = 'Step Chart'
                    self.plottype.setCurrentIndex(self.plottype.findText(value))
                elif key == 'maximum' + choice:
                    try:
                        self.maxSpin.setValue(int(value))
                    except:
                        self.maxSpin.setValue(0)
                elif key == 'sheet' + choice:
                    isheet = value
                elif key == 'spill_label' + choice:
                    try:
                        self.spill_label.setText(value)
                    except:
                        pass
                elif key == 'target' + choice:
                    try:
                        self.target = value
                    except:
                        pass
                elif key == 'suptitle' + choice:
                    self.suptitle.setText(value)
                elif key == 'title' + choice:
                    self.title.setText(value)
        except:
             pass
        if ifile != '':
            if self.book is not None:
                self.book.release_resources()
                self.book = None
                self.toprow = None
            self.file.setText(ifile)
            if os.path.exists(ifile):
                self.setSheet(ifile, isheet)
            elif os.path.exists(self.scenarios + ifile):
                self.setSheet(self.scenarios + ifile, isheet)
            else:
                self.log.setText('File not found - ' + ifile)
                # Maybe delete dodgy entry
                self.error = True
                return
            self.setColumns(isheet, columns=columns, breakdowns=breakdowns)
            for column in columns:
                self.check_colour(column, config, add=False)

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
        self.files.clear()
        for i in range(len(self.history)):
            try:
                self.files.addItem(ifiles[self.history[i]])
            except:
                pass
        self.files.setCurrentIndex(0)
        self.setup[1] = False

    def fileChanged(self):
        if os.path.exists(self.file.text()):
            curfile = self.file.text()
        else:
            curfile = self.scenarios + self.file.text()
        newfile = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file', curfile)[0]
        if newfile != '':
            if self.book is not None:
                self.book.release_resources()
                self.book = None
            self.toprow = None
            isheet = self.sheet.currentText()
            self.setSheet(newfile, isheet)
            if newfile[: len(self.scenarios)] == self.scenarios:
                self.file.setText(newfile[len(self.scenarios):])
            else:
                self.file.setText(newfile)
            self.popfileslist(self.file.text())
            self.setup[1] = False
            self.updated = True

    def filesChanged(self):
        if self.setup[1]:
            return
        self.setup[0] = True
        self.saveConfig()
        self.get_file_config(self.history[self.files.currentIndex()])
        if self.error:
            self.error = False
            return
        self.popfileslist(self.files.currentText())
        self.log.setText("File 'loaded'")
        self.setup[0] = False

    def periodChanged(self):
        if self.select_day:
            self.adaylbl.setText('(Diurnal profile for a day of ' + self.period.currentText() + ')')
        if not self.setup[0]:
            self.updated = True

    def somethingChanged(self):
       # if self.plottype.currentText() == 'Bar Chart' and self.period.currentText() == '<none>' and 1 == 2:
       #     self.plottype.setCurrentIndex(self.plottype.currentIndex() + 1) # set to something else
       # el
        if self.plottype.currentText() == 'Line Chart' and self.percentage.isChecked():
            self.percentage.setCheckState(QtCore.Qt.Unchecked)
        if not self.setup[0]:
            self.updated = True

    def setSheet(self, ifile, isheet):
        if self.book is None:
            try:
                self.book = WorkBook()
                self.book.open_workbook(ifile)
            except:
                self.log.setText("Can't open file - " + ifile)
                return
        ndx = 0
        self.sheet.clear()
        j = -1
        for sht in self.book.sheet_names():
            j += 1
            self.sheet.addItem(sht)
            if sht == isheet:
                ndx = j
        self.sheet.setCurrentIndex(ndx)

    def sheetChanged(self):
        self.toprow = None
        if self.book is None:
            self.book = WorkBook()
            self.book.open_workbook(newfile)
        isheet = self.sheet.currentText()
        if isheet not in self.book.sheet_names():
            self.log.setText("Can't find sheet - " + isheet)
            return
        self.setColumns(isheet)
        self.updated = True

    def overlayChanged(self):
        self.log.setText('')
        overlay = self.overlays.currentText()
        if overlay != self.overlay:
            self.overlay = overlay
            self.updated = True

    def targetChanged(self):
        if self.error:
            self.error = False
            return
        target = self.targets.currentText()
        if target != self.target:
            if target == '<none>':
                if self.target != '':
                    self.ignore.addItem(self.target)
                    try:
                        self.ignore.item(self.ignore.count() - 1).setBackground(QtGui.QColor(tgt))
                    except:
                        pass
                self.target = target
            else:
                items = self.order.findItems(target, QtCore.Qt.MatchExactly)
                for item in items:
                    self.order.takeItem(self.order.row(item))
                items = self.ignore.findItems(target, QtCore.Qt.MatchExactly)
                for item in items:
                    self.ignore.takeItem(self.ignore.row(item))
                if self.target != '<none>':
                    self.ignore.addItem(self.target)
                    tgt = self.target.lower()
                    for alph in self.alpha_word:
                        tgt.replace(alph, '')
                    tgt.replace('  ', ' ')
                    try:
                        self.ignore.item(self.ignore.count() - 1).setBackground(QtGui.QColor(tgt))
                    except:
                        pass
                self.target = target
        self.updated = True

    def setColumns(self, isheet, columns=[], breakdowns=[]):
        try:
            ws = self.book.sheet_by_name(isheet)
        except:
            self.log.setText("Can't find sheet - " + isheet)
            return
        tech_row = -1
        row = 0
        self.zone_row = -1
        self.breakdown_row = -1
        ignores = []
        while row < ws.nrows:
            if ws.cell_value(row, 0) in ['Split', 'Breakdown']:
                self.breakdown_row = row
                self.brk_order.clear()
                self.brk_ignore.clear()
                brk_orders = []
                for col in range(2, ws.ncols):
                    if ws.cell_value(row, col) is not None and ws.cell_value(row, col) != '':
                        try:
                            brk_orders.index(ws.cell_value(row, col))
                        except:
                            brk_orders.append(ws.cell_value(row, col))
                for brk_order in brk_orders:
                    if brk_order in breakdowns:
                        self.brk_order.addItem(brk_order)
                    else:
                        self.brk_ignore.addItem(brk_order)
                self.brk_label.setHidden(False)
                self.brk_order.setHidden(False)
                self.brk_ignore.setHidden(False)
            elif ws.cell_value(row, 0) == 'Technology':
                tech_row = row
            elif ws.cell_value(row, 0) == 'Zone':
                self.zone_row = row
            elif ws.cell_value(row, 0) in ['Hour', 'Interval', 'Trading Interval']:
                if ws.cell_value(row, 1) != 'Period':
                    self.log.setText(isheet + ' sheet format incorrect')
                    return
                if ws.cell_value(row, 0) in ['Interval', 'Trading Interval']:
                    self.interval = 48
                else:
                    self.interval = 24
                if tech_row >= 0:
                    self.toprow = [tech_row, row]
                else:
                    self.toprow = [row, row]
                self.rows = ws.nrows - (row + 1)
                oldcolumns = []
                if len(columns) == 0:
                    for col in range(self.order.count()):
                        oldcolumns.append(self.order.item(col).text())
                self.order.clear()
                self.ignore.clear()
                self.targets.clear()
                self.targets.addItem('<none>')
                self.targets.setCurrentIndex(0)
                if tech_row >= 0:
                    the_row = tech_row
                else:
                    the_row = row
                for col in range(ws.ncols -1, 1, -1):
                    try:
                        column = ws.cell_value(the_row, col).replace('\n',' ')
                    except:
                        column = str(ws.cell_value(the_row, col))
                    if self.breakdown_row >= 0:
                        try:
                            i = brk_orders.index(ws.cell_value(self.breakdown_row, col))
                            if i > 0:
                                if self.zone_row < 0:
                                    continue
                                if ws.cell_value(self.zone_row, col) is not None and ws.cell_value(self.zone_row, col) != '':
                                    continue
                        except:
                            pass
                    if self.zone_row > 0 and ws.cell_value(self.zone_row, col) != '' and ws.cell_value(self.zone_row, col) is not None:
                        column = ws.cell_value(self.zone_row, col).replace('\n',' ') + '.' + column
                    if column in oldcolumns:
                        if column not in columns:
                            columns.append(column)
                    if self.targets.findText(column, QtCore.Qt.MatchExactly) >= 0:
                        pass
                    else:
                        self.targets.addItem(column)
                        if column == self.target:
                            itm = self.targets.findText(column, QtCore.Qt.MatchExactly)
                            self.targets.setCurrentIndex(itm)
                    if column in columns:
                        pass
                    else:
                        if column != self.target:
                            if column in ignores:
                                continue
                            ignores.append(column)
                            self.ignore.addItem(column)
                            try:
                                self.ignore.item(self.ignore.count() - \
                                 1).setBackground(QtGui.QColor(self.set_colour(column)))
                            except:
                                pass
                for column in columns:
                    if column != self.target:
                        self.order.addItem(column)
                        try:
                            self.order.item(self.order.count() - 1).setBackground(QtGui.QColor(self.set_colour(column)))
                        except:
                            pass
                break
            row += 1
        else:
            self.log.setText(isheet + ' sheet format incorrect')
        if self.breakdown_row < 0:
            self.brk_label.setHidden(True)
            self.brk_order.setHidden(True)
            self.brk_ignore.setHidden(True)

    def helpClicked(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.help,
                 title='Help for powerplot (' + fileVersion() + ')', section='powerplot')
        dialog.exec()

    def quitClicked(self):
        if self.book is not None:
            self.book.release_resources()
        if not self.updated and not self.colours_updated:
            self.close()
        self.saveConfig()
        self.close()

    def saveConfig(self):
        updates = {}
        if self.updated:
            config = configparser.RawConfigParser()
            config.read(self.config_file)
            choice = self.history[0]
            save_file = self.file.text().replace(getUser(), '$USER$')
            try:
                self.max_files = int(config.get('Powerplot', 'file_choices'))
            except:
                pass
            lines = []
            if len(self.history) > 0:
                line = ''
                for itm in self.history:
                    line += itm + ','
                line = line[:-1]
                lines.append('file_history=' + line)
            lines.append('file' + choice + '=' + self.file.text().replace(getUser(), '$USER$'))
            lines.append('grid' + choice + '=' + self.gridtype.currentText())
            lines.append('sheet' + choice + '=' + self.sheet.currentText())
            if self.overlay == '<none>':
                overlay = ''
            else:
                overlay = self.overlay
            lines.append('overlay' + choice + '=' + overlay)
            lines.append('period' + choice + '=')
            if self.period.currentText() != '<none>':
                lines[-1] = lines[-1] + self.period.currentText()
            lines.append('cperiod' + choice + '=')
            if self.cperiod.currentText() != '<none>':
                lines[-1] = lines[-1] + self.cperiod.currentText()
            lines.append('spill_label' + choice + '=' + self.spill_label.text())
            lines.append('suptitle' + choice + '=' + self.suptitle.text())
            lines.append('target' + choice + '=' + self.target)
            lines.append('title' + choice + '=' + self.title.text())
            lines.append('maximum' + choice + '=')
            if self.maxSpin.value() != 0:
                lines[-1] = lines[-1] + str(self.maxSpin.value())
            lines.append('percentage' + choice + '=')
            if self.percentage.isChecked():
                lines[-1] = lines[-1] + 'True'
            lines.append('plot' + choice + '=' + self.plottype.currentText())
            cols = 'columns' + choice + '='
            for col in range(self.order.count()):
                try:
                    if self.order.item(col).text().index(',') >= 0:
                        try:
                            if self.order.item(col).text().index("'") >= 0:
                                qte = '"'
                        except:
                            qte = "'"
                except:
                    qte = ''
                cols += qte + self.order.item(col).text() + qte + ','
            if cols[-1] != '=':
                cols = cols[:-1]
            lines.append(cols)
            cols = 'breakdown' + choice + '='
            for col in range(self.brk_order.count()):
                try:
                    if self.brk_order.item(col).text().index(',') >= 0:
                        try:
                            if self.brk_order.item(col).text().index("'") >= 0:
                                qte = '"'
                        except:
                            qte = "'"
                except:
                    qte = ''
                cols += qte + self.brk_order.item(col).text() + qte + ','
            if cols[-1] != '=':
                cols = cols[:-1]
            lines.append(cols)
            updates['Powerplot'] = lines
        if self.colours_updated:
            lines = []
            for key, value in self.colours.items():
                if value != '':
                    lines.append(key.replace(' ', '_') + '=' + value)
            updates['Plot Colors'] = lines
        SaveIni(updates, ini_file=self.config_file)
        self.updated = False
        self.colours_updated = False

    def editColours(self, color=False):
        # if they've selected some items I'll create a palette of colours for them
        palette = []
        if self.palette:
            for item in self.order.selectedItems():
                palette.append(item.text())
            for item in self.ignore.selectedItems():
                palette.append(item.text())
        dialr = Colours(section='Plot Colors', ini_file=self.config_file, add_colour=color,
                        palette=palette)
        dialr.exec()
        self.colours = {}
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        try:
            colours = config.items('Plot Colors')
            for item, colour in colours:
                itm = item.replace('_', ' ')
                self.colours[itm] = colour
        except:
            pass
        for c in range(self.order.count()):
            col = self.order.item(c).text()
            try:
                self.order.item(c).setBackground(QtGui.QColor(self.set_colour(col)))
            except:
                pass
        for c in range(self.ignore.count()):
            col = self.ignore.item(c).text()
            try:
                self.ignore.item(c).setBackground(QtGui.QColor(self.set_colour(col)))
            except:
                pass

    def check_colour(self, colour, config, add=True):
        colr = colour.lower()
        if colr in self.colours.keys():
            return True
        try:
            colr2 = self.set_colour(colr)
            return True
        except:
            pass
        colr2 = colr.replace(' ', '_')
        if config is not None:
            try:
                amap = config.get('Map', 'map_choice')
                tgt_colr = config.get('Colors' + amap, colr2)
                self.colours[colr] = tgt_colr
                self.colours_updated = True
                return True
            except:
                pass
            try:
                tgt_colr = config.get('Colors', colr2)
                self.colours[colr] = tgt_colr
                self.colours_updated = True
                return True
            except:
                pass
            if not add:
                return False
            self.editColours(color=colour)
            if colr not in self.colours.keys():
                self.log.setText('No colour for ' + colour)
                return False
            return True
        if not add:
            return False
        self.editColours(color=colour)
        if colr not in self.colours.keys():
            self.log.setText('No colour for ' + colour)
            return False
        return True

    def editIniFile(self):
        dialr = EdtDialog(self.config_file, section='[Powermatch]')
        dialr.exec_()
     #   self.get_config()   # refresh config values
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        self.log.setText(self.config_file + ' edited. Reload may be required.')

    def ppClicked(self):
        if self.book is None:
            self.log.setText('Error accessing Workbook.')
            return
        if self.order.count() == 0 and self.target == '<none>' and self.plottype.currentText() != 'Line Chart':
            self.log.setText('Nothing to plot.')
            return
        isheet = self.sheet.currentText()
        if isheet == '':
            self.log.setText('Sheet not set.')
            return
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
            config = configparser.RawConfigParser()
            config.read(self.config_file)
        else:
            config = None
        for c in range(self.order.count()):
            if not self.check_colour(self.order.item(c).text(), config):
                return
        if self.target != '<none>':
            if not self.check_colour(self.target, config):
                return
            if not self.check_colour('shortfall', config):
                return
        del config
        self.log.setText('')
        i = self.file.text().rfind('/')
        if i > 0:
            matplotlib.rcParams['savefig.directory'] = self.file.text()[:i + 1]
        else:
            matplotlib.rcParams['savefig.directory'] = self.scenarios
        if self.gridtype.currentText() == 'Both':
            gridtype = 'both'
        elif self.gridtype.currentText() == 'Horizontal':
            gridtype = 'y'
        elif self.gridtype.currentText() == 'Vertical':
            gridtype = 'x'
        else:
            gridtype = ''
        ws = self.book.sheet_by_name(isheet)
        if self.toprow is None:
            tech_row = -1
            row = 0
            self.zone_row =  -1
            self.breakdown_row = -1
            while row < ws.nrows:
                if ws.cell_value(row, 0) in ['Split', 'Breakdown']:
                    self.breakdown_row = row
                elif ws.cell_value(row, 0) == 'Technology':
                    tech_row = row
                elif ws.cell_value(row, 0) == 'Zone':
                    self.zone_row = row
                elif ws.cell_value(row, 0) in ['Hour', 'Interval', 'Trading Interval']:
                    if ws.cell_value(row, 1) != 'Period':
                        self.log.setText(isheet + ' sheet format incorrect')
                        return
                    if tech_row >= 0:
                        self.toprow = [tech_row, row]
                    else:
                        self.toprow = [row, row]
                    self.rows = ws.nrows - (row + 1)
                    break
                row += 1
        try:
            year = int(ws.cell_value(self.toprow[1] + 1, 1)[:4])
            if year % 4 == 0 and year % 100 != 0 or year % 400 == 0:
                self.leapyear = True
            else:
                self.leapyear = False
        except:
            self.leapyear = False
            year = ''
        the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if self.leapyear: #rows == 8784: # leap year
            the_days[1] = 29
        mth_labels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        hr_labels = ['0:00', '4:00', '8:00', '12:00', '16:00', '20:00', '23:00']
        figname = self.plottype.currentText().lower().replace(' ','') + '_' + str(year)
        breakdowns = []
        suptitle = self.suptitle.text()
        if self.breakdown_row >= 0:
            for c in range(self.brk_order.count()):
                breakdowns.append(self.brk_order.item(c).text())
        if self.period.currentText() == '<none>' \
          or (self.period.currentText() == 'Year' and self.plottype.currentText() == 'Heat Map'): # full year of hourly figures
            m = 0
            d = 1
            day_labels = []
            days_per_label = 1 # set to 1 and flex_on to True to have some flexibilty in x_labels
            flex_on = True
            while m < len(the_days):
                day_labels.append('%s %s' % (str(d), mth_labels[m]))
                d += days_per_label
                if d > the_days[m]:
                    d = d - the_days[m]
                    m += 1
            x = []
            len_x = self.rows
            for i in range(len_x):
                x.append(i)
            load = []
            tgt_col = -1
            overlay_cols = []
            overlay = []
            data = []
            label = []
            maxy = 0
            miny = 0
            titl = self.replace_words('y', self.title.text(), str(year))
            titl = self.replace_words('m', titl, '')
            titl = titl.replace('  ', ' ')
            titl = titl.replace('Diurnal ', '')
            titl = titl.replace('Diurnal', '')
            titl = self.replace_words('s', titl, isheet)
            if suptitle != '':
                suptitle = self.replace_words('y', suptitle, str(year))
                suptitle = self.replace_words('m', suptitle, '')
                suptitle = suptitle.replace('  ', ' ')
                suptitle = suptitle.replace('Diurnal ', '')
                suptitle = suptitle.replace('Diurnal', '')
                suptitle = self.replace_words('s', suptitle, isheet)
            for c2 in range(2, ws.ncols):
                try:
                    column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                except:
                    column = str(ws.cell_value(self.toprow[0], c2))
                if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                    column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                if column == self.target:
                    tgt_col = c2
                if self.overlay != '<none>':
                    if column == self.target or column[:len(self.overlay)] == self.overlay:
                        overlay_cols.append(c2)
            if len(breakdowns) > 0:
                for c in range(self.order.count() -1, -1, -1):
                    col = self.order.item(c).text()
                    for c2 in range(2, ws.ncols):
                        try:
                            brkdown = ws.cell_value(self.breakdown_row, c2)
                            if brkdown is not None and brkdown != '' and brkdown != breakdowns[0]:
                                continue
                        except:
                            pass
                        try:
                            column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                        except:
                            column = str(ws.cell_value(self.toprow[0], c2))
                        if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                            column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                        if column == col:
                            data.append([])
                            label.append(column)
                            for row in range(self.toprow[1] + 1, self.toprow[1] + self.rows + 1):
                                data[-1].append(ws.cell_value(row, c2))
                                try:
                                    maxy = max(maxy, data[-1][-1])
                                    miny = min(miny, data[-1][-1])
                                except:
                                    self.log.setText("Invalid data - '" + column + "' (" + ss_col(c2) + ') row ' + str(row + 1) + " is '" + str(data[-1][-1]) + "'")
                                    return
                for breakdown in breakdowns[1:]:
                    for c in range(self.order.count() -1, -1, -1):
                        col = self.order.item(c).text()
                        for c2 in range(2, ws.ncols):
                            try:
                                brkdown = ws.cell_value(self.breakdown_row, c2)
                                if brkdown != breakdown:
                                    continue
                            except:
                                continue
                            try:
                                column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                            except:
                                column = str(ws.cell_value(self.toprow[0], c2))
                            if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                                column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                            if column == col:
                                data.append([])
                                label.append(column + ' ' + breakdown)
                                for row in range(self.toprow[1] + 1, self.toprow[1] + self.rows + 1):
                                    data[-1].append(ws.cell_value(row, c2))
                                    try:
                                        maxy = max(maxy, data[-1][-1])
                                        miny = min(miny, data[-1][-1])
                                    except:
                                        self.log.setText("Invalid data - '" + column + "' (" + ss_col(c2) + ') row ' + str(row + 1) + " is '" + str(data[-1][-1]) + "'")
                                        return
            else:
                for c in range(self.order.count() -1, -1, -1):
                    col = self.order.item(c).text()
                    for c2 in range(2, ws.ncols):
                        try:
                            column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                        except:
                            column = str(ws.cell_value(self.toprow[0], c2))
                        if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                            column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                        if column == col:
                            data.append([])
                            label.append(column)
                            for row in range(self.toprow[1] + 1, self.toprow[1] + self.rows + 1):
                                data[-1].append(ws.cell_value(row, c2))
                                try:
                                    maxy = max(maxy, data[-1][-1])
                                    miny = min(miny, data[-1][-1])
                                except:
                                    self.log.setText("Invalid data - '" + column + "' (" + ss_col(c2) + ') row ' + str(row + 1) + " is '" + str(data[-1][-1]) + "'")
                                    return
                            break
            if tgt_col >= 0:
                for row in range(self.toprow[1] + 1, self.toprow[1] + self.rows + 1):
                    load.append(ws.cell_value(row, tgt_col))
                    maxy = max(maxy, load[-1])
            if len(overlay_cols) > 0:
                for row in range(self.toprow[1] + 1, self.toprow[1] + self.rows + 1):
                    overlay.append(0.)
                if self.overlay == 'Charge':
                    for row in range(self.toprow[1] + 1, self.toprow[1] + self.rows + 1):
                        for col in overlay_cols:
                            overlay[row - self.toprow[1] - 1] += ws.cell_value(row, col)
                else:
                    for row in range(self.toprow[1] + 1, self.toprow[1] + self.rows + 1):
                        overlay[row - self.toprow[1] - 1] = ws.cell_value(row, overlay_cols[-1])
                maxy = max(maxy, overlay[row - self.toprow[1] - 1])
            if self.plottype.currentText() == 'Line Chart':
                fig = plt.figure(figname, constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontsize=16)
                if gridtype != '':
                    plt.grid(axis=gridtype)
                lc1 = plt.subplot(111)
                plt.title(titl)
                for c in range(len(data)):
                    lc1.plot(x, data[c], linewidth=1.5, label=label[c], color=self.set_colour(label[c]))
                if len(load) > 0:
                    lc1.plot(x, load, linewidth=2.5, label=self.target, color=self.set_colour(self.target))
                if len(overlay) > 0:
                    lc1.plot(x, overlay, linewidth=1.5, label=self.overlay, color='black', linestyle='dotted')
                if self.maxSpin.value() > 0:
                    maxy = self.maxSpin.value()
                else:
                    try:
                        rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                        maxy = ceil(maxy / rndup) * rndup
                    except:
                        pass
             #   miny = 0
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                lc1.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 1),
                          prop=lbl_font)
                plt.ylim([miny, maxy])
                plt.xlim([0, len(x)])
                xticks = list(range(0, len(x), self.interval * days_per_label))
                lc1.set_xticks(xticks)
                lc1.set_xticklabels(day_labels[:len(xticks)], rotation='vertical')
                lc1.set_xlabel('Period')
                lc1.set_ylabel('Power (MW)') # MWh?
                zp = ZoomPanX()
                f = zp.zoom_pan(lc1, base_scale=1.2, flex_ticks=flex_on) # enable scrollable zoom
                plt.show()
                del zp
            elif self.plottype.currentText() in ['Cumulative', 'Step Chart']:
                if self.plottype.currentText() == 'Cumulative':
                    step = None
                else:
                    step = 'pre'
                fig = plt.figure(figname, constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontsize=16)
                if gridtype != '':
                    plt.grid(axis=gridtype)
                cu1 = plt.subplot(111)
                plt.title(titl)
                if self.percentage.isChecked():
                    totals = [0.] * len(x)
                    bottoms = [0.] * len(x)
                    values = [0.] * len(x)
                    for c in range(len(data)):
                        for h in range(len(data[c])):
                            totals[h] = totals[h] + data[c][h]
                    for h in range(len(data[0])):
                        values[h] = data[0][h] / totals[h] * 100.
                    cu1.fill_between(x, 0, values, label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]), step=step)
                    for c in range(1, len(data)):
                        for h in range(len(data[c])):
                            bottoms[h] = values[h]
                            values[h] = values[h] + data[c][h] / totals[h] * 100.
                        cu1.fill_between(x, bottoms, values, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                        step=step)
                    maxy = 100
                    cu1.set_ylabel('Power (%)')
                else:
                    if self.target == '<none>':
                        cu1.fill_between(x, 0, data[0], label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                        step=step)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                data[c][h] = data[c][h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                            cu1.fill_between(x, data[c - 1], data[c], label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                            step=step)
                        top = data[0][:]
                        for d in range(1, len(data)):
                            for h in range(len(top)):
                                top[h] = max(top[h], data[d][h])
                        if self.plottype.currentText() == 'Cumulative':
                            cu1.plot(x, top, color='white')
                        else:
                            cu1.step(x, top, color='white')
                    else:
                 #       pattern = ['-', '+', 'x', '\\', '|', '/', '*', 'o', 'O', '.']
                 #       pat = 0
                        full = []
                        for h in range(len(load)):
                           full.append(min(load[h], data[0][h]))
                        cu1.fill_between(x, 0, full, label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                        step=step)
                        for h in range(len(data[0])):
                            if data[0][h] > full[h]:
                                if self.spill_label.text() != '':
                                    cu1.fill_between(x, full, data[0], alpha=self.alpha, color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                        label=label[0] + ' ' + self.spill_label.text(), step=step)
                                else:
                                    cu1.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                    step=step)
                                break
                        for c in range(1, len(data)):
                            full = []
                            for h in range(len(data[c])):
                                data[c][h] = data[c][h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                                full.append(max(min(load[h], data[c][h]), data[c - 1][h]))
                            cu1.fill_between(x, data[c - 1], full, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                            step=step)
                            for h in range(len(data[c])):
                                if data[c][h] > full[h] + self.margin_of_error:
                                    if self.spill_label.text() != '':
                                        cu1.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                        label=label[c] + ' ' + self.spill_label.text(), step=step)
                                    else:
                                        cu1.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                        step=step)
                                    break
                        top = data[0][:]
                        for d in range(1, len(data)):
                            for h in range(len(top)):
                                top[h] = max(top[h], data[d][h])
                        if self.plottype.currentText() == 'Cumulative':
                            if self.alpha == 0:
                                cu1.plot(x, top, color='gray', linestyle='dashed')
                            else:
                                cu1.plot(x, top, color='gray')
                        else:
                            if self.alpha == 0:
                                cu1.step(x, top, color='gray', linestyle='dashed')
                            else:
                                cu1.step(x, top, color='gray')
                        short = []
                        do_short = False
                        c = len(data) - 1
                        for h in range(len(load)):
                            if load[h] > data[c][h] + self.margin_of_error:
                                do_short = True
                            short.append(max(data[c][h], load[h]))
                        if do_short:
                            cu1.fill_between(x, data[c], short, label='Shortfall', color=self.set_colour('shortfall'),
                                            step=step)
                        if self.plottype.currentText() == 'Cumulative':
                            cu1.plot(x, load, linewidth=2.5, label=self.target, color=self.set_colour(self.target))
                        else:
                            cu1.step(x, load, linewidth=2.5, label=self.target, color=self.set_colour(self.target))
                    if len(overlay) > 0:
                        if self.plottype.currentText() == 'Cumulative':
                            cu1.plot(x, overlay, linewidth=1.5, label=self.overlay, color='black', linestyle='dotted')
                        else:
                            cu1.step(x, overlay, linewidth=1.5, label=self.overlay, color='black', linestyle='dotted')
                    if self.maxSpin.value() > 0:
                        maxy = self.maxSpin.value()
                    else:
                        try:
                            rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                            maxy = ceil(maxy / rndup) * rndup
                        except:
                            pass
                    cu1.set_ylabel('Power (MW)')
        #        miny = 0
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                cu1.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=lbl_font)
                plt.ylim([miny, maxy])
                plt.xlim([0, len(x)])
                xticks = list(range(0, len(x), self.interval * days_per_label))
                cu1.set_xticks(xticks)
                cu1.set_xticklabels(day_labels[:len(xticks)], rotation='vertical')
                cu1.set_xlabel('Period')
                zp = ZoomPanX()
                f = zp.zoom_pan(cu1, base_scale=1.2, flex_ticks=flex_on) # enable scrollable zoom
                plt.show()
                del zp
            elif self.plottype.currentText() == 'Bar Chart':
                fig = plt.figure(figname, constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontsize=16)
                if gridtype != '':
                    plt.grid(axis=gridtype)
                bc1 = plt.subplot(111)
                plt.title(titl)
                if self.percentage.isChecked():
                    miny = 0
                    totals = [0.] * len(x)
                    bottoms = [0.] * len(x)
                    values = [0.] * len(x)
                    for c in range(len(data)):
                        for h in range(len(data[c])):
                            totals[h] = totals[h] + data[c][h]
                    for h in range(len(data[0])):
                        values[h] = data[0][h] / totals[h] * 100.
                    bc1.bar(x, values, label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]))
                    for c in range(1, len(data)):
                        for h in range(len(data[c])):
                            bottoms[h] = bottoms[h] + values[h]
                            values[h] = data[c][h] / totals[h] * 100.
                        bc1.bar(x, values, bottom=bottoms, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]))
                    maxy = 100
                    bc1.set_ylabel('Power (%)')
                else:
                    if self.target == '<none>':
                        bc1.bar(x, data[0], label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]))
                        bottoms = [0.] * len(x)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                bottoms[h] = bottoms[h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                            bc1.bar(x, data[c], bottom=bottoms, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]))
                    else:
                        bc1.bar(x, data[0], label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]))
                        bottoms = [0.] * len(x)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                bottoms[h] = bottoms[h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                            bc1.bar(x, data[c], bottom=bottoms, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]))
                        bc1.plot(x, load, linewidth=2.5, label=self.target, color=self.set_colour(self.target))
                        bc1.plot(x, overlay, linewidth=1.5, label=self.overlay, color='black', linestyle='dotted')
                    if self.maxSpin.value() > 0:
                        maxy = self.maxSpin.value()
                    else:
                        try:
                            rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                            maxy = ceil(maxy / rndup) * rndup
                        except:
                            pass
                    bc1.set_ylabel('Power (MW)')
        #        miny = 0
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                bc1.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=lbl_font)
                plt.ylim([miny, maxy])
                plt.xlim([0, len(x)])
                xticks = list(range(0, len(x), self.interval * days_per_label))
                bc1.set_xticks(xticks)
                bc1.set_xticklabels(day_labels[:len(xticks)], rotation='vertical')
                bc1.set_xlabel('Period')
                zp = ZoomPanX()
                f = zp.zoom_pan(bc1, base_scale=1.2, flex_ticks=flex_on) # enable scrollable zoom
                plt.show()
                del zp
            elif self.plottype.currentText() == 'Heat Map':
                fig = plt.figure(figname)
                if suptitle != '':
                    fig.suptitle(suptitle, fontsize=16)
                if gridtype != '':
                    plt.grid(axis=gridtype)
                hm1 = plt.subplot(111)
                plt.title(titl)
                hmdata = []
                for hr in range(self.interval):
                    hmdata.append([])
                    for dy in range(365):
                        hmdata[-1].append(0)
                for col in range(len(data)):
                    x = 0
                    y = -1
                    for hr in range(len(data[col])):
                        y += 1
                        if y > self.interval - 1:
                            x += 1
                            y = 0
                        hmdata[y][x] += data[col][hr]
                if tgt_col >= 0:
                    x = 0
                    y = -1
                    for hr in range(len(load)):
                        y += 1
                        if y > self.interval - 1:
                            x += 1
                            y = 0
                        hmdata[y][x] = hmdata[y][x] / load[hr]
                miny = 999999
                maxy = 0
                for y in range(len(hmdata)):
                    for x in range(len(hmdata[0])):
                        miny = min(miny, hmdata[y][x])
                        maxy = max(maxy, hmdata[y][x])
                if tgt_col >= 0:
                    fmt_str = '{:.3f}'
                    vmax = 1
                    vmin = 0
                    if maxy > 1:
                        maxy = 1
                else:
                    fmt_str = '{:,.0f}'
                    if self.cbar2:
                        vmax = maxy
                        vmin = miny
                    else:
                        vmax = None
                        vmin = None
                im1 = hm1.imshow(hmdata, cmap=self.cmap, interpolation='nearest', aspect='auto', vmax=vmax, vmin=vmin)
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                day = -0.5
                xticks = [day]
                for m in range(len(the_days) - 1):
                    day += the_days[m]
                    xticks.append(day)
                hm1.set_xticks(xticks)
                hm1.set_xticklabels(mth_labels, ha='left')
                yticks = [-0.5]
                if self.interval == 24:
                    for y in range(0, 24, 4):
                        yticks.append(y + 2.5)
                else:
                    for y in range(0, 48, 8):
                        yticks.append(y + 6.5)
                hm1.invert_yaxis()
                hm1.set_yticks(yticks)
                hm1.set_yticklabels(hr_labels, va='bottom')
                hm1.set_xlabel('Period')
                hm1.set_ylabel('Hour')
                if self.cbar:
                    fig.subplots_adjust(bottom=0.1, right=0.8, top=0.9)
                    cb_ax = fig.add_axes([0.85, 0.1, 0.025, 0.8])
                    cbar = fig.colorbar(im1, cax=cb_ax)
                    if self.cbar2:
                        curcticks = cbar.get_ticks()
                        cticks = [vmin, vmax]
                        if miny != vmin:
                            cticks.append(miny)
                        if maxy != vmax:
                            cticks.append(maxy)
                        for y in range(len(curcticks)):
                            if curcticks[y] > vmin and curcticks[y] < vmax:
                                cticks.append(curcticks[y])
                        cticks.sort()
                        cbar.set_ticks(cticks)
                self.log.setText('Heat map value range: {} to {}'.format(fmt_str, fmt_str).format(miny, maxy))
                QtCore.QCoreApplication.processEvents()
                plt.show()
        else: # diurnal average
            if self.interval == 24:
                ticks = list(range(0, 21, 4))
                ticks.append(23)
            else:
                ticks = list(range(0, 41, 8))
                ticks.append(47)
            titl = self.title.text()
            titl = self.replace_words('s', self.title.text(), isheet)
            if suptitle != '':
                suptitle = self.replace_words('s', suptitle, isheet)
            if self.select_day and self.aday.currentText() != '':
                if self.period.currentText() == 'Year':
                    i = 0
                else:
                    try:
                        i = self.seasons[self.period.currentText()][0]
                    except:
                        i = mth_labels.index(self.period.currentText())
                if self.aday.currentIndex() > the_days[i]:
                    self.aday.setCurrentIndex(self.aday.currentIndex() - the_days[i])
                    i += 1
                self.period.setCurrentIndex(i + 2)
                d = 0
                for m in range(i):
                    d += the_days[m]
                d += self.aday.currentIndex() - 1
                strt_row = [d * self.interval + self.toprow[1]]
                todo_rows = [self.interval]
                y = self.replace_words('y', self.title.text(), 'find')
                m = self.replace_words('m', self.title.text(), 'find')
                if y > m:
                    txt = self.aday.currentText() + ' ' + self.period.currentText()
                else:
                    txt = self.period.currentText() + ' ' + self.aday.currentText()
                titl = self.replace_words('y', self.title.text(), str(year))
                titl = self.replace_words('m', titl, txt)
                if suptitle != '':
                    suptitle = self.replace_words('y', suptitle, str(year))
                    suptitle = self.replace_words('m', suptitle, txt)
            else:
                titl = self.replace_words('y', self.title.text(), str(year))
                if suptitle != '':
                    suptitle = self.replace_words('y', suptitle, str(year))
                todo_mths = []
                if self.period.currentText() == 'Year':
                    titl = self.replace_words('m', titl, '')
                    titl = titl.replace('  ', ' ')
                    if suptitle != '':
                        suptitle = self.replace_words('m', suptitle, '')
                        suptitle = suptitle.replace('  ', ' ')
                    strt_row = [self.toprow[1]]
                    todo_rows = [self.rows]
                else:
                    strt_row = []
                    todo_rows = []
                    if self.period.currentText() in self.seasons.keys():
                        titl = self.replace_words('m', titl, self.period.currentText())
                        if suptitle != '':
                            suptitle = self.replace_words('m', suptitle, '')
                        for s in self.seasons[self.period.currentText()]:
                            m = 0
                            strt_row.append(0)
                            while m < s:
                                strt_row[-1] = strt_row[-1] + the_days[m] * self.interval
                                m += 1
                            todo_mths.append(m)
                            strt_row[-1] = strt_row[-1] + self.toprow[1]
                            todo_rows.append(the_days[s] * self.interval)
                    else:
                         i = mth_labels.index(self.period.currentText())
                         todo_mths = [i]
                         if self.cperiod.currentText() == '<none>':
                             titl = self.replace_words('m', titl, self.period.currentText())
                             if suptitle != '':
                                 suptitle = self.replace_words('m', suptitle, self.period.currentText())
                         else:
                             titl = self.replace_words('m', titl, self.period.currentText() + \
                                    ' to ' + self.cperiod.currentText())
                             if suptitle != '':
                                 suptitle = self.replace_words('m', suptitle, self.period.currentText() + \
                                            ' to ' + self.cperiod.currentText())
                             j = mth_labels.index(self.cperiod.currentText())
                             if j == i:
                                 pass
                             elif j > i:
                                 for k in range(i + 1, j + 1):
                                     todo_mths.append(k)
                             else:
                                 for k in range(i + 1, 12):
                                     todo_mths.append(k)
                                 for k in range(j + 1):
                                     todo_mths.append(k)
                         for s in todo_mths:
                             m = 0
                             strt_row.append(0)
                             while m < s:
                                 strt_row[-1] = strt_row[-1] + the_days[m] * self.interval
                                 m += 1
                             strt_row[-1] = strt_row[-1] + self.toprow[1]
                             todo_rows.append(the_days[s] * self.interval)
            load = []
            tgt_col = -1
            data = []
            label = []
            if self.plottype.currentText() == 'Heat Map':
                fig = plt.figure(figname)
                if suptitle != '':
                    fig.suptitle(suptitle, fontsize=16)
                if gridtype != '':
                    plt.grid(axis=gridtype)
                hm2 = plt.subplot(111)
                plt.title(titl)
                hmdata = []
                days = int(sum(todo_rows) / self.interval)
                for hr in range(self.interval):
                    hmdata.append([])
                    for dy in range(days):
                        hmdata[-1].append(0)
                for c2 in range(2, ws.ncols):
                    try:
                        column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                    except:
                        column = str(ws.cell_value(self.toprow[0], c2))
                    if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                        column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                    if column == self.target:
                        tgt_col = c2
                for c in range(self.order.count() -1, -1, -1):
                    col = self.order.item(c).text()
                    for c2 in range(2, ws.ncols):
                        try:
                            column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                        except:
                            column = str(ws.cell_value(self.toprow[0], c2))
                        if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                            column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                        if column == col:
                            dy = 0
                            hr = -1
                            for s in range(len(strt_row)):
                                for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                    hr += 1
                                    if hr > self.interval - 1:
                                        dy += 1
                                        hr = 0
                                    try:
                                        hmdata[hr][dy] += ws.cell_value(row, c2)
                                    except:
                                        pass
                            break
                if tgt_col >= 0:
                    dy = 0
                    hr = -1
                    for s in range(len(strt_row)):
                        for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                            hr += 1
                            if hr > self.interval - 1:
                                dy += 1
                                hr = 0
                            try:
                                hmdata[hr][dy] = hmdata[hr][dy] / ws.cell_value(row, tgt_col)
                            except:
                                pass
                miny = 999999
                maxy = 0
                tgap = 0
                for y in range(len(hmdata)):
                    for x in range(len(hmdata[0])):
                        miny = min(miny, hmdata[y][x])
                        maxy = max(maxy, hmdata[y][x])
                if tgt_col >= 0:
                    fmt_str = '{:.3f}'
                    vmax = 1
                    vmin = 0
                    if self.cbar2:
                        tgap = .03
                    if maxy > 1:
                        maxy = 1
                else:
                    fmt_str = '{:,.0f}'
                    if self.cbar2:
                        tgap = (maxy - miny) * .03
                        vmax = maxy
                        vmin = miny
                    else:
                        vmax = None
                        vmin = None
                im2 = hm2.imshow(hmdata, cmap=self.cmap, interpolation='nearest', aspect='auto', vmax=vmax, vmin=vmin)
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                yticks = [-0.5]
                if self.interval == 24:
                    for y in range(0, 24, 4):
                        yticks.append(y + 2.5)
                else:
                    for y in range(0, 48, 8):
                        yticks.append(y + 6.5)
                hm2.invert_yaxis()
                hm2.set_yticks(yticks)
                hm2.set_yticklabels(hr_labels, va='bottom')
                if len(hmdata[0]) > 31:
                    day = -0.5
                    xticks = []
                    xticklbls = []
                    for m in todo_mths:
                        xticks.append(day)
                        xticklbls.append(mth_labels[m])
                        day = day + the_days[m]
                    hm2.set_xticks(xticks)
                    hm2.set_xticklabels(xticklbls, ha='left')
                else:
                    curxticks = hm2.get_xticks()
                    xticks = []
                    xticklabels = []
                    for x in range(len(curxticks)):
                        if curxticks[x] >= 0 and curxticks[x] < len(hmdata[0]):
                            xticklabels.append(str(int(curxticks[x] + 1)))
                            xticks.append(curxticks[x])
                    hm2.set_xticks(xticks)
                    hm2.set_xticklabels(xticklabels)
                hm2.set_xlabel('Day')
                hm2.set_ylabel('Hour')
                if self.cbar:
                    fig.subplots_adjust(bottom=0.1, right=0.8, top=0.9)
                    cb_ax = fig.add_axes([0.85, 0.1, 0.025, 0.8])
                    cbar = fig.colorbar(im2, cax=cb_ax)
                    if self.cbar2:
                        curcticks = cbar.get_ticks()
                        cticks = [vmin, vmax]
                        if miny != vmin:
                            cticks.append(miny)
                        if maxy != vmax:
                            cticks.append(maxy)
                        for y in range(len(curcticks)):
                            if curcticks[y] < vmin or curcticks[y] > vmax:
                                continue
                            if self.cbar2 and tgap != 0:
                                if miny - tgap < curcticks[y] < miny + tgap:
                                    continue
                                if maxy - tgap < curcticks[y] < maxy + tgap:
                                    continue
                            cticks.append(curcticks[y])
                        cticks.sort()
                        cbar.set_ticks(cticks)
                self.log.setText('Heat map value range: {} to {}'.format(fmt_str, fmt_str).format(miny, maxy))
                QtCore.QCoreApplication.processEvents()
                plt.show()
                return
            overlay_cols = []
            overlay = []
            miny = 0
            maxy = 0
            hs = []
            for h in range(self.interval):
                hs.append(h)
            x = hs[:]
            for c2 in range(2, ws.ncols):
                try:
                    column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                except:
                    column = str(ws.cell_value(self.toprow[0], c2))
                if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                    column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                if column == self.target:
                    tgt_col = c2
                if self.overlay != '<none>':
                    if column == self.target or column[:len(self.overlay)] == self.overlay:
                        overlay_cols.append(c2)
            if len(breakdowns) > 0:
                for c in range(self.order.count() -1, -1, -1):
                    col = self.order.item(c).text()
                    for c2 in range(2, ws.ncols):
                        try:
                            brkdown = ws.cell_value(self.breakdown_row, c2)
                            if brkdown is not None and brkdown != '' and brkdown != breakdowns[0]:
                                continue
                        except:
                            pass
                        try:
                            column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                        except:
                            column = str(ws.cell_value(self.toprow[0], c2))
                        if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                            column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                        if column == col:
                            data.append([])
                            data[-1] = [0] * len(hs)
                            label.append(column)
                            tot_rows = 0
                            for s in range(len(strt_row)):
                                h = 0
                                for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                    try:
                                        data[-1][h] = data[-1][h] + ws.cell_value(row, c2)
                                    except:
                                        break # part period
                                    h += 1
                                    if h >= self.interval:
                                        h = 0
                                tot_rows += todo_rows[s]
                            for h in range(self.interval):
                                data[-1][h] = data[-1][h] / (tot_rows / self.interval)
                                maxy = max(maxy, data[-1][h])
                                miny = min(miny, data[-1][h])
                            break
                for breakdown in breakdowns[1:]:
                    for c in range(self.order.count() -1, -1, -1):
                        col = self.order.item(c).text()
                        for c2 in range(2, ws.ncols):
                            try:
                                brkdown = ws.cell_value(self.breakdown_row, c2)
                                if brkdown != breakdown:
                                    continue
                            except:
                                continue
                            try:
                                column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                            except:
                                column = str(ws.cell_value(self.toprow[0], c2))
                            if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                                column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                            if column == col:
                                data.append([])
                                data[-1] = [0] * len(hs)
                                label.append(column + ' ' + breakdown)
                                tot_rows = 0
                                for s in range(len(strt_row)):
                                    h = 0
                                    for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                        try:
                                            data[-1][h] = data[-1][h] + ws.cell_value(row, c2)
                                        except:
                                            break # part period
                                        h += 1
                                        if h >= self.interval:
                                            h = 0
                                    tot_rows += todo_rows[s]
                                for h in range(self.interval):
                                    data[-1][h] = data[-1][h] / (tot_rows / self.interval)
                                    maxy = max(maxy, data[-1][h])
                                    miny = min(miny, data[-1][h])
                                break
            else:
                for c in range(self.order.count() -1, -1, -1):
                    col = self.order.item(c).text()
                    for c2 in range(2, ws.ncols):
                        try:
                            column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                        except:
                            column = str(ws.cell_value(self.toprow[0], c2))
                        if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                            column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                        if column == col:
                            data.append([])
                            data[-1] = [0] * len(hs)
                            label.append(column)
                            tot_rows = 0
                            for s in range(len(strt_row)):
                                h = 0
                                for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                    try:
                                        data[-1][h] = data[-1][h] + ws.cell_value(row, c2)
                                    except:
                                        break # part period
                                    h += 1
                                    if h >= self.interval:
                                        h = 0
                                tot_rows += todo_rows[s]
                            for h in range(self.interval):
                                data[-1][h] = data[-1][h] / (tot_rows / self.interval)
                                maxy = max(maxy, data[-1][h])
                                miny = min(miny, data[-1][h])
                            break
            if tgt_col >= 0:
                load = [0] * len(hs)
                tot_rows = 0
                for s in range(len(strt_row)):
                    h = 0
                    for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                        load[h] = load[h] + ws.cell_value(row, tgt_col)
                        h += 1
                        if h >= self.interval:
                           h = 0
                    tot_rows += todo_rows[s]
                for h in range(self.interval):
                    load[h] = load[h] / (tot_rows / self.interval)
                    maxy = max(maxy, load[h])
            if len(overlay_cols) > 0:
                if self.overlay == 'Underlying Load':
                    overlay_cols = [overlay_cols[-1]]
                overlay = [0] * len(hs)
                tot_rows = 0
                for s in range(len(strt_row)):
                    h = 0
                    for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                        for col in overlay_cols:
                            overlay[h] += ws.cell_value(row, col)
                        h += 1
                        if h >= self.interval:
                           h = 0
                    tot_rows += todo_rows[s]
                for h in range(self.interval):
                    overlay[h] = overlay[h] / (tot_rows / self.interval)
                    maxy = max(maxy, overlay[h])
            if self.plottype.currentText() == 'Line Chart':
                fig = plt.figure(figname + '_' + self.period.currentText().lower(),
                                 constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontsize=16)
                if gridtype != '':
                    plt.grid(axis=gridtype)
                lc2 = plt.subplot(111)
                plt.title(titl)
                for c in range(len(data)):
                    lc2.plot(x, data[c], linewidth=1.5, label=label[c], color=self.set_colour(label[c]))
                if len(load) > 0:
                    lc2.plot(x, load, linewidth=2.5, label=self.target, color=self.set_colour(self.target))
                if len(overlay) > 0:
                    lc2.plot(x, overlay, linewidth=1.5, label=self.overlay, color='black', linestyle='dotted')
                if self.maxSpin.value() > 0:
                    maxy = self.maxSpin.value()
                else:
                    try:
                        rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                        maxy = ceil(maxy / rndup) * rndup
                    except:
                        pass
           #     miny = 0
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                lc2.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 1),
                              prop=lbl_font)
                plt.ylim([miny, maxy])
                plt.xlim([0, self.interval - 1])
                lc2.set_xticks(ticks)
                lc2.set_xticklabels(hr_labels)
                lc2.set_xlabel('Hour of the Day')
                lc2.set_ylabel('Power (MW)')
                zp = ZoomPanX()
                f = zp.zoom_pan(lc2, base_scale=1.2) # enable scrollable zoom
                plt.show()
                del zp
            elif self.plottype.currentText() in ['Cumulative', 'Step Chart']:
                if self.plottype.currentText() == 'Cumulative':
                    step = None
                else:
                    step = 'pre' # 'post'
                fig = plt.figure(figname + '_' + self.period.currentText().lower(),
                                 constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontsize=16)
                if gridtype != '':
                    plt.grid(axis=gridtype)
                cu2 = plt.subplot(111)
                plt.title(titl)
                if self.percentage.isChecked():
                    totals = [0.] * len(x)
                    bottoms = [0.] * len(x)
                    values = [0.] * len(x)
                    for c in range(len(data)):
                        for h in range(len(data[c])):
                            totals[h] = totals[h] + data[c][h]
                    for h in range(len(data[0])):
                        values[h] = data[0][h] / totals[h] * 100.
                    cu2.fill_between(x, 0, values, label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]), step=step)
                    for c in range(1, len(data)):
                        for h in range(len(data[c])):
                            bottoms[h] = values[h]
                            values[h] = values[h] + data[c][h] / totals[h] * 100.
                        cu2.fill_between(x, bottoms, values, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]), step=step)
                    maxy = 100
                    cu2.set_ylabel('Power (%)')
                else:
                    if self.target == '<none>':
                        cu2.fill_between(x, 0, data[0], label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]), step=step)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                data[c][h] = data[c][h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                            cu2.fill_between(x, data[c - 1], data[c], label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]), step=step)
                    else:
                #        pattern = ['-', '+', 'x', '\\', '*', 'o', 'O', '.']
                #        pat = 0
                        full = []
                        for h in range(len(load)):
                           full.append(min(load[h], data[0][h]))
                        cu2.fill_between(x, 0, full, label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]), step=step)
                        for h in range(len(full)):
                            if data[0][h] > full[h]:
                                if self.spill_label.text() != '':
                                    cu2.fill_between(x, full, data[0], alpha=self.alpha, color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                        label=label[0] + ' ' + self.spill_label.text(), step=step)
                                else:
                                    cu2.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]), step=step)
                                break
                        for c in range(1, len(data)):
                            full = []
                            for h in range(len(data[c])):
                                data[c][h] = data[c][h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                                full.append(max(min(load[h], data[c][h]), data[c - 1][h]))
                            cu2.fill_between(x, data[c - 1], full, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]), step=step)
                 #           pat += 1
                 #           if pat >= len(pattern):
                 #               pat = 0
                            for h in range(len(full)):
                                if data[c][h] > full[h] + self.margin_of_error:
                                    if self.spill_label.text() != '':
                                        cu2.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                        label=label[c] + ' ' + self.spill_label.text(), step=step)
                                    else:
                                        cu2.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                        step=step)
                                    break
                        top = data[0][:]
                        for d in range(1, len(data)):
                            for h in range(len(top)):
                                top[h] = max(top[h], data[d][h])
                        if self.plottype.currentText() == 'Cumulative':
                            if self.alpha == 0:
                                cu2.plot(x, top, color='gray', linestyle='dashed')
                            else:
                                cu2.plot(x, top, color='gray')
                        else:
                            if self.alpha == 0:
                                cu2.step(x, top, color='gray', linestyle='dashed')
                            else:
                                cu2.step(x, top, color='gray')
                        short = []
                        do_short = False
                        for h in range(len(load)):
                            if load[h] > data[c][h] + self.margin_of_error:
                                do_short = True
                            short.append(max(data[c][h], load[h]))
                        if do_short:
                            cu2.fill_between(x, data[c], short, label='Shortfall', color=self.set_colour('shortfall'), step=step)
                        if self.plottype.currentText() == 'Cumulative':
                            cu2.plot(x, load, linewidth=2.5, label=self.target, color=self.set_colour(self.target))
                        else:
                            cu2.step(x, load, linewidth=2.5, label=self.target, color=self.set_colour(self.target))
                    if len(overlay) > 0:
                        if self.plottype.currentText() == 'Cumulative':
                            cu2.plot(x, overlay, linewidth=1.5, label=self.overlay, color='black', linestyle='dotted')
                        else:
                            cu2.step(x, overlay, linewidth=1.5, label=self.overlay, color='black', linestyle='dotted')
                    if self.maxSpin.value() > 0:
                        maxy = self.maxSpin.value()
                    else:
                        try:
                            rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                            maxy = ceil(maxy / rndup) * rndup
                        except:
                            pass
                    cu2.set_ylabel('Power (MW)')
         #       miny = 0
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                cu2.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=lbl_font)
                plt.ylim([miny, maxy])
                plt.xlim([0, self.interval - 1])
                cu2.set_xticks(ticks)
                cu2.set_xticklabels(hr_labels)
                cu2.set_xlabel('Hour of the Day')
                zp = ZoomPanX()
                f = zp.zoom_pan(cu2, base_scale=1.2) # enable scrollable zoom
                plt.show()
                del zp
            elif self.plottype.currentText() == 'Bar Chart':
                fig = plt.figure(figname + '_' + self.period.currentText().lower(),
                                 constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontsize=16)
                if gridtype != '':
                    plt.grid(axis=gridtype)
                bc2 = plt.subplot(111)
                plt.title(titl)
                if self.percentage.isChecked():
                    miny = 0
                    totals = [0.] * len(x)
                    bottoms = [0.] * len(x)
                    values = [0.] * len(x)
                    for c in range(len(data)):
                        for h in range(len(data[c])):
                            totals[h] = totals[h] + data[c][h]
                    for h in range(len(data[0])):
                        values[h] = data[0][h] / totals[h] * 100.
                    bc2.bar(x, values, label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]))
                    for c in range(1, len(data)):
                        for h in range(len(data[c])):
                            bottoms[h] = bottoms[h] + values[h]
                            values[h] = data[c][h] / totals[h] * 100.
                        bc2.bar(x, values, bottom=bottoms, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]))
                    maxy = 100
                    bc2.set_ylabel('Power (%)')
                else:
                    if self.target == '<none>':
                        bc2.bar(x, data[0], label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]))
                        bottoms = [0.] * len(x)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                bottoms[h] = bottoms[h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h] + bottoms[h])
                            bc2.bar(x, data[c], bottom=bottoms, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]))
                    else:
                        bc2.bar(x, data[0], label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]))
                        bottoms = [0.] * len(x)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                bottoms[h] = bottoms[h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h] + bottoms[h])
                            bc2.bar(x, data[c], bottom=bottoms, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]))
                        bc2.plot(x, load, linewidth=2.5, label=self.target, color=self.set_colour(self.target))
                    if len(overlay) > 0:
                        bc2.plot(x, overlay, linewidth=1.5, label=self.overlay, color='black', linestyle='dotted')
                    if self.maxSpin.value() > 0:
                        maxy = self.maxSpin.value()
                    else:
                        try:
                            rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                            maxy = ceil(maxy / rndup) * rndup
                        except:
                            pass
                    bc2.set_ylabel('Power (MW)')
        #        miny = 0
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                bc2.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=lbl_font)
                plt.ylim([miny, maxy])
                plt.xlim([0, self.interval - 1])
                bc2.set_xticks(ticks)
                bc2.set_xticklabels(hr_labels)
                bc2.set_xlabel('Hour of the Day')
                zp = ZoomPanX()
                f = zp.zoom_pan(bc2, base_scale=1.2) # enable scrollable zoom
                plt.show()
                del zp

    def corrClicked(self):
        if self.show_contribution:
            corrtxt = ['month', 'contribution']
            sortby = ''
        else:
            corrtxt = ['coefficient', 'correlation']
            sortby = corrtxt[0]
        if self.book is None:
            self.log.setText('Error accessing Workbook.')
            return
        if self.order.count() == 0:
            self.log.setText('No columns chosen.')
            return
        isheet = self.sheet.currentText()
        if isheet == '':
            self.log.setText('Sheet not set.')
            return
        if self.target == '<none>':
            self.log.setText('Target not set.')
            return
        ws = self.book.sheet_by_name(isheet)
        if self.toprow is None:
            tech_row = -1
            row = 0
            self.zone_row =  -1
            self.breakdown_row = -1
            while row < ws.nrows:
                if ws.cell_value(row, 0) in ['Split', 'Breakdown']:
                    self.breakdown_row = row
                elif ws.cell_value(row, 0) == 'Technology':
                    tech_row = row
                elif ws.cell_value(row, 0) == 'Zone':
                    self.zone_row = row
                elif ws.cell_value(row, 0) in ['Hour', 'Interval', 'Trading Interval']:
                    if ws.cell_value(row, 1) != 'Period':
                        self.log.setText(isheet + ' sheet format incorrect')
                        return
                    if tech_row >= 0:
                        self.toprow = [tech_row, row]
                    else:
                        self.toprow = [row, row]
                    self.rows = ws.nrows - (row + 1)
                    break
        try:
            year = int(ws.cell_value(self.toprow[1] + 1, 1)[:4])
            if year % 4 == 0 and year % 100 != 0 or year % 400 == 0:
                self.leapyear = True
            else:
                self.leapyear = False
        except:
            self.leapyear = False
            year = ''
        the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if self.leapyear: #rows == 8784: # leap year
            the_days[1] = 29
        mth_labels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        periods = []
        objects = []
        per1 = self.period.currentText()
        per2 = self.cperiod.currentText()
        if self.show_contribution and per1 in  ['<none>', 'Year']:
            per1 = 'Jan'
            per2 = 'Dec'
        if per1 in ['<none>', 'Year']: # full year of hourly figures
            strt_row = [self.toprow[1]]
            todo_rows = [self.rows]
            per2 = ''
            if self.show_contribution:
                periods = mth_labels[:]
        else:
            strt_row = []
            todo_rows = []
            if per1 in self.seasons.keys():
                strt_row = []
                todo_rows = []
                for s in self.seasons[per1]:
                    m = 0
                    strt_row.append(0)
                    while m < s:
                        strt_row[-1] = strt_row[-1] + the_days[m] * self.interval
                        m += 1
                    strt_row[-1] = strt_row[-1] + self.toprow[1]
                    todo_rows.append(the_days[s] * self.interval)
                    if self.show_contribution:
                        periods.append(mth_labels[m])
                per2 = ''
            else:
                i = mth_labels.index(per1)
                todo_mths = [i]
                if per2 != '<none>':
                    j = mth_labels.index(per2)
                    if j == i:
                        pass
                    elif j > i:
                        for k in range(i + 1, j + 1):
                            todo_mths.append(k)
                    else:
                        for k in range(i + 1, 12):
                            todo_mths.append(k)
                        for k in range(j + 1):
                            todo_mths.append(k)
                for s in todo_mths:
                    m = 0
                    strt_row.append(0)
                    while m < s:
                        strt_row[-1] = strt_row[-1] + the_days[m] * self.interval
                        m += 1
                    strt_row[-1] = strt_row[-1] + self.toprow[1]
                    todo_rows.append(the_days[s] * self.interval)
                    if self.show_contribution:
                        periods.append(mth_labels[m])
        tgt = []
        tgt_col = -1
        data = []
        best_corr = 0
        best_fac = ''
        total_tgt = 0
        total_data = []
        for c2 in range(2, ws.ncols):
            try:
                column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
            except:
                column = str(ws.cell_value(self.toprow[0], c2))
            if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
            if column == self.target:
                tgt_col = c2
        if tgt_col < 0:
            self.log.setText('Target not found.')
            return
        base_tgt = []
        for row in range(self.toprow[1] + 1, self.toprow[1] + self.rows + 1):
            base_tgt.append(ws.cell_value(row, tgt_col))
        table = []
        if self.show_contribution:
            fields = ['month']
            decpts = [0]
        else:
            fields = ['aggregation', 'object'] + corrtxt
            decpts = [0, 0, 4]
        rows = []
        if self.rows > 8784:
            mult = 2
        else:
            mult = 1
        if self.show_contribution:
            aggregates = {'Month': 24 * 31 * mult}
        else:
            aggregates = {'Hour': 1, 'Day': 24 * mult, 'Week': 24 * 7 * mult,
                          'Fortnight': 24 * 14 * mult, 'Month': 24 * 31 * mult}
        debug = False
        if debug:
            cf = open('debug_corr.csv', 'w')
        for key, agg in aggregates.items():
            tgt = []
            tagg = 0
            atgt = 0
            for s in range(len(strt_row)):
                strt = strt_row[s] - self.toprow[1]
                for row in range(strt, strt + todo_rows[s]):
                    atgt += base_tgt[row]
                    tagg += 1
                    if tagg == agg:
                        tgt.append(atgt) # / tagg)
                        tagg = 0
                        atgt = 0
                if tagg > 0:
                    tgt.append(atgt) # / tagg)
                    tagg = 0
                    atgt = 0
            if not self.show_contribution and len(tgt) < 2:
                continue
            data = []
            for c in range(self.order.count() -1, -1, -1):
                col = self.order.item(c).text()
                for c2 in range(2, ws.ncols):
                    try:
                        column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                    except:
                        column = str(ws.cell_value(self.toprow[0], c2))
                    if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                        column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                    if column == col:
                        data.append([])
                        total_data.append(0.)
                        sagg = 0
                        asrc = 0
                        for s in range(len(strt_row)):
                            for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                asrc += ws.cell_value(row, c2)
                                sagg += 1
                                if sagg == agg:
                                    data[-1].append(asrc) # / tagg)
                                    sagg = 0
                                    asrc = 0
                            if sagg > 0:
                                data[-1].append(asrc) # / tagg)
                                sagg = 0
                                asrc = 0
                        if self.show_contribution:
                            fields.append(col.lower())
                            decpts.append(4)
                            continue
                            tgt_total = 0
                            src_total = 0
                            for value in tgt:
                                tgt_total += value
                            for value in data[-1]:
                                src_total += value
                            try:
                                corr = src_total / tgt_total
                            except:
                                corr = 0
                        else:
                            try:
                                corr = np.corrcoef(tgt, data[-1])
                                if np.isnan(corr.item((0, 1))):
                                    cor = 'n/a'
                                    corr = 0
                            except:
                                cor = 'n/a'
                                corr = 0
                            else:
                                corr = corr.item((0, 1))
                        if corr == 0:
                            cor = 'n/a'
                        elif abs(corr) < 0.1:
                            cor = 'None'
                        elif abs(corr) < 0.3:
                            cor = 'Little if any'
                        elif abs(corr) < 0.5:
                            cor = 'Low'
                        elif abs(corr) < 0.7:
                            cor = 'Moderate'
                        elif abs(corr) < 0.9:
                            cor = 'High'
                        else:
                            cor = 'Very high'
                        if abs(corr) > best_corr:
                            best_corr = corr
                            best_fac = key + ': ' + self.order.item(c).text()
                        if self.show_contribution:
                            rows.append([self.order.item(c).text(), corr, cor])
                        else:
                            rows.append([key, self.order.item(c).text(), corr, cor])
                       #     break
            if debug:
                for i in range(len(tgt)):
                    line = '%s,%d,%0.2f,' % (key, agg, tgt[i])
                    for d in range(len(data)):
                        line += '%0.2f,' % data[d][i]
                    cf.write(line + '\n')
        if self.show_contribution:
            for m in range(len(periods)):
                if debug:
                    line = periods[m] + ',' + str(tgt[m]) + ','
                total_tgt += tgt[m]
                rows.append([periods[m]])
                for o in range(len(fields) - 1):
                    if debug:
                        line += str(data[o][m]) + ','
                    total_data[o] += data[o][m]
                    pct = '{:.1f}%'.format(data[o][m] * 100. / tgt[m])
                    rows[-1].append(pct)
                if debug:
                    cf.write(line + '\n')
            rows.append(['Total'])
            for o in range(len(fields) - 1):
                pct = '{:.1f}%'.format(total_data[o] * 100. / total_tgt)
                rows[-1].append(pct)
        if debug:
            cf.close()
        title = '%s %s %s' % (corrtxt[1].title(), self.target, per1)
        if per2 != '' and per2 != '<none>':
            title += ' to ' + per2
        dialog = Table(rows, fields=fields, title=title, decpts=decpts, sortby=sortby, reverse=True, save_folder=self.scenarios)
        dialog.exec()
        self.log.setText('Best %s: %s %.4f' % (corrtxt[1], best_fac, best_corr))


if "__main__" == __name__:
    app = QtWidgets.QApplication(sys.argv)
    ex = PowerPlot()
    app.exec()
    app.deleteLater()
    sys.exit()
