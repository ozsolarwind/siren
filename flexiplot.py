#!/usr/bin/python3
#
#  Copyright (C) 2020-2025 Sustainable Energy Now Inc., Angus King
#
#  flexiplot.py - This file is possibly part of SIREN.
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
import subprocess
import sys
from math import log10, ceil, sqrt
import matplotlib
if matplotlib.__version__ > '3.5.1':
    matplotlib.use('Qt5Agg')
else:
    matplotlib.use('TkAgg')
from matplotlib.font_manager import FontProperties
import matplotlib.pyplot as plt
import random
import displayobject
import displaytable
from colours import Colours, PlotPalette
from credits import fileVersion
from editini import EdtDialog, SaveIni
from getmodels import getModelFile
from powerplot import MyQDialog, ChangeFontProp
try:
    from senplot3d import TablePlot3D
except:
    TablePlot3D = None
from senutils import ClickableQLabel, getParents, getUser, ListWidget, ssCol, strSplit, techClean, WorkBook
from zoompan import ZoomPanX

col_letters = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ'

def font_props(fontin, fontdict=True):
    font_dict = {'family': None, 'style': None, 'variant': None, 'weight': None,
                 'stretch': None, 'size': None, 'fname': None}
    bits = fontin.split(',')
    for bit in bits:
        value = bit.split('=')
        try:
            font_dict[value[0].lower()] = value[1]
        except:
            pass
    if fontdict:
        del font_dict['fname']
        return font_dict
    else:
        return FontProperties(family=font_dict['family'], style=font_dict['style'],
                              variant=font_dict['variant'], weight=font_dict['weight'],
                              stretch=font_dict['stretch'], size=font_dict['size'],
                              fname=font_dict['fname'])


class CustomCombo(QtWidgets.QComboBox):
    def __init__(self, parent=None):
        self.last_key = ''
        super().__init__(parent)

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Delete:
            if self.last_key == QtCore.Qt.Key_Shift:
                self.removeItem(self.currentIndex())
            self.last_key = event.key()
        else:
            self.last_key = event.key()
            QtWidgets.QComboBox.keyPressEvent(self, event)

class FlexiPlot(QtWidgets.QWidget):

    def set_fontdict(self, item=None, who=None):
        if item is None:
            bits = 'sans-serif:style=normal:variant=normal:weight=normal:stretch=normal:size=10.0'.split(':')
        else:
            bits = item.split(':')
        fontdict = {}
        for bit in bits:
            bits2 = bit.split('=')
            if len(bits2) == 1:
                bit2 = bits2[0].replace("'", '')
                bit2 = bit2.replace("\\", '')
                fontdict['family'] = bit2
            else:
                try:
                    fontdict[bits2[0]] = float(bits2[1])
                except:
                    fontdict[bits2[0]] = bits2[1]
        return fontdict

    def get_name_row_col(self, series):
        i = series.rfind(' (')
        label = series[:i]
        rc = series[i + 2:-1]
        for r in range(len(rc) -1, -1, -1):
            if not rc[r].isdigit():
                break
        row = int(rc[r + 1:]) - 1
        col = 0
        for c in range(r + 1):
            col = col * 26 + col_letters.index(rc[c])
        col -= 1
        return label, row, col

    def __init__(self, help='help.html'):
        super(FlexiPlot, self).__init__()
        self.help = help
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            self.config_file = sys.argv[1]
        else:
            self.config_file = getModelFile('flexiplot.ini')
        config.read(self.config_file)
        if not config.has_section('Flexiplot'): # new file set windows section
            self.restorewindows = True
        else:
            self.restorewindows = False
        try:
            rw = config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
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
        ifile = ''
        isheet = ''
        self.setup = [False, False]
        self.details = True
        self.book = None
        self.rows = None
        self.leapyear = False
        iper = '<none>'
        imax = 0
        self.alpha = 0.25
        self.legend_on_pie = True
        self.percent_on_pie = True
        self.legend_side = 'None'
        self.legend_ncol = 1
        self.legend_font = FontProperties()
        self.legend_font.set_size('medium')
        self.fontprops = {}
        self.fontprops['Label']= self.set_fontdict()
        self.fontprops['Label']['size'] = 10.
        self.fontprops['Ticks']= self.set_fontdict()
        self.fontprops['Ticks']['size'] = 10.
        self.fontprops['Title']= self.set_fontdict()
        self.fontprops['Title']['size'] = 14.
        self.constrained_layout = False
        self.max_xoffset = 5
        self.yseries = []
        self.palette = True
        self.history = None
        self.max_files = 10
        self.current_file = ''
        ifiles = self.get_flex_config()
        if len(ifiles) > 0:
            if self.history is None:
                self.history = sorted(ifiles.keys(), reverse=True)
            ifile = ifiles[self.history[0]]
        matplotlib.rcParams['savefig.directory'] = os.getcwd()
        self.grid = QtWidgets.QGridLayout()
        self.updated = False
        self.colours_updated = False
        self.log = QtWidgets.QLabel('')
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
        openfile = QtWidgets.QPushButton('Open File')
        self.grid.addWidget(openfile, rw, 3)
        openfile.clicked.connect(self.openfileClicked)
        listfiles = QtWidgets.QPushButton('List Files')
        self.grid.addWidget(listfiles, rw, 4)
        listfiles.clicked.connect(self.listfilesClicked)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Title:'), rw, 0)
        self.title = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.title, rw, 1, 1, 2)
        ttlfClicked = QtWidgets.QPushButton('Title Font', self)
        self.grid.addWidget(ttlfClicked, rw, 3)
        ttlfClicked.clicked.connect(self.doFont)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Abscissa (X):'), rw, 0)
        self.data_order = QtWidgets.QComboBox()
        self.data_order.addItem('Data in Rows (indicate Label column)')
        self.data_order.addItem('Data in Columns (indicate Label row)')
        self.grid.addWidget(self.data_order, rw, 1)
        self.data_order.currentIndexChanged.connect(self.dorderChanged)
        self.data_in_rows = True
        self.xoffset = QtWidgets.QComboBox()
        self.grid.addWidget(self.xoffset, rw, 2)
        for col in range(self.max_xoffset):
            self.xoffset.addItem(ssCol(col, base=0))
        self.xoffset.setCurrentIndex(0)
        self.xoffset.currentIndexChanged.connect(self.xoffsetChanged)
        self.grid.addWidget(QtWidgets.QLabel('(one set of variables for the common axis)'), rw, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Abscissa Series:'), rw, 0)
        self.xseries = QtWidgets.QComboBox()
        self.grid.addWidget(self.xseries, rw, 1, 1, 2)
        self.xextra = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.xextra, rw, 4)
        showseries = QtWidgets.QPushButton('Show X Series')
        self.grid.addWidget(showseries, rw, 3)
        showseries.clicked.connect(self.showClicked)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Abscissa Label:'), rw, 0)
        self.xlabel = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.xlabel, rw, 1, 1, 2)
        lblfClicked = QtWidgets.QPushButton('Label Font', self)
        self.grid.addWidget(lblfClicked, rw, 3)
        lblfClicked.clicked.connect(self.doFont)
        ticfClicked = QtWidgets.QPushButton('Ticks Font', self)
        self.grid.addWidget(ticfClicked, rw, 4)
        ticfClicked.clicked.connect(self.doFont)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Ordinate (Y):'), rw, 0)
        self.grid.addWidget(QtWidgets.QLabel('(multiple sets of variables, one for each surface)'), rw, 1)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Ordinate Series:\n(move to right\nto exclude)'), rw, 0)
        self.order = ListWidget(self)
        self.grid.addWidget(self.order, rw, 1, 1, 2)
        self.ignore = ListWidget(self)
        self.grid.addWidget(self.ignore, rw, 3, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel(' '), rw, 5)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Ordinate Label:'), rw, 0)
        self.ylabel = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.ylabel, rw, 1, 1, 2)
        showseries = QtWidgets.QPushButton('Show Y Series')
        self.grid.addWidget(showseries, rw, 3)
        showseries.clicked.connect(self.showClicked)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Applicate (Z) Label:'), rw, 0)
        self.zlabel = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.zlabel, rw, 1, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(for 3D the table of X / Y values)'), rw, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Maximum:'), rw, 0)
        self.maxSpin = QtWidgets.QSpinBox()
        self.maxSpin.setRange(0, 100000)
        self.maxSpin.setSingleStep(500)
        self.grid.addWidget(self.maxSpin, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('(Handy if you want to produce a series of charts)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Type of Chart:'), rw, 0)
        plots = ['Bar Chart', 'Cumulative', 'Line Chart', 'Pie Chart', 'Step Chart']
        if TablePlot3D is not None:
            plots.append('3D Surface Chart')
        self.plottype = QtWidgets.QComboBox()
        for plot in plots:
             self.plottype.addItem(plot)
        self.grid.addWidget(self.plottype, rw, 1) #, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Type of chart - stacked except for Line Chart)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Percentage:'), rw, 0)
        self.percentage = QtWidgets.QCheckBox()
        self.percentage.setCheckState(QtCore.Qt.Unchecked)
        self.grid.addWidget(self.percentage, rw, 1) #, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Check for percentage distribution)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Show Grid:'), rw, 0)
        grids = ['Both', 'Horizontal', 'Vertical', 'None']
        self.gridtype = QtWidgets.QComboBox()
        for grid in grids:
             self.gridtype.addItem(grid)
        self.grid.addWidget(self.gridtype, rw, 1) #, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Choose gridlines)'), rw, 3, 1, 3)
        if ifile != '':
            self.get_file_config(self.history[0])
        self.files.currentIndexChanged.connect(self.filesChanged)
        self.file.clicked.connect(self.fileChanged)
        self.xseries.activated[str].connect(self.xseriesChanged)
        self.xseries.currentIndexChanged.connect(self.xseriesChanged)
        self.xextra.textChanged.connect(self.xseriesChanged)
  #      self.xvalues.textChanged.connect(self.somethingChanged)
        self.files.currentIndexChanged.connect(self.xseriesChanged)
        self.sheet.currentIndexChanged.connect(self.sheetChanged)
        self.quiet = False
        self.title.textChanged.connect(self.somethingChanged)
        self.xlabel.textChanged.connect(self.somethingChanged)
        self.ylabel.textChanged.connect(self.somethingChanged)
        self.zlabel.textChanged.connect(self.somethingChanged)
        self.maxSpin.valueChanged.connect(self.somethingChanged)
        self.plottype.currentIndexChanged.connect(self.plotChanged)
        self.gridtype.currentIndexChanged.connect(self.somethingChanged)
        self.percentage.stateChanged.connect(self.somethingChanged)
        rw += 1
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.log.setPalette(msg_palette)
        self.grid.addWidget(self.log, rw, 1, 1, 4)
        rw += 1
        done = QtWidgets.QPushButton('Done', self)
        self.grid.addWidget(done, rw, 0)
        done.clicked.connect(self.doneClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.doneClicked)
        pp = QtWidgets.QPushButton('Chart', self)
        self.grid.addWidget(pp, rw, 1)
        pp.clicked.connect(self.ppClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('p'), self, self.ppClicked)
        cb = QtWidgets.QPushButton('Colours', self)
        self.grid.addWidget(cb, rw, 2)
        cb.clicked.connect(self.editColours)
        legendClicked = QtWidgets.QPushButton('Legend Properties', self)
        self.grid.addWidget(legendClicked, rw, 3)
        legendClicked.clicked.connect(self.doFont)
        ep = QtWidgets.QPushButton('Preferences', self)
        self.grid.addWidget(ep, rw, 4)
        ep.clicked.connect(self.editIniFile)
        help = QtWidgets.QPushButton('Help', self)
        self.grid.addWidget(help, rw, 5)
        help.clicked.connect(self.helpClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - flexiplot (' + fileVersion() + ') - FlexiPlot')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        if self.restorewindows:
            try:
                rw = config.get('Windows', 'flexiplot_size').split(',')
                self.resize(int(rw[0]), int(rw[1]))
                mp = config.get('Windows', 'flexiplot_pos').split(',')
                self.move(int(mp[0]), int(mp[1]))
            except:
                pass
        else:
            self.center()
            self.resize(int(self.sizeHint().width() * 1.2), int(self.sizeHint().height() * 1.2))
        self.log.setText('Preferences file: ' + self.config_file)
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        centerPoint = QtWidgets.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def get_flex_config(self):
        ifiles = {}
        self.constrained_layout = False
        self.palette = True
        self.sparse_ticks = []
        self.plot_palette = 'random'
        self.minimum_3d = 10
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        try: # get defaults and list of files if any
            items = config.items('Flexiplot')
            for key, value in items:
                if key == 'alpha':
                    try:
                        self.alpha = float(value)
                    except:
                        pass
                elif key == 'constrained_layout':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.constrained_layout = True
                elif key == 'file_history':
                    self.history = value.split(',')
                elif key == 'file_choices':
                    self.max_files = int(value)
                elif key[:4] == 'file':
                    ifiles[key[4:]] = value.replace('$USER$', getUser())
                elif key == 'label_font':
                    try:
                        self.fontprops['Label'] = self.set_fontdict(value)
                    except:
                        pass
                elif key == 'legend_font':
                    try:
                        self.legend_font.set_fontconfig_pattern(value)
                    except:
                        pass
                elif key == 'legend_on_pie':
                    if value.lower() in ['pct', 'percentage', '%', '%age']:
                        self.legend_on_pie = False
                    elif value.lower() in ['false', 'no', 'off']:
                        self.legend_on_pie = False
                        self.percent_on_pie = False
                elif key == 'legend_ncol':
                    try:
                        self.legend_ncol = int(value)
                    except:
                        pass
                elif key == 'legend_side':
                    try:
                        self.legend_side = value
                    except:
                        pass
                elif key == 'max_xoffset':
                    try:
                        self.max_xoffset = int(value)
                    except:
                        pass
                elif key == 'minimum_3d':
                    try:
                        self.minimum_3d = int(value)
                    except:
                        pass
                elif key == 'palette':
                    if value.lower() in ['false', 'no', 'off']:
                        self.palette = False
                elif key == 'plot_palette':
                    if value.lower() in ['false', 'no', 'off']:
                        self.plot_palette = ''
                    else:
                        self.plot_palette = value
                elif key == 'sparse_ticks':
                    try:
                        self.sparse_ticks = [int(value)]
                    except:
                        if value.lower() in ['true', 'yes', 'on']:
                            self.sparse_ticks = True
                        else:
                            if value.find(':') > 0:
                                self.sparse_ticks = value.split(':')
                            else:
                                self.sparse_ticks = value.split(',')
                elif key == 'title_font':
                    try:
                        self.fontprops['Title'] = self.set_fontdict(value)
                    except:
                        pass
                elif key == 'ticks_font':
                    try:
                        self.fontprops['Ticks'] = self.set_fontdict(value)
                    except:
                        pass
        except:
            pass
        return ifiles

    def get_file_config(self, choice=''):
        ifile = ''
        ignore = True
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        yseries = []
        isheet = ''
        ixseries = ''
        try: # get list of files if any
            ifile = config.get('Flexiplot', 'file' + choice).replace('$USER$', getUser())
        except:
            pass
        if not os.path.exists(ifile) and not os.path.exists(self.scenarios + ifile):
            if self.book is not None:
                self.book.close()
                self.book = None
            self.log.setText("Can't find file - " + ifile)
            msgbox = QtWidgets.QMessageBox()
            msgbox.setWindowTitle('SIREN - flexiplot file not found')
            msgbox.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
            fname = ifile[ifile.rfind('/') + 1:]
            msgbox.setText("Can't find '" + fname + "'\nDo you want to remove it from file history (Y)?")
            msgbox.setDetailedText(self.log.text())
            msgbox.setIcon(QtWidgets.QMessageBox.Question)
            msgbox.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            reply = msgbox.exec_()
            if reply == QtWidgets.QMessageBox.Yes:
                self.file_cleanup(choice)
                self.files.removeItem(self.files.currentIndex())
            return
        try: # get list of files if any
            self.maxSpin.setValue(0)
            self.gridtype.setCurrentIndex(0)
            self.percentage.setCheckState(QtCore.Qt.Unchecked)
            self.title.setText('')
            self.xextra.setText('')
            self.xlabel.setText('')
            self.ylabel.setText('')
            self.zlabel.setText('')
            self.title.setText('')
            self.data_order.setCurrentIndex(0)
            self.xseries.clear()
            self.order.clear()
            self.ignore.clear()
            items = config.items('Flexiplot')
            for key, value in items:
                if key == 'file' + choice:
                    ifile = value.replace('$USER$', getUser())
                elif key == 'grid' + choice:
                    self.gridtype.setCurrentIndex(self.gridtype.findText(value))
                elif key == 'percentage' + choice:
                    if value.lower() in ['true', 'yes', 'on']:
                        self.percentage.setCheckState(QtCore.Qt.Checked)
                    else:
                        self.percentage.setCheckState(QtCore.Qt.Unchecked)
                elif key == 'plot' + choice:
                    self.plottype.setCurrentIndex(self.plottype.findText(value))
                elif key == 'maximum' + choice:
                    try:
                        self.maxSpin.setValue(int(value))
                    except:
                        self.maxSpin.setValue(0)
                elif key == 'sheet' + choice:
                    isheet = value
                elif key == 'title' + choice:
                    self.title.setText(value)
                elif key == 'xextra' + choice:
                    self.xextra.setText(value)
                elif key == 'xlabel' + choice:
                    self.xlabel.setText(value)
                elif key == 'xoffset' + choice:
                    self.xoffset.setCurrentIndex(self.xoffset.findText(value))
                    if value.isdigit():
                        self.data_order.setCurrentIndex(1)
                    else:
                        self.data_order.setCurrentIndex(0)
                elif key == 'xseries' + choice:
                    ixseries = value.replace('\\n', '\n')
                elif key == 'ylabel' + choice:
                    self.ylabel.setText(value.replace('\\n', '\n'))
                elif key == 'yseries' + choice:
                    yseries = strSplit(value.replace('\\n', '\n'))
                elif key == 'zlabel' + choice:
                    self.zlabel.setText(value)
        except:
             pass
        nocolour = []
        if ifile != '':
            if self.book is not None:
                self.book.close()
                self.book = None
            self.file.setText(ifile)
            if os.path.exists(ifile):
                self.setSheet(ifile, isheet)
            else:
                self.setSheet(self.scenarios + ifile, isheet)
            self.setSeries(isheet, xseries=ixseries, yseries=yseries)
            for series in yseries:
                if not self.check_colour(series, config, add=False):
                    nocolour.append(series.lower())
            if ixseries != '':
                self.xseries.setCurrentIndex(self.xseries.findText(ixseries))
        if len(nocolour) > 0:
            more_colour = PlotPalette(nocolour, palette=self.plot_palette)
            self.colours = {**self.colours, **more_colour}
        ignore = False

    def popfileslist(self, ifile, ifiles=None):
        self.setup[1] = True
        self.current_file = ifile
        if ifiles is None:
             ifiles = {}
             for i in range(self.files.count()):
                 ifiles[self.history[i]] = self.files.itemText(i)
        if self.history is None:
            self.history = ['']
            ifiles = {'': ifile}
        else:
            for i in range(len(self.history) - 1, -1, -1):
                try:
                    if ifile == ifiles[self.history[i]]:
                        ihist = self.history[i]
                        del self.history[i]
                        self.history.insert(0, ihist) # make this entry first
                        break
                except KeyError:
                    del self.history[i]
                except:
                    pass
            else:
            # find new entry
                if len(self.history) >= self.max_files:
                    self.history.insert(0, self.history.pop(-1)) # make last entry first
                else:
                    hist = sorted(self.history)
                    if hist[0] == '':
                        hist = [int(x) for x in hist[1:]]
                        hist.sort()
                        hist.insert(0, '')
                    else:
                        hist = [int(x) for x in hist[1:]]
                        hist.sort()
                    hist = [str(x) for x in hist]
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
        self.log.setText('')
        if os.path.exists(self.file.text()):
            curfile = self.file.text()
        else:
            curfile = self.scenarios + self.file.text()
        newfile = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file', curfile)[0]
        if newfile != '':
            if self.book is not None:
                self.book.close()
                self.book = None
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
        self.log.setText('')
        self.saveConfig()
        self.get_file_config(self.history[self.files.currentIndex()])
        if self.book is None:
            return
        self.popfileslist(self.files.currentText())
        self.log.setText('File "loaded"')
        self.setup[0] = False

    def doFont(self):
        if self.sender().text()[:6] == 'Legend':
            if self.legend_on_pie and self.percent_on_pie:
                lop = 'True'
            elif not self.legend_on_pie and self.percent_on_pie:
                lop = 'Pct'
            else:
                lop = 'False'
            legend_properties = {'legend_ncol': str(self.legend_ncol),
                                 'legend_on_pie': lop,
                                 'legend_side': self.legend_side}
            legend = ChangeFontProp('Legend', self.legend_font, legend_properties)
            legend.exec_()
            values, font_values = legend.getValues()
            if values is None:
                return
            for key, value in values.items():
                if key == 'legend_ncol':
                    self.legend_ncol = int(value)
                elif key == 'legend_on_pie':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.legend_on_pie = True
                        self.percent_on_pie = True
                    if value.lower() in ['pct', 'percentage', '%', '%age']:
                        self.legend_on_pie = False
                        self.percent_on_pie = True
                    elif value.lower() in ['false', 'no', 'off']:
                        self.legend_on_pie = False
                        self.percent_on_pie = False
                else:
                    setattr(self, key, value)
            self.legend_font = legend.getFont()
            self.log.setText('Legend properties updated')
            self.updated = True
        else:
            what = self.sender().text().split(' ')[0]
            font = ChangeFontProp(what, self.fontprops[what])
            font.exec()
            if font.getFont() is None:
                return
            self.fontprops[what] = self.set_fontdict(font.getFontDict())
            self.log.setText(what + ' font properties updated')
            self.updated = True
        return

    def plotChanged(self):
        if not self.setup[0]:
            self.updated = True
        self.setSeries(self.sheet.currentText())

    def somethingChanged(self):
        if not self.setup[0]:
            self.updated = True

    def setSheet(self, ifile, isheet):
        if self.book is None:
            try:
                self.book = WorkBook()
                self.book.open_workbook(ifile)
            except:
                self.book = None
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
        self.log.setText('')
        if self.book is None:
            try:
                self.book = WorkBook()
                self.book.open_workbook(newfile)
            except:
                self.book = None
                self.log.setText("Can't open file - " + newfile)
                return
        isheet = self.sheet.currentText()
        if isheet not in self.book.sheet_names():
            self.log.setText("Can't find sheet - " + isheet)
            return
        self.setSeries(isheet)
        self.updated = True

    def seriesChanged(self):
        if self.setup[0]:
            return
        self.log.setText('')
        self.setSeries(self.sheet.currentText())
        self.updated = True

    def dorderChanged(self):
        self.xoffset.clear()
        if self.data_order.currentIndex() == 0:
            for col in range(self.max_xoffset):
                self.xoffset.addItem(ssCol(col, base=0))
            self.data_in_rows = True
        else:
            for row in range(self.max_xoffset):
                self.xoffset.addItem(f'{row + 1}')
            self.data_in_rows = False
        self.xoffset.setCurrentIndex(0)

    def xoffsetChanged(self):
        if self.setup[0]:
            return
        self.log.setText('')
        self.setSeries(self.sheet.currentText())
        self.updated = True

    def xseriesChanged(self):
        if self.setup[0]:
            return
        self.log.setText('')
        self.setSeries(self.sheet.currentText())
        self.updated = True

    def xvaluesChanged(self):
        self.log.setText('')
        xvalues = self.xvaluesi.currentText()
        self.xvalues = [xvalues]
        for i in range(self.xvaluesi.count()):
            if self.xvaluesi.itemText(i) != xvalues:
                self.xvalues.append(self.xvaluesi.itemText(i))
        self.updated = True

    def setSeries(self, isheet, xseries='', yseries=[]):
        try:
            ws = self.book.sheet_by_name(isheet)
        except:
            self.log.setText("Can't find sheet - " + isheet)
            return
        self.setup[0] = True
        if xseries == '':
            xseries = self.xseries.currentText()
        self.xseries.clear()
        xaxes = []
        xaxes3d = []
        if self.xoffset.currentText().isdigit():
            row = int(self.xoffset.currentText()) - 1
            self.data_in_rows = False
            # find unique columns
            for col in range(ws.ncols):
                xaxes.append(f"{str(ws.cell_value(row, col))} ({ssCol(col, base=0)}{self.xoffset.currentText()})")
                xaxes3d.append(f"{str(ws.cell_value(row, col))}")
            if self.plottype.currentText() == '3D Surface Chart':
                rows = []
                keys = []
                for x in range(len(xaxes3d)):
                    try:
                        i = keys.index(xaxes3d[x])
                        rows[i].append(x)
                    except:
                        keys.append(xaxes3d[x])
                        rows.append([x])
                dels = []
                for row in rows:
                    if len(row) > 1:
                        dels.extend(row)
                dels = sorted(dels, reverse=True)
                if len(dels) > self.minimum_3d:
                    for d in dels:
                        del xaxes[d]
                    self.log.setText(f'Possibly table data ({len(dels)} table data rows)')
            if len(yseries) == 0:
                for col in range(self.order.count()):
                    yseries.append(self.order.item(col).text())
        else:
            col = col_letters.find(self.xoffset.currentText()) - 1
            self.data_in_rows = True
            # find unique rows
            for row in range(ws.nrows):
                xaxes.append(f"{str(ws.cell_value(row, col))} ({self.xoffset.currentText()}{row + 1})")
                xaxes3d.append(f"{str(ws.cell_value(row, col))}")
            if self.plottype.currentText() == '3D Surface Chart':
                rows = []
                keys = []
                for x in range(len(xaxes3d)):
                    try:
                        i = keys.index(xaxes3d[x])
                        rows[i].append(x)
                    except:
                        keys.append(xaxes3d[x])
                        rows.append([x])
                dels = []
                for row in rows:
                    if len(row) > 1:
                        dels.extend(row)
                dels = sorted(dels, reverse=True)
                if len(dels) > self.minimum_3d:
                    for d in dels:
                        del xaxes[d]
                    self.log.setText(f'Possibly table data ({len(dels)} table data rows)')
            if len(yseries) == 0:
                for col in range(self.order.count()):
                    yseries.append(self.order.item(col).text())
        self.order.clear()
        self.ignore.clear()
        c = 0
     #   for d in rows:
      #      if len(d) == 1:
       #         if self.data_in_rows:
        #            xaxes[c] = f'{xaxes[c]} ({self.xoffset.currentText()}{d[0] + 1})'
         #       else:
          #          xaxes[c] = f'{xaxes[c]} ({ssCol(d[0],base=0)}{self.xoffset.currentText()})'
           #     c += 1
        if self.xextra.text() != '':
            self.xseries.addItem(f' ({self.xextra.text()})')
        for item in xaxes:
            self.xseries.addItem(item)
            try:
                itm = item[:item.rfind(' (')]
                item_colour = QtGui.QColor(self.colours[itm.lower()])
            except:
                item_colour = None
            if item in yseries:
                self.order.addItem(item)
                try:
                    self.order.item(self.order.count() - 1).setBackground(item_colour)
                except:
                    pass
            else:
                self.ignore.addItem(item)
                try:
                    self.ignore.item(self.ignore.count() - 1).setBackground(item_colour)
                except:
                    pass
        try:
            self.xseries.setCurrentIndex(self.xseries.findText(xseries))
        except:
            pass
        self.setup[0] = False
        self.updated = True
        return

    def helpClicked(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.help,
                 title='Help for flexiplot (' + fileVersion() + ')', section='flexiplot')
        dialog.exec_()

    def listfilesClicked(self):
        ifiles = {}
        for i in range(self.files.count()):
             if self.history[i] == '':
                 ifiles[' '] = self.files.itemText(i)
             else:
                 ifiles[self.history[i]] = self.files.itemText(i)
        fields = ['Choice', 'File name']
        dialog = displaytable.Table(ifiles, title='Files', fields=fields, sortby='', edit='delete')
        dialog.exec_()
        if dialog.getValues() is not None:
            ofiles = dialog.getValues()
            dels = []
            for key in ifiles.keys():
                if key not in ofiles.keys():
                    dels.append(key)
            for choice in dels:
                self.file_cleanup(choice)
                for i in range(self.files.count()):
                     if self.files.itemText(i) == ifiles[choice]:
                         self.files.removeItem(i)
        del dialog
        self.filesChanged()

    def file_cleanup(self, choice):
        if choice == ' ':
            choice = ''
        line = ''
        for i in range(len(self.history) -1, -1, -1):
            if self.history[i] == choice:
                del self.history[i]
            else:
                line += self.history[i] + ','
        line = line[:-1]
        lines = ['file_history=' + line]
        for prop in ['file', 'grid', 'maximum', 'percentage', 'plot', 'sheet', 'title',
                     'xextra', 'xlabel', 'xoffset', 'xseries', 'ylabel', 'yseries', 'zlabel']:
            lines.append(prop + choice + '=')
        for prop in ['columns', 'series', 'xvalues']: # old properties
            lines.append(prop + choice + '=')
        updates = {'Flexiplot': lines}
        SaveIni(updates, ini_file=self.config_file)
        self.log.setText('File removed from file history')

    def openfileClicked(self):
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            os.startfile(self.file.text())
            self.log.setText('File opened')
        elif sys.platform in ['darwin', 'linux', 'linux2']:
            pid = subprocess.Popen(['open', self.file.text()]).pid
            self.log.setText('File opened (pid ' + str(pid) + ')')

    def showClicked(self):
        if self.book is None:
            self.log.setText('Error accessing Workbook.')
            return
        isheet = self.sheet.currentText()
        if isheet == '':
            self.log.setText('Sheet not set.')
            return
        ws = self.book.sheet_by_name(isheet)
        x = []
        series = self.xseries.currentText()
        if series == '':
            self.log.setText('No Abscissa (X) Series chosen.')
            return
        label, row, col = self.get_name_row_col(series)
        values = []
        if self.xoffset.currentText().isdigit():
            fields = ['Row', '??']
            if label == 'None' or label == '':
                values.append([row + 1, '??'])
                row += 1
            for row in range(row, ws.nrows):
                if ws.cell_value(row, col) is None:
                    break
                values.append([row + 1, str(ws.cell_value(row, col))])
            max_row = row + 1
            if self.sender().text()[-8:] == 'Y Series':
                for o in range(self.order.count()):
                    label, row1, col = self.get_name_row_col(self.order.item(o).text())
                    fields.append(label)
                    c = 0
                    for row in range(row1, max_row):
                        if ws.cell_value(row, col) is None:
                            values[c].append('')
                        else:
                            values[c].append(str(ws.cell_value(row, col)))
                        c += 1
                        if c >= len(values):
                            break
        else:
            fields = ['Col', '??']
            if label == 'None' or label == '':
                values.append([ssCol(col, base=0), '??'])
                col += 1
            for col in range(col, ws.ncols):
                if ws.cell_value(row, col) is None:
                    break
                values.append([ssCol(col, base=0), str(ws.cell_value(row, col))])
            max_col = col + 1
            if self.sender().text()[-8:] == 'Y Series':
                for o in range(self.order.count()):
                    label, row, col = self.get_name_row_col(self.order.item(o).text())
                    fields.append(label)
                    c = 0
                    for col in range(col, max_col):
                        if ws.cell_value(row, col) is None:
                            values[c].append('')
                        else:
                            values[c].append(str(ws.cell_value(row, col)))
                        c += 1
                        if c >= len(values):
                            break
        dialog = displaytable.Table(values, title='Series values', fields=fields, sortby='')
        dialog.exec_()
        del dialog

    def doneClicked(self):
        if self.book is not None:
            self.book.close()
        if not self.updated and not self.colours_updated:
            self.close()
        self.saveConfig()
        self.close()

    def closeEvent(self, event):
        plt.close('all')
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)   # need to account for border
            lines.append('flexiplot_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('flexiplot_size=%s,%s' % (str(self.width()), str(self.height())))
            updates['Windows'] = lines
            SaveIni(updates, ini_file=self.config_file)
        event.accept()

    def editIniFile(self):
        curfldr = self.file.text()[:self.file.text().rfind('/')]
        dialr = EdtDialog(self.config_file, section='[Flexiplot]')
        dialr.exec_()
        ifiles = self.get_flex_config()

    def saveConfig(self):
        updates = {}
        if self.order.updated:
            self.updated = True
            self.order.updated = False
        if self.updated:
            config = configparser.RawConfigParser()
            config.read(self.config_file)
            try:
                choice = self.history[0]
            except:
                choice = ''
            save_file = self.current_file.replace(getUser(), '$USER$')
            try:
                self.max_files = int(config.get('Flexiplot', 'file_choices'))
            except:
                pass
            lines = []
            lines.append('legend_font=' + self.legend_font.get_fontconfig_pattern())
            lines.append('legend_on_pie=')
            if not self.legend_on_pie and self.percent_on_pie:
                lines[-1] += 'Pct'
            elif not self.legend_on_pie and not self.percent_on_pie:
                lines[-1] += 'False'
            lines.append('legend_ncol=')
            if self.legend_ncol != 1:
                lines[-1] += str(self.legend_ncol)
            lines.append('legend_side=')
            if self.legend_side != 'None':
                lines[-1] += self.legend_side
            for key, value in self.fontprops.items():
                lines.append(key.lower() + '_font=')
                for key2, value2 in value.items():
                    lines[-1] += f'{key2}={value2}:'
                lines[-1] = lines[-1][:-1]
            if len(self.history) > 0:
                line = ''
                for itm in self.history:
                    line += itm + ','
                line = line[:-1]
                lines.append('file_history=' + line)
            lines.append('file' + choice + '=' + save_file)
            lines.append('grid' + choice + '=' + self.gridtype.currentText())
            lines.append('maximum' + choice + '=')
            if self.maxSpin.value() != 0:
                lines[-1] = lines[-1] + str(self.maxSpin.value())
            lines.append('percentage' + choice + '=')
            if self.percentage.isChecked():
                lines[-1] = lines[-1] + 'True'
            lines.append('plot' + choice + '=' + self.plottype.currentText())
            lines.append('sheet' + choice + '=' + self.sheet.currentText())
            lines.append('title' + choice + '=' + self.title.text())
            lines.append('xextra' + choice + '=' + self.xextra.text())
            lines.append('xlabel' + choice + '=' + self.xlabel.text())
            lines.append('xoffset' + choice + '=' + self.xoffset.currentText())
            lines.append('xseries' + choice + '=' + self.xseries.currentText().replace('\n', '\\n'))
            lines.append('ylabel' + choice + '=' + self.ylabel.text())
            line = 'yseries' + choice + '='
            for y in range(self.order.count()):
                try:
                    if self.order.item(y).text().index(',') >= 0:
                        try:
                            if self.order.item(y).text().index("'") >= 0:
                                qte = '"'
                        except:
                            qte = "'"
                except:
                    qte = ''
                line += qte + self.order.item(y).text().replace('\n', '\\n') + qte + ','
            if line[-1] != '=':
                line = line[:-1]
            lines.append(line)
            lines.append('zlabel' + choice + '=' + self.zlabel.text())
            for prop in ['columns', 'series', 'xvalues']: # old properties
                lines.append(prop + choice + '=')
            updates['Flexiplot'] = lines
            if self.restorewindows and not config.has_section('Windows'): # new file set windows section
                updates['Windows'] = ['restorewindows=True']
        if self.colours_updated:
            lines = []
            for key, value in self.colours.items():
                if value != '':
                    lines.append(key.strip().replace(' ', '_') + '=' + value)
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
                        palette=palette, underscore=True)
        if not dialr.cancelled:
            dialr.exec_()
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
            colr = self.order.item(c).text()[:self.order.item(c).text().rfind(' (')]
            try:
                self.order.item(c).setBackground(QtGui.QColor(self.colours[colr.lower()]))
            except:
                pass
        for c in range(self.ignore.count()):
            colr = self.ignore.item(c).text()[:self.ignore.item(c).text().rfind(' (')]
            try:
                self.ignore.item(c).setBackground(QtGui.QColor(self.colours[colr.lower()]))
            except:
                pass

    def check_colour(self, colour, config, add=True):
        colr = colour.lower()
        if colr in self.colours.keys():
            return True
        elif self.plot_palette[0].lower() == 'r':
            # new approach to generate random colour if not in [Plot Colors]
            r = lambda: random.randint(0,255)
            new_colr = '#%02X%02X%02X' % (r(),r(),r())
            self.colours[colr] = new_colr
            return True
        # previous approach may ask for new colours to be stored in .ini file
        colr2 = colr.replace('_', ' ')
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

    def ppClicked(self):
        if self.book is None:
            self.log.setText('Error accessing Workbook.')
            return
        if self.xseries.currentText() == '' or self.order.count() == 0:
            self.log.setText('Nothing to plot.')
            return
        isheet = self.sheet.currentText()
        if isheet == '':
            self.log.setText('Sheet not set.')
            return
        config = configparser.RawConfigParser()
        config.read(self.config_file)
  #      for c in range(self.order.count()):
   #         if not self.check_colour(self.order.item(c).text(), config):
    #            return
        if self.plottype.currentText() == '3D Surface Chart':
            html_3d = False
            contours_3d = False
            background_3d = True
            colorbars_3d = []
            aspectmode_3d = 'auto'
            try:
                aspectmode_3d = config.get('Flexiplot', 'aspectmode_3d').lower()
            except:
                pass
            try:
                chk = config.get('Flexiplot', 'background_3d')
                if chk.lower() in ['false', 'off', 'no']:
                    background_3d = False
            except:
                pass
            try:
                chk = config.get('Flexiplot', 'colorbars_3d')
                bits = chk.split('(')
                for bit in bits:
                    if len(bit) > 0:
                        bit = bit.rstrip('),')
                        bits2 = bit.split(',')
                        colorbars_3d.append([])
                        for bit2 in bits2:
                            colorbars_3d[-1].append(bit2)
            except:
                pass
            try:
                chk = config.get('Flexiplot', 'contours_3d')
                if chk.lower() in ['true', 'on', 'yes']:
                    contours_3d = True
            except:
                pass
            try:
                chk = config.get('Flexiplot', 'html_3d')
                if chk.lower() in ['true', 'on', 'yes']:
                    html_3d = True
            except:
                pass
        del config
        self.log.setText('')
        titl = self.title.text() #.replace('$YEAR$', str(year))
        titl = titl.replace('$MTH$', '')
        titl = titl.replace('$MONTH$', '')
        titl = titl.replace('  ', '')
        titl = titl.replace('Diurnal ', '')
        titl = titl.replace('Diurnal', '')
        titl = titl.replace('$SHEET$', isheet)
        ws = self.book.sheet_by_name(isheet)
        if self.plottype.currentText() == '3D Surface Chart':
            if html_3d:
                saveit = self.file.text()[:self.file.text().rfind('.')] + '.html'
            else:
                saveit = None
            yseries = []
            colours = {}
            nocolour = []
            for c in range(self.order.count()):
                yseries.append(self.order.item(c).text()[:self.order.item(c).text().rfind(' (')])
                try:
                    colours[yseries[-1].lower()] = self.colours[yseries[-1].lower()]
                except:
                    nocolour.append(yseries[-1].lower())
            if len(nocolour) > 0:
                for bar in colorbars_3d:
                    colours[nocolour[0].lower()] = bar
                    del nocolour[0]
                    if len(nocolour) == 0:
                        break
                if len(nocolour) > 0:
                    more_colour = PlotPalette(nocolour, palette=self.plot_palette)
                    colours = {**colours, **more_colour}
            # colours, title, ws, x_offset, x_name, y_label, y_names, z_label, html=None, background=False, contours=False
            x_name = self.xseries.currentText()[:self.xseries.currentText().rfind(' (')]
            if x_name == '':
                label, row, col = self.get_name_row_col(self.xseries.currentText())
                x_name = row
            TablePlot3D(colours, titl, ws, self.xoffset.currentText(), x_name, self.ylabel.text(), yseries, self.zlabel.text(),
                        html=saveit, background=background_3d, contours=contours_3d, aspectmode=aspectmode_3d)
            self.log.setText("3D Chart complete (You'll need to close the browser window yourself).")
            return
        i = self.file.text().rfind('/')
        if i > 0:
            matplotlib.rcParams['savefig.directory'] = self.file.text()[:i + 1]
        else:
            matplotlib.rcParams['savefig.directory'] = self.scenarios
        x = []
        xlabels = []
        series = self.xseries.currentText()
        label, row, col = self.get_name_row_col(series)
        values = []
        if self.xoffset.currentText().isdigit():
            for row in range(row + 1, ws.nrows):
                if ws.cell_value(row, col) is None:
                    row -= 1
                    break
                xlabels.append(ws.cell_value(row, col))
           # print('(1372)', row, ws.nrows, xlabels[0], xlabels[-1])
            max_row = row + 1
            for i in range(len(xlabels)):
                x.append(i)
        else:
            for col in range(col + 1, ws.ncols):
                if ws.cell_value(row, col) is None:
                    col -= 1
                    break
                xlabels.append(ws.cell_value(row, col))
           # print('(1382)', col, ws.ncols, xlabels[0], xlabels[-1])
            max_col = col + 1
            for i in range(len(xlabels)):
                x.append(i)
        data = []
        labels = []
        try:
            self.tick_color = self.fontprops['Ticks']['color']
        except:
            self.tick_color = '#000000'
        miny = 0
        maxy = 0
        nocolour = []
        if self.xoffset.currentText().isdigit():
            for o in range(self.order.count()):
                label, row, col = self.get_name_row_col(self.order.item(o).text())
                labels.append(label)
                if label.lower() in self.colours.keys():
                    pass
                else:
                    nocolour.append(label.lower())
                data.append([])
                for row in range(row + 1, max_row):
                    if ws.cell_value(row, col) is None or ws.cell_value(row, col) == '':
                        data[-1].append(0.)
                    else:
                        data[-1].append(ws.cell_value(row, col))
                    try:
                        miny = min(miny, data[-1][-1])
                        maxy = max(maxy, data[-1][-1])
                    except:
                        self.log.setText('Data in cell ' + ssCol(col, base=0) + str(row + 1) + ' seems wrong - ' + data[-1][-1])
                        return
        else:
            for o in range(self.order.count()):
                label, row, col = self.get_name_row_col(self.order.item(o).text())
                labels.append(label)
                if label.lower() in self.colours.keys():
                    pass
                else:
                    nocolour.append(label.lower())
                data.append([])
                for col in range(col + 1, max_col):
                    if ws.cell_value(row, col) is None or ws.cell_value(row, col) == '':
                        data[-1].append(0.)
                    else:
                        data[-1].append(ws.cell_value(row, col))
                    try:
                        miny = min(miny, data[-1][-1])
                        maxy = max(maxy, data[-1][-1])
                    except:
                        self.log.setText('Data in cell ' + ssCol(col, base=0) + str(row + 1) + ' seems wrong - ' + data[-1][-1])
                        return
        if len(nocolour) > 0:
            more_colour = PlotPalette(nocolour, palette=self.plot_palette)
            self.colours = {**self.colours, **more_colour}
        if self.gridtype.currentText() == 'Both':
            gridtype = 'both'
        elif self.gridtype.currentText() == 'Horizontal':
            gridtype = 'y'
        elif self.gridtype.currentText() == 'Vertical':
            gridtype = 'x'
        else:
            gridtype = ''
        figname = self.plottype.currentText().lower().replace(' ','')  # + str(year))
        fig = plt.figure(figname, constrained_layout=self.constrained_layout)
        if gridtype != '':
            plt.grid(axis=gridtype)
        loc = 'lower right'
        graph = plt.subplot(111)
        if self.legend_side == 'Right':
            plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
        elif self.legend_side == 'Left':
            plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
            loc = 'lower left'
        plt.title(titl, fontdict=self.fontprops['Title'])
        if self.plottype.currentText() == 'Pie Chart':
            fig = plt.figure(figname, constrained_layout=self.constrained_layout)
            plt.title(titl, fontdict=self.fontprops['Title'])
            if gridtype != '':
                plt.grid(axis=gridtype)
            plabels = labels[:]
            datasum = []
            colors = []
            for dat in data:
                datasum.append(sum(dat))
                colors.append(self.colours[labels[len(datasum) - 1].lower()])
            for d in range(len(datasum) -1, -1, -1):
                if datasum[d] == 0:
                    del datasum[d]
                    del colors[d]
                    del plabels[d]
            tot = sum(datasum)
            # https://stackoverflow.com/questions/3942878/how-to-decide-font-color-in-white-or-black-depending-on-background-color
            whites = []
            threshold = sqrt(1.05 * 0.05) - 0.05
            if self.percent_on_pie:
                for c in range(len(colors)):
                    other = 0
                    intensity = 0.
                    for i in range(1, 5, 2):
                        colnum = int(colors[c][i : i + 2], 16)
                        colr = colnum / 255.0
                        if colr <= 0.04045:
                            colr = colr / 12.92
                        else:
                            colr = pow((colr + 0.055) / 1.055, 2.4)
                        if i == 1: # red
                            intensity += colr * 0.216
                            other += colnum * 0.299
                        elif i == 3: # green
                            intensity += colr * 0.7152
                            other += colnum * 0.587
                        else: # blue
                            intensity += colr * 0.0722
                            other += colnum * 0.114
                    if intensity < threshold:
                        whites.append(c)
            if self.legend_on_pie: # legend on chart
                if self.percent_on_pie:
                    patches, texts, autotexts = graph.pie(datasum, labels=plabels, autopct='%1.1f%%', colors=colors, startangle=90, pctdistance=.70)
                    for text in autotexts:
                        text.set_family(self.legend_font.get_family()[0])
                        text.set_style(self.legend_font.get_style())
                        text.set_variant(self.legend_font.get_variant())
                        text.set_stretch(self.legend_font.get_stretch())
                        text.set_weight(self.legend_font.get_weight())
                        text.set_size(self.legend_font.get_size())
                else:
                    patches, texts = graph.pie(datasum, labels=plabels, colors=colors, startangle=90, pctdistance=.70)
                for text in texts:
                    text.set_family(self.legend_font.get_family()[0])
                    text.set_style(self.legend_font.get_style())
                    text.set_variant(self.legend_font.get_variant())
                    text.set_stretch(self.legend_font.get_stretch())
                    text.set_weight(self.legend_font.get_weight())
                    text.set_size(self.legend_font.get_size())
            else:
                loc = 'lower right'
                if self.percent_on_pie:
                    patches, texts, autotexts = graph.pie(datasum, labels=None, autopct='%1.1f%%', colors=colors, startangle=90, pctdistance=.70)
                    for text in autotexts:
                        text.set_family(self.legend_font.get_family()[0])
                        text.set_style(self.legend_font.get_style())
                        text.set_variant(self.legend_font.get_variant())
                        text.set_stretch(self.legend_font.get_stretch())
                        text.set_weight(self.legend_font.get_weight())
                        text.set_size(self.legend_font.get_size())
                    for text in texts:
                        text.set_family(self.legend_font.get_family()[0])
                        text.set_style(self.legend_font.get_style())
                        text.set_variant(self.legend_font.get_variant())
                        text.set_stretch(self.legend_font.get_stretch())
                        text.set_weight(self.legend_font.get_weight())
                        text.set_size(self.legend_font.get_size())
                else:
                    for i in range(len(datasum)):
                        label[i] = f'{plabels[i]} ({datasum[i]/tot*100:0.1f}%)'
                    patches, texts = graph.pie(datasum, colors=colors, startangle=90)
                if self.legend_side == 'Right':
                    plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                elif self.legend_side == 'Left':
                    plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                    loc = 'lower left'
                fig.legend(patches, plabels, loc=loc, ncol=self.legend_ncol, prop=self.legend_font).set_draggable(True)
            for c in whites:
                autotexts[c].set_color('white')
            p = plt.gcf()
            p.gca().add_artist(plt.Circle((0, 0), 0.40, color='white'))
            plt.show()
            return
        elif self.plottype.currentText() in ['Cumulative', 'Step Chart']:
            if self.plottype.currentText() == 'Cumulative':
                step = None
            else:
                step = 'pre'
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
                graph.fill_between(x, 0, values, label=labels[0], color=self.colours[labels[0].lower()], step=step)
                for c in range(1, len(data)):
                    for h in range(len(data[c])):
                        bottoms[h] = values[h]
                        values[h] = values[h] + data[c][h] / totals[h] * 100.
                    graph.fill_between(x, bottoms, values, label=labels[c], color=self.colours[labels[c].lower()], step=step)
                maxy = 100
            else:
                graph.fill_between(x, miny, data[0], label=labels[0], color=self.colours[labels[0].lower()], step=step)
                for c in range(1, len(data)):
                    for h in range(len(data[c])):
                        data[c][h] = data[c][h] + data[c - 1][h]
                        maxy = max(maxy, data[c][h])
                    graph.fill_between(x, data[c - 1], data[c], label=labels[c], color=self.colours[labels[c].lower()], step=step)
                top = data[0][:]
                for d in range(1, len(data)):
                    for h in range(len(top)):
                        top[h] = max(top[h], data[d][h])
                if self.plottype.currentText() == 'Cumulative':
                    graph.plot(x, top, color='white')
                else:
                    graph.step(x, top, color='white')
                if self.maxSpin.value() > 0:
                    maxy = self.maxSpin.value()
                else:
                    try:
                        rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                        maxy = ceil(maxy / rndup) * rndup
                    except:
                        pass
        elif self.plottype.currentText() == 'Bar Chart':
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
                graph.bar(x, values, label=labels[0], color=self.colours[labels[0].lower()])
                for c in range(1, len(data)):
                    for h in range(len(data[c])):
                        bottoms[h] = bottoms[h] + values[h]
                        values[h] = data[c][h] / totals[h] * 100.
                    graph.bar(x, values, bottom=bottoms, label=labels[c], color=self.colours[labels[c].lower()])
                maxy = 100
            else:
                graph.bar(x, data[0], label=labels[0], color=self.colours[labels[0].lower()])
                bottoms = [0.] * len(x)
                for c in range(1, len(data)):
                    for h in range(len(data[c])):
                        bottoms[h] = bottoms[h] + data[c - 1][h]
                        maxy = max(maxy, data[c][h] + bottoms[h])
                    graph.bar(x, data[c], bottom=bottoms, label=labels[c], color=self.colours[labels[c].lower()])
                if self.maxSpin.value() > 0:
                    maxy = self.maxSpin.value()
                else:
                    try:
                        rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                        maxy = ceil(maxy / rndup) * rndup
                    except:
                        pass
        else: # Line Chart
            for c in range(len(data)):
                mx = max(data[c])
                maxy = max(maxy, mx)
                graph.plot(x, data[c], linewidth=2.0, label=labels[c], color=self.colours[labels[c].lower()])
            if self.maxSpin.value() > 0:
                maxy = self.maxSpin.value()
            else:
                try:
                    rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                    maxy = ceil(maxy / rndup) * rndup
                except:
                    pass
        graph.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=self.legend_font)
        plt.ylim([miny, maxy])
        plt.xlim([0, len(x) - 1])
        if self.sparse_ticks or len(self.sparse_ticks) > 0:
            # if self.plottype.currentText() == 'Line Chart' and len(x) > 24:
            tick_labels = [str(xlabels[0])]
            xticks = [0]
            if self.sparse_ticks and isinstance(self.sparse_ticks, bool):
                for l in range(1, len(xlabels)):
                    if xlabels[l] != tick_labels[-1]:
                        xticks.append(l)
                        tick_labels.append(xlabels[l])
            elif len(self.sparse_ticks) == 1:
                for l in range(self.sparse_ticks[0] + 1, len(xlabels), self.sparse_ticks[0]):
                    xticks.append(l)
                    tick_labels.append(xlabels[l])
            else:
                try:
                    fr = int(self.sparse_ticks[0]) - 1
                    to = int(self.sparse_ticks[1])
                except:
                    fr = 0
                    to = 3
                for l in range(1, len(xlabels)):
                    if str(xlabels[l])[fr : to] != tick_labels[-1][fr : to]:
                        xticks.append(l)
                        tick_labels.append(xlabels[l])
            xticks.append(len(xlabels) - 1)
            tick_labels.append(xlabels[-1])
            plt.xticks(xticks)
            graph.set_xticklabels(tick_labels, rotation='vertical', fontdict=self.fontprops['Ticks'])
            graph.tick_params(colors=self.tick_color, which='both')
        else:
            plt.xticks(x, xlabels)
            graph.set_xticklabels(xlabels, rotation='vertical', fontdict=self.fontprops['Ticks'])
        xlabel = self.xlabel.text()
        if xlabel == '':
            xlabel = self.xseries.currentText()
            xlabel = xlabel[:xlabel.rfind(' (')]
        graph.set_xlabel(xlabel, fontdict=self.fontprops['Label'])
        graph.set_ylabel(self.ylabel.text(), fontdict=self.fontprops['Label'])
        graph.set_yticks(graph.get_yticks().tolist())
        yticks = graph.get_yticklabels()
        graph.set_yticklabels(yticks, fontdict=self.fontprops['Ticks'])
        graph.tick_params(colors=self.tick_color, which='both')
        if self.percentage.isChecked():
            formatter = plt.FuncFormatter(lambda y, pos: '{:.0f}%'.format(y))
        else:
            formatter = plt.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y))
        graph.yaxis.set_major_formatter(formatter)
        zp = ZoomPanX(yformat=formatter)
        f = zp.zoom_pan(graph, base_scale=1.2, dropone=True) # enable scrollable zoom
        plt.show()
        del zp


if "__main__" == __name__:
    app = QtWidgets.QApplication(sys.argv)
    ex = FlexiPlot()
    app.exec_()
    app.deleteLater()
    sys.exit()
