#!/usr/bin/python3
#
#  Copyright (C) 2020-2021 Sustainable Energy Now Inc., Angus King
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
import sys
from math import log10, ceil
import matplotlib
from matplotlib.font_manager import FontProperties
import pylab as plt
import pyexcel as pxl
import random
import displayobject
from colours import Colours
from credits import fileVersion
from editini import EditSect, SaveIni
from getmodels import getModelFile
from senutils import ClickableQLabel, getParents, getUser, techClean
from zoompan import ZoomPanX

def charSplit(string, char=',', dropquote=True):
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

def get_range(text, alphabet=None, base=0):
    if len(text) < 1:
        return None
    if text[0].isdigit():
        roco = [int(x) - 1 for x in text.split(',') if x.strip().isdigit()]
        for i in roco:
            if len(roco) < 4:
                return None
        return roco
    if alphabet is None:
        alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    if alphabet[0] == ' ':
        alphabet = alphabet[1:]
    alphabet = alphabet.upper()
    bits = ['', '', '', '']
    b = 0
    in_char = True
    for char in text:
        if char.isdigit():
            if in_char:
                in_char = False
                b += 1
        else:
            if alphabet.find(char.upper()) < 0:
                continue
            if not in_char:
                in_char = True
                b += 1
        if b >= len(bits):
            break
        bits[b] += char.upper()
    try:
        bits[1] = int(bits[1]) - (1 - base)
        bits[3] = int(bits[3]) - (1 - base)
    except:
        pass
    for b in (0, 2):
        row = 0
        ndx = 1
        for c in range(len(bits[b]) -1, -1, -1):
            ndx1 = alphabet.index(bits[b][c]) + 1
            row = row + ndx1 * ndx
            ndx = ndx * len(alphabet)
        bits[b] = row - (1 - base)
    for c in bits:
        if c == '':
            return None
    return [bits[1], bits[0], bits[3], bits[2]]

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


class ListWidget(QtWidgets.QListWidget):
    def decode_data(self, bytearray):
        data = []
        ds = QtCore.QDataStream(bytearray)
        while not ds.atEnd():
            row = ds.readInt32()
            column = ds.readInt32()
            map_items = ds.readInt32()
            for i in range(map_items):
                key = ds.readInt32()
                value = QtCore.QVariant()
                ds >> value
                data.append(value.value())
        return data

    def __init__(self, parent=None):
        super(ListWidget, self).__init__(parent)
        self.setDragDropMode(self.DragDrop)
        self.setSelectionMode(self.ExtendedSelection)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            super(ListWidget, self).dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(QtCore.Qt.CopyAction)
            event.accept()
        else:
            super(ListWidget, self).dragMoveEvent(event)

    def dropEvent(self, event):
        if event.source() == self:
            event.setDropAction(QtCore.Qt.MoveAction)
            QtWidgets.QListWidget.dropEvent(self, event)
        else:
            ba = event.mimeData().data('application/x-qabstractitemmodeldatalist')
            data_items = self.decode_data(ba)
            event.setDropAction(QtCore.Qt.MoveAction)
            event.source().deleteItems(data_items)
            super(ListWidget, self).dropEvent(event)

    def deleteItems(self, items):
        for row in range(self.count() -1, -1, -1):
            if self.item(row).text() in items:
             #   r = self.row(item)
                self.takeItem(row)


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
        columns = []
        self.setup = [False, False]
        self.details = True
        self.book = None
        self.rows = None
        self.leapyear = False
        iper = '<none>'
        imax = 0
        self.alpha = 0.25
        self.title_font = 'size=x-large' #'size=15'
        self.label_font = '' #'size=x-large'
        self.legend_font = '' #'size=x-large'
        self.ticks_font = '' #'size=large'
        self.constrained_layout = False
        self.series = []
        self.xvalues = []
        self.palette = True
        self.history = None
        self.max_files = 10
        ifiles = self.get_flex_config()
        if len(ifiles) > 0:
            if self.history is None:
                self.history = sorted(ifiles.keys(), reverse=True)
            while ifile == '' and len(self.history) > 0:
                try:
                    ifile = ifiles[self.history[0]]
                except:
                    self.history.pop(0)
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
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Title:'), rw, 0)
        self.title = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.title, rw, 1, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Series:'), rw, 0)
        self.seriesi = CustomCombo()
        for series in self.series:
            self.seriesi.addItem(series)
        self.seriesi.setEditable(True)
        self.grid.addWidget(self.seriesi, rw, 1, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Cells for Series Categories; A1:B2 or r1,c1,r2,c2 format)'), rw, 3, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Series Label:'), rw, 0)
        self.ylabel = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.ylabel, rw, 1, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('X Values:'), rw, 0)
        self.xvaluesi = CustomCombo()
        for xvalues in self.xvalues:
            self.xvaluesi.addItem(xvalues)
        self.xvaluesi.setEditable(True)
        self.grid.addWidget(self.xvaluesi, rw, 1, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Cells for X values; A1:B2 or r1,c1,r2,c2 format)'), rw, 3, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('X Label:'), rw, 0)
        self.xlabel = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.xlabel, rw, 1, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Maximum:'), rw, 0)
        self.maxSpin = QtWidgets.QSpinBox()
        self.maxSpin.setRange(0, 100000)
        self.maxSpin.setSingleStep(500)
        self.grid.addWidget(self.maxSpin, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('(Handy if you want to produce a series of plots)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Type of Plot:'), rw, 0)
        plots = ['Bar Chart', 'Cumulative', 'Linegraph', 'Step Plot']
        self.plottype = QtWidgets.QComboBox()
        for plot in plots:
             self.plottype.addItem(plot)
        self.grid.addWidget(self.plottype, rw, 1) #, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Type of plot - stacked except for Linegraph)'), rw, 3, 1, 3)
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
     #   self.seriesi.textChanged.connect(self.seriesChanged)
        self.seriesi.activated[str].connect(self.seriesChanged)
        self.seriesi.currentIndexChanged.connect(self.seriesChanged)
        self.xvaluesi.activated[str].connect(self.xvaluesChanged)
        self.xvaluesi.currentIndexChanged.connect(self.xvaluesChanged)
  #      self.xvalues.textChanged.connect(self.somethingChanged)
        self.files.currentIndexChanged.connect(self.seriesChanged)
        self.sheet.currentIndexChanged.connect(self.sheetChanged)
        self.title.textChanged.connect(self.somethingChanged)
        self.maxSpin.valueChanged.connect(self.somethingChanged)
        self.plottype.currentIndexChanged.connect(self.somethingChanged)
        self.gridtype.currentIndexChanged.connect(self.somethingChanged)
        self.percentage.stateChanged.connect(self.somethingChanged)
        self.order.itemSelectionChanged.connect(self.somethingChanged)
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
        pp = QtWidgets.QPushButton('Plot', self)
        self.grid.addWidget(pp, rw, 1)
        pp.clicked.connect(self.ppClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('p'), self, self.ppClicked)
        cb = QtWidgets.QPushButton('Colours', self)
        self.grid.addWidget(cb, rw, 2)
        cb.clicked.connect(self.editColours)
        ep = QtWidgets.QPushButton('Preferences', self)
        self.grid.addWidget(ep, rw, 3)
        ep.clicked.connect(self.editIniFile)
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
        self.random_colours = True
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
                    self.label_font = value
                elif key == 'legend_font':
                    self.legend_font = value
                elif key == 'palette':
                    if value.lower() in ['false', 'no', 'off']:
                        self.palette = False
                elif key == 'random_colours' or key == 'random_colors':
                    if value.lower() in ['false', 'no', 'off']:
                        self.random_colours = False
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
                elif key == 'ticks_font':
                    self.ticks_font = value
                elif key == 'title_font':
                    self.title_font = value
        except:
            pass
        return ifiles

    def get_file_config(self, choice=''):
        ifile = ''
        ignore = True
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        columns = []
        isheet = ''
        try: # get list of files if any
            self.maxSpin.setValue(0)
            self.gridtype.setCurrentIndex(0)
            self.percentage.setCheckState(QtCore.Qt.Unchecked)
            items = config.items('Flexiplot')
            for key, value in items:
                if key == 'columns' + choice:
                    columns = charSplit(value)
                elif key == 'file' + choice:
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
                elif key == 'series' + choice:
                    self.seriesi.clear()
                    self.series = charSplit(value)
                    for series in self.series:
                        self.seriesi.addItem(series)
                elif key == 'maximum' + choice:
                    try:
                        self.maxSpin.setValue(int(value))
                    except:
                        self.maxSpin.setValue(0)
                elif key == 'sheet' + choice:
                    isheet = value
                elif key == 'title' + choice:
                    self.title.setText(value)
                elif key == 'xlabel' + choice:
                    self.xlabel.setText(value)
                elif key == 'xvalues' + choice:
                    self.xvaluesi.clear()
                    self.xvalues = charSplit(value)
                    for xvalues in self.xvalues:
                        self.xvaluesi.addItem(xvalues)
                elif key == 'ylabel' + choice:
                    self.ylabel.setText(value)
        except:
             pass
        self.columns = []
        if ifile != '':
            if self.book is not None:
                pxl.free_resources()
                self.book = None
            self.file.setText(ifile)
            if os.path.exists(ifile):
                self.setSheet(ifile, isheet)
            else:
                self.setSheet(self.scenarios + ifile, isheet)
            self.setColumns(isheet, columns=columns)
            for column in self.columns:
                self.check_colour(column, config, add=False)
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
            # find new entry
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
        self.log.setText('')
        if os.path.exists(self.file.text()):
            curfile = self.file.text()
        else:
            curfile = self.scenarios + self.file.text()
        newfile = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file', curfile)[0]
        if newfile != '':
            if self.book is not None:
                pxl.free_resources()
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
        self.popfileslist(self.files.currentText())
        self.log.setText('File "loaded"')
        self.setup[0] = False

    def somethingChanged(self):
        if not self.setup[0]:
            self.updated = True

    def setSheet(self, ifile, isheet):
        if self.book is None:
            try:
                self.book = pxl.get_book(file_name=ifile)
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
        self.log.setText('')
        if self.book is None:
            self.book = pxl.get_book(file_name=ifile)
        isheet = self.sheet.currentText()
        if isheet not in self.book.sheet_names():
            self.log.setText("Can't find sheet - " + isheet)
            return
        self.setColumns(isheet)
        self.updated = True

    def seriesChanged(self):
        self.log.setText('')
        series = self.seriesi.currentText()
        self.series = [series]
        for i in range(self.seriesi.count()):
            if self.seriesi.itemText(i) != series:
                self.series.append(self.seriesi.itemText(i))
        self.setColumns(self.sheet.currentText())
        self.updated = True

    def xvaluesChanged(self):
        self.log.setText('')
        xvalues = self.xvaluesi.currentText()
        self.xvalues = [xvalues]
        for i in range(self.xvaluesi.count()):
            if self.xvaluesi.itemText(i) != xvalues:
                self.xvalues.append(self.xvaluesi.itemText(i))
        self.updated = True

    def setColumns(self, isheet, columns=[]):
        try:
            ws = getattr(self.book, isheet.replace(' ', '_'))
        except:
            self.log.setText("Can't find sheet - " + isheet)
            return
        self.columns = []
        oldcolumns = []
        for col in range(self.order.count()):
            oldcolumns.append(self.order.item(col).text())
        self.order.clear()
        self.ignore.clear()
        try:
            roco = get_range(self.series[0])
        except:
            return
        if roco is None:
            return
        for row in range(roco[0], roco[2] + 1):
            for col in range(roco[1], roco[3] + 1):
                try:
                    column = str(ws[row, col]).replace('\n', ' ')
                except:
                    continue
                self.columns.append(column) # need order of columns
                if column in oldcolumns and column not in columns:
                    columns.append(column)
                if column in columns:
                    pass
                else:
                    self.ignore.addItem(column)
                    try:
                        self.ignore.item(self.ignore.count() - 1) \
                            .setBackground(QtGui.QColor(self.colours[column.lower()]))
                    except:
                        pass
        for column in columns:
            if column in self.columns:
                self.order.addItem(column)
                try:
                    self.order.item(self.order.count() - 1) \
                        .setBackground(QtGui.QColor(self.colours[column.lower()]))
                except:
                    pass
        self.updated = True

    def helpClicked(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.help,
                 title='Help for flexiplot (' + fileVersion() + ')', section='flexiplot')
        dialog.exec_()

    def doneClicked(self):
        if self.book is not None:
            pxl.free_resources()
        if not self.updated and not self.colours_updated:
            self.close()
        self.saveConfig()
        self.close()

    def closeEvent(self, event):
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
        dialr = EditSect('Flexiplot', curfldr, ini_file=self.config_file)
        ifiles = self.get_flex_config()

    def saveConfig(self):
        updates = {}
        if self.updated:
            config = configparser.RawConfigParser()
            config.read(self.config_file)
            try:
                choice = self.history[0]
            except:
                choice = ''
            save_file = self.file.text().replace(getUser(), '$USER$')
            try:
                self.max_files = int(config.get('Flexiplot', 'file_choices'))
            except:
                pass
            lines = []
            fix_history = []
            try:
                if len(self.history) > 0:
                    check_history = []
                    line = ''
                    for i in range(len(self.history)):
                        itm = self.history[i]
                        if self.files.itemText(i) in check_history:
                            fix_history.append([i, itm])
                        else:
                            check_history.append(self.files.itemText(i))
                            line += itm + ','
                    line = line[:-1]
                    lines.append('file_history=' + line)
            except:
                pass
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
            lines.append('file' + choice + '=' + self.file.text().replace(getUser(), '$USER$'))
            lines.append('grid' + choice + '=' + self.gridtype.currentText())
            lines.append('maximum' + choice + '=')
            if self.maxSpin.value() != 0:
                lines[-1] = lines[-1] + str(self.maxSpin.value())
            lines.append('percentage' + choice + '=')
            if self.percentage.isChecked():
                lines[-1] = lines[-1] + 'True'
            lines.append('plot' + choice + '=' + self.plottype.currentText())
            lines.append('sheet' + choice + '=' + self.sheet.currentText())
            line = 'series' + choice + '='
            for series in self.series:
                if series.find(',') >= 0:
                    line += "'" + series + "',"
                else:
                    line += series + ','
            lines.append(line[:-1])
            lines.append('title' + choice + '=' + self.title.text())
            lines.append('xlabel' + choice + '=' + self.xlabel.text())
            line = 'xvalues' + choice + '='
            for xvalues in self.xvalues:
                if xvalues.find(',') >= 0:
                    line += "'" + xvalues + "',"
                else:
                    line += xvalues + ','
            lines.append(line[:-1])
            lines.append('ylabel' + choice + '=' + self.ylabel.text())
            props = ['columns', 'file', 'grid', 'maximum', 'percentage', 'plot', 'sheet', 'series',
                     'title', 'xlabel', 'xvalues', 'ylabel']
            for i in range(len(fix_history) -1, -1, -1):
                self.files.removeItem(fix_history[i][0])
                self.history.pop(fix_history[i][0])
                for prop in props:
                    lines.append(prop + fix_history[i][1] + '=')
            updates['Flexiplot'] = lines
            if self.restorewindows and not config.has_section('Windows'): # new file set windows section
                updates['Windows'] = ['restorewindows=True']
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
                        palette=palette, underscore=True)
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
            col = self.order.item(c).text()
            try:
                self.order.item(c).setBackground(QtGui.QColor(self.colours[col.lower()]))
            except:
                pass
        for c in range(self.ignore.count()):
            col = self.ignore.item(c).text()
            try:
                self.ignore.item(c).setBackground(QtGui.QColor(self.colours[col.lower()]))
            except:
                pass

    def check_colour(self, colour, config, add=True):
        colr = colour.lower()
        if colr in self.colours.keys():
            return True
        elif self.random_colours:
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
        if self.order.count() == 0:
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
        del config
        self.log.setText('')
        i = self.file.text().rfind('/')
        if i > 0:
            matplotlib.rcParams['savefig.directory'] = self.file.text()[:i + 1]
        else:
            matplotlib.rcParams['savefig.directory'] = self.scenarios
        ws = getattr(self.book, isheet.replace(' ', '_'))
        x = []
        xlabels = []
        rocox = get_range(self.xvalues[0])
      #  print(rocox)
        data_in_cols = True
        if rocox[1] == rocox[3]:
            pass
        elif rocox[0] == rocox[2]:
            data_in_cols = False
        else:
            print('Assume in columns')
        ctr = 0
        if data_in_cols:
            if rocox[2] >= ws.number_of_rows():
                rocox[2] = ws.number_of_rows() - 1
            for row in range(rocox[0], rocox[2] + 1):
                x.append(ctr)
                ctr += 1
                try:
                    xlabels.append(str(int(ws[row, rocox[1]])))
                except:
                    xlabels.append(ws[row, rocox[1]])
        else:
            if rocox[3] >= ws.number_of_columns():
                rocox[3] = ws.number_of_columns() - 1
            for col in range(rocox[1], rocox[3] + 1):
                x.append(ctr)
                ctr += 1
                try:
                    xlabels.append(str(int(ws[rocox[0], col])))
                except:
                    xlabels.append(ws[rocox[0], col])
        data = []
        label = []
        miny = 0
        maxy = 0
        titl = self.title.text() #.replace('$YEAR$', str(year))
        titl = titl.replace('$MTH$', '')
        titl = titl.replace('$MONTH$', '')
        titl = titl.replace('  ', '')
        titl = titl.replace('Diurnal ', '')
        titl = titl.replace('Diurnal', '')
        titl = titl.replace('$SHEET$', isheet)
        roco = get_range(self.series[0])
        for c in range(self.order.count() -1, -1, -1):
            try:
                column = self.order.item(c).text()
                ndx = self.columns.index(column)
                if data_in_cols:
                    col = roco[1] + ndx
                else:
                    row = roco[0] + ndx
            except:
                continue
            data.append([])
            label.append(column)
            if data_in_cols:
                for row in range(rocox[0], rocox[2] + 1):
                    if ws[row, col] == '':
                        data[-1].append(0.)
                    else:
                        data[-1].append(ws[row, col])
                    miny = min(miny, data[-1][-1])
                    maxy = max(maxy, data[-1][-1])
            else:
                for col in range(rocox[1], rocox[3] + 1):
                    if ws[row, col] == '':
                        data[-1].append(0.)
                    else:
                        data[-1].append(ws[row, col])
                    miny = min(miny, data[-1][-1])
                    maxy = max(maxy, data[-1][-1])
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
        graph = fig.add_subplot(111)
        plt.title(titl, fontdict=font_props(self.title_font))
        if self.plottype.currentText() in ['Cumulative', 'Step Plot']:
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
                graph.fill_between(x, 0, values, label=label[0], color=self.colours[label[0].lower()], step=step)
                for c in range(1, len(data)):
                    for h in range(len(data[c])):
                        bottoms[h] = values[h]
                        values[h] = values[h] + data[c][h] / totals[h] * 100.
                    graph.fill_between(x, bottoms, values, label=label[c], color=self.colours[label[c].lower()], step=step)
                maxy = 100
            else:
                graph.fill_between(x, miny, data[0], label=label[0], color=self.colours[label[0].lower()], step=step)
                for c in range(1, len(data)):
                    for h in range(len(data[c])):
                        data[c][h] = data[c][h] + data[c - 1][h]
                        maxy = max(maxy, data[c][h])
                    graph.fill_between(x, data[c - 1], data[c], label=label[c], color=self.colours[label[c].lower()], step=step)
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
                graph.bar(x, values, label=label[0], color=self.colours[label[0].lower()])
                for c in range(1, len(data)):
                    for h in range(len(data[c])):
                        bottoms[h] = bottoms[h] + values[h]
                        values[h] = data[c][h] / totals[h] * 100.
                    graph.bar(x, values, bottom=bottoms, label=label[c], color=self.colours[label[c].lower()])
                maxy = 100
            else:
                graph.bar(x, data[0], label=label[0], color=self.colours[label[0].lower()])
                bottoms = [0.] * len(x)
                for c in range(1, len(data)):
                    for h in range(len(data[c])):
                        bottoms[h] = bottoms[h] + data[c - 1][h]
                        maxy = max(maxy, data[c][h] + bottoms[h])
                    graph.bar(x, data[c], bottom=bottoms, label=label[c], color=self.colours[label[c].lower()])
                if self.maxSpin.value() > 0:
                    maxy = self.maxSpin.value()
                else:
                    try:
                        rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                        maxy = ceil(maxy / rndup) * rndup
                    except:
                        pass
        else: #Linegraph
            for c in range(len(data)):
                mx = max(data[c])
                maxy = max(maxy, mx)
                graph.plot(x, data[c], linewidth=2.0, label=label[c], color=self.colours[label[c].lower()])
            if self.maxSpin.value() > 0:
                maxy = self.maxSpin.value()
            else:
                try:
                    rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                    maxy = ceil(maxy / rndup) * rndup
                except:
                    pass
        leg_font = font_props(self.legend_font, fontdict=False)
        graph.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=leg_font)
        plt.ylim([miny, maxy])
        plt.xlim([0, len(x) - 1])
        if self.sparse_ticks or len(self.sparse_ticks) > 0:
            # if self.plottype.currentText() == 'Linegraph' and len(x) > 24:
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
            graph.set_xticklabels(tick_labels, rotation='vertical', fontdict=font_props(self.ticks_font))
        else:
            plt.xticks(x, xlabels)
            graph.set_xticklabels(xlabels, rotation='vertical', fontdict=font_props(self.ticks_font))
        graph.set_xlabel(self.xlabel.text(), fontdict=font_props(self.label_font))
        graph.set_ylabel(self.ylabel.text(), fontdict=font_props(self.label_font))
        yticks = graph.get_yticklabels()
        graph.set_yticklabels(yticks, fontdict=font_props(self.ticks_font))
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
