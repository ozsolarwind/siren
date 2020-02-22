#!/usr/bin/python3
#
#  Copyright (C) 2020 Sustainable Energy Now Inc., Angus King
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
from PyQt4 import QtCore, QtGui
import sys
from math import log10, ceil
import matplotlib
from matplotlib.font_manager import FontProperties
import pylab as plt
import xlrd
import displayobject
from colours import Colours
from credits import fileVersion
from editini import SaveIni
from getmodels import getModelFile
from parents import getParents
from senuser import getUser, techClean
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

def get_range(text, alphabet=None):
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
        bits[1] = int(bits[1]) - 1
        bits[3] = int(bits[3]) - 1
    except:
        pass
    for b in (0, 2):
        row = 0
        ndx = 1
        for c in range(len(bits[b]) -1, -1, -1):
            ndx1 = alphabet.index(bits[b][c]) + 1
            row = row + ndx1 * ndx
            ndx = ndx * len(alphabet)
        bits[b] = row - 1
    for c in bits:
        if c == '':
            return None
    return [bits[1], bits[0], bits[3], bits[2]]


class ThumbListWidget(QtGui.QListWidget):
    def __init__(self, type, parent=None):
        super(ThumbListWidget, self).__init__(parent)
        self.setIconSize(QtCore.QSize(124, 124))
        self.setDragDropMode(QtGui.QAbstractItemView.DragDrop)
        self.setSelectionMode(QtGui.QAbstractItemView.ExtendedSelection)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            super(ThumbListWidget, self).dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(QtCore.Qt.CopyAction)
            event.accept()
        else:
            super(ThumbListWidget, self).dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(QtCore.Qt.CopyAction)
            event.accept()
            links = []
            for url in event.mimeData().urls():
                links.append(str(url.toLocalFile()))
            self.emit(QtCore.SIGNAL('dropped'), links)
        else:
            event.setDropAction(QtCore.Qt.MoveAction)
            super(ThumbListWidget, self).dropEvent(event)


class ClickableQLabel(QtGui.QLabel):
    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtGui.QApplication.widgetAt(event.globalPos()).setFocus()
        self.emit(QtCore.SIGNAL('clicked()'))


class FlexiPlot(QtGui.QWidget):

    def __init__(self, help='help.html'):
        super(FlexiPlot, self).__init__()
        self.help = help
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            self.config_file = sys.argv[1]
        else:
            self.config_file = getModelFile('flexiplot.ini')
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
        self.series = '0,0,0,0'
        self.ylabel = ''
        self.xvalues = '0,0,0,0'
        self.palette = True
        self.history = None
        self.max_files = 10
        ifiles = {}
        try: # get alpha
            self.alpha = float(config.get('Powerplot', 'alpha'))
        except:
            pass
        try: # get list of files if any
            items = config.items('Powerplot')
            for key, value in items:
                if key == 'file_history':
                    self.history = value.split(',')
                elif key == 'file_choices':
                    self.max_files = int(value)
                elif key[:4] == 'file':
                    ifiles[key[4:]] = value.replace('$USER$', getUser())
                elif key == 'palette':
                    if value.lower() in ['false', 'no', 'off']:
                        self.palette = False
        except:
            pass
        if len(ifiles) > 0:
            if self.history is None:
                self.history = sorted(ifiles.keys(), reverse=True)
            ifile = ifiles[self.history[0]]
        matplotlib.rcParams['savefig.directory'] = os.getcwd()
        self.grid = QtGui.QGridLayout()
        self.updated = False
        self.colours_updated = False
        self.log = QtGui.QLabel('')
        rw = 0
        self.grid.addWidget(QtGui.QLabel('Recent Files:'), rw, 0)
        self.files = QtGui.QComboBox()
        if ifile != '':
            self.popfileslist(ifile, ifiles)
        self.grid.addWidget(self.files, rw, 1, 1, 5)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('File:'), rw, 0)
        self.file = ClickableQLabel()
        self.file.setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
        self.file.setText('')
        self.grid.addWidget(self.file, rw, 1, 1, 5)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Sheet:'), rw, 0)
        self.sheet = QtGui.QComboBox()
        self.grid.addWidget(self.sheet, rw, 1, 1, 2)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Series:'), rw, 0)
        self.seriesi = QtGui.QLineEdit(self.series)
        self.grid.addWidget(self.seriesi, rw, 1, 1, 2)
        self.grid.addWidget(QtGui.QLabel('(Cells for Series Categories; A1:B2 or r1,c1,r2,c2 format)'), rw, 3, 1, 2)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('X Values:'), rw, 0)
        self.xvalues = QtGui.QLineEdit('')
        self.grid.addWidget(self.xvalues, rw, 1, 1, 2)
        self.grid.addWidget(QtGui.QLabel('(Cells for X values; A1:B2 or r1,c1,r2,c2 format)'), rw, 3, 1, 2)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Title:'), rw, 0)
        self.title = QtGui.QLineEdit('')
        self.grid.addWidget(self.title, rw, 1, 1, 2)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Maximum:'), rw, 0)
        self.maxSpin = QtGui.QSpinBox()
        self.maxSpin.setRange(0, 6000)
        self.maxSpin.setSingleStep(500)
        self.grid.addWidget(self.maxSpin, rw, 1)
        self.grid.addWidget(QtGui.QLabel('(Handy if you want to produce a series of plots)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Cumulative:'), rw, 0)
        self.cumulative = QtGui.QCheckBox()
        self.cumulative.setCheckState(QtCore.Qt.Checked)
        self.grid.addWidget(self.cumulative, rw, 1) #, 1, 2)
        self.grid.addWidget(QtGui.QLabel('(Check for hourly generation profile)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Percentage:'), rw, 0)
        self.percentage = QtGui.QCheckBox()
        self.percentage.setCheckState(QtCore.Qt.Unchecked)
        self.grid.addWidget(self.percentage, rw, 1) #, 1, 2)
        self.grid.addWidget(QtGui.QLabel('(Check for percentage distribution)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Column Order:\n(move to right\nto exclude)'), rw, 0)
        self.order = ThumbListWidget(self)
        self.grid.addWidget(self.order, rw, 1, 1, 2)
        self.ignore = ThumbListWidget(self)
        self.grid.addWidget(self.ignore, rw, 3, 1, 2)
        self.grid.addWidget(QtGui.QLabel(' '), rw, 5)
        if ifile != '':
            self.get_file_config(self.history[0])
        self.files.currentIndexChanged.connect(self.filesChanged)
        self.connect(self.file, QtCore.SIGNAL('clicked()'), self.fileChanged)
        self.seriesi.textChanged.connect(self.seriesChanged)
        self.xvalues.textChanged.connect(self.somethingChanged)
        self.files.currentIndexChanged.connect(self.seriesChanged)
        self.sheet.currentIndexChanged.connect(self.sheetChanged)
        self.title.textChanged.connect(self.somethingChanged)
        self.maxSpin.valueChanged.connect(self.somethingChanged)
        self.cumulative.stateChanged.connect(self.somethingChanged)
        self.percentage.stateChanged.connect(self.somethingChanged)
        self.order.itemSelectionChanged.connect(self.somethingChanged)
        rw += 1
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.log.setPalette(msg_palette)
        self.grid.addWidget(self.log, rw, 1, 1, 4)
        rw += 1
        quit = QtGui.QPushButton('Done', self)
        self.grid.addWidget(quit, rw, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        pp = QtGui.QPushButton('Plot', self)
        self.grid.addWidget(pp, rw, 1)
        pp.clicked.connect(self.ppClicked)
        QtGui.QShortcut(QtGui.QKeySequence('p'), self, self.ppClicked)
        cb = QtGui.QPushButton('Colours', self)
        self.grid.addWidget(cb, rw, 2)
        cb.clicked.connect(self.editColours)
        help = QtGui.QPushButton('Help', self)
        self.grid.addWidget(help, rw, 4)
        help.clicked.connect(self.helpClicked)
        QtGui.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - flexiplot (' + fileVersion() + ') - FlexiPlot')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.center()
        self.resize(int(self.sizeHint().width() * 1.07), int(self.sizeHint().height() * 1.07))
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
        centerPoint = QtGui.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def get_file_config(self, choice=''):
        ifile = ''
        ignore = True
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        columns = []
        try: # get list of files if any
            items = config.items('Powerplot')
            for key, value in items:
                if key == 'columns' + choice:
                    columns = charSplit(value)
                elif key == 'cumulative' + choice:
                    if value.lower() in ['false', 'no', 'off']:
                        self.cumulative.setCheckState(QtCore.Qt.Unchecked)
                    else:
                        self.cumulative.setCheckState(QtCore.Qt.Checked)
                elif key == 'file' + choice:
                    ifile = value.replace('$USER$', getUser())
                elif key == 'percentage' + choice:
                    if value.lower() in ['true', 'yes', 'on']:
                        self.percentage.setCheckState(QtCore.Qt.Checked)
                    else:
                        self.percentage.setCheckState(QtCore.Qt.Unchecked)
                elif key == 'series' + choice:
                    self.series = value
                    self.seriesi.setText(value)
                elif key == 'maximum' + choice:
                    try:
                        self.maxSpin.setValue(int(value))
                    except:
                        self.maxSpin.setValue(0)
                elif key == 'sheet' + choice:
                    isheet = value
                elif key == 'xvalues' + choice:
                    self.xvalues.setText(value)
                elif key == 'ylabel' + choice:
                    self.ylabel = value
                elif key == 'title' + choice:
                    self.title.setText(value)
        except:
             pass
        if ifile != '':
            if self.book is not None:
                self.book.release_resources()
                self.book = None
            self.file.setText(ifile)
            if os.path.exists(ifile):
                self.setSheet(ifile, isheet)
            else:
                self.setSheet(self.scenarios + ifile, isheet)
            self.setColumns(isheet, columns=columns)
            for column in columns:
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
                if len(self.history) >= self.max_files:
                    self.history.insert(0, self.history.pop(-1)) # make last entry first
                else:
                    self.history.insert(0, str(len(self.history)))
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
        newfile = str(QtGui.QFileDialog.getOpenFileName(self, 'Open file', curfile))
        if newfile != '':
            if self.book is not None:
                self.book.release_resources()
                self.book = None
            isheet = str(self.sheet.currentText())
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
                self.book = xlrd.open_workbook(ifile, on_demand=True)
            except:
                self.log.setText("Can't open file - " + ifile)
                return
        ndx = 0
        self.sheet.clear()
        j = -1
        for sht in self.book.sheet_names():
            j += 1
            self.sheet.addItem(str(sht))
            if str(sht) == isheet:
                ndx = j
        self.sheet.setCurrentIndex(ndx)

    def sheetChanged(self):
        self.log.setText('')
        if self.book is None:
            self.book = xlrd.open_workbook(newfile, on_demand=True)
        isheet = str(self.sheet.currentText())
        if isheet not in self.book.sheet_names():
            self.log.setText("Can't find sheet - " + isheet)
            return
        self.setColumns(isheet)
        self.updated = True

    def seriesChanged(self):
        self.log.setText('')
        series = self.seriesi.text()
        if series != self.series:
            self.series = series
            self.setColumns(str(self.sheet.currentText()))
        self.updated = True

    def setColumns(self, isheet, columns=[]):
        try:
            ws = self.book.sheet_by_name(isheet)
        except:
            self.log.setText("Can't find sheet - " + isheet)
            return
        self.columns = []
        oldcolumns = []
        for col in range(self.order.count()):
            oldcolumns.append(self.order.item(col).text())
        self.order.clear()
        self.ignore.clear()
        roco = get_range(self.series)
        if roco is None:
            return
        for row in range(roco[0], roco[2] + 1):
            for col in range(roco[1], roco[3] + 1):
                column = str(ws.cell_value(row, col)).replace('\n',' ')
                self.columns.append(column) # need order of columns
                if column in oldcolumns and column not in columns:
                    columns.append(column)
                if column in columns:
                    pass
                else:
                    self.ignore.addItem(column)
                    try:
                        self.ignore.item(self.ignore.count() - \
                         1).setBackground(QtGui.QColor(self.colours[column.lower()]))
                    except:
                        pass
        for column in columns:
            self.order.addItem(column)
            try:
                self.order.item(self.order.count() - 1).setBackground(QtGui.QColor(self.colours[column.lower()]))
            except:
                pass
        self.updated = True

    def helpClicked(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.help,
                 title='Help for flexiplot (' + fileVersion() + ')', section='flexiplot')
        dialog.exec_()

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
            save_file = str(self.file.text()).replace(getUser(), '$USER$')
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
            cols = 'columns' + choice + '='
            for col in range(self.order.count()):
                try:
                    if str(self.order.item(col).text()).index(',') >= 0:
                        try:
                            if str(self.order.item(col).text()).index("'") >= 0:
                                qte = '"'
                        except:
                            qte = "'"
                except:
                    qte = ''
                cols += qte + str(self.order.item(col).text()) + qte + ','
            if cols[-1] != '=':
                cols = cols[:-1]
            lines.append(cols)
            lines.append('cumulative' + choice + '=')
            if not self.cumulative.isChecked():
                lines[-1] = lines[-1] + 'False'
            lines.append('file' + choice + '=' + str(self.file.text()).replace(getUser(), '$USER$'))
            lines.append('maximum' + choice + '=')
            if self.maxSpin.value() != 0:
                lines[-1] = lines[-1] + str(self.maxSpin.value())
            lines.append('percentage' + choice + '=')
            if self.percentage.isChecked():
                lines[-1] = lines[-1] + 'True'
            lines.append('sheet' + choice + '=' + str(self.sheet.currentText()))
            lines.append('series' + choice + '=' + self.series)
            lines.append('title' + choice + '=' + self.title.text())
            lines.append('xvalues' + choice + '=' + self.xvalues.text())
            lines.append('ylabel' + choice + '=' + self.ylabel)
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
            col = str(self.order.item(c).text())
            try:
                self.order.item(c).setBackground(QtGui.QColor(self.colours[col.lower()]))
            except:
                pass
        for c in range(self.ignore.count()):
            col = str(self.ignore.item(c).text())
            try:
                self.ignore.item(c).setBackground(QtGui.QColor(self.colours[col.lower()]))
            except:
                pass

    def check_colour(self, colour, config, add=True):
        colr = colour.lower()
        if colr in self.colours.keys():
            return True
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

    def ppClicked(self):
        def format_func(value, tick_number):
            if self.percentage.isChecked():
                return '{:.0f}%'.format(value)
            else:
                return '{:,.0f}'.format(value)

        if self.book is None:
            self.log.setText('Error accessing Workbook.')
            return
        if self.order.count() == 0:
            self.log.setText('Nothing to plot.')
            return
        isheet = str(self.sheet.currentText())
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
        ws = self.book.sheet_by_name(isheet)
        x = []
        xlabels = []
        rocox = get_range(self.xvalues.text())
        ctr = 0
        for row in range(rocox[0], rocox[2] + 1):
            x.append(ctr)
            ctr += 1
            xlabels.append(str(int(ws.cell_value(row, rocox[1]))))
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
        roco = get_range(self.series)
        for c in range(self.order.count() -1, -1, -1):
            try:
                column = str(self.order.item(c).text())
                ndx = self.columns.index(column)
                col = roco[1] + ndx
            except:
                continue
            data.append([])
            label.append(column)
            for row in range(rocox[0], rocox[2] + 1):
                if ws.cell_value(row, col) == '':
                    data[-1].append(0.)
                else:
                    data[-1].append(ws.cell_value(row, col))
                miny = min(miny, data[-1][-1])
                maxy = max(maxy, data[-1][-1])
        if not self.cumulative.isChecked():
            fig = plt.figure('linegraph')
            plt.grid(True)
            ax = fig.add_subplot(111)
            plt.title(titl)
            for c in range(len(data)):
                ax.plot(x, data[c], linewidth=1.0, label=label[c], color=self.colours[label[c].lower()])
            if self.maxSpin.value() > 0:
                maxy = self.maxSpin.value()
            else:
                try:
                    rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                    maxy = ceil(maxy / rndup) * rndup
                except:
                    pass
            lbl_font = FontProperties()
            lbl_font.set_size('small')
            ax.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 1),
                          prop=lbl_font)
            plt.ylim([miny, maxy])
            plt.xlim([0, len(x) - 1])
            plt.xticks(x, xlabels)
            ax.set_xticklabels(xlabels, rotation='vertical')
            ax.yaxis.set_major_formatter(plt.FuncFormatter(format_func))
            ax.set_ylabel(self.ylabel)
            zp = ZoomPanX()
            f = zp.zoom_pan(ax, base_scale=1.2) # enable scrollable zoom
            plt.show()
            del zp
        else:
            fig = plt.figure('cumulative') # + str(year))
            plt.grid(True)
            bx = fig.add_subplot(111)
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
                bx.fill_between(x, 0, values, label=label[0], color=self.colours[label[0].lower()])
                for c in range(1, len(data)):
                    for h in range(len(data[c])):
                        bottoms[h] = values[h]
                        values[h] = values[h] + data[c][h] / totals[h] * 100.
                    bx.fill_between(x, bottoms, values, label=label[c], color=self.colours[label[c].lower()])
                maxy = 100
            else:
                bx.fill_between(x, miny, data[0], label=label[0], color=self.colours[label[0].lower()])
                for c in range(1, len(data)):
                    for h in range(len(data[c])):
                        data[c][h] = data[c][h] + data[c - 1][h]
                        maxy = max(maxy, data[c][h])
                    bx.fill_between(x, data[c - 1], data[c], label=label[c], color=self.colours[label[c].lower()])
                top = data[0][:]
                for d in range(1, len(data)):
                    for h in range(len(top)):
                        top[h] = max(top[h], data[d][h])
                bx.plot(x, top, color='white')
                if self.maxSpin.value() > 0:
                    maxy = self.maxSpin.value()
                else:
                    try:
                        rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                        maxy = ceil(maxy / rndup) * rndup
                    except:
                        pass
            bx.set_ylabel(self.ylabel)
            lbl_font = FontProperties()
            lbl_font.set_size('small')
            bx.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=lbl_font)
            plt.ylim([miny, maxy])
            plt.xlim([0, len(x) - 1])
            plt.xticks(x, xlabels)
            bx.set_xticklabels(xlabels, rotation='vertical')
            bx.yaxis.set_major_formatter(plt.FuncFormatter(format_func))
            zp = ZoomPanX()
            f = zp.zoom_pan(bx, base_scale=1.2) # enable scrollable zoom
            plt.show()
            del zp


if "__main__" == __name__:
    app = QtGui.QApplication(sys.argv)
    ex = FlexiPlot()
    app.exec_()
    app.deleteLater()
    sys.exit()
