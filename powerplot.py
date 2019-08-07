#!/usr/bin/python3
#
#  Copyright (C) 2019 Sustainable Energy Now Inc., Angus King
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
from parents import getParents
from senuser import getUser, techClean
from zoompan import ZoomPanX


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
            self.emit(QtCore.SIGNAL("dropped"), links)
        else:
            event.setDropAction(QtCore.Qt.MoveAction)
            super(ThumbListWidget, self).dropEvent(event)


class ClickableQLabel(QtGui.QLabel):
    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtGui.QApplication.widgetAt(event.globalPos()).setFocus()
        self.emit(QtCore.SIGNAL('clicked()'))


class PowerPlot(QtGui.QWidget):

    def __init__(self, help='help.html'):
        super(PowerPlot, self).__init__()
        self.help = help
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'powerplot.ini'
        config.read(config_file)
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
        self.details = True
        self.book = None
        self.toprow = None
        self.rows = None
        self.leapyear = False
        iper = '<none>'
        imax = 0
        self.alpha = 0.25
        self.target = ''
        ititle = ''
        icum = True
        try:
             items = config.items('Powerplot')
             for key, value in items:
                 if key == 'alpha':
                     try:
                         self.alpha = float(value)
                     except:
                         pass
                 elif key == 'columns':
                     bits = value.split(',')
                     columns = []
                     for bit in bits:
                         columns.append(bit)
                 elif key == 'cumulative':
                     if value.lower() in ['false', 'no', 'off']:
                         icum = False
                 elif key == 'file':
                     ifile = value.replace('$USER$', getUser())
                 elif key == 'period':
                     iper = value
                 elif key == 'maximum':
                     try:
                         imax = int(value)
                     except:
                         pass
                 elif key == 'sheet':
                     isheet = value
                 elif key == 'target':
                     try:
                         self.target = value
                     except:
                         pass
                 elif key == 'title':
                     ititle = value
        except:
            pass
        for column in columns:
            self.check_colour(column, config, add=False)
        matplotlib.rcParams['savefig.directory'] = os.getcwd()
        self.grid = QtGui.QGridLayout()
        self.updated = False
        self.colours_updated = False
        self.log = QtGui.QLabel('')
        rw = 0
        self.grid.addWidget(QtGui.QLabel('File:'), rw, 0)
        self.file = ClickableQLabel()
        self.file.setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
        self.file.setText(ifile)
        self.connect(self.file, QtCore.SIGNAL('clicked()'), self.fileChanged)
        self.grid.addWidget(self.file, rw, 1, 1, 5)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Sheet:'), 1, 0)
        self.sheet = QtGui.QComboBox()
        self.grid.addWidget(self.sheet, rw, 1, 1, 2)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Period:'), rw, 0)
        self.period = QtGui.QComboBox()
        self.period.addItem('<none>')
    #    if iper == (self.period.currentText()))
        self.period.addItem('Year')
        for mth in mth_labels:
            self.period.addItem(mth)
        i = self.period.findText(iper, QtCore.Qt.MatchExactly)
        if i >=0 :
            self.period.setCurrentIndex(i)
        self.grid.addWidget(self.period, rw, 1, 1, 2)
        self.period.currentIndexChanged.connect(self.somethingChanged)
        self.grid.addWidget(QtGui.QLabel('(Diurnal profile for Period)'), rw, 3, 1, 2)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Target:'), rw, 0)
        self.targets = QtGui.QComboBox()
        self.grid.addWidget(self.targets, rw, 1, 1, 2)
        self.grid.addWidget(QtGui.QLabel('(e.g. Load)'), rw, 3, 1, 2)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Title:'), rw, 0)
        self.title = QtGui.QLineEdit(ititle)
        self.title.textChanged.connect(self.somethingChanged)
        self.grid.addWidget(self.title, rw, 1, 1, 2)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Maximum:'), rw, 0)
        self.maxSpin = QtGui.QSpinBox()
        self.maxSpin.setRange(0, 6000)
        self.maxSpin.setSingleStep(500)
        self.maxSpin.setValue(imax)
        self.grid.addWidget(self.maxSpin, rw, 1)
        self.maxSpin.valueChanged.connect(self.somethingChanged)
        self.grid.addWidget(QtGui.QLabel('(Handy if you want to produce a series of plots)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Cumulative:'), rw, 0)
        self.cumulative = QtGui.QCheckBox()
        if icum:
            self.cumulative.setCheckState(QtCore.Qt.Checked)
        self.cumulative.stateChanged.connect(self.somethingChanged)
        self.grid.addWidget(self.cumulative, rw, 1, 1, 2)
        self.grid.addWidget(QtGui.QLabel('(Check for hourly generation profile)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtGui.QLabel('Column Order:\n(move to right\nto exclude)'), rw, 0)
        self.order = ThumbListWidget(self)
        self.order.itemSelectionChanged.connect(self.somethingChanged)
        self.grid.addWidget(self.order, rw, 1, 1, 2)
        self.ignore = ThumbListWidget(self)
        self.grid.addWidget(self.ignore, rw, 3, 1, 2)
        self.grid.addWidget(QtGui.QLabel(' '), rw, 5)
        if ifile != '':
            if os.path.exists(ifile):
                self.setSheet(ifile, isheet)
            else:
                self.setSheet(self.scenarios + ifile, isheet)
            self.setColumns(isheet, columns=columns)
        self.sheet.currentIndexChanged.connect(self.sheetChanged)
        self.targets.currentIndexChanged.connect(self.targetChanged)
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
        self.setWindowTitle('SIREN - powerplot (' + fileVersion() + ') - PowerPlot')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.center()
        self.resize(int(self.sizeHint().width()* 1.07), int(self.sizeHint().height() * 1.07))
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
        centerPoint = QtGui.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

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
                self.toprow = None
            isheet = str(self.sheet.currentText())
            self.setSheet(newfile, isheet)
            if newfile[: len(self.scenarios)] == self.scenarios:
                self.file.setText(newfile[len(self.scenarios):])
            else:
                self.file.setText(newfile)
            self.updated = True

    def somethingChanged(self):
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
        self.toprow = None
        if self.book is None:
            self.book = xlrd.open_workbook(newfile, on_demand=True)
        isheet = str(self.sheet.currentText())
        if isheet not in self.book.sheet_names():
            self.log.setText("Can't find sheet - " + isheet)
            return
        self.setColumns(isheet)
        self.updated = True

    def targetChanged(self):
        self.log.setText('')
        target = str(self.targets.currentText())
        if target != self.target:
            if target == '<none>':
                if self.target != '':
                    self.ignore.addItem(self.target)
                    try:
                        self.ignore.item(self.ignore.count() - 1).setBackground(QtGui.QColor(self.colours[self.target.lower()]))
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
                    try:
                        self.ignore.item(self.ignore.count() - 1).setBackground(QtGui.QColor(self.colours[self.target.lower()]))
                    except:
                        pass
                self.target = target
        self.updated = True

    def setColumns(self, isheet, columns=[]):
        try:
            ws = self.book.sheet_by_name(isheet)
        except:
            self.log.setText("Can't find sheet - " + isheet)
            return
        row = 0
        while row < ws.nrows:
            if str(ws.cell_value(row, 0)) == 'Hour':
                if str(ws.cell_value(row, 1)) != 'Period':
                    self.log.setText(isheet + ' sheet format incorrect')
                    return
                self.toprow = row
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
                for col in range(2, ws.ncols):
                    column = str(ws.cell_value(row, col)).replace('\n',' ')
                    if column in oldcolumns:
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
                            self.ignore.addItem(column)
                            try:
                                self.ignore.item(self.ignore.count() - 1).setBackground(QtGui.QColor(self.colours[column.lower()]))
                            except:
                                pass
                for column in columns:
                    if column != self.target:
                        self.order.addItem(column)
                        try:
                            self.order.item(self.order.count() - 1).setBackground(QtGui.QColor(self.colours[column.lower()]))
                        except:
                            pass
                break
            row += 1
        else:
            self.log.setText(isheet + ' sheet format incorrect')

    def helpClicked(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.help,
                 title='Help for powerplot (' + fileVersion() + ')', section='powerplot')
        dialog.exec_()

    def quitClicked(self):
        if self.book is not None:
            self.book.release_resources()
        if not self.updated and not self.colours_updated:
            self.close()
        updates = {}
        if self.updated:
            lines = []
            lines.append('file=' + str(self.file.text()).replace(getUser(), '$USER$'))
            lines.append('sheet=' + str(self.sheet.currentText()))
            lines.append('period=')
            if str(self.period.currentText()) != '<none>':
                lines[-1] = lines[-1] + str(self.period.currentText())
            lines.append('target=' + self.target)
            lines.append('title=' + self.title.text())
            lines.append('maximum=')
            if self.maxSpin.value() != 0:
                lines[-1] = lines[-1] + str(self.maxSpin.value())
            lines.append('cumulative=')
            if not self.cumulative.isChecked():
                lines[-1] = lines[-1] + 'False'
            cols = 'columns='
            for col in range(self.order.count()):
                cols += str(self.order.item(col).text()) + ','
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
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'powerplot.ini'
        SaveIni(updates, ini_file=config_file)
        self.close()

    def editColours(self, color=False):
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'powerplot.ini'
        dialr = Colours(section='Plot Colors', ini_file=config_file, add_colour=color)
        dialr.exec_()
        self.colours = {}
        config = configparser.RawConfigParser()
        config.read(config_file)
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
            config.read(config_file)
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
        ws = self.book.sheet_by_name(isheet)
        if self.toprow is None:
            row = 0
            while row < ws.nrows:
                if str(ws.cell_value(row, 0)) == 'Hour':
                    if str(ws.cell_value(row, 1)) != 'Period':
                        self.log.setText(isheet + ' sheet format incorrect')
                        return
                    self.toprow = row
                    self.rows = ws.nrows - row + 1
                    break
        try:
            year = int(ws.cell_value(self.toprow + 1, 1)[:4])
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
        if str(self.period.currentText()) == '<none>': # full year of hourly figures
            m = 0
            d = 1
            day_labels = []
            while m < len(the_days):
                day_labels.append('%s %s' % (str(d), mth_labels[m]))
                d += 7
                if d > the_days[m]:
                    d = d - the_days[m]
                    m += 1
            x = []
            len_x = self.rows
            for i in range(len_x):
                x.append(i)
            load = []
            tgt_col = -1
            data = []
            label = []
            maxy = 0
            titl = self.title.text().replace('$YEAR$', str(year))
            for c in range(self.order.count()):
                col = str(self.order.item(c).text())
                for c2 in range(2, ws.ncols):
                    column = str(ws.cell_value(self.toprow, c2)).replace('\n',' ')
                    if column == col:
                        data.append([])
                        label.append(column)
                        for row in range(self.toprow + 1, self.toprow + self.rows + 1):
                            data[-1].append(ws.cell_value(row, c2))
                            maxy = max(maxy, data[-1][-1])
                        break
                    elif column == self.target:
                        tgt_col = c2
            if tgt_col >= 0:
                for row in range(self.toprow + 1, self.toprow + self.rows + 1):
                    load.append(ws.cell_value(row, tgt_col))
                    maxy = max(maxy, load[-1])
            if not self.cumulative.isChecked():
                fig = plt.figure('linegraph_' + str(year))
                plt.grid(True)
                ax = fig.add_subplot(111)
                plt.title(titl)
                for c in range(len(data)):
                    ax.plot(x, data[c], linewidth=1.0, label=label[c], color=self.colours[label[c].lower()])
                if len(load) > 0:
                    ax.plot(x, load, linewidth=2.0, label=self.target, color=self.colours[self.target.lower()])
                if self.maxSpin.value() > 0:
                    maxy = self.maxSpin.value()
                else:
                    try:
                        rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                        maxy = ceil(maxy / rndup) * rndup
                    except:
                        pass
                miny = 0
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                ax.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 1),
                              prop=lbl_font)
                plt.ylim([miny, maxy])
                plt.xlim([0, len(x)])
                plt.xticks(list(range(12, len(x), 168)))
                ax.set_xticklabels(day_labels, rotation='vertical')
                ax.set_xlabel('Month of the year')
                ax.set_ylabel('Power (MW)')
                zp = ZoomPanX()
                f = zp.zoom_pan(ax, base_scale=1.2) # enable scrollable zoom
                plt.show()
                del zp
            else:
                fig = plt.figure('cumulative_' + str(year))
                plt.grid(True)
                bx = fig.add_subplot(111)
                plt.title(titl)
                if self.target == '<none>':
                    bx.fill_between(x, 0, data[0], label=label[0], color=self.colours[label[0].lower()])
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
                else:
                    pattern = ['-', '+', 'x', '\\', '*', 'o', 'O', '.']
                    pat = 0
                    full = []
                    for h in range(len(load)):
                       full.append(min(load[h], data[0][h]))
                    bx.fill_between(x, 0, full, label=label[0], color=self.colours[label[0].lower()])
                    bx.fill_between(x, full, data[0], alpha=self.alpha, color=self.colours[label[0].lower()]) #, hatch=pattern[0])
                    for c in range(1, len(data)):
                        full = []
                        for h in range(len(data[c])):
                            data[c][h] = data[c][h] + data[c - 1][h]
                            maxy = max(maxy, data[c][h])
                            full.append(max(min(load[h], data[c][h]), data[c - 1][h]))
                        bx.fill_between(x, data[c - 1], full, label=label[c], color=self.colours[label[c].lower()])
                        pat += 1
                        if pat >= len(pattern):
                            pat = 0
                        bx.fill_between(x, full, data[c], alpha=self.alpha, color=self.colours[label[c].lower()]) #,
                                     #   hatch=pattern[pat])
                    top = data[0][:]
                    for d in range(1, len(data)):
                        for h in range(len(top)):
                            top[h] = max(top[h], data[d][h])
                    bx.plot(x, top, color='white')
                    short = []
                    for h in range(len(load)):
                        short.append(max(data[c][h], load[h]))
                    bx.fill_between(x, data[c], short, label='Shortfall', color=self.colours['shortfall'])
                    bx.plot(x, load, linewidth=2.0, label=self.target, color=self.colours[self.target.lower()])
                    if self.maxSpin.value() > 0:
                        maxy = self.maxSpin.value()
                    else:
                        try:
                            rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                            maxy = ceil(maxy / rndup) * rndup
                        except:
                            pass
                miny = 0
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                bx.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=lbl_font)
                plt.ylim([miny, maxy])
                plt.xlim([0, len(x)])
                plt.xticks(list(range(12, len(x), 168)))
                bx.set_xticklabels(day_labels, rotation='vertical')
                bx.set_xlabel('Month of the year')
                bx.set_ylabel('Power (MW)')
                zp = ZoomPanX()
                f = zp.zoom_pan(bx, base_scale=1.2) # enable scrollable zoom
                plt.show()
                del zp
        else: # diurnal average
            hr_labels = ['0:00', '4:00', '8:00', '12:00', '16:00', '20:00', '23:00']
            ticks = list(range(0, 21, 4))
            ticks.append(23)
            titl = self.title.text().replace('$YEAR$', str(year))
            if self.period.currentText() == 'Year':
                strt_row = self.toprow
                todo_rows = self.rows
            else:
                titl = titl.replace('$MTH$', self.period.currentText())
                titl = titl.replace('$MONTH$', self.period.currentText())
                i = mth_labels.index(self.period.currentText())
                m = 0
                strt_row = 0
                while m < i:
                    strt_row = strt_row + the_days[m] * 24
                    m += 1
                strt_row = self.toprow + strt_row
                todo_rows = the_days[i] * 24
            load = []
            tgt_col = -1
            data = []
            label = []
            maxy = 0
            hs = []
            for h in range(24):
                hs.append(h)
            x = hs[:]
            for c in range(self.order.count()):
                col = str(self.order.item(c).text())
                for c2 in range(2, ws.ncols):
                    column = str(ws.cell_value(self.toprow, c2)).replace('\n',' ')
                    if column == col:
                        data.append([])
                        data[-1] = hs[:]
                        label.append(column)
                        h = 0
                        for row in range(strt_row + 1, strt_row + todo_rows + 1):
                            data[-1][h] = data[-1][h] + ws.cell_value(row, c2)
                            h += 1
                            if h >= 24:
                                h = 0
                        for h in range(24):
                            data[-1][h] = data[-1][h] / (todo_rows / 24)
                            maxy = max(maxy, data[-1][h])
                        break
                    elif column == self.target:
                        tgt_col = c2
            if tgt_col >= 0:
                load = hs[:]
                h = 0
                for row in range(strt_row + 1, strt_row + todo_rows + 1):
                    load[h] = load[h] + ws.cell_value(row, tgt_col)
                    h += 1
                    if h >= 24:
                        h = 0
                for h in range(24):
                    load[h] = load[h] / (todo_rows / 24)
                    maxy = max(maxy, load[h])
            if not self.cumulative.isChecked():
                fig = plt.figure('linegraph_' + str(year) + '_' + self.period.currentText())
                plt.grid(True)
                cx = fig.add_subplot(111)
                plt.title(titl)
                for c in range(len(data)):
                    cx.plot(x, data[c], linewidth=1.0, label=label[c], color=self.colours[label[c].lower()])
                if len(load) > 0:
                    cx.plot(x, load, linewidth=2.0, label=self.target, color=self.colours[self.target.lower()])
                if self.maxSpin.value() > 0:
                    maxy = self.maxSpin.value()
                else:
                    try:
                        rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                        maxy = ceil(maxy / rndup) * rndup
                    except:
                        pass
                miny = 0
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                cx.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 1),
                              prop=lbl_font)
                plt.ylim([miny, maxy])
                plt.xlim([0, 23])
                plt.xticks(ticks)
                cx.set_xticklabels(hr_labels)
                cx.set_xlabel('Hour of the Day')
                cx.set_ylabel('Power (MW)')
                zp = ZoomPanX()
                f = zp.zoom_pan(cx, base_scale=1.2) # enable scrollable zoom
                plt.show()
                del zp
            else:
                fig = plt.figure('cumulative_' + str(year))
                plt.grid(True)
                dx = fig.add_subplot(111)
                plt.title(titl)
                if self.target == '<none>':
                    dx.fill_between(x, 0, data[0], label=label[0], color=self.colours[label[0].lower()])
                    for c in range(1, len(data)):
                        for h in range(len(data[c])):
                            data[c][h] = data[c][h] + data[c - 1][h]
                            maxy = max(maxy, data[c][h])
                        dx.fill_between(x, data[c - 1], data[c], label=label[c], color=self.colours[label[c].lower()])
                else:
                    pattern = ['-', '+', 'x', '\\', '*', 'o', 'O', '.']
                    pat = 0
                    full = []
                    for h in range(len(load)):
                       full.append(min(load[h], data[0][h]))
                    dx.fill_between(x, 0, full, label=label[0], color=self.colours[label[0].lower()])
                    dx.fill_between(x, full, data[0], alpha=self.alpha, color=self.colours[label[0].lower()]) #, hatch=pattern[0])
                    for c in range(1, len(data)):
                        full = []
                        for h in range(len(data[c])):
                            data[c][h] = data[c][h] + data[c - 1][h]
                            maxy = max(maxy, data[c][h])
                            full.append(max(min(load[h], data[c][h]), data[c - 1][h]))
                        dx.fill_between(x, data[c - 1], full, label=label[c], color=self.colours[label[c].lower()])
                        pat += 1
                        if pat >= len(pattern):
                            pat = 0
                        dx.fill_between(x, full, data[c], alpha=self.alpha, color=self.colours[label[c].lower()]) #,
                                     #   hatch=pattern[pat])
                    top = data[0][:]
                    for d in range(1, len(data)):
                        for h in range(len(top)):
                            top[h] = max(top[h], data[d][h])
                    dx.plot(x, top, color='white')
                    short = []
                    for h in range(len(load)):
                        short.append(max(data[c][h], load[h]))
                    dx.fill_between(x, data[c], short, label='Shortfall', color=self.colours['shortfall'])
                    dx.plot(x, load, linewidth=2.0, label=self.target, color=self.colours[self.target.lower()])
                if self.maxSpin.value() > 0:
                    maxy = self.maxSpin.value()
                else:
                    try:
                        rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                        maxy = ceil(maxy / rndup) * rndup
                    except:
                        pass
                miny = 0
                lbl_font = FontProperties()
                lbl_font.set_size('small')
                dx.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=lbl_font)
                plt.ylim([miny, maxy])
                plt.xlim([0, 23])
                plt.xticks(ticks)
                dx.set_xticklabels(hr_labels)
                dx.set_xlabel('Hour of the Day')
                dx.set_ylabel('Power (MW)')
                zp = ZoomPanX()
                f = zp.zoom_pan(dx, base_scale=1.2) # enable scrollable zoom
                plt.show()
                del zp


if "__main__" == __name__:
    app = QtGui.QApplication(sys.argv)
    ex = PowerPlot()
    app.exec_()
    app.deleteLater()
    sys.exit()
