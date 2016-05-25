#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
#
#  viewresource.py - This file is part of SIREN.
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
from PyQt4 import QtCore, QtGui
import ConfigParser   # decode .ini file

import displayobject
from editini import SaveIni


def gradient(lo, hi, steps=10):
    colours = []
    grad = []
    for i in range(3):
        lo_1 = int(lo[i * 2 + 1:i * 2 + 3], base=16)
        hi_1 = int(hi[i * 2 + 1:i * 2 + 3], base=16)
        incr = int((hi_1 - lo_1) / steps)
        grad.append([lo_1])
        for j in range(steps):
            lo_1 += incr
            if lo_1 < 0:
                lo_1 = 0
            elif lo_1 > 255:
                lo_1 = 255
            grad[-1].append(lo_1)
    for j in range(len(grad[0])):
        colr = '#'
        for i in range(3):
            colr += str(hex(grad[i][j]))[2:].zfill(2)
        colours.append(colr)
    return colours


class Resource(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, year=None, resource_grid=None):
        super(Resource, self).__init__()
        self.year = year
        self.resource_grid = resource_grid
        self.initUI()

    def initUI(self):
        self.be_open = True
        self.colours = {'dhi': ['DHI (Diffuse)', '#717100', '#ffff00', None], \
                        'dni': ['DNI (Normal)', '#734c00', '#ff5500', None], \
                        'ghi': ['GHI (Direct)', '#8b0000', '#ff0000', None], \
                        'temp': ['Temperature', '#0d52e7', '#e8001f', None], \
                        'wind': ['Wind Speed', '#82b2ff', '#0000b6', None], \
                        'wind50': ['Wind @ 50m', '#82b2ff', '#0000b6', None]}
        self.hourly = False
        self.daily = False
        self.detail = ['Daily By Month']
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        if self.year is None:
            try:
                self.year = config.get('Base', 'year')
            except:
                pass
        self.map = ''
        try:
            tmap = config.get('Map', 'map_choice')
        except:
            tmap = ''
        for key in self.colours:
            try:
                self.colours[key][1] = config.get('Map', key + '_low')
            except:
                pass
            try:
                self.colours[key][2] = config.get('Map', key + '_high')
            except:
                pass
        if tmap != '':
            for key in self.colours:
                try:
                    self.colours[key][1] = config.get('Map' + tmap, key + '_low')
                    self.map = tmap
                except:
                    pass
                try:
                    self.colours[key][2] = config.get('Map' + tmap, key + '_high')
                    self.map = tmap
                except:
                    pass
        try:
            self.seasons = []
            items = config.items('Power')
            for item, values in items:
                if item[:6] == 'season':
                    if item == 'season':
                        continue
                    bits = values.split(',')
                    self.seasons.append(bits[0])
        except:
            self.seasons = ['Summer', 'Autumn', 'Winter', 'Spring']
        try:
            self.helpfile = config.get('Files', 'help')
        except:
            self.helpfile = ''
        try:
            opacity = float(config.get('View', 'resource_opacity'))
        except:
            opacity = .5
        try:
            period = config.get('View', 'resource_period')
        except:
            period = self.year
        try:
            steps = int(config.get('View', 'resource_steps'))
        except:
            steps = 2
        try:
            max_steps = int(config.get('View', 'resource_max_steps'))
        except:
            max_steps = 10
        try:
            variable = config.get('View', 'resource_variable')
        except:
            variable = self.colours['ghi'][0]
        try:
            variable = config.get('View', 'resource_rainfall')
            if variable.lower() in ['true', 'yes', 'on']:
                self.colours['rain'] = ['Rainfall', '#e6e7e6', '#005500', None]
        except:
            pass
        self.restorewindows = False
        try:
            rw = config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
        if self.resource_grid is None:
             # get path variables to create resource grid
            try:
                rf = config.get('Files', 'resource_grid')
                self.resource_grid = rf
            except:
                pass
        else:
            yrf = self.resource_grid.replace('$YEAR$', self.year)
            if not os.path.exists(yrf):
                return
            mrf = self.resource_grid.replace('$YEAR$', self.year + '-01')
            if os.path.exists(mrf):
                self.hourly = True
                self.detail.append('Hourly by Month')
            drf = self.resource_grid.replace('$YEAR$', self.year + '-01-01')
            if os.path.exists(drf):
                self.daily = True
                self.detail.append('Hourly by Day')
        self.btn = []
        self.grid = QtGui.QGridLayout()
        row = 0
        self.detailCombo = QtGui.QComboBox()
        self.skipCombo = QtGui.QComboBox()
        self.skipdayCombo = QtGui.QComboBox()
        for det in self.detail:
            self.detailCombo.addItem(det)
        if len(self.detail) > 1:
            self.grid.addWidget(QtGui.QLabel('Weather Detail:'), row, 0)
            self.grid.addWidget(self.detailCombo, row, 1, 1, 2)
            self.grid.addWidget(self.skipCombo, row, 3, 1, 2)
            self.skipCombo.hide()
            self.grid.addWidget(self.skipdayCombo, row, 5)
            self.skipdayCombo.hide()
            row += 1
        self.grid.addWidget(QtGui.QLabel('Weather Period:'), row, 0)
        self.periodCombo = QtGui.QComboBox()
        self.setPeriod(period)
        self.periodCombo.currentIndexChanged[str].connect(self.periodChange)
        self.skipCombo.currentIndexChanged[str].connect(self.skip)
        self.skipdayCombo.currentIndexChanged[str].connect(self.skipday)
        self.grid.addWidget(self.periodCombo, row, 1, 1, 2)
        loop = QtGui.QPushButton('Next', self)
        self.grid.addWidget(loop, row, 3, 1, 4)
        loop.clicked.connect(self.loopClicked)
      #   self.grid.addWidget(QtGui.QLabel('Period Loop:'), 1, 0)
      #   self.loopSpin = QtGui.QSpinBox()
      #   self.loopSpin.setRange(0, 10)
      #   self.loopSpin.setValue(0)
      #   self.loopSpin.valueChanged[str].connect(self.loopChanged)
      #   self.grid.addWidget(self.loopSpin, 1, 1, 1, 2)
      #   self.grid.addWidget(QtGui.QLabel('(seconds)'), 1, 3, 1, 3)
        row += 1
        self.grid.addWidget(QtGui.QLabel('Weather Variable:'), row, 0)
        self.whatCombo = QtGui.QComboBox()
        for key in sorted(self.colours):
            self.whatCombo.addItem(self.colours[key][0])
            if variable == self.colours[key][0]:
                self.whatCombo.setCurrentIndex(self.whatCombo.count() - 1)
        self.whatCombo.currentIndexChanged[str].connect(self.periodChange)
        self.grid.addWidget(self.whatCombo, row, 1, 1, 2)
        if len(self.detail) > 1:
            self.detailCombo.currentIndexChanged[str].connect(self.changeDetail)
        row += 1
        self.grid.addWidget(QtGui.QLabel('Colour Steps:'), row, 0)
        self.stepSpin = QtGui.QSpinBox()
        self.stepSpin.setRange(0, max_steps)
        self.stepSpin.setValue(steps)
        self.stepSpin.valueChanged[str].connect(self.stepChanged)
        self.grid.addWidget(self.stepSpin, row, 1, 1, 2)
        row += 1
        self.grid.addWidget(QtGui.QLabel('Opacity:'), row, 0)
        self.opacitySpin = QtGui.QDoubleSpinBox()
        self.opacitySpin.setRange(0, 1.)
        self.opacitySpin.setDecimals(2)
        self.opacitySpin.setSingleStep(.05)
        self.opacitySpin.setValue(opacity)
        self.opacitySpin.valueChanged[str].connect(self.stepChanged)
        self.grid.addWidget(self.opacitySpin, row, 1, 1, 2)
        row += 1
        self.grid.addWidget(QtGui.QLabel('Low Colour'), row, 1)
        self.grid.addWidget(QtGui.QLabel('High Colour'), row, 2)
        row += 1
        self.gradients = []
        value = self.palette().color(QtGui.QPalette.Window)
        self.first_row = row
        for key in sorted(self.colours):
            self.gradients.append([])
            for i in range(self.stepSpin.maximum() + 1):
                self.gradients[-1].append(QtGui.QLabel('__'))
                self.gradients[-1][-1].setStyleSheet('QLabel {background-color: %s; color: %s;}' % \
                                                    (value.name(), value.name()))
                self.grid.addWidget(self.gradients[-1][-1], row, i + 3)
            row += 1
        row = self.first_row
        for key in sorted(self.colours):
            self.grid.addWidget(QtGui.QLabel(self.colours[key][0]), row, 0)
            if self.colours[key][1] != '':
                value = QtGui.QColor(self.colours[key][1])
                self.btn.append(QtGui.QPushButton(key + '_1', self))
                self.btn[-1].clicked.connect(self.colourChanged)
                self.btn[-1].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % \
                                 (value.name(), value.name()))
                self.grid.addWidget(self.btn[-1], row, 1)
            if self.colours[key][2] != '':
                value = QtGui.QColor(self.colours[key][2])
                self.btn.append(QtGui.QPushButton(key + '_2', self))
                self.btn[-1].clicked.connect(self.colourChanged)
                self.btn[-1].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % \
                                 (value.name(), value.name()))
                self.grid.addWidget(self.btn[-1], row, 2)
            if self.stepSpin.value() > 0:
                colors = gradient(self.colours[key][1], self.colours[key][2], self.stepSpin.value())
                for i in range(len(colors)):
                    value = QtGui.QColor(colors[i])
                    self.gradients[row - self.first_row][i].setStyleSheet('QLabel {background-color: %s; color: %s;}' % \
                                    (value.name(), value.name()))
            row += 1
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, row, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        doit = QtGui.QPushButton('Show', self)
        self.grid.addWidget(doit, row, 1)
        doit.clicked.connect(self.periodChange)
        hide = QtGui.QPushButton('Hide', self)
        self.grid.addWidget(hide, row, 2)
        hide.clicked.connect(self.hideClicked)
        save = QtGui.QPushButton('Save', self)
        self.grid.addWidget(save, row, 3, 1, 4)
        save.clicked.connect(self.saveClicked)
        help = QtGui.QPushButton('Help', self)
        self.grid.addWidget(help, row, 7, 1, 4)
        help.clicked.connect(self.helpClicked)
        QtGui.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Renewable Resource Overlay')
        self.stepChanged('a')
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            move_right = False
        else:
            move_right = True
        if self.restorewindows:
            try:
                rw = config.get('Windows', 'resource_size').split(',')
                self.resize(int(rw[0]), int(rw[1]))
                mp = config.get('Windows', 'resource_pos').split(',')
                self.move(int(mp[0]), int(mp[1]))
                move_right = False
            except:
                pass
        if move_right:
            frameGm = self.frameGeometry()
            screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
            trPoint = QtGui.QApplication.desktop().availableGeometry(screen).topRight()
            frameGm.moveTopRight(trPoint)
            self.move(frameGm.topRight())
        QtGui.QShortcut(QtGui.QKeySequence('pgup'), self, self.prevClicked)
        QtGui.QShortcut(QtGui.QKeySequence('pgdown'), self, self.loopClicked)
        QtGui.QShortcut(QtGui.QKeySequence('Ctrl++'), self, self.loopClicked)
        QtGui.QShortcut(QtGui.QKeySequence('Ctrl+-'), self, self.prevClicked)
        self.show()

    def loopChanged(self, val):
        return

    def changeDetail(self, val):
        self.setPeriod()

    def helpClicked(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.helpfile, title='Help', section='resource')
        dialog.exec_()

    def loopClicked(self):
        if self.periodCombo.currentIndex() == (self.periodCombo.count() - 1):
            self.periodCombo.setCurrentIndex(0)
        else:
            self.periodCombo.setCurrentIndex(self.periodCombo.currentIndex() + 1)
        return

    def prevClicked(self):
        if self.periodCombo.currentIndex() == 0:
            self.periodCombo.setCurrentIndex(self.periodCombo.count() - 1)
        else:
            self.periodCombo.setCurrentIndex(self.periodCombo.currentIndex() - 1)
        return

    def skip(self):
        if self.skipCombo.currentText() == '':
            return
        mth = int(str(self.skipCombo.currentText())[-2:])
        if mth == 2:
            max_day = 28
        elif mth in [4, 6, 9, 11]:
            max_day = 30
        else:
            max_day = 31
        self.skipdayCombo.clear()
        for d in range(max_day):
            self.skipdayCombo.addItem('{0:02d}'.format(d + 1))
        if self.detailCombo.currentText() == 'Hourly by Month':
            self.periodCombo.setCurrentIndex(self.periodCombo.findText(self.skipCombo.currentText() + '_01:00'))
        elif self.detailCombo.currentText() == 'Hourly by Day':
            self.periodCombo.setCurrentIndex(self.periodCombo.findText(self.skipCombo.currentText() + \
                '-' + self.skipdayCombo.currentText() + '_01:00'))

    def skipday(self):
        if self.skipdayCombo.currentText() == '':
            return
        self.periodCombo.setCurrentIndex(self.periodCombo.findText(self.skipCombo.currentText() + \
            '-' + self.skipdayCombo.currentText() + '_01:00'))

    def setPeriod(self, period=None):
        if self.periodCombo.count() > 0:
            self.periodCombo.clear()
        if self.skipCombo.count() > 0:
            self.skipCombo.clear()
        if self.skipdayCombo.count() > 0:
            self.skipdayCombo.clear()
        if self.detailCombo.currentText() == 'Daily By Month':
            self.periodCombo.addItem(self.year)
            for i in range(12):
                self.periodCombo.addItem(self.year + '-' + '{0:02d}'.format(i + 1))
            for i in range(len(self.seasons)):
                self.periodCombo.addItem(self.year + '-' + self.seasons[i])
            if period == self.year + '-' + self.seasons[i]:
                self.periodCombo.setCurrentIndex(self.periodCombo.count() - 1)
            self.skipCombo.hide()
            self.skipdayCombo.hide()
        elif self.detailCombo.currentText() == 'Hourly by Month':
            for i in range(12):
                self.skipCombo.addItem(self.year + '-{0:02d}'.format(i + 1))
                for j in range(24):
                    self.periodCombo.addItem(self.year + '-{0:02d}'.format(i + 1) + \
                                             '_{0:02d}'.format(j + 1) + ':00')
                if period == self.year + '-' + '{0:02d}'.format(i):
                    self.periodCombo.setCurrentIndex(i)
            self.skipCombo.show()
            self.skipdayCombo.hide()
        elif self.detailCombo.currentText() == 'Hourly by Day':
            the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
            for i in range(12):
                self.skipCombo.addItem(self.year + '-{0:02d}'.format(i + 1))
                for d in range(the_days[i]):
                    for j in range(24):
                        self.periodCombo.addItem(self.year + '-{0:02d}'.format(i + 1) + \
                                             '-{0:02d}'.format(d + 1) + \
                                             '_{0:02d}'.format(j + 1) + ':00')
            self.skipCombo.show()
           #  for d in range(31):
            #     self.skipdayCombo.addItem('{0:02d}'.format(d + 1))
            self.skipdayCombo.show()
        self.periodCombo.setCurrentIndex(0)

    def stepChanged(self, val):
        value = self.palette().color(QtGui.QPalette.Window)
        for i in range(len(self.gradients)):
            for j in range(len(self.gradients[i])):
                self.gradients[i][j].setStyleSheet('QLabel {background-color: %s; color: %s;}' % \
                                                    (value.name(), value.name()))
        if self.stepSpin.value() > 0:
            row = self.first_row
            for key in sorted(self.colours):
                self.colours[key][3] = gradient(self.colours[key][1], self.colours[key][2], self.stepSpin.value())
                for i in range(len(self.colours[key][3])):
                    value = QtGui.QColor(self.colours[key][3][i])
                    value.setAlphaF(self.opacitySpin.value())
                    self.gradients[row - self.first_row][i].setStyleSheet( \
                                  'QLabel {background-color: %s; color: %s; opacity: 50}' % \
                                  (value.name(), value.name()))
                    val = []
                    for j in range(3):
                        val.append(str(int(self.colours[key][3][i][j * 2 + 1:j * 2 + 3], base=16)))
                    val.append(str(self.opacitySpin.value() * 100))
                    it = 'QLabel {background-color: rgba(%s,%s,%s,%s%%); color: rgba(%s,%s,%s,%s%%)}' % \
                         (val[0], val[1], val[2], val[3], val[0], val[1], val[2], val[3])
                    self.gradients[row - self.first_row][i].setStyleSheet(it)
                row += 1
        self.procStart.emit('show')

    @QtCore.pyqtSlot()
    def exit(self):
        self.be_open = False
        self.close()

    def closeEvent(self, event):
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)   # need to account for border
            lines.append('resource_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('resource_size=%s,%s' % (str(self.width()), str(self.height())))
            updates['Windows'] = lines
            SaveIni(updates)
        if self.be_open:
            self.procStart.emit('goodbye')
        event.accept()

    def quitClicked(self):
        self.close()

    def periodChange(self):
        if self.periodCombo.currentText() == '':
            return
        self.procStart.emit('show')

    def hideClicked(self):
        self.procStart.emit('hide')

    def saveClicked(self):
        updates = {}
        colour_lines = []
        for key in self.colours:
            colour_lines.append('%s_high=%s' % (key, self.colours[key][2]))
            colour_lines.append('%s_low=%s' % (key, self.colours[key][1]))
        updates['Colors' + self.map] = colour_lines
        view_lines = []
        view_lines.append('resource_opacity=%s' % str(self.opacitySpin.value()))
        view_lines.append('resource_period=$YEAR$%s' % str(self.periodCombo.currentText())[4:])
        view_lines.append('resource_steps=%s' % str(self.stepSpin.value()))
        view_lines.append('resource_variable=%s' % str(self.whatCombo.currentText()))
        updates['View'] = view_lines
        SaveIni(updates)

    def Results(self):
        for key in self.colours:
            if self.colours[key][0] == self.whatCombo.currentText():
                colours = self.colours[key][3]
                break
        return str(self.periodCombo.currentText()), str(self.whatCombo.currentText()), \
               self.stepSpin.value(), self.opacitySpin.value(), colours

    def colourChanged(self):
        sender = str(self.sender().text()).split('_')
        key = sender[0]
        ndx = int(sender[1])
        if self.colours[key][ndx] != '':
            value = QtGui.QColor(self.colours[key][ndx])
            col = QtGui.QColorDialog.getColor(value)
        else:
            col = QtGui.QColorDialog.getColor(QtGui.QColor(''))
        if col.isValid():
            self.colours[key][ndx] = str(col.name())
            for i in range(len(self.btn)):
                if self.btn[i] == self.sender():
                    self.btn[i].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % (col.name(), col.name()))
                    break
            self.stepChanged('a')
