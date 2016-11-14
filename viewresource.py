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
import time
from PyQt4 import QtCore, QtGui
import ConfigParser   # decode .ini file
import xlrd

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

def rreplace(s, old, new, occurrence):
    li = s.rsplit(old, occurrence)
    return new.join(li)


class Resource(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)


    def __init__(self, year=None, scene=None):
        super(Resource, self).__init__()
        self.year = year
        self.scene = scene
        self.resource_items = []
        self.be_open = True
        self.colours = {'dhi': ['DHI (Diffuse)', '#717100', '#ffff00', None],
                        'dni': ['DNI (Normal)', '#734c00', '#ff5500', None],
                        'ghi': ['GHI (Direct)', '#8b0000', '#ff0000', None],
                        'temp': ['Temperature', '#0d52e7', '#e8001f', None],
                        'wind': ['Wind Speed', '#82b2ff', '#0000b6', None],
                        'wind50': ['Wind @ 50m', '#82b2ff', '#0000b6', None]}
        self.hourly = False
        self.daily = False
        self.ignore = False
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
      # get default colours
        for key in self.colours:
            try:
                self.colours[key][1] = config.get('Colors', key + '_low')
            except:
                pass
            try:
                self.colours[key][2] = config.get('Colors', key + '_high')
            except:
                pass
        if tmap != '':
          # get this maps colours
            for key in self.colours:
                try:
                    self.colours[key][1] = config.get('Colors' + tmap, key + '_low')
                    self.map = tmap
                except:
                    pass
                try:
                    self.colours[key][2] = config.get('Colors' + tmap, key + '_high')
                    self.map = tmap
                except:
                    pass
        self.seasons = []
        self.periods = []
        try:
            items = config.items('Power')
            for item, values in items:
                if item[:6] == 'season':
                    if item == 'season':
                        continue
                    bits = values.split(',')
                    self.seasons.append(bits[0])
                elif item[:6] == 'period':
                    if item == 'period':
                        continue
                    bits = values.split(',')
                    self.periods.append(bits[0])
        except:
            pass
        if len(self.seasons) == 0:
            self.seasons = ['Summer', 'Autumn', 'Winter', 'Spring']
        if len(self.periods) == 0:
            self.periods = ['Winter', 'Summer']
        for i in range(len(self.periods)):
            for j in range(len(self.seasons)):
                if self.periods[i] == self.seasons[j]:
                    self.periods[i] += '2'
                    break
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
        yrf = rreplace(self.scene.resource_grid, '$YEAR$', self.year, 1)
        yrf = yrf.replace('$YEAR$', self.year)
        if not os.path.exists(yrf):
            return
        mrf = rreplace(self.scene.resource_grid, '$YEAR$', self.year + '-01', 1)
        mrf = mrf.replace('$YEAR$', self.year)
        if os.path.exists(mrf):
            self.hourly = True
            self.detail.append('Hourly by Month')
        drf = rreplace(self.scene.resource_grid, '$YEAR$', self.year + '-01-01', 1)
        drf = drf.replace('$YEAR$', self.year)
        if os.path.exists(drf):
            self.daily = True
            self.detail.append('Hourly by Day')
        self.resource_file = ""
        self.the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        self.mth_index = [0]
        for i in range(len(self.the_days) - 1):
            self.mth_index.append(self.mth_index[-1] + self.the_days[i] * 24)
        self.btn = []
        self.grid = QtGui.QGridLayout()
        row = 0
        self.detailCombo = QtGui.QComboBox()
        self.periodCombo = QtGui.QComboBox()
        self.dayCombo = QtGui.QComboBox()
        self.hourCombo = QtGui.QComboBox()
        for det in self.detail:
            self.detailCombo.addItem(det)
        if len(self.detail) > 1:
            self.grid.addWidget(QtGui.QLabel('Weather Detail:'), row, 0)
            self.grid.addWidget(self.detailCombo, row, 1, 1, 2)
            self.detailCombo.currentIndexChanged[str].connect(self.changeDetail)
            row += 1
        self.grid.addWidget(QtGui.QLabel('Weather Period:'), row, 0)
        self.periodCombo.currentIndexChanged[str].connect(self.periodChange)
        self.grid.addWidget(self.periodCombo, row, 1, 1, 2)
        self.dayCombo.currentIndexChanged[str].connect(self.periodChange)
        self.dayCombo.hide()
        self.grid.addWidget(self.dayCombo, row, 3, 1, 2)
        self.hourCombo.currentIndexChanged[str].connect(self.periodChange)
        self.hourCombo.hide()
        self.grid.addWidget(self.hourCombo, row, 5, 1, 3)
        row += 1
        prev = QtGui.QPushButton('<', self)
        width = prev.fontMetrics().boundingRect('<').width() + 10
        prev.setMaximumWidth(width)
        self.grid.addWidget(prev, row, 0)
        prev.clicked.connect(self.prevClicked)
        self.slider = QtGui.QSlider(QtCore.Qt.Horizontal)
        self.slider.valueChanged.connect(self.slideChanged)
        self.grid.addWidget(self.slider, row, 1, 1, 5)
        next = QtGui.QPushButton('>', self)
        width = next.fontMetrics().boundingRect('>').width() + 10
        next.setMaximumWidth(width)
        self.grid.addWidget(next, row, 7)
        next.clicked.connect(self.nextClicked)
        row += 1
        self.do_loop = False
        self.grid.addWidget(QtGui.QLabel('Period Loop (secs):'), row, 0)
        self.loopSpin = QtGui.QDoubleSpinBox()
        self.loopSpin.setRange(0., 10.)
        self.loopSpin.setDecimals(1)
        self.loopSpin.setSingleStep(.2)
        self.loopSpin.setValue(0.)
        self.loopSpin.valueChanged[str].connect(self.loopChanged)
        self.grid.addWidget(self.loopSpin, row, 1, 1, 2)
        self.loop = QtGui.QPushButton('Loop', self)
        self.grid.addWidget(self.loop, row, 3, 1, 4)
        self.loop.clicked.connect(self.loopClicked)
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
                self.gradients[-1][-1].setStyleSheet('QLabel {background-color: %s; color: %s;}' %
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
                self.btn[-1].setStyleSheet('QPushButton {background-color: %s; color: %s;}' %
                                 (value.name(), value.name()))
                self.grid.addWidget(self.btn[-1], row, 1)
            if self.colours[key][2] != '':
                value = QtGui.QColor(self.colours[key][2])
                self.btn.append(QtGui.QPushButton(key + '_2', self))
                self.btn[-1].clicked.connect(self.colourChanged)
                self.btn[-1].setStyleSheet('QPushButton {background-color: %s; color: %s;}' %
                                 (value.name(), value.name()))
                self.grid.addWidget(self.btn[-1], row, 2)
            if self.stepSpin.value() > 0:
                colors = gradient(self.colours[key][1], self.colours[key][2], self.stepSpin.value())
                for i in range(len(colors)):
                    value = QtGui.QColor(colors[i])
                    self.gradients[row - self.first_row][i].setStyleSheet('QLabel {background-color: %s; color: %s;}' %
                                    (value.name(), value.name()))
            row += 1
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, row, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        doit = QtGui.QPushButton('Show', self)
        self.grid.addWidget(doit, row, 1)
        doit.clicked.connect(self.showClicked)
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
        QtGui.QShortcut(QtGui.QKeySequence('pgdown'), self, self.nextClicked)
        QtGui.QShortcut(QtGui.QKeySequence('Ctrl++'), self, self.nextClicked)
        QtGui.QShortcut(QtGui.QKeySequence('Ctrl+-'), self, self.prevClicked)
        self.setPeriod()
        self.resourceGrid()
        self.show()

    def loopChanged(self, val):
        return

    def changeDetail(self):
        self.setPeriod()
        self.resourceGrid()

    def helpClicked(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.helpfile, title='Help', section='resource')
        dialog.exec_()

    def nextClicked(self):
        if self.slider.value() == self.slider.maximum():
            self.slider.setValue(0)
        else:
            self.slider.setValue(self.slider.value() + 1)

    def prevClicked(self):
        if self.slider.value() == self.slider.minimum():
            self.slider.setValue(self.slider.maximum())
        else:
            self.slider.setValue(self.slider.value() - 1)

    def slideChanged(self, val):
        self.ignore = True
        if self.detailCombo.currentText() == 'Daily By Month':
            self.periodCombo.setCurrentIndex(val)
        else:
            period = divmod(val, 24)
            if self.detailCombo.currentText() == 'Hourly by Month':
                self.periodCombo.setCurrentIndex(period[0])
            elif self.detailCombo.currentText() == 'Hourly by Day':
                for m in range(len(self.mth_index) -1, -1, -1):
                    if val >= self.mth_index[m]:
                        break
                self.periodCombo.setCurrentIndex(m)
                period = divmod(val - self.mth_index[m], 24)
                if m != self.dayCombo.currentIndex:
                    self.setDayItems(m)
                self.dayCombo.setCurrentIndex(period[0])
            self.hourCombo.setCurrentIndex(period[1])
        self.ignore = False
        self.resourceGrid()

    def loopClicked(self):
        if self.loop.text() == 'Stop':
            self.do_loop = False
            self.scene.exitLoop = False
            self.detailCombo.setEnabled(True)
            self.periodCombo.setEnabled(True)
            self.dayCombo.setEnabled(True)
            self.hourCombo.setEnabled(True)
            self.loop.setText('Loop')
        else:
            self.do_loop = True
            self.detailCombo.setEnabled(False)
            self.periodCombo.setEnabled(False)
            self.dayCombo.setEnabled(False)
            self.hourCombo.setEnabled(False)
            self.loop.setText('Stop')
            self.periodChange()

    def periodChange(self, value=None):
        if self.ignore:
            return
        if self.periodCombo.currentIndex() < 0:
            per_ndx = 0
        else:
            per_ndx = self.periodCombo.currentIndex()
        if self.detailCombo.currentText() != 'Hourly by Day' or self.dayCombo.currentIndex() < 0:
            day_ndx = 0
        else:
            day_ndx = self.dayCombo.currentIndex()
            if day_ndx > self.the_days[per_ndx] - 1:
                day_ndx = self.the_days[per_ndx] - 1
        if self.hourCombo.currentIndex() < 0:
            hr_ndx = 0
        else:
            hr_ndx = self.hourCombo.currentIndex()
        if self.detailCombo.currentText() == 'Daily By Month':
            val = per_ndx
        elif self.detailCombo.currentText() == 'Hourly by Month':
            val = per_ndx * 24 + hr_ndx
        elif self.detailCombo.currentText() == 'Hourly by Day':
            val = day_ndx * 24 + self.mth_index[per_ndx] + hr_ndx
        while val <= self.slider.maximum():
            self.slider.setValue(val)
            QtCore.QCoreApplication.processEvents()
            QtCore.QCoreApplication.flush()
            if not self.do_loop or self.scene.exitLoop:
                break
            val += 1
            if val < self.slider.maximum():
                self.slider.setValue(val)
        if self.do_loop:
            self.do_loop = False
            self.detailCombo.setEnabled(True)
            self.periodCombo.setEnabled(True)
            self.dayCombo.setEnabled(True)
            self.hourCombo.setEnabled(True)
            self.loop.setText('Loop')
            if not self.scene.exitLoop:
                self.slider.setValue(0)
        self.scene.exitLoop = False

    def setDayItems(self, mth):
        if self.dayCombo.count() > 0:
            self.dayCombo.clear()
        if mth < 0 or mth >= len(self.the_days):
            rnge = 31
        else:
            rnge = self.the_days[mth]
        for d in range(rnge):
            self.dayCombo.addItem('{0:02d}'.format(d + 1))

    def setPeriod(self):
        if self.periodCombo.count() > 0:
            self.periodCombo.clear()
        if self.dayCombo.count() > 0:
            self.dayCombo.clear()
        if self.hourCombo.count() > 0:
            self.hourCombo.clear()
        if self.detailCombo.currentText() == 'Daily By Month':
            self.periodCombo.addItem(self.year)
        for i in range(12):
            self.periodCombo.addItem(self.year + '-' + '{0:02d}'.format(i + 1))
        if self.detailCombo.currentText() == 'Daily By Month':
            for i in range(len(self.seasons)):
                self.periodCombo.addItem(self.year + '-' + self.seasons[i])
            for i in range(len(self.periods)):
                self.periodCombo.addItem(self.year + '-' + self.periods[i])
            self.dayCombo.hide()
            self.hourCombo.hide()
            self.slider.setMaximum(self.periodCombo.count() - 1)
        else:
            for j in range(24):
                self.hourCombo.addItem('{0:02d}'.format(j + 1) + ':00')
                self.hourCombo.show()
            if self.detailCombo.currentText() == 'Hourly by Month':
                self.slider.setMaximum(self.periodCombo.count() * 24 - 1)
                self.dayCombo.hide()
            elif self.detailCombo.currentText() == 'Hourly by Day':
                self.setDayItems(0)
                self.slider.setMaximum(365 * 24 - 1)
                self.dayCombo.show()
                self.dayCombo.setCurrentIndex(0)
            self.hourCombo.show()
            self.hourCombo.setCurrentIndex(0)
        self.periodCombo.setCurrentIndex(0)
        self.slider.setValue(0)

    def stepChanged(self, val):
        value = self.palette().color(QtGui.QPalette.Window)
        for i in range(len(self.gradients)):
            for j in range(len(self.gradients[i])):
                self.gradients[i][j].setStyleSheet('QLabel {background-color: %s; color: %s;}' %
                                                    (value.name(), value.name()))
        if self.stepSpin.value() > 0:
            row = self.first_row
            for key in sorted(self.colours):
                self.colours[key][3] = gradient(self.colours[key][1], self.colours[key][2], self.stepSpin.value())
                for i in range(len(self.colours[key][3])):
                    value = QtGui.QColor(self.colours[key][3][i])
                    value.setAlphaF(self.opacitySpin.value())
                    self.gradients[row - self.first_row][i].setStyleSheet(
                                  'QLabel {background-color: %s; color: %s; opacity: 50}' %
                                  (value.name(), value.name()))
                    val = []
                    for j in range(3):
                        val.append(str(int(self.colours[key][3][i][j * 2 + 1:j * 2 + 3], base=16)))
                    val.append(str(self.opacitySpin.value() * 100))
                    it = 'QLabel {background-color: rgba(%s,%s,%s,%s%%); color: rgba(%s,%s,%s,%s%%)}' % \
                         (val[0], val[1], val[2], val[3], val[0], val[1], val[2], val[3])
                    self.gradients[row - self.first_row][i].setStyleSheet(it)
                row += 1
        self.resourceGrid()

    @QtCore.pyqtSlot()
    def exit(self):
        self.be_open = False
        self.close()

    def closeEvent(self, event):
        self.clear_Resource()
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

    def showClicked(self):
        self.resourceGrid()

    def hideClicked(self):
        self.clear_Resource()

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

    def clear_Resource(self):
        try:
            for i in range(len(self.resource_items)):
                self.scene.removeItem(self.resource_items[i])
            del self.resource_items
        except:
            pass

    def resourceGrid(self):
        def in_map(a_min, a_max, b_min, b_max):
            return (a_min <= b_max) and (b_min <= a_max)
        self.clear_Resource()
        for key in self.colours:
            if self.colours[key][0] == self.whatCombo.currentText():
                colours = self.colours[key][3]
                break
        period = str(self.periodCombo.currentText())
        if self.detailCombo.currentText() == 'Daily By Month':
            pass
        else:
            if self.detailCombo.currentText() == 'Hourly by Day':
                period += '-' + str(self.dayCombo.currentText())
            period += '_' + str(self.hourCombo.currentText())
        variable = str(self.whatCombo.currentText())
        steps = self.stepSpin.value()
        opacity = self.opacitySpin.value()
        i = period.find('_')
        if i > 0:
            new_file = rreplace(self.scene.resource_grid, '$YEAR$', period[:i], 1)
            new_file = new_file.replace('$YEAR$', period[:4])
        else:
            new_file = self.scene.resource_grid.replace('$YEAR$', period[:4])
        self.resource_items = []
        if new_file[-4:] == '.xls' or new_file[-5:] == '.xlsx':
            if new_file != self.resource_file:
                if os.path.exists(new_file):
                    self.resource_var = {}
                    workbook = xlrd.open_workbook(new_file)
                    self.resource_worksheet = workbook.sheet_by_index(0)
                    num_cols = self.resource_worksheet.ncols - 1
#                   get column names
                    curr_col = -1
                    while curr_col < num_cols:
                        curr_col += 1
                        self.resource_var[self.resource_worksheet.cell_value(0, curr_col)] = curr_col
                    self.resource_file = new_file
                else:
                    return
            num_rows = self.resource_worksheet.nrows - 1
            num_cols = self.resource_worksheet.ncols - 1
            lo_valu = 99999.
            hi_valu = 0.
            calc_minmax = True
            cells = []
            curr_row = 1
            if self.resource_worksheet.cell_value(curr_row, self.resource_var['Period']) == 'Min.' and \
              self.resource_worksheet.cell_value(curr_row + 1, self.resource_var['Period']) == 'Max.':
                calc_minmax = False
                lo_valu = self.resource_worksheet.cell_value(curr_row, self.resource_var[variable])
                hi_valu = self.resource_worksheet.cell_value(curr_row + 1, self.resource_var[variable])
                curr_row += 1
            while curr_row < num_rows:
                curr_row += 1
                a_lat = float(self.resource_worksheet.cell_value(curr_row, self.resource_var['Latitude']))
                a_lon = float(self.resource_worksheet.cell_value(curr_row, self.resource_var['Longitude']))
                if in_map(a_lat - 0.25, a_lat + 0.25, self.scene.map_lower_right[0],
                          self.scene.map_upper_left[0]) \
                  and in_map(a_lon - 0.3333, a_lon + 0.3333,
                          self.scene.map_upper_left[1], self.scene.map_lower_right[1]):
                    try:
                        if self.resource_worksheet.cell_value(curr_row, self.resource_var['Period']) == period:
                            cells.append([float(self.resource_worksheet.cell_value(curr_row, self.resource_var['Latitude'])),
                                         float(self.resource_worksheet.cell_value(curr_row, self.resource_var['Longitude'])),
                                         self.resource_worksheet.cell_value(curr_row, self.resource_var[variable])])
                        if calc_minmax:
                            if self.resource_worksheet.cell_value(curr_row, self.resource_var[variable]) < lo_valu:
                                lo_valu = self.resource_worksheet.cell_value(curr_row, self.resource_var[variable])
                            if self.resource_worksheet.cell_value(curr_row, self.resource_var[variable]) > hi_valu:
                                hi_valu = self.resource_worksheet.cell_value(curr_row, self.resource_var[variable])
                    except:
                        pass
        else:
            if os.path.exists(new_file):
                resource = open(new_file)
                things = csv.DictReader(resource)
                for cell in things:
                    a_lat = float(cell['Latitude'])
                    a_lon = float(cell['Longitude'])
                    if in_map(a_lat - 0.25, a_lat + 0.25, self.scene().map_lower_right[0],
                              self.scene().map_upper_left[0]) \
                      and in_map(a_lon - 0.3333, a_lon + 0.3333,
                              self.scene().map_upper_left[1], self.scene().map_lower_right[1]):
                        if cell['Period'] == period:
                            cells.append(cell['Latitude'], cell['Longitude'], cell[variable])
                        if cell[variable] < lo_valu:
                            lo_valu = cell[variable]
                        if cell[variable] > hi_valu:
                            hi_valu = cell[variable]
                resource.close()
            else:
                return
        if steps > 0:
            incr = (hi_valu - lo_valu) / steps
        else:
            lo_colour = []
            hi_colour = []
            for i in range(3):
                lo_1 = int(colours[0][i * 2 + 1:i * 2 + 3], base=16)
                lo_colour.append(lo_1)
                hi_1 = int(colours[-1][i * 2 + 1:i * 2 + 3], base=16)
                hi_colour.append(hi_1)
        lo_per = 99999.
        hi_per = 0.
        lons = []
        lon_cell = .3125
        for cell in cells:
            if cell[1] not in lons:
                lons.append(cell[1])
        if len(lons) > 1:
            lons = sorted(lons)
            lon_cell = (lons[1] - lons[0]) / 2.
            del lons
        for cell in cells:
            p = self.scene.mapFromLonLat(QtCore.QPointF(cell[1] - lon_cell, cell[0] + .25))
            pe = self.scene.mapFromLonLat(QtCore.QPointF(cell[1] + lon_cell, cell[0] + .25))
            ps = self.scene.mapFromLonLat(QtCore.QPointF(cell[1] - lon_cell, cell[0] - .25))
            x_d = pe.x() - p.x()
            y_d = ps.y() - p.y()
            self.resource_items.append(QtGui.QGraphicsRectItem(p.x(), p.y(), x_d, y_d))
            if steps > 0:
                step = int(round((cell[2] - lo_valu) / incr))
                a_colour = QtGui.QColor(colours[step])
            else:
                colr = []
                pct = (cell[2] - lo_valu) / (hi_valu - lo_valu)
                for i in range(3):
                    colr.append(((hi_colour[i] - lo_colour[i]) * pct + lo_colour[i]) / 255.)
                a_colour = QtGui.QColor()
                a_colour.setRgbF(colr[0], colr[1], colr[2])
            if cell[2] >= lo_valu:
                if cell[2] > 0 or variable == 'temp':
                    self.resource_items[-1].setBrush(a_colour)
            self.resource_items[-1].setPen(a_colour)
            self.resource_items[-1].setOpacity(opacity)
            self.resource_items[-1].setZValue(1)
            self.scene.addItem(self.resource_items[-1])
            if cell[2] < lo_per:
                lo_per = cell[2]
            if cell[2] > hi_per:
                hi_per = cell[2]
        QtCore.QCoreApplication.processEvents()
        QtCore.QCoreApplication.flush()
        if self.do_loop and not self.scene.exitLoop:
            if self.loopSpin.value() > 0:
                time.sleep(self.loopSpin.value())
    QtCore.QCoreApplication.processEvents()
    QtCore.QCoreApplication.flush()
