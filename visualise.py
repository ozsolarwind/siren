#!/usr/bin/python
#
#  Copyright (C) 2016 Sustainable Energy Now Inc., Angus King
#
#  visualise.py - This file is part of SIREN.
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

import math
import os
import sys
import time
from PyQt4 import QtCore, QtGui
import ConfigParser   # decode .ini file


class Visualise(QtGui.QDialog):
    def destinationxy(self, lon1, lat1, bearing, distance):
        """
        Given a start point, initial bearing, and distance, calculate
        the destination point and final bearing travelling along a
        (shortest distance) great circle arc
        """
        radius = 6367.  # km is the radius of the Earth
      # convert decimal degrees to radians
        ln1, lt1, baring = map(math.radians, [lon1, lat1, bearing])
      # "reverse" haversine formula
        lat2 = math.asin(math.sin(lt1) * math.cos(distance / radius) +
               math.cos(lt1) * math.sin(distance / radius) * math.cos(baring))
        lon2 = ln1 + math.atan2(math.sin(baring) * math.sin(distance / radius) * math.cos(lt1),
               math.cos(distance / radius) - math.sin(lt1) * math.sin(lat2))
        return QtCore.QPointF(math.degrees(lon2), math.degrees(lat2))

    def __init__(self, stations, powers, main, year=None):
        super(Visualise, self).__init__()
        self.year = year
        self.stations = stations
        self.powers = powers
        self.scene = main.view.scene()
        self.visual_group = QtGui.QGraphicsItemGroup()
        self.visual_items = []
        self.scene.addItem(self.visual_group)
        self.be_open = True
        self.ignore = False
        self.detail = ['Diurnal', 'Hourly']
        self.the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        self.mth_index = [0]
        for i in range(len(self.the_days) - 1):
            self.mth_index.append(self.mth_index[-1] + self.the_days[i] * 24)
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
        self.seasons = []
        self.periods = []
        self.daily = True
        try:
            items = config.items('Power')
            for item, values in items:
                if item[:6] == 'season':
                    if item == 'season':
                        continue
                    i = int(item[6:]) - 1
                    if i >= len(self.seasons):
                        self.seasons.append([])
                    self.seasons[i] = values.split(',')
                    for j in range(1, len(self.seasons[i])):
                        self.seasons[i][j] = int(self.seasons[i][j]) - 1
                elif item[:6] == 'period':
                    if item == 'period':
                        continue
                    i = int(item[6:]) - 1
                    if i >= len(self.periods):
                        self.periods.append([])
                    self.periods[i] = values.split(',')
                    for j in range(1, len(self.periods[i])):
                        self.periods[i][j] = int(self.periods[i][j]) - 1
        except:
            pass
        if len(self.seasons) == 0:
            self.seasons = [['Summer', 11, 0, 1], ['Autumn', 2, 3, 4], ['Winter', 5, 6, 7], ['Spring', 8, 9, 10]]
        if len(self.periods) == 0:
            self.periods = [['Winter', 4, 5, 6, 7, 8, 9], ['Summer', 10, 11, 0, 1, 2, 3]]
        for i in range(len(self.periods)):
            for j in range(len(self.seasons)):
                if self.periods[i][0] == self.seasons[j][0]:
                    self.periods[i][0] += '2'
                    break
        self.stn_items = []
        for i in range(len(self.stations)):
            self.stn_items.append(-1)
        for i in range(len(self.scene._stations.stations)):
            try:
                j = self.stations.index(self.scene._stations.stations[i].name)
                self.stn_items[j] = i
            except:
                pass
        self.grid = QtGui.QGridLayout()
        row = 0
        self.detailCombo = QtGui.QComboBox()
        self.dayCombo = QtGui.QComboBox()
        self.hourCombo = QtGui.QComboBox()
        for det in self.detail:
            self.detailCombo.addItem(det)
        self.grid.addWidget(QtGui.QLabel('Detail:'), row, 0)
        self.grid.addWidget(self.detailCombo, row, 1, 1, 2)
        self.detailCombo.currentIndexChanged[str].connect(self.changeDetail)
        row += 1
        self.grid.addWidget(QtGui.QLabel('Period:'), row, 0)
        self.periodCombo = QtGui.QComboBox()
        self.grid.addWidget(self.periodCombo, row, 1) #, 1, 2)
        self.grid.addWidget(self.dayCombo, row, 2) #, 1, 2)
        self.grid.addWidget(self.hourCombo, row, 3)
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
        self.grid.addWidget(QtGui.QLabel('Repeat Loop:'), row, 0)
        self.repeat = QtGui.QCheckBox()
        self.repeat.setCheckState(QtCore.Qt.Unchecked)
        self.grid.addWidget(self.repeat, row, 1)
        self.loopSpin = QtGui.QDoubleSpinBox()
        row += 1
        self.do_loop = False
        self.grid.addWidget(QtGui.QLabel('Period Loop (secs):'), row, 0)
        self.loopSpin = QtGui.QDoubleSpinBox()
        self.loopSpin.setRange(0., 10.)
        self.loopSpin.setDecimals(1)
        self.loopSpin.setSingleStep(.1)
        self.loopSpin.setValue(0.)
        self.loopSpin.valueChanged[str].connect(self.loopChanged)
        self.grid.addWidget(self.loopSpin, row, 1, 1, 2)
        self.loop = QtGui.QPushButton('Loop', self)
        self.grid.addWidget(self.loop, row, 3, 1, 4)
        self.loop.clicked.connect(self.loopClicked)
        row += 1
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, row, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Visualise generation')
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            move_right = False
        else:
            move_right = True
        QtGui.QShortcut(QtGui.QKeySequence('pgup'), self, self.prevClicked)
        QtGui.QShortcut(QtGui.QKeySequence('pgdown'), self, self.nextClicked)
        QtGui.QShortcut(QtGui.QKeySequence('Ctrl++'), self, self.nextClicked)
        QtGui.QShortcut(QtGui.QKeySequence('Ctrl+-'), self, self.prevClicked)
        QtGui.QShortcut(QtGui.QKeySequence('Ctrl+Z'), self, self.loopClicked)
        self.setPeriod()
        self.periodCombo.currentIndexChanged[str].connect(self.periodChange)
        self.dayCombo.currentIndexChanged[str].connect(self.periodChange)
        self.hourCombo.currentIndexChanged[str].connect(self.periodChange)
        self.showPeriod(0)
        self.show()

    def loopChanged(self, val):
        return

    def changeDetail(self, val):
        self.setPeriod()

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
        period = divmod(val, 24)
        if self.detailCombo.currentText() == 'Diurnal':
             self.periodCombo.setCurrentIndex(period[0])
        else:
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
        self.showPeriod(val)

    def loopClicked(self):
        if self.loop.text() == 'Stop':
            self.do_loop = False
            self.scene.exitLoop = False
            self.loop.setText('Loop')
        else:
            self.do_loop = True
            self.loop.setText('Stop')
            self.periodChange()

    def periodChange(self):
        if self.ignore:
            return
        if self.periodCombo.currentIndex() < 0:
            per_ndx = 0
        else:
            per_ndx = self.periodCombo.currentIndex()
        if self.detailCombo.currentText() == 'Diurnal' or self.dayCombo.currentIndex() < 0:
            day_ndx = 0
        else:
            day_ndx = self.dayCombo.currentIndex()
            if day_ndx > self.the_days[per_ndx] - 1:
                day_ndx = self.the_days[per_ndx] - 1
        if self.hourCombo.currentIndex() < 0:
            hr_ndx = 0
        else:
            hr_ndx = self.hourCombo.currentIndex()
        if self.detailCombo.currentText() == 'Diurnal':
            val = (per_ndx + day_ndx) * 24 + hr_ndx
        else:
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
            else:
                if self.repeat.isChecked():
                    val = 0
        if self.do_loop:
            self.do_loop = False
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

    def setPeriod(self, period=None):
        if self.periodCombo.count() > 0:
            self.periodCombo.clear()
        if self.hourCombo.count() > 0:
            self.hourCombo.clear()
        for h in range(24):
            self.hourCombo.addItem('{0:02d}:00'.format(h))
        self.hourCombo.setCurrentIndex(0)
        strt = [0]
        stop = []
        for i in range(len(self.the_days)):
            if i > 0:
                strt.append(stop[-1])
            stop.append(strt[-1] + self.the_days[i] * 24)
        self.data = []
        for s in range(len(self.stations)):
            self.data.append([])
        if self.detailCombo.currentText() == 'Diurnal':
            for s in range(len(self.data)):
                for h in range(24):
                    self.data[s].append(0.)
                if stop[-1] > len(self.powers[s]):
                    continue
                for d in range(strt[0], stop[-1], 24):
                    for h in range(24):
                        self.data[s][h] += self.powers[s][d + h]
                for h in range(24):
                    self.data[s][h] = self.data[s][h] / 365
            self.periodCombo.addItem(self.year)
            ndx = 0
            for m in range(12):
                ndx += 1
                for s in range(len(self.data)):
                    for h in range(24):
                        self.data[s].append(0.)
                    for d in range(strt[m], stop[m], 24):
                        if stop[m] > len(self.powers[s]):
                            break
                        for h in range(24):
                            self.data[s][ndx * 24 + h] += self.powers[s][d + h]
                    for h in range(24):
                        self.data[s][ndx * 24 + h] = self.data[s][ndx * 24 + h] / self.the_days[m]
                self.periodCombo.addItem(self.year + '-' + '{0:02d}'.format(m + 1))
            for i in range(len(self.seasons)):
                ndx += 1
                for s in range(len(self.data)):
                    dys = 0
                    for h in range(24):
                        self.data[s].append(0.)
                    for j in range(1, len(self.seasons[i])):
                        m = self.seasons[i][j] - 1
                        if stop[m] > len(self.powers[s]):
                            break
                        dys += self.the_days[m]
                        for d in range(strt[m], stop[m], 24):
                            for h in range(24):
                                self.data[s][ndx * 24 + h] += self.powers[s][d + h]
                    if dys > 0:
                        for h in range(24):
                            self.data[s][ndx * 24 + h] = self.data[s][ndx * 24 + h] / dys
                self.periodCombo.addItem(self.year + '-' + self.seasons[i][0])
            for i in range(len(self.periods)):
                ndx += 1
                for s in range(len(self.data)):
                    dys = 0
                    for h in range(24):
                        self.data[s].append(0.)
                    for j in range(1, len(self.periods[i])):
                        m = self.periods[i][j] - 1
                        if stop[m] > len(self.powers[s]):
                            break
                        dys += self.the_days[m]
                        for d in range(strt[m], stop[m], 24):
                            for h in range(24):
                                self.data[s][ndx * 24 + h] += self.powers[s][d + h]
                    if dys > 0:
                        for h in range(24):
                            self.data[s][ndx * 24 + h] = self.data[s][ndx * 24 + h] / dys
                self.periodCombo.addItem(self.year + '-' + self.periods[i][0])
            self.dayCombo.hide()
        elif self.detailCombo.currentText() == 'Hourly':
            for m in range(12):
                self.periodCombo.addItem(self.year + '-' + '{0:02d}'.format(m + 1))
            self.setDayItems(0)
            self.data = self.powers[:]
            self.dayCombo.show()
        self.periodCombo.setCurrentIndex(0)
        self.dayCombo.setCurrentIndex(0)
        self.hourCombo.setCurrentIndex(0)
        self.slider.setMaximum(len(self.data[0]) - 1)
        self.periodChange()

    def closeEvent(self, event):
        self.visual_group.setVisible(False)
        self.scene.removeItem(self.visual_group)
        event.accept()

    def quitClicked(self):
        self.close()

    def showPeriod(self, period):
        self.visual_group.setVisible(False)
        while len(self.visual_items) > 0:
            self.visual_group.removeFromGroup(self.visual_items[-1])
            self.visual_items.pop()
        for i in range(len(self.stations)):
            if period >= len(self.data[i]):
                continue
            if self.data[i][period] <= 0:
                continue
            st = self.scene._stations.stations[self.stn_items[i]]
            p = self.scene.mapFromLonLat(QtCore.QPointF(st.lon, st.lat))
            size = math.sqrt(self.data[i][period] * self.scene.capacity_area / math.pi)
            east = self.destinationxy(st.lon, st.lat, 90., size)
            pe = self.scene.mapFromLonLat(QtCore.QPointF(east.x(), st.lat))
            north = self.destinationxy(st.lon, st.lat, 0., size)
            pn = self.scene.mapFromLonLat(QtCore.QPointF(st.lon, north.y()))
            x_d = p.x() - pe.x()
            y_d = pn.y() - p.y()
            el = QtGui.QGraphicsEllipseItem(p.x() - x_d / 2, p.y() - y_d / 2, x_d, y_d)
            el.setBrush(QtGui.QColor(self.scene.colors[st.technology]))
            el.setOpacity(1)
            if self.scene.colors['border'] != '':
                el.setPen(QtGui.QColor(self.scene.colors['border']))
            else:
                el.setPen(QtGui.QColor(self.scene.colors[st.technology]))
            el.setZValue(1)
            self.visual_items.append(el)
            self.visual_group.addToGroup(self.visual_items[-1])
        if  self.detailCombo.currentText() == 'Diurnal':
            txt = self.periodCombo.currentText() + ' ' + self.hourCombo.currentText()
        else:
            txt = self.periodCombo.currentText() + ' ' + self.dayCombo.currentText() + ' ' + self.hourCombo.currentText()
        itm = QtGui.QGraphicsSimpleTextItem(txt)
        new_font = itm.font()
        new_font.setPointSizeF(self.scene.width() / 50)
        itm.setFont(new_font)
        itm.setBrush(QtGui.QColor(self.scene.colors['station_name']))
        fh = int(QtGui.QFontMetrics(new_font).height() * 1.1)
        p = QtCore.QPointF(self.scene.upper_left[0] + fh / 2, self.scene.upper_left[1] + fh / 2)
        frll = self.scene.mapToLonLat(p)
        p = self.scene.mapFromLonLat(QtCore.QPointF(frll.x(), frll.y()))
        p = QtCore.QPoint(p.x(), p.y())
        itm.setPos(p.x(), p.y())
        itm.setZValue(1)
        self.visual_items.append(itm)
        self.visual_group.addToGroup(self.visual_items[-1])
        self.visual_group.setVisible(True)
        QtCore.QCoreApplication.processEvents()
        QtCore.QCoreApplication.flush()
        if self.do_loop and not self.scene.exitLoop:
            if self.loopSpin.value() > 0:
                time.sleep(self.loopSpin.value())
    QtCore.QCoreApplication.processEvents()
    QtCore.QCoreApplication.flush()
