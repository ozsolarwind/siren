#!/usr/bin/python
#
#  Copyright (C) 2016-2019 Sustainable Energy Now Inc., Angus King
#
#  floatlegend.py - This file is part of SIREN.
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

import sys
from PyQt4 import QtCore, QtGui
import ConfigParser   # decode .ini file

from editini import SaveIni


class FloatLegend(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, techdata, stations, flags):
        super(FloatLegend, self).__init__()
        self.stations = stations
        self.techdata = techdata
        self.flags = flags
        self.initUI()

    def initUI(self):
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        self.restorewindows = False
        try:
            rw = config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
        self.be_open = True
        self.grid = QtGui.QGridLayout()
        tech_sizes = {}
        txt_lens = [10, 0]
        for st in self.stations:
            if st.technology[:6] == 'Fossil' and not self.flags[3]:
                continue
            if st.technology not in tech_sizes:
                tech_sizes[st.technology] = 0.
            if st.technology == 'Wind':
                tech_sizes[st.technology] += self.techdata[st.technology][0] * float(st.no_turbines) * pow((st.rotor * .001), 2)
            else:
                tech_sizes[st.technology] += self.techdata[st.technology][0] * float(st.capacity)
            txt_lens[0] = max(txt_lens[0], len(st.technology))
        total_area = 0.
        row = 0
        for key, value in iter(sorted(tech_sizes.iteritems())):
            labl = QtGui.QLabel('__')
            colr = QtGui.QColor(self.techdata[key][1])
            labl.setStyleSheet('QLabel {background-color: %s; color: %s;}' % (colr.name(), colr.name()))
            self.grid.addWidget(labl, row, 0)
            self.grid.addWidget(QtGui.QLabel(key), row, 1)
            total_area += value
            area = '%s sq. Km' % '{:0.1f}'.format(value)
            txt_lens[1] = max(txt_lens[1], len(area))
            self.grid.addWidget(QtGui.QLabel(area), row, 2)
            self.grid.setRowStretch(row, 0)
            row += 1
        self.grid.setColumnStretch(0, 0)
        self.grid.setColumnStretch(1, 1)
        self.grid.addWidget(QtGui.QLabel('Total Area'), row, 1)
        area = '%s sq. Km' % '{:0.1f}'.format(total_area)
        txt_lens[1] = max(txt_lens[1], len(area))
        self.grid.addWidget(QtGui.QLabel(area), row, 2)
        self.grid.setRowStretch(row, 1)
        row += 1
        if self.flags[0]:
            txt = 'Circles show relative capacity in MW'
        elif self.flags[1]:
            txt = 'Circles show relative generation in MWh'
        else:
            txt = 'Station '
            if self.flags[2]:
                txt += 'Squares'
            else:
                txt += 'Circles'
            txt += ' show estimated area in sq. Km'
        if len(txt) > (txt_lens[0] + txt_lens[1]):
            txts = txt.split(' ')
            txt = txts[0]
            ln = len(txt)
            for i in range(1, len(txts)):
                if ln + len(txts[i]) > (txt_lens[0] + txt_lens[1]):
                    txt += '\n'
                    txt += txts[i]
                    ln = len(txts[i])
                else:
                    txt += ' ' + txts[i]
                    ln += len(txts[i]) + 1
        self.grid.addWidget(QtGui.QLabel(txt), row, 1, 1, 2)
        self.grid.setRowStretch(row, 1)
        row += 1
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, row, 1)
        self.grid.setRowStretch(row, 1)
        self.grid.setVerticalSpacing(10)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Legend')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        screen = QtGui.QApplication.desktop().primaryScreen()
        scr_right = QtGui.QApplication.desktop().availableGeometry(screen).right()
        scr_bottom = QtGui.QApplication.desktop().availableGeometry(screen).bottom()
        win_width = self.sizeHint().width()
        win_height = self.sizeHint().height()
        if self.restorewindows:
            try:
                rw = config.get('Windows', 'legend_size').split(',')
                lst_width = int(rw[0])
                lst_height = int(rw[1])
                mp = config.get('Windows', 'legend_pos').split(',')
                lst_left = int(mp[0])
                lst_top = int(mp[1])
                lst_right = lst_left + lst_width
                lst_bottom = lst_top + lst_height
                screen = QtGui.QApplication.desktop().screenNumber(QtCore.QPoint(lst_left, lst_top))
                scr_right = QtGui.QApplication.desktop().availableGeometry(screen).right()
                scr_left = QtGui.QApplication.desktop().availableGeometry(screen).left()
                if lst_right < scr_right:
                    if (lst_right - win_width) >= scr_left:
                        scr_right = lst_right
                    else:
                        scr_right = scr_left + win_width
                scr_bottom = QtGui.QApplication.desktop().availableGeometry(screen).bottom()
                scr_top = QtGui.QApplication.desktop().availableGeometry(screen).top()
                if lst_bottom < scr_bottom:
                    if (lst_bottom - win_height) >= scr_top:
                        scr_bottom = lst_bottom
                    else:
                        scr_bottom = scr_top + win_height
            except:
                pass
        win_left = scr_right - win_width
        win_top = scr_bottom - win_height
        self.resize(win_width, win_height)
        self.move(win_left, win_top)
        self.show()

    @QtCore.pyqtSlot()
    def exit(self):
        self.be_open = False
        self.close()

    def closeEvent(self, event):
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)   # need to account for border
            lines.append('legend_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('legend_size=%s,%s' % (str(self.width()), str(self.height())))
            updates['Windows'] = lines
            SaveIni(updates)
        if self.be_open:
            self.procStart.emit('goodbye')
        event.accept()

    def quitClicked(self):
        self.close()
