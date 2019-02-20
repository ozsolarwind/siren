#!/usr/bin/python
#
#  Copyright (C) 2015-2019 Sustainable Energy Now Inc., Angus King
#
#  floatmenu.py - This file is part of SIREN.
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
import ConfigParser   # decode .ini file
from PyQt4 import QtGui, QtCore

from editini import SaveIni


class FloatMenu(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)
    procAction = QtCore.pyqtSignal(QtGui.QAction)

    def __init__(self, menubar):
        super(FloatMenu, self).__init__()
        self.menubar = menubar
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
        try:
            open_menus = config.get('Windows', 'menu_open').split(',')
        except:
            open_menus = ''
        self.be_open = True
        self.menus = []
        for lvl1 in self.menubar.actions():
            if str(lvl1.text()).find('&') >= 0:
                try:
                    self.menus.append([str(lvl1.text()).replace('&', ''), lvl1.icon(), None, []])
                except:
                    self.menus.append([str(lvl1.text()).replace('&', ''), None, None, []])
                try:
                    for lvl2 in lvl1.menu().actions():
                        if str(lvl2.text()).find('&') >= 0:
                            try:
                                self.menus[-1][-1].append([str(lvl2.text()).replace('&', ''), lvl2.icon(), None, []])
                            except:
                                self.menus[-1][-1].append([str(lvl2.text()).replace('&', ''), None, None, []])
                            try:
                                for lvl3 in lvl2.menu().actions():
                                    self.menus[-1][-1][-1][-1].append([str(lvl3.text()),
                                        lvl3.icon(), lvl3, '3'])
                            except:
                                pass
                        else:
                            self.thisaction = lvl2
                            self.menus[-1][-1].append([str(lvl2.text()), lvl2.icon(), lvl2, '2'])
                except:
                    pass
            else:
                self.menus.append([str(lvl1.text()), lvl1.icon(), lvl1.actions()[0], '1'])
        self.grid = QtGui.QGridLayout()
        ctr = 0
        self.butn = []
        self.topmenus = {}
        self.buttons = {}
        for i in range(len(self.menus)):
            self.butn.append(QtGui.QPushButton(self.menus[i][0], self))
            self.butn[-1].setIcon(self.menus[i][1])
            self.butn[-1].setStyleSheet('QPushButton {color: #005fb6; border: 2px solid #e65900;' +
                               ' border-radius: 6px;}')
            self.butn[-1].clicked.connect(self.menuClicked)
            self.grid.addWidget(self.butn[-1], ctr, 0, 1, 3)
            ctr += 1
            self.topmenus[self.menus[i][0]] = [False, ctr, 0]
            if type(self.menus[i][3]) is list:
                for j in range(len(self.menus[i][3])):
                    self.butn.append(QtGui.QPushButton(self.menus[i][3][j][0], self))
                    self.butn[-1].setIcon(self.menus[i][3][j][1])
                #     self.butn[-1].setStyleSheet('QPushButton {Text-align:left;}')
                    self.butn[-1].clicked.connect(self.buttonClicked)
                    self.butn[-1].hide()
                    self.buttons[self.menus[i][3][j][0]] = self.menus[i][3][j][2]
                    self.grid.addWidget(self.butn[-1], ctr, 1, 1, 2)
                    ctr += 1
                    if type(self.menus[i][3][j][3]) is list:
                        for k in range(len(self.menus[i][3][j][3])):
                            self.butn.append(QtGui.QPushButton(self.menus[i][3][j][3][k][0], self))
                            self.butn[-1].setIcon(self.menus[i][3][j][3][k][1])
                          #   self.butn[-1].setStyleSheet('QPushButton {Text-align:left;}')
                            self.butn[-1].clicked.connect(self.buttonClicked)
                            self.butn[-1].hide()
                            self.buttons[self.menus[i][3][j][3][k][0]] = self.menus[i][3][j][3][k][2]
                            self.grid.addWidget(self.butn[-1], ctr, 2)
                            ctr += 1
            self.topmenus[self.menus[i][0]][2] = ctr
        quit = QtGui.QPushButton('Quit Menu', self)
        self.grid.addWidget(quit, ctr, 0)
        quit.clicked.connect(self.close)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.close)
        for mnu in open_menus:
            strt = self.topmenus[mnu][1]
            stop = self.topmenus[mnu][2]
            for j in range(strt, stop):
                self.butn[j].show()
            self.topmenus[mnu][0] = True
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Menu')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        if self.restorewindows:
            try:
                rw = config.get('Windows', 'menu_size').split(',')
                self.resize(int(rw[0]), int(rw[1]))
                mp = config.get('Windows', 'menu_pos').split(',')
                self.move(int(mp[0]), int(mp[1]))
            except:
                pass
        self.show()

    def menuClicked(self, event):
        sender = self.sender()
        strt = self.topmenus[str(sender.text())][1]
        stop = self.topmenus[str(sender.text())][2]
        if self.topmenus[str(sender.text())][0]:
            for j in range(strt, stop):
                self.butn[j].hide()
            self.topmenus[str(sender.text())][0] = False
        else:
            for j in range(strt, stop):
                self.butn[j].show()
            self.topmenus[str(sender.text())][0] = True
    #     self.topmenus[str(sender.text())][0] = self.topmenus[str(sender.text())][0] != True

    @QtCore.pyqtSlot(QtGui.QAction)
    def buttonClicked(self, event):
        sender = self.sender()
        self.thetext = str(sender.text())
        self.procAction.emit(self.buttons[str(sender.text())])

    @QtCore.pyqtSlot()
    def exit(self):
        self.be_open = False
        self.close()

    def text(self):
        return self.thetext

    def closeEvent(self, event):
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)   # need to account for border
            lines.append('menu_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('menu_size=%s,%s' % (str(self.width()), str(self.height())))
            open_menus = ''
            for key in self.topmenus.keys():
                if self.topmenus[key][0]:
                    open_menus += ',' + key
            if open_menus != '':
                lines.append('menu_open=%s' % open_menus[1:])
            updates['Windows'] = lines
            SaveIni(updates)
        if self.be_open:
            self.procStart.emit('goodbye')
        event.accept()
