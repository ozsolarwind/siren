#!/usr/bin/python
#
#  Copyright (C) 2015-2017 Sustainable Energy Now Inc., Angus King
#
#  credits.py - This file is part of SIREN.
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
from datetime import datetime
import os
import sys
from PyQt4 import QtGui, QtCore
import ConfigParser  # decode .ini file
if sys.platform == 'win32' or sys.platform == 'cygwin':
    from win32api import GetFileVersionInfo, LOWORD, HIWORD

from editini import SaveIni


def fileVersion():
    ver = '1.0.?.?'
    if sys.argv[0][-3:] == '.py':
        modtime = datetime.fromtimestamp(os.path.getmtime(sys.argv[0]))
        ver = '1.0.%04d.%02d.%02d' % (modtime.year, modtime.month, modtime.day)
    else:
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            try:
                if sys.argv[0].find('\\') >= 0:  # if full path
                    info = GetFileVersionInfo(sys.argv[0], '\\')
                else:
                    info = GetFileVersionInfo(os.getcwd() + '\\' + sys.argv[0], '\\')
                ms = info['ProductVersionMS']
              #  ls = info['FileVersionLS']
                ls = info['ProductVersionLS']
                ver = str(HIWORD(ms)) + '.' + str(LOWORD(ms)) + '.' + str(HIWORD(ls)) + '.' + str(LOWORD(ls))
            except:
                pass
    return ver


class Credits(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, initial=False):
        super(Credits, self).__init__()
        self.initial = initial
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
        pixmap = QtGui.QPixmap('sen_logo_2.png')
        lbl = QtGui.QLabel(self)
        lbl.setPixmap(pixmap)
        self.grid.addWidget(lbl)
        if os.path.exists('credits.html'):
            web = QtGui.QTextEdit()
            htf = open('credits.html', 'r')
            html = htf.read()
            html = html.replace('[VERSION]', fileVersion())
            htf.close()
            if html[:5] == '<html' or html[:15] == '<!DOCTYPE html>':
                web.setHtml(QtCore.QString(html))
            else:
                web.setPlainText(QtCore.QString(html))
            web.setReadOnly(True)
            small = QtGui.QFont()
            small.setPointSize(10)
            web.setFont(small)
            self.grid.addWidget(web, 1, 0)
        else:
            bold = QtGui.QFont()
            bold.setBold(True)
            labels = []
            labels.append("SIREN - SEN's Interactive Renewable Energy Network tool")
            labels.append('Copyright (C) 2015-2017 Sustainable Energy Now Inc., Angus King')
            labels.append('Release:' + fileVersion())
            labels.append('SIREN is free software: you can redistribute it and/or modify it under the terms of the' +
                          ' GNU Affero General Public License. The program is distributed WITHOUT ANY WARRANTY')
            labels.append('The SEN SAM simulation is used to calculate energy generation for renewable energy' +
                          ' power stations using SAM models.')
            labels.append('To get started press F1 (menu option Help -> Help) to view the SIREN Help file.')
            labels.append('Capabilities, assumptions, limitations (ie transmission, geographic capabilities,' +
                          ' etc), verification')
            labels.append('Contact angus@ozsolarwind.com for more information or alternative modelling' +
                          ' capabilities/suggestions needed')
            labels.append('We acknowledge US DOE, NASA and OpenStreetMap as follows (press Ctrl+I, menu option' +
                          ' Help -> About, for more details):')
            labels.append('SIREN uses System Advisor Model (SAM) modules for energy calculations. SAM is provided' +
                          ' by NREL for the US DOE')
            labels.append('SIREN may use weather data obtained from MERRA data, a NASA atmospheric data set')
            labels.append('SIREN may use a map derived from OpenStreetMap (MapQuest) Open Aerial Tiles')
            for i in range(len(labels)):
                labl = QtGui.QLabel(labels[i])
                labl.setWordWrap(True)
                if i == 0:
                    labl.setFont(bold)
                self.grid.addWidget(labl)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.close)
        QtGui.QShortcut(QtGui.QKeySequence('F1'), self, self.help)
        QtGui.QShortcut(QtGui.QKeySequence('Ctrl+I'), self, self.about)
        screen = QtGui.QDesktopWidget().availableGeometry()
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            pct = 0.85
        else:
            pct = 0.90
        h = int(screen.height() * pct)
        self.resize(pixmap.width() + 55, h)
        frameGm = self.frameGeometry()
        screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
        tlPoint = QtGui.QApplication.desktop().availableGeometry(screen).topLeft()
        frameGm.moveTopLeft(tlPoint)
        self.move(frameGm.topLeft())
        self.setWindowTitle('SIREN - Credits')
        if self.restorewindows:
            try:
                rw = config.get('Windows', 'credits_size').split(',')
                self.resize(int(rw[0]), int(rw[1]))
                mp = config.get('Windows', 'credits_pos').split(',')
                self.move(int(mp[0]), int(mp[1]))
            except:
                pass
        self.show()

    @QtCore.pyqtSlot()
    def help(self):
        self.procStart.emit('help')

    def about(self):
        self.procStart.emit('about')

    def exit(self):
        self.be_open = False
        self.close()

    def closeEvent(self, event):
        if self.be_open:
            reply = QtGui.QMessageBox.question(self, 'SIREN Credits',
                    'Do you want to close Credits window?', QtGui.QMessageBox.Yes |
                    QtGui.QMessageBox.No, QtGui.QMessageBox.No)
            if reply == QtGui.QMessageBox.Yes:
                pass
            else:
                event.ignore()
                return
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)  # need to account for border
            lines.append('credits_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('credits_size=%s,%s' % (str(self.width()), str(self.height())))
            updates['Windows'] = lines
            SaveIni(updates)
        if self.be_open:
            self.procStart.emit('goodbye')
        event.accept()
