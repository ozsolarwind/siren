#!/usr/bin/python3
#
#  Copyright (C) 2015-2022 Sustainable Energy Now Inc., Angus King
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
from PyQt5 import QtCore, QtGui, QtWidgets
import configparser  # decode .ini file
from getmodels import getModelFile
if sys.platform == 'win32' or sys.platform == 'cygwin':
    from win32api import GetFileVersionInfo, LOWORD, HIWORD

#from editini import SaveIni
import editini

def fileVersion(program=None, year=False):
    ver = '?'
    ver_yr = '????'
    if program == None:
        check = sys.argv[0]
    else:
        s = program.rfind('.')
        if s < 0:
            check = program + sys.argv[0][sys.argv[0].rfind('.'):]
        else:
            check = program
    if check[-3:] == '.py':
        try:
            modtime = datetime.fromtimestamp(os.path.getmtime(check))
            ver = '4.0.%04d.%d%02d' % (modtime.year, modtime.month, modtime.day)
            ver_yr = '%04d' % modtime.year
        except:
            pass
    elif check[-5:] != '.exe':
        try:
            modtime = datetime.fromtimestamp(os.path.getmtime(check))
            ver = '4.0.%04d.%d%02d' % (modtime.year, modtime.month, modtime.day)
            ver_yr = '%04d' % modtime.year
        except:
            pass
    else:
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            try:
                if check.find('\\') >= 0:  # if full path
                    info = GetFileVersionInfo(check, '\\')
                else:
                    info = GetFileVersionInfo(os.getcwd() + '\\' + check, '\\')
                ms = info['ProductVersionMS']
              #  ls = info['FileVersionLS']
                ls = info['ProductVersionLS']
                ver = str(HIWORD(ms)) + '.' + str(LOWORD(ms)) + '.' + str(HIWORD(ls)) + '.' + str(LOWORD(ls))
                ver_yr = str(HIWORD(ls))
            except:
                try:
                    info = os.path.getmtime(os.getcwd() + '\\' + check)
                    ver = '4.0.' + datetime.fromtimestamp(info).strftime('%Y.%m%d')
                    ver_yr = datetime.fromtimestamp(info).strftime('%Y')
                    if ver[9] == '0':
                        ver = ver[:9] + ver[10:]
                except:
                    pass
    if year:
        return ver_yr
    else:
        return ver


class Credits(QtWidgets.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, initial=False):
        super(Credits, self).__init__()
        self.initial = initial
        self.initUI()

    def initUI(self):
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = getModelFile('SIREN.ini')
        config.read(config_file)
        self.restorewindows = False
        try:
            rw = config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
        self.be_open = True
        self.grid = QtWidgets.QGridLayout()
        pixmap = QtGui.QPixmap('sen_logo_2.png')
        lbl = QtWidgets.QLabel(self)
        lbl.setPixmap(pixmap)
        self.grid.addWidget(lbl)
        if os.path.exists('credits.html'):
            web = QtWidgets.QTextEdit()
            htf = open('credits.html', 'r')
            html = htf.read()
            html = html.replace('[VERSION]', fileVersion())
            html = html.replace('[VERSION-YEAR]', fileVersion(year=True))
            htf.close()
            if html[:5] == '<html' or html[:15] == '<!DOCTYPE html>':
                web.setHtml(html)
            else:
                web.setPlainText(html)
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
            labels.append('Copyright (C) 2015-' + fileVersion(year=True) + ' Sustainable Energy Now Inc., Angus King')
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
            labels.append('SIREN may use weather data obtained from MERRA-2 data, a NASA atmospheric data set')
            labels.append('SIREN may use a map derived from OpenStreetMap (MapQuest) Open Aerial Tiles')
            for i in range(len(labels)):
                labl = QtWidgets.QLabel(labels[i])
                labl.setWordWrap(True)
                if i == 0:
                    labl.setFont(bold)
                self.grid.addWidget(labl)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.close)
        QtWidgets.QShortcut(QtGui.QKeySequence('F1'), self, self.help)
        QtWidgets.QShortcut(QtGui.QKeySequence('Ctrl+I'), self, self.about)
        screen = QtWidgets.QDesktopWidget().availableGeometry()
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            pct = 0.85
        else:
            pct = 0.90
        h = int(screen.height() * pct)
        self.resize(pixmap.width() + 55, h)
        frameGm = self.frameGeometry()
        screen = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        tlPoint = QtWidgets.QApplication.desktop().availableGeometry(screen).topLeft()
        frameGm.moveTopLeft(tlPoint)
        self.move(frameGm.topLeft())
        self.setWindowTitle('SIREN - Credits')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
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
            reply = QtWidgets.QMessageBox.question(self, 'SIREN Credits',
                    'Do you want to close Credits window?', QtWidgets.QMessageBox.Yes |
                    QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.Yes:
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
            editini.SaveIni(updates)
        if self.be_open:
            self.procStart.emit('goodbye')
        event.accept()
