#!/usr/bin/python
#
#  Copyright (C) 2016 Sustainable Energy Now Inc., Angus King
#
#  indexweather.py - This file is part of SIREN.
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

import datetime
from math import *
import os
import sys
import time
from PyQt4 import QtCore, QtGui
import ConfigParser   # decode .ini file
import xlwt

import displayobject
from credits import fileVersion
from senuser import getUser

class makeIndex():

    def close(self):
        return

    def getLog(self):
        return self.log

    def __init__(self, what, src_dir, tgt_fil):
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        self.log = ''
        files = []
        fils = os.listdir(src_dir)
        for fil in fils:
            if (what[0].lower() == 's' and (fil[-4:] == '.csv' or fil[-4:] == '.smw')) \
              or (what[0].lower() == 'w' and fil[-4:] == '.srw'):
                tf = open(src_dir + '/' + fil, 'r')
                lines = tf.readlines()
                tf.close()
                if fil[-4:] == '.smw':
                    bits = lines[0].split(',')
                    src_lat = float(bits[4])
                    src_lon = float(bits[5])
                elif fil[-4:] == '.srw':
                    bits = lines[0].split(',')
                    src_lat = float(bits[5])
                    src_lon = float(bits[6])
                elif fil[-10:] == '(TMY2).csv' or fil[-10:] == '(TMY3).csv' \
                  or fil[-10:] == '(INTL).csv' or fil[-4:] == '.csv':
                    fst_row = len(lines) - 8760
                    if fst_row < 3:
                        bits = lines[0].split(',')
                        src_lat = float(bits[4])
                        src_lon = float(bits[5])
                    else:
                        cols = lines[fst_row - 3].split(',')
                        bits = lines[fst_row - 2].split(',')
                        for i in range(len(cols)):
                            if cols[i].lower() in ['latitude', 'lat']:
                                src_lat = float(bits[i])
                            elif cols[i].lower() in ['longitude', 'lon', 'long', 'lng']:
                                src_lon = float(bits[i])
                else:
                    continue
                files.append([src_lat, src_lon, fil])
        if tgt_fil[-4:] == '.xls' or tgt_fil[-5:] == '.xlsx':
            wb = xlwt.Workbook()
            fnt = xlwt.Font()
            fnt.bold = True
            styleb = xlwt.XFStyle()
            styleb.font = fnt
            ws = wb.add_sheet('Index')
            ws.write(0, 0, 'Latitude')
            ws.write(0, 1, 'Longitude')
            ws.write(0, 2, 'Filename')
            lens = [8, 9, 8]
            for i in range(len(files)):
                for c in range(3):
                    ws.write(i + 1, c, files[i][c])
                    lens[c] = max(len(str(files[i][c])), lens[c])
            for c in range(len(lens)):
                if lens[c] * 275 > ws.col(c).width:
                    ws.col(c).width = lens[c] * 275
            ws.set_panes_frozen(True)   # frozen headings instead of split panes
            ws.set_horz_split_pos(1)   # in general, freeze after last heading row
            ws.set_remove_splits(True)   # if user does unfreeze, don't leave a split there
            wb.save(tgt_fil)
        else:
            tf = open(tgt_fil, 'w')
            hdr = 'Latitude,Longitude,Filename\n'
            tf.write(hdr)
            for i in range(len(files)):
                line = '%s,%s,"%s"\n' % (files[i][0], files[i][1], files[i][2])
                tf.write(line)
            tf.close()
        self.log += '%s created' % tgt_fil[tgt_fil.rfind('/') + 1:]


class ClickableQLabel(QtGui.QLabel):
    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtGui.QApplication.widgetAt(event.globalPos()).setFocus()
        self.emit(QtCore.SIGNAL('clicked()'))

class getParms(QtGui.QWidget):

    def __init__(self, help='help.html'):
        super(getParms, self).__init__()
        self.help = help
        self.initUI()

    def initUI(self):
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        try:
            self.base_year = config.get('Base', 'year')
        except:
            self.base_year = '2012'
        self.parents = []
        try:
            self.parents = config.items('Parents')
        except:
            pass
        try:
            self.solarfiles = config.get('Files', 'solar_files')
            for key, value in self.parents:
                self.solarfiles = self.solarfiles.replace(key, value)
            self.solarfiles = self.solarfiles.replace('$USER$', getUser())
            self.solarfiles = self.solarfiles.replace('$YEAR$', self.base_year)
        except:
            self.solarfiles = ''
        try:
            self.solarindex = config.get('Files', 'solar_index')
            for key, value in self.parents:
                self.solarindex = self.solarindex.replace(key, value)
            self.solarindex = self.solarindex.replace('$USER$', getUser())
            self.solarindex = self.solarindex.replace('$YEAR$', self.base_year)
        except:
            self.solarindex = ''
        if self.solarindex == '':
            self.solarindex = self.solarfiles + '/solar_index.xls'
        try:
            self.windfiles = config.get('Files', 'wind_files')
            for key, value in self.parents:
                self.windfiles = self.windfiles.replace(key, value)
            self.windfiles = self.windfiles.replace('$USER$', getUser())
            self.windfiles = self.windfiles.replace('$YEAR$', self.base_year)
        except:
            self.windfiles = ''
        try:
            self.windindex = config.get('Files', 'wind_index')
            for key, value in self.parents:
                self.windindex = self.windindex.replace(key, value)
            self.windindex = self.windindex.replace('$USER$', getUser())
            self.windindex = self.windindex.replace('$YEAR$', self.base_year)
        except:
            self.windindex = ''
        if self.windindex == '':
            self.windindex = self.windfiles + '/wind_index.xls'
        self.grid = QtGui.QGridLayout()
        self.grid.addWidget(QtGui.QLabel('Solar Folder:'), 1, 0)
        self.ssource = ClickableQLabel()
        self.ssource.setText(self.solarfiles)
        self.ssource.setFrameStyle(6)
        self.connect(self.ssource, QtCore.SIGNAL('clicked()'), self.sdirChanged)
        self.grid.addWidget(self.ssource, 1, 1, 1, 4)
        self.grid.addWidget(QtGui.QLabel('Solar Index:'), 2, 0)
        self.starget = ClickableQLabel()
        self.starget.setText(self.solarindex)
        self.starget.setFrameStyle(6)
        self.connect(self.starget, QtCore.SIGNAL('clicked()'), self.stgtChanged)
        self.grid.addWidget(self.starget, 2, 1, 1, 4)
        self.grid.addWidget(QtGui.QLabel('Wind Folder:'), 3, 0)
        self.wsource = ClickableQLabel()
        self.wsource.setText(self.windfiles)
        self.wsource.setFrameStyle(6)
        self.connect(self.wsource, QtCore.SIGNAL('clicked()'), self.wdirChanged)
        self.grid.addWidget(self.wsource, 3, 1, 1, 4)
        self.grid.addWidget(QtGui.QLabel('Wind Index:'), 4, 0)
        self.wtarget = ClickableQLabel()
        self.wtarget.setText(self.windindex)
        self.wtarget.setFrameStyle(6)
        self.connect(self.wtarget, QtCore.SIGNAL('clicked()'), self.wtgtChanged)
        self.grid.addWidget(self.wtarget, 4, 1, 1, 4)
        self.grid.addWidget(QtGui.QLabel('Properties:'), 5, 0)
        self.properties = QtGui.QLineEdit()
        self.properties.setReadOnly(True)
        self.grid.addWidget(self.properties, 5, 1, 1, 4)
        self.log = QtGui.QLabel(' ')
        self.grid.addWidget(self.log, 6, 1, 1, 4)
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, 7, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        dosolar = QtGui.QPushButton('Produce Solar Index', self)
        wdth = dosolar.fontMetrics().boundingRect(dosolar.text()).width() + 9
        dosolar.setMaximumWidth(wdth)
        self.grid.addWidget(dosolar, 7, 1)
        dosolar.clicked.connect(self.dosolarClicked)
        dowind = QtGui.QPushButton('Produce Wind Index', self)
        dowind.setMaximumWidth(wdth)
        self.grid.addWidget(dowind, 7, 2)
        dowind.clicked.connect(self.dowindClicked)
        help = QtGui.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        self.grid.addWidget(help, 7, 3)
        help.clicked.connect(self.helpClicked)
        QtGui.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        self.grid.setColumnStretch(3, 5)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN indexweather (' + fileVersion() + ') - Make resource grid file')
        self.center()
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
        centerPoint = QtGui.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def sdirChanged(self):
        curdir = self.ssource.text()
        newdir = str(QtGui.QFileDialog.getExistingDirectory(self, 'Choose Solar Folder',
                 curdir, QtGui.QFileDialog.ShowDirsOnly))
        if newdir != '':
            self.ssource.setText(newdir)

    def wdirChanged(self):
        curdir = self.wsource.text()
        newdir = str(QtGui.QFileDialog.getExistingDirectory(self, 'Choose Wind Folder',
                 curdir, QtGui.QFileDialog.ShowDirsOnly))
        if newdir != '':
            self.wsource.setText(newdir)

    def stgtChanged(self):
        curtgt = self.starget.text()
        newtgt = str(QtGui.QFileDialog.getSaveFileName(self, 'Choose Solar Index',
                 curtgt))
        if newtgt != '':
            self.starget.setText(newtgt)

    def wtgtChanged(self):
        curtgt = self.starget.text()
        newtgt = str(QtGui.QFileDialog.getSaveFileName(self, 'Choose Wind Index',
                 curtgt))
        if newtgt != '':
            self.wtarget.setText(newtgt)

    def helpClicked(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.help, \
                 title='Help for SIREN indexweather (' + fileVersion() + ')', section='index')
        dialog.exec_()

    def quitClicked(self):
        self.close()

    def dosolarClicked(self):
        solar = makeIndex('Solar', str(self.ssource.text()), str(self.starget.text()))
        self.log.setText(solar.getLog())
        prop = 'solar_index=' + str(self.starget.text())
        self.properties.setText(prop)

    def dowindClicked(self):
        wind = makeIndex('Wind', str(self.wsource.text()), str(self.wtarget.text()))
        self.log.setText(wind.getLog())
        prop = 'wind_index=' + str(self.wtarget.text())
        self.properties.setText(prop)


if "__main__" == __name__:
    app = QtGui.QApplication(sys.argv)
    if len(sys.argv) > 2:   # arguments
        src_dir_s = ''
        src_dir_w = ''
        tgt_fil = ''
        for i in range(1, len(sys.argv)):
            if sys.argv[i][:6] == 'solar=':
                src_dir_s = sys.argv[i][6:]
            elif sys.argv[i][:5] == 'wind=':
                src_dir_w = sys.argv[i][5:]
            elif sys.argv[i][:7] == 'target=' or sys.argv[i][:7] == 'tgtfil=':
                tgt_fil = sys.argv[i][7:]
        if src_dir_s != '':
            files = makeIndex('Solar', src_dir_s, tgt_fil)
        elif src_dir_w != '':
            files = makeIndex('Wind', src_dir_w, tgt_fil)
        else:
            print 'No source directory specified'
    else:
        ex = getParms()
        app.exec_()
        app.deleteLater()
        sys.exit()
