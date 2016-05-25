#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
#
#  updateswis.py - This file is part of SIREN.
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

import csv
import httplib
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
from station import Stations

class makeFile():

    def close(self):
        return

    def getLog(self):
        return self.log

    def __init__(self, host, url, tgt_fil, excel=''):
        self.host = host
        self.url = url
        self.tgt_fil = tgt_fil
        self.excel = excel
        http_host = 'data.wa.aemo.com.au'
        url = self.url
        conn = httplib.HTTPConnection(http_host)
        conn.request("GET", url)
        response = conn.getresponse()
        if response.status != 200:
            self.log = 'Error Response: ' + str(response.status) + ' ' + response.reason
            return
        datas = response.read().split('\n')
        conn.close()
        common_fields = ['Facility Code', 'Participant Name', 'Participant Code', \
                         'Facility Type', 'Balancing Status', 'Capacity Credits (MW)', \
                         'Maximum Capacity (MW)', 'Registered From']
        fac_file = self.tgt_fil
        facile = open(fac_file)
        facilities = csv.DictReader(facile)
        new_facilities = csv.DictReader(datas)
        self.log = ''
        for new_field in new_facilities.fieldnames:
            for field in common_fields:
                if new_field == field:
                    break
            else:
                if new_field != 'Extracted At':
                    self.log += 'New field:', new_field + '\n'
        changes = 0
        for new_facility in new_facilities:
            if new_facility['Extracted At'] != '':
                info = 'Extracted At: ' + new_facility['Extracted At']
                break
        for new_facility in new_facilities:
            facile.seek(0)
            for facility in facilities:
                if new_facility['Facility Code'] == facility['Facility Code']:
                    for field in common_fields:
                        if new_facility[field] != facility[field]:
                            if field == 'Registered From' and facility[field][0] != '2':
                                new_time = time.strptime(new_facility[field], '%Y-%m-%d 00:00:00')
                                new_date = time.strftime('%B %d %Y', new_time)
                                new_date = new_date.replace(' 0', ' ')
                                if new_date == facility[field]:
                                    continue
                            try:
                                field_was = float(facility[field])
                                field_new = float(new_facility[field])
                                if field_was == field_new:
                                    continue
                            except:
                                pass
                            self.log += "Changed field in '%s:'\n    '%s' was '%s', now '%s'\n" % \
                                         (facility['Facility Code'], field, facility[field], \
                                          new_facility[field])
                            changes += 1
                    break
            else:
                self.log += 'New facility', new_facility['Facility Code'] + '\n'
        facile.seek(0)
        for facility in facilities:
            new_facilities = csv.DictReader(datas)
            for new_facility in new_facilities:
                if new_facility['Facility Code'] == facility['Facility Code']:
                    break
            else:
                self.log += "Facility '%s' no longer on file\n" % facility['Facility Code']
        if changes > 0:
            msgbox = QtGui.QMessageBox()
            msgbox.setWindowTitle('updateswis Status')
            msgbox.setText(str(changes) + ' changes. Do you want to replace existing file (Y)?')
            msgbox.setInformativeText(info)
            msgbox.setDetailedText(self.log)
            msgbox.setIcon(QtGui.QMessageBox.Question)
            msgbox.setStandardButtons(QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
            reply = msgbox.exec_()
            if reply == QtGui.QMessageBox.Yes:
                extra_fields = ['Latitude', 'Longitude', 'Facility Name', 'Turbine', 'No. turbines']
                upd_file = open(fac_file + '.csv', 'wb')
                upd_writer = csv.writer(upd_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                upd_writer.writerow(common_fields + extra_fields)
                new_facilities = csv.DictReader(datas)
                for new_facility in new_facilities:
                    if new_facility['Facility Code'] == 'Facility Code':   # for some reason headers are duplicated
                        continue
                    new_line = []
                    for field in common_fields:
                        new_line.append(new_facility[field])
                    facile.seek(0)
                    for facility in facilities:
                        if new_facility['Facility Code'] == facility['Facility Code']:
                            for field in extra_fields:
                                new_line.append(facility[field])
                            break
                    else:
                        for field in extra_fields:
                            new_line.append('')
                    upd_writer.writerow(new_line)
                upd_file.close()
                if os.path.exists(fac_file + '~'):
                    os.remove(fac_file + '~')
                os.rename(fac_file, fac_file + '~')
                os.rename(fac_file + '.csv', fac_file)
            self.log = '%s updated' % tgt_fil[tgt_fil.rfind('/') + 1:]
        else:
            self.log = 'No changes to existing file. No update required.'
        if self.excel != '':
            if self.excel[-4:] == '.csv' or self.excel[-4:] == '.xls' or self.excel[-5:] == '.xlsx':
                pass
            else:
                self.excel += '.xls'
            if os.path.exists(self.excel):
                if os.path.exists(self.excel + '~'):
                    os.remove(self.excel + '~')
                os.rename(self.excel, self.excel + '~')
            ctr = 0
            d = 0
            fields = ['Station Name', 'Technology', 'Latitude', 'Longitude', 'Maximum Capacity (MW)', \
                      'Turbine', 'Rotor Diam', 'No. turbines', 'Area']
            stations = Stations(stations2=False)
            if self.excel[-4:] == '.csv':
                upd_file = open(self.excel, 'wb')
                upd_file.write('Description:,"SWIS Existing Stations"\n')
                upd_writer = csv.writer(upd_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                upd_writer.writerow(fields)
                for stn in stations.stations:
                    new_line = []
                    new_line.append(stn.name)
                    new_line.append(stn.technology)
                    new_line.append(str(stn.lat))
                    new_line.append(str(stn.lon))
                    new_line.append(str(stn.capacity))
                    if stn.technology == 'Wind':
                        new_line.append(stn.turbine)
                        new_line.append(str(stn.rotor))
                        new_line.append(str(stn.no_turbines))
                    else:
                        new_line.append('')
                        new_line.append('0')
                        new_line.append('0')
                    new_line.append(str(stn.area))
                    upd_writer.writerow(new_line)
                    ctr += 1
                upd_file.close()
            else:
                wb = xlwt.Workbook()
                fnt = xlwt.Font()
                fnt.bold = True
                styleb = xlwt.XFStyle()
                styleb.font = fnt
                lens = []
                for i in range(len(fields)):
                    lens.append(len(fields[i]))
                ws = wb.add_sheet('Existing')
                d = 0
                ws.write(ctr, 0, 'Description:')
                ws.write_merge(ctr, ctr, 1, 9, 'SWIS Existing Stations')
                d = -1
                ctr += 1
                for i in range(len(fields)):
                    ws.write(ctr, i, fields[i])
                for stn in stations.stations:
                    ctr += 1
                    ws.write(ctr, 0, stn.name)
                    lens[0] = max(lens[0], len(stn.name))
                    ws.write(ctr, 1, stn.technology)
                    lens[1] = max(lens[1], len(stn.technology))
                    ws.write(ctr, 2, stn.lat)
                    lens[2] = max(lens[2], len(str(stn.lat)))
                    ws.write(ctr, 3, stn.lon)
                    lens[3] = max(lens[3], len(str(stn.lon)))
                    ws.write(ctr, 4, stn.capacity)
                    lens[4] = max(lens[4], len(str(stn.capacity)))
                    if stn.technology == 'Wind':
                        ws.write(ctr, 5, stn.turbine)
                        lens[5] = max(lens[5], len(stn.turbine))
                        ws.write(ctr, 6, stn.rotor)
                        lens[6] = max(lens[6], len(str(stn.rotor)))
                        ws.write(ctr, 7, stn.no_turbines)
                        lens[7] = max(lens[7], len(str(stn.no_turbines)))
                    else:
                        ws.write(ctr, 6, 0)
                        ws.write(ctr, 7, 0)
                    ws.write(ctr, 8, stn.area)
                    lens[8] = max(lens[8], len(str(stn.area)))
                for c in range(9):
                    if lens[c] * 275 > ws.col(c).width:
                        ws.col(c).width = lens[c] * 275
                ws.set_panes_frozen(True)   # frozen headings instead of split panes
                ws.set_horz_split_pos(1 - d)   # in general, freeze after last heading row
                ws.set_remove_splits(True)   # if user does unfreeze, don't leave a split there
                wb.save(self.excel)


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
        self.yrndx = -1
        try:
            self.years = []
            years = config.get('Base', 'years')
            bits = years.split(',')
            for i in range(len(bits)):
                rngs = bits[i].split('-')
                if len(rngs) > 1:
                    for j in range(int(rngs[0].strip()), int(rngs[1].strip()) + 1):
                        if str(j) == self.base_year:
                            self.yrndx = len(self.years)
                        self.years.append(str(j))
                else:
                    if rngs[0].strip() == self.base_year:
                        self.yrndx = len(self.years)
                    self.years.append(rngs[0].strip())
        except:
            self.years = self.base_year
            self.yrndx = 0
        if self.yrndx < 0:
            self.yrndx = len(self.years)
            self.years.append(self.base_year)
        parents = []
        try:
            parents = config.items('Parents')
        except:
            pass
        self.grid_stations = ''
        try:
            fac_file = config.get('Files', 'grid_stations')
            for key, value in parents:
                fac_file = fac_file.replace(key, value)
            fac_file = fac_file.replace('$USER$', getUser())
            fac_file = fac_file.replace('$YEAR$', self.base_year)
            self.grid_stations = fac_file
        except:
            pass
        self.grid = QtGui.QGridLayout()
        self.grid.addWidget(QtGui.QLabel('Host site:'), 0, 0)
        self.host = QtGui.QLineEdit()
        self.host.setText('data.wa.aemo.com.au')
        self.grid.addWidget(self.host, 0, 1, 1, 2)
        self.grid.addWidget(QtGui.QLabel('File location:'), 1, 0)
        self.url = QtGui.QLineEdit()
        self.url.setText('/datafiles/facilities/facilities.csv')
        self.grid.addWidget(self.url, 1, 1, 1, 2)
        self.grid.addWidget(QtGui.QLabel('Target file:'), 2, 0)
        self.target = ClickableQLabel()
        self.target.setText(self.grid_stations)
        self.target.setFrameStyle(6)
        self.connect(self.target, QtCore.SIGNAL('clicked()'), self.tgtChanged)
        self.grid.addWidget(self.target, 2, 1, 1, 3)
        self.grid.addWidget(QtGui.QLabel('Excel file:'), 3, 0)
        self.excel = ClickableQLabel()
        self.excel.setText('')
        self.excel.setFrameStyle(6)
        self.connect(self.excel, QtCore.SIGNAL('clicked()'), self.excelChanged)
        self.grid.addWidget(self.excel, 3, 1, 1, 3)
        self.log = QtGui.QLabel(' ')
        self.grid.addWidget(self.log, 4, 1, 1, 3)
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, 5, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        dofile = QtGui.QPushButton('Update Existing Stations', self)
        self.grid.addWidget(dofile, 5, 1)
        dofile.clicked.connect(self.dofileClicked)
        help = QtGui.QPushButton('Help', self)
        self.grid.addWidget(help, 5, 2)
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
        self.setWindowTitle('SIREN updateswis (' + fileVersion() + ') - Update SWIS Stations')
        self.center()
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
        centerPoint = QtGui.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def tgtChanged(self):
        curtgt = self.target.text()
        newtgt = str(QtGui.QFileDialog.getSaveFileName(self, 'Choose Target File',
                 curtgt))
        if newtgt != '':
            self.target.setText(newtgt)

    def excelChanged(self):
        curtgt = self.excel.text()
        newtgt = str(QtGui.QFileDialog.getSaveFileName(self, 'Choose Excel File',
                 curtgt))
        if newtgt != '':
            self.excel.setText(newtgt)

    def helpClicked(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.help, \
                 title='Help for SIREN updateswis (' + fileVersion() + ')', section='updateswis')
        dialog.exec_()

    def quitClicked(self):
        self.close()

    def dofileClicked(self):
        resource = makeFile(str(self.host.text()), str(self.url.text()), str(self.target.text()),
                            str(self.excel.text()))
        log = resource.getLog()
        self.log.setText(log)


if "__main__" == __name__:
    app = QtGui.QApplication(sys.argv)
    if len(sys.argv) > 2:   # arguments
        host = 'data.wa.aemo.com.au'
        url = '/datafiles/facilities/facilities.csv'
        tgt_fil = ''
        excel = ''
        for i in range(1, len(sys.argv)):
            if sys.argv[i][:5] == 'host=':
                host = int(sys.argv[i][5:])
            elif sys.argv[i][:4] == 'url=':
                url = sys.argv[i][4:]
            elif sys.argv[i][:7] == 'target=' or sys.argv[i][:7] == 'tgtfil=':
                tgt_fil = sys.argv[i][7:]
            elif sys.argv[i][:6] == 'excel=':
                excel = sys.argv[i][6:]
        files = makeFile(host, url, tgt_fil, excel)
    else:
        ex = getParms()
        app.exec_()
        app.deleteLater()
        sys.exit()
