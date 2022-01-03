#!/usr/bin/python3
#
#  Copyright (C) 2017-2022 Sustainable Energy Now Inc., Angus King
#
#  sirenupd.py - This file is part of SIREN.
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
import os
import configparser   # decode .ini file
from PyQt5 import QtCore, QtGui, QtWidgets
import subprocess
import sys

import credits

def get_response(outputs):
    chk_str = 'HTTP request sent, awaiting response... '
    response = '200 OK'
    output = outputs.splitlines()
    for l in range(len(output) -1, -1, -1):
        if len(output[l]) > len(chk_str):
            if output[l][:len(chk_str)] == chk_str:
                response = output[l][len(chk_str):]
                break
    return response


class UpdDialog(QtWidgets.QDialog):

    def __init__(self, ini_file='getfiles.ini', parent=None):
        self.debug = False
        QtWidgets.QDialog.__init__(self, parent)
        self.setWindowTitle('SIREN Update (' + credits.fileVersion() + ') - Check for new versions')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        config = configparser.RawConfigParser()
        config_file = ini_file
        config.read(config_file)
        self.wget_cmd = 'wget --no-check-certificate -O'
        try:
            self.wget_cmd = config.get('sirenupd', 'wget_cmd')
        except:
            pass
        self.host = 'https://sourceforge.net/projects/sensiren/files/'
        try:
            self.host = config.get('sirenupd', 'url')
        except:
            pass
        row = 0
        newgrid = QtWidgets.QGridLayout()
        versions_file = 'siren_versions.csv'
        command = '%s %s %s%s' % (self.wget_cmd, versions_file, self.host, versions_file)
        command = command.split()
        if self.debug:
            response = '200 OK'
        else:
            try:
                pid = subprocess.Popen(command, stderr=subprocess.PIPE)
            except:
                command[0] = command[0] + '.exe'
                pid = subprocess.Popen(command, stderr=subprocess.PIPE)
            response = get_response(pid.communicate()[1])
        if response != '200 OK':
            newgrid.addWidget(QtWidgets.QLabel('Error encountered accessing siren_versions.csv\n\n' + \
                                           response), 0, 0, 1, 4)
            row = 1
        elif os.path.exists(versions_file):
            new_versions = []
            versions = open(versions_file)
            programs = csv.DictReader(versions)
            for program in programs:
                version = credits.fileVersion(program=program['Program'])
             #   print(version, program['Program'], program['Version'])
                if version != '?' and version != program['Version']:
                    cur = version.split('.')
                    new = program['Version'].split('.')
                    if new[0] == cur[0]: # must be same major release
                        curt = ''
                        newt = ''
                        for b in range (1, len(cur)):
                            curt += cur[b].rjust(4, '0')
                            newt += new[b].rjust(4, '0')
                        if newt > curt:
                            new_versions.append([program['Program'], version, program['Version']])
                    elif new[0] > cur[0]:
                        new_versions.append([program['Program'], version, 'New Release'])
            versions.close()
            if len(new_versions) > 0:
                newgrid.addWidget(QtWidgets.QLabel('New versions are available for the following programs.' + \
                                               '\nChoose those you wish to update.' + \
                                               '\nUpdates can take a while so please be patient.' + \
                                               '\nContact siren@sen.asn.au with any issues.'), 0, 0, 1, 4)
                self.table = QtWidgets.QTableWidget()
                self.table.setRowCount(len(new_versions))
                self.table.setColumnCount(4)
                hdr_labels = ['', 'Program', 'New Ver.', 'Current / Status']
                self.table.setHorizontalHeaderLabels(hdr_labels)
                self.table.verticalHeader().setVisible(False)
                self.newbox = []
                self.newprog = []
                rw = -1
                for new_version in new_versions:
                     rw += 1
                     self.newbox.append(QtWidgets.QTableWidgetItem())
                     self.newprog.append(new_version[0])
                     if new_version[2] != 'New Release':
                         self.newbox[-1].setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                         self.newbox[-1].setCheckState(QtCore.Qt.Unchecked)
                     self.table.setItem(rw, 0, self.newbox[-1])
                     self.table.setItem(rw, 1, QtWidgets.QTableWidgetItem(new_version[0]))
                     self.table.setItem(rw, 2, QtWidgets.QTableWidgetItem(new_version[2]))
                     self.table.setItem(rw, 3, QtWidgets.QTableWidgetItem(new_version[1]))
                self.table.resizeColumnsToContents()
                self.table.setColumnWidth(0, 29)
                newgrid.addWidget(self.table, 1, 0, 1, 4)
                doit = QtWidgets.QPushButton('Update')
                doit.clicked.connect(self.doitClicked)
                newgrid.addWidget(doit, 2, 1)
                row = 2
            else:
                newgrid.addWidget(QtWidgets.QLabel('No new versions available.'), 0, 0, 1, 4)
                row = 1
        else:
            newgrid.addWidget(QtWidgets.QLabel('No versions file available.'), 0, 0, 1, 4)
            row = 1
        quit = QtWidgets.QPushButton('Quit')
        quit.clicked.connect(self.quit)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
        newgrid.addWidget(quit, row, 0)
        self.setLayout(newgrid)

    def doitClicked(self):
        default_suffix = sys.argv[0][sys.argv[0].rfind('.'):]
        for p in range(len(self.newbox)):
            if self.newbox[p].checkState() == QtCore.Qt.Checked:
                self.newbox[p].setCheckState(QtCore.Qt.Unchecked)
                self.newbox[p].setFlags(self.newbox[p].flags() ^ QtCore.Qt.ItemIsUserCheckable)
                s = self.newprog[p].rfind('.')
                if s < len(self.newprog[p]) - 4 and self.newprog[p][s:] != '.html':
                    suffix = default_suffix
                else:
                    suffix = ''
                command = '%s %snew%s %s%s%s' % (self.wget_cmd,
                          self.newprog[p], suffix, self.host, self.newprog[p], suffix)
                command = command.split()
                if self.debug:
                    response = '200 OK'
                else:
                    pid = subprocess.Popen(command, stderr=subprocess.PIPE)
                    response = get_response(pid.communicate()[1])
                    if response != '200 OK':
                        errmsg = 'Error ' + response
                        self.table.setItem(p, 3, QtWidgets.QTableWidgetItem(errmsg))
                        continue
                if suffix == '.exe' or suffix == '.py':
                    if os.path.exists(self.newprog[p] + 'new' + suffix):
                        version = credits.fileVersion(program=self.newprog[p] + 'new')
                        if version != '?':
                            if version < self.table.item(p, 3).text():
                                if not self.debug:
                                    os.rename(self.newprog[p] + 'new' + suffix,
                                              self.newprog[p] + '.' + version + suffix)
                                self.table.setItem(p, 3, QtWidgets.QTableWidgetItem('Current is newer'))
                                continue
                if not self.debug:
                    if os.path.exists(self.newprog[p] + suffix + '~'):
                        os.remove(self.newprog[p] + suffix + '~')
                    os.rename(self.newprog[p] + suffix, self.newprog[p] + suffix +'~')
                    os.rename(self.newprog[p] + 'new' + suffix, self.newprog[p] + suffix)
                self.table.setItem(p, 3, QtWidgets.QTableWidgetItem('Updated'))
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(0, 29)

    def quit(self):
        self.close()


if '__main__' == __name__:
    ini_file = 'getfiles.ini'
    if len(sys.argv) > 1:  # arguments
        if sys.argv[1][-4:] == '.ini':
            ini_file = sys.argv[1]
    app = QtWidgets.QApplication(sys.argv)
    upddialog = UpdDialog(ini_file=ini_file)
 #    app.exec_()
 #   app.deleteLater()
 #   sys.exit()
    upddialog.show()
    sys.exit(app.exec_())
