#!/usr/bin/python3
#
#  Copyright (C) 2017-2020 Sustainable Energy Now Inc., Angus King
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
from PyQt4 import QtCore, QtGui
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


class UpdDialog(QtGui.QDialog):

    def __init__(self, parent=None):
        self.debug = False
        QtGui.QDialog.__init__(self, parent)
        self.setWindowTitle('SIREN Update (' + credits.fileVersion() + ') - Check for new versions')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        row = 0
        newgrid = QtGui.QGridLayout()
        host = 'https://sourceforge.net/projects/sensiren/files/'
        versions_file = 'siren_versions.csv'
        command = 'wget -O %s %s%s' % (versions_file, host, versions_file)
        command = command.split(' ')
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
            newgrid.addWidget(QtGui.QLabel('Error encountered accessing siren_versions.csv\n\n' + \
                                           response), 0, 0, 1, 4)
            row = 1
        elif os.path.exists(versions_file):
            new_versions = []
            versions = open(versions_file)
            programs = csv.DictReader(versions)
            for program in programs:
                version = credits.fileVersion(program=program['Program'])
                if version != '?' and version != program['Version']:
                    cur = version.split('.')
                    new = program['Version'].split('.')
                    if new[0] == cur[0]: # must be same major release
                        cur[-1] = cur[-1].rjust(4, '0')
                        new[-1] = new[-1].rjust(4, '0')
                        for b in range(len(cur)):
                            if new[b] > cur[b]:
                                new_versions.append([program['Program'], version, program['Version']])
                                break
                    elif new[0] > cur[0]:
                        new_versions.append([program['Program'], version, 'New Release'])
            versions.close()
            if len(new_versions) > 0:
                newgrid.addWidget(QtGui.QLabel('New versions are available for the following programs.' + \
                                               '\nChoose those you wish to update.' + \
                                               '\nUpdates can take a while so please be patient.' + \
                                               '\nContact siren@sen.asn.au with any issues.'), 0, 0, 1, 4)
                self.table = QtGui.QTableWidget()
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
                     self.newbox.append(QtGui.QTableWidgetItem())
                     self.newprog.append(new_version[0])
                     if new_version[2] != 'New Release':
                         self.newbox[-1].setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                         self.newbox[-1].setCheckState(QtCore.Qt.Unchecked)
                     self.table.setItem(rw, 0, self.newbox[-1])
                     self.table.setItem(rw, 1, QtGui.QTableWidgetItem(new_version[0]))
                     self.table.setItem(rw, 2, QtGui.QTableWidgetItem(new_version[2]))
                     self.table.setItem(rw, 3, QtGui.QTableWidgetItem(new_version[1]))
                self.table.resizeColumnsToContents()
                self.table.setColumnWidth(0, 29)
                newgrid.addWidget(self.table, 1, 0, 1, 4)
                doit = QtGui.QPushButton('Update')
                doit.clicked.connect(self.doitClicked)
                newgrid.addWidget(doit, 2, 1)
                row = 2
            else:
                newgrid.addWidget(QtGui.QLabel('No new versions available.'), 0, 0, 1, 4)
                row = 1
        else:
            newgrid.addWidget(QtGui.QLabel('No versions file available.'), 0, 0, 1, 4)
            row = 1
        quit = QtGui.QPushButton('Quit')
        quit.clicked.connect(self.quit)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
        newgrid.addWidget(quit, row, 0)
        self.setLayout(newgrid)

    def doitClicked(self):
        host = 'https://sourceforge.net/projects/sensiren/files/'
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
                command = 'wget -O %snew%s %s%s%s' % (self.newprog[p], suffix, host,
                          self.newprog[p], suffix)
                command = command.split(' ')
                if self.debug:
                    response = '200 OK'
                else:
                    pid = subprocess.Popen(command, stderr=subprocess.PIPE)
                    response = get_response(pid.communicate()[1])
                    if response != '200 OK':
                        errmsg = 'Error ' + response
                        self.table.setItem(p, 3, QtGui.QTableWidgetItem(errmsg))
                        continue
                if suffix == '.exe' or suffix == '.py':
                    if os.path.exists(self.newprog[p] + 'new' + suffix):
                        version = credits.fileVersion(program=self.newprog[p] + 'new')
                        if version != '?':
                            if version < str(self.table.item(p, 3).text()):
                                if not self.debug:
                                    os.rename(self.newprog[p] + 'new' + suffix,
                                              self.newprog[p] + '.' + version + suffix)
                                self.table.setItem(p, 3, QtGui.QTableWidgetItem('Current is newer'))
                                continue
                if not self.debug:
                    if os.path.exists(self.newprog[p] + suffix + '~'):
                        os.remove(self.newprog[p] + suffix + '~')
                    os.rename(self.newprog[p] + suffix, self.newprog[p] + suffix +'~')
                    os.rename(self.newprog[p] + 'new' + suffix, self.newprog[p] + suffix)
                self.table.setItem(p, 3, QtGui.QTableWidgetItem('Updated'))
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(0, 29)

    def quit(self):
        self.close()


if '__main__' == __name__:
    app = QtGui.QApplication(sys.argv)
    upddialog = UpdDialog()
 #    app.exec_()
 #   app.deleteLater()
 #   sys.exit()
    upddialog.show()
    sys.exit(app.exec_())
