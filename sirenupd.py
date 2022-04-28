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
# Note: At present sirenupd only allows updates for existing files
#       (or zip files for Python). In the future it may be
#       expanded to allow new files to be distributed

import csv
from datetime import datetime
import os
import configparser   # decode .ini file
from PyQt5 import QtCore, QtGui, QtWidgets
from shutil import copy2
import subprocess
import sys
import tempfile
import time
import zipfile

import credits
from senuser import getUser

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

    def get_new_versions(self, versions_file, local=''):
        versions = open(versions_file)
        programs = csv.DictReader(versions)
        for program in programs:
            version = credits.fileVersion(program=program['Program'].replace('.zip', '.py'))
            if version != '?' and version != program['Version']:
                cur = version.split('.')
                new = program['Version'].split('.')
                if new[0] == cur[0]: # must be same major release
                    curt = ''
                    newt = ''
                    for b in range (1, len(cur)):
                        curt += cur[b].rjust(4, '0')
                        newt += new[b].rjust(4, '0')
                    if newt > curt or self.debug:
                        self.new_versions.append([program['Program'], program['Version'] + local, version])
                elif new[0] > cur[0]:
                    self.new_versions.append([program['Program'], 'New Release' + local, version])
        versions.close()

    def __init__(self, ini_file='getfiles.ini', parent=None):
        QtWidgets.QDialog.__init__(self, parent)
        self.setWindowTitle('SIREN Update (' + credits.fileVersion() + ') - Check for new versions')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        config = configparser.RawConfigParser()
        config_file = ini_file
        config.read(config_file)
        self.debug = False
        try:
            debug = config.get('sirenupd', 'debug')
            if debug.lower() in ['true', 'yes', 'on']:
                self.debug = True
        except:
            pass
        self.wget_cmd = 'wget --no-check-certificate -O'
        try:
            self.wget_cmd = config.get('sirenupd', 'wget_cmd')
        except:
            pass
        try:
            self.local = config.get('sirenupd', 'local')
            self.local = self.local.replace('$USER$', getUser())
        except:
            self.local = ''
        self.host = 'https://sourceforge.net/projects/sensiren/files/'
        try:
            self.host = config.get('sirenupd', 'url')
        except:
            pass
        self.new_versions = []
        row = 0
        newgrid = QtWidgets.QGridLayout()
        versions_file = 'siren_versions.csv'
        if self.local != '': # local location?
            versions_file = self.local + versions_file
            if os.path.exists(versions_file):
                self.get_new_versions(versions_file, local=' (local)')
        # now remote versions
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
            self.get_new_versions(versions_file)
        else:
            newgrid.addWidget(QtWidgets.QLabel('No versions file available.'), 0, 0, 1, 4)
            row = 1
        if len(self.new_versions) > 0:
            self.new_versions.sort()
            if self.debug:
                newgrid.addWidget(QtWidgets.QLabel('Debug mode.'), row, 0, 1, 4)
            else:
                newgrid.addWidget(QtWidgets.QLabel('New versions are available for the following programs.' + \
                                           '\nChoose those you wish to update.' + \
                                           '\nUpdates can take a while so please be patient.' + \
                                           '\nContact siren@sen.asn.au with any issues.'), row, 0, 1, 4)
            self.table = QtWidgets.QTableWidget()
            self.table.setRowCount(len(self.new_versions))
            self.table.setColumnCount(4)
            hdr_labels = ['', 'Program', 'New Ver.', 'Current / Status']
            self.table.setHorizontalHeaderLabels(hdr_labels)
            self.table.verticalHeader().setVisible(False)
            self.newbox = []
            self.newprog = []
            rw = -1
            for new_version in self.new_versions:
                 rw += 1
                 self.newbox.append(QtWidgets.QTableWidgetItem())
                 self.newprog.append(new_version[0])
                 if new_version[1][:11] != 'New Release':
                     self.newbox[-1].setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                     self.newbox[-1].setCheckState(QtCore.Qt.Unchecked)
                 self.table.setItem(rw, 0, self.newbox[-1])
                 self.table.setItem(rw, 1, QtWidgets.QTableWidgetItem(new_version[0]))
                 self.table.setItem(rw, 2, QtWidgets.QTableWidgetItem(new_version[1]))
                 self.table.setItem(rw, 3, QtWidgets.QTableWidgetItem(new_version[2]))
            self.table.resizeColumnsToContents()
            newgrid.addWidget(self.table, 1, 0, 1, 4)
            doit = QtWidgets.QPushButton('Update')
            doit.clicked.connect(self.doitClicked)
            newgrid.addWidget(doit, 2, 1)
            row = 2
        else:
            newgrid.addWidget(QtWidgets.QLabel('No new versions available.'), 0, 0, 1, 4)
            row = 1
        quit = QtWidgets.QPushButton('Quit')
        quit.clicked.connect(self.quit)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
        newgrid.addWidget(quit, row, 0)
        self.setLayout(newgrid)
        self.resize(500, int(self.sizeHint().height()))

    def doitClicked(self):
        def do_zipfile(zip_in, temp_dir): # copy multiple files
            zf = zipfile.ZipFile(zip_in, 'r')
            ctr = 0
            for zi in zf.infolist():
                if zi.filename[:-4] == '.exe':
                    if temp_dir is None:
                         temp_dir = tempfile.gettempdir() + '/'
                    zf.extract(zi, temp_dir)
                    newver = credits.fileVersion(program=temp_dir + zi.filename)
                    if newver != '?':
                        curver = credits.fileVersion(program=zi.filename)
                        if newver > curver or curver == '?':
                            if not self.debug:
                                if os.path.exists(zi.filename + '~'):
                                    os.remove(zi.filename + '~')
                                if cur_verson != '?':
                                    os.rename(zi.filename, zi_filename + '~')
                                os.rename(temp_dir + zi.filename, zi.filename)
                            ctr += 1
                else:
                    newtime = datetime.fromtimestamp(time.mktime(zi.date_time + (0, 0, -1)))
                    try:
                        curtime = datetime.fromtimestamp(int(os.path.getmtime(zi.filename)))
                    except:
                        curtime = 0
                    if curtime == 0 or newtime > curtime:
                        # print(zi.filename, newtime, curtime)
                        if not self.debug:
                            if os.path.exists(zi.filename + '~'):
                                os.remove(zi.filename + '~')
                            if curtime != 0:
                                os.rename(zi.filename, zi.filename + '~')
                            zf.extract(zi)
                            date_time = time.mktime(zi.date_time + (0, 0, -1))
                            os.utime(zi.filename, (date_time, date_time))
                        ctr += 1
            if ctr == 0:
                msg = 'Current is newer'
            else:
                msg = f'Updated {ctr} files'
            return msg

        default_suffix = sys.argv[0][sys.argv[0].rfind('.'):]
        temp_dir = None
        for p in range(len(self.newbox)):
            if self.newbox[p].checkState() == QtCore.Qt.Checked:
                self.newbox[p].setCheckState(QtCore.Qt.Unchecked)
                self.newbox[p].setFlags(self.newbox[p].flags() ^ QtCore.Qt.ItemIsUserCheckable)
                s = self.newprog[p].rfind('.')
                if s < 0:
                    newprog = self.newprog[p]
                    suffix = default_suffix
                else:
                    newprog = self.newprog[p][:s]
                    suffix = self.newprog[p][s:]
                if self.table.item(p, 2).text().find('(local)') > 0:
                    if suffix == '.zip': # copy multiple files
                        msg = do_zipfile(self.local + newprog + suffix, temp_dir)
                        self.table.setItem(p, 3, QtWidgets.QTableWidgetItem(msg))
                        continue
                    else:
                        curver = credits.fileVersion(program=newprog + suffix)
                        #  print(newprog, suffix, curver, self.table.item(p, 2).text())
                        if curver != '?':
                            if self.table.item(p, 2).text() < curver:
                                self.table.setItem(p, 3, QtWidgets.QTableWidgetItem('Current is newer'))
                                continue
                    if not self.debug:
                        if os.path.exists(newprog + suffix + '~'):
                            os.remove(newprog + suffix + '~')
                        os.rename(newprog + suffix, newprog + suffix + '~')
                        copy2(self.local + newprog + suffix, newprog + suffix)
                    self.table.setItem(p, 3, QtWidgets.QTableWidgetItem('Updated'))
                else:
                    command = '%s %snew%s %s%s%s' % (self.wget_cmd,
                              newprog, suffix, self.host, newprog, suffix)
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
                    if suffix == '.zip': # copy multiple files
                        msg = do_zipfile(newprog + 'new' + suffix, temp_dir)
                        self.table.setItem(p, 3, QtWidgets.QTableWidgetItem(msg))
                        continue
                    else:
                        if os.path.exists(newprog + 'new' + suffix):
                            newver = credits.fileVersion(program=newprog + 'new')
                            if newver != '?':
                                if newver < self.table.item(p, 3).text():
                                    if not self.debug:
                                        os.rename(newprog + 'new' + suffix,
                                                  newprog + '.' + newver + suffix)
                                    self.table.setItem(p, 3, QtWidgets.QTableWidgetItem('Current is newer'))
                                    continue
                    if not self.debug:
                        if os.path.exists(newprog + suffix + '~'):
                            os.remove(newprog + suffix + '~')
                        os.rename(newprog + suffix, newprog + suffix + '~')
                        os.rename(newprog + 'new' + suffix, newprog + suffix)
                    self.table.setItem(p, 3, QtWidgets.QTableWidgetItem('Updated'))
        self.table.resizeColumnsToContents()

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
