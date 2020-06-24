#!/usr/bin/python3
#
#  Copyright (C) 2016-2020 Sustainable Energy Now Inc., Angus King
#
#  siren.py - This file is part of SIREN.
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

import configparser   # decode .ini file
import datetime
from functools import partial
import os
from PyQt5 import QtCore, QtGui, QtWidgets
from shutil import copy
import subprocess
import sys
import time
import webbrowser

import displayobject
from credits import fileVersion
from editini import EdtDialog
from getmodels import getModelFile, commonprefix
from senuser import getUser


class TabDialog(QtWidgets.QDialog):

    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent)
        self.siren_dir = '.'
        if len(sys.argv) > 1:
            if sys.argv[1][-4:] == '.ini':
                self.invoke(sys.argv[1])
                sys.exit()
            elif os.path.isdir(sys.argv[1]):
                self.siren_dir = sys.argv[1]
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                if self.siren_dir[-1] != '\\' and self.siren_dir[-1] != '/':
                    self.siren_dir += '\\'
            elif self.siren_dir[-1] != '/':
                self.siren_dir += '/'
        else:
            self.siren_dir = getModelFile()
        self.entries = []
        fils = os.listdir(self.siren_dir)
        self.help = ''
        self.about = ''
        config = configparser.RawConfigParser()
        ignore = ['flexiplot.ini', 'getfiles.ini', 'powerplot.ini', 'siren_default.ini',
                  'siren_windows_default.ini']
        errors = ''
        for fil in sorted(fils):
            if fil[-4:] == '.ini':
                if fil in ignore:
                    continue
                mod_time = time.strftime('%Y-%m-%d %H:%M:%S',
                           time.localtime(os.path.getmtime(self.siren_dir + fil)))
                try:
                    config.read(self.siren_dir + fil)
                except configparser.DuplicateOptionError as err:
                    errors += 'DuplicateOptionError ' + str(err) + '\n'
                    continue
                except:
                    err = sys.exc_info()[0]
                    errors += 'Error: ' + str(err) + ' While reading from ' + self.siren_dir + fil + '\n'
                    continue
                try:
                    model_name = config.get('Base', 'name')
                except:
                    model_name = ''
                self.entries.append([fil, model_name, mod_time])
                if self.about == '':
                    try:
                        self.about = config.get('Files', 'about')
                        if not os.path.exists(self.about):
                            self.about = ''
                    except:
                        pass
                if self.help == '':
                    try:
                        self.help = config.get('Files', 'help')
                        if not os.path.exists(self.help):
                            self.help = ''
                    except:
                        pass
        if len(errors) > 0:
            dialog = displayobject.AnObject(QtWidgets.QDialog(), errors,
                     title='SIREN (' + fileVersion() + ') - Preferences file errors')
            dialog.exec_()
        if len(self.entries) == 0:
            self.new()
     #    if len(entries) == 1:
     #        self.invoke(entries[0][0])
     #        sys.exit()
        self.setWindowTitle('SIREN (' + fileVersion() + ') - Select SIREN Model')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        buttonLayout = QtWidgets.QHBoxLayout()
        self.quitButton = QtWidgets.QPushButton(self.tr('&Quit'))
        buttonLayout.addWidget(self.quitButton)
        self.quitButton.clicked.connect(self.quit)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
        self.newButton = QtWidgets.QPushButton(self.tr('&New Model'))
        buttonLayout.addWidget(self.newButton)
        self.newButton.clicked.connect(self.new)
        buttons = QtWidgets.QFrame()
        buttons.setLayout(buttonLayout)
        layout = QtWidgets.QGridLayout()
        self.table = QtWidgets.QTableWidget()
        self.table.setRowCount(len(self.entries))
        self.table.setColumnCount(3)
        hdr_labels = ['Preference File', 'SIREN Model', 'Date modified']
        self.table.setHorizontalHeaderLabels(hdr_labels)
        self.headers = self.table.horizontalHeader()
        self.headers.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.headers.customContextMenuRequested.connect(self.header_click)
        self.headers.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        max_row = 30
        for rw in range(len(self.entries)):
            ln = 0
            for cl in range(3):
                self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(self.entries[rw][cl]))
                ln += len(self.entries[rw][cl])
            if ln > max_row:
                max_row = ln
        self.sort_asc = True # start in date descending order
        self.sort_col = 2
        self.order(2)
        self.table.resizeColumnsToContents()
        self.table.itemClicked.connect(self.Clicked)
        fnt = self.table.fontMetrics()
        ln = max_row * max(9, fnt.averageCharWidth())
        ln2 = (len(self.entries) + 8) * (fnt.xHeight() + fnt.lineSpacing())
        screen = QtWidgets.QDesktopWidget().screenGeometry()
        if ln > screen.width() * .9:
            ln = int(screen.width() * .9)
        if ln2 > screen.height() * .9:
            ln2 = int(screen.height() * .9)
        layout.addWidget(QtWidgets.QLabel('Click on row for Desired Model; Right click column header to sort'), 0, 0)
        layout.addWidget(self.table, 1, 0)
        layout.addWidget(buttons, 2, 0)
        menubar = QtWidgets.QMenuBar()
        utilities = ['flexiplot', 'getmap', 'getmerra2', 'makeweatherfiles', 'powerplot', 'sirenupd']
        utilicon = ['line.png', 'map.png', 'download.png', 'weather.png', 'line.png', 'download.png']
        spawns = []
        icons = []
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            for i in range(len(utilities)):
                if os.path.exists(utilities[i] + '.exe'):
                    spawns.append(utilities[i] + '.exe')
                    icons.append(utilicon[i])
                else:
                    if os.path.exists(utilities[i] + '.py'):
                        spawns.append(utilities[i] + '.py')
                        icons.append(utilicon[i])
        else:
            for i in range(len(utilities)):
                if os.path.exists(utilities[i] + '.py'):
                    spawns.append(utilities[i] + '.py')
                    icons.append(utilicon[i])
        if len(spawns) > 0:
            spawnitem = []
            spawnMenu = menubar.addMenu('&Tools')
            for i in range(len(spawns)):
                if type(spawns[i]) is list:
                    who = spawns[i][0][:spawns[i][0].find('.')]
                else:
                    who = spawns[i][:spawns[i].find('.')]
                spawnitem.append(QtWidgets.QAction(QtGui.QIcon(icons[i]), who, self))
                spawnitem[-1].triggered.connect(partial(self.spawn, spawns[i]))
                spawnMenu.addAction(spawnitem[-1])
            layout.setMenuBar(menubar)
        help = QtWidgets.QAction(QtGui.QIcon('help.png'), 'Help', self)
        help.setShortcut('F1')
        help.setStatusTip('Help')
        help.triggered.connect(self.showHelp)
        about = QtWidgets.QAction(QtGui.QIcon('about.png'), 'About', self)
        about.setShortcut('Ctrl+I')
        about.setStatusTip('About')
        about.triggered.connect(self.showAbout)
        helpMenu = menubar.addMenu('&Help')
        helpMenu.addAction(help)
        helpMenu.addAction(about)
        self.setLayout(layout)
        size = QtCore.QSize(ln, ln2)
        self.resize(size)

    def Clicked(self):
        for i, row in enumerate(self.table.selectionModel().selectedRows()):
            ent = self.table.item(row.row(), 0).text()
            self.invoke(self.siren_dir + ent)
        self.quit()

    def invoke(self, ent):
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            if os.path.exists('sirenm.exe'):
                pid = subprocess.Popen(['sirenm.exe', ent]).pid
            elif os.path.exists('sirenm.py'):
                pid = subprocess.Popen(['sirenm.py', ent], shell=True).pid
        else:
            pid = subprocess.Popen(['python3', 'sirenm.py', ent]).pid # python3

    def new(self):
        do_new = makeNew(self.siren_dir)
        do_new.exec_()
        if do_new.ini_file != '':
            self.invoke(do_new.ini_file)
            self.quit()

    def showAbout(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.about, title='About SENs SAM Model')
        dialog.exec_()

    def showHelp(self):
        webbrowser.open_new(self.help)

    def spawn(self, who):
        if type(who) is list:
            if os.path.exists(who[0]):
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if who[0][-3:] == '.py':
                        pid = subprocess.Popen([who], shell=True).pid
                    else:
                        pid = subprocess.Popen([who]).pid
                else:
                    pid = subprocess.Popen(['python3', who[0], who[1]]).pid
        else:
            if os.path.exists(who):
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if who[-3:] == '.py':
                        pid = subprocess.Popen([who], shell=True).pid
                    else:
                        pid = subprocess.Popen([who]).pid
                else:
                    pid = subprocess.Popen(['python3', who]).pid
        return

    def header_click(self, position):
        column = self.headers.logicalIndexAt(position)
        self.order(column)

    def order(self, col):
        rw = 0
        step = 1
        if col == self.sort_col:
            if self.sort_asc:
                rw = self.table.rowCount() - 1
                step = -1
                self.sort_asc = False
        else:
            self.sort_asc = True
            self.sort_col = col
        for item in sorted(self.entries, key=lambda x: x[col]):
            for cl in range(3):
                self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(item[cl]))
            rw += step

    def quit(self):
        self.close()


class ClickableQLabel(QtWidgets.QLabel):
    clicked = QtCore.pyqtSignal()

    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtWidgets.QApplication.widgetAt(event.globalPos()).setFocus()
        self.clicked.emit()


class makeNew(QtWidgets.QDialog):

    def __init__(self, siren_dir=None, help='help.html'):
        super(makeNew, self).__init__()
        self.siren_dir = siren_dir
        self.help = help
        self.ini_file = ''
        self.initUI()

    def initUI(self):
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            ini_file = 'siren_windows_default.ini'
        else:
            ini_file = 'siren_default.ini'
        if os.path.exists(self.siren_dir + ini_file):
            ini_file = self.siren_dir + ini_file
        else:
            if not os.path.exists(ini_file):
                return
        file_sects = ['[Parents]', '[Files]', '[SAM Modules]']
        dir_props = ['pow_files', 'sam_sdk', 'scenarios', 'solar_files', 'variable_files', 'wind_files']
        field_props = ['scenario_prefix']
        inf = open(ini_file, 'r')
        lines = inf.readlines()
        inf.close()
        sections = {}
        props = []
        for line in lines:
            if line[0] == ';' or line[0] == '#':
                continue
            if line[0] == '[':
                if len(props) > 0:
                    sections[section] = props
                j = line.find(']')
                section = line[:j + 1]
                props = []
            else:
                j = line.find('=')
                if line[j + 1] == '<' or section == '[Parents]':
                    prop = line[j + 1:].lstrip('<').strip().rstrip('>')
                    props.append([line[:j], prop])
        if len(props) > 0:
            sections[section] = props
        row = 0
        self.fields = []
        self.fields.append(['section', 'typ', 'name', 'value', QtWidgets.QLineEdit()])
        self.grid = QtWidgets.QGridLayout()
        self.grid.addWidget(QtWidgets.QLabel('New file name:'), row, 0)
        self.fields[row][4].setText('siren_new.ini')
        self.grid.addWidget(self.fields[row][4], row, 1)
        self.fields[row][4].textChanged.connect(self.filenameChanged)
        self.msg = QtWidgets.QLabel('')
        self.grid.addWidget(self.msg, row, 2, 1, 3)
        now = datetime.datetime.now()
        if '[Base]' in list(sections.keys()):
            for key, value in sections['[Base]']:
                row += 1
                self.fields.append(['[Base]', 'txt', key, value, QtWidgets.QLineEdit()])
                if key == 'year':
                    self.fields[-1][3] = str((now.year - 1))
                if self.fields[row][0] != self.fields[row - 1][0]:
                    self.grid.addWidget(QtWidgets.QLabel(self.fields[row][0]), row, 0)
                self.grid.addWidget(QtWidgets.QLabel(self.fields[row][2]), row, 1)
                self.fields[row][4].setText(self.fields[row][3])
                self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
        self.parents = {}
        sections['[Parents]'].append(['$USER$', getUser()])
        sections['[Parents]'].append(['$YEAR$', str(now.year - 1)])
        if '[Parents]' in list(sections.keys()):
            for key, value in sections['[Parents]']:
                self.parents[key] = value
                row += 1
                self.fields.append(['[Parents]', '?', key, value, '?'])
                if self.fields[row][0] != self.fields[row - 1][0]:
                    self.grid.addWidget(QtWidgets.QLabel(self.fields[row][0]), row, 0)
                self.grid.addWidget(QtWidgets.QLabel(self.fields[row][2]), row, 1)
                if key == '$USER$' or key == '$YEAR$':
                    self.fields[row][1] = 'txt'
                    self.fields[row][4] = QtWidgets.QLabel(self.fields[row][3])
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
                else:
                    self.fields[row][1] = 'dir'
                    self.fields[row][4] = ClickableQLabel()
                    self.fields[row][4].setText(self.fields[row][3])
                    self.fields[row][4].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
                    self.fields[row][4].clicked.connect(self.itemClicked)
        for section, props in iter(sections.items()):
            if section == '[Base]' or section == '[Parents]':
                continue
            elif section == '[Map]':
                for prop in props:
                    row += 1
                    self.fields.append([section, '?', prop[0], prop[1], '?'])
                    if self.fields[row][0] != self.fields[row - 1][0]:
                        self.grid.addWidget(QtWidgets.QLabel(self.fields[row][0]), row, 0)
                    self.grid.addWidget(QtWidgets.QLabel(self.fields[row][2]), row, 1)
                    if prop[0] == 'map' or (prop[0][:3] == 'map' and prop[0][3] != '_'):
                        self.fields[row][3] = prop[1]
                        self.fields[row][1] = 'fil'
                        self.fields[row][4] = ClickableQLabel()
                        self.fields[row][4].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
                        self.fields[row][4].clicked.connect(self.itemClicked)
                    else:
                        self.fields[row][1] = 'txt'
                        self.fields[row][4] = QtWidgets.QLineEdit()
                    self.fields[row][4].setText(self.fields[row][3])
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
            elif section in file_sects:
                for prop in props:
                    row += 1
                    self.fields.append([section, '?', prop[0], prop[1], '?'])
                    if self.fields[row][0] != self.fields[row - 1][0]:
                        self.grid.addWidget(QtWidgets.QLabel(self.fields[row][0]), row, 0)
                    self.grid.addWidget(QtWidgets.QLabel(self.fields[row][2]), row, 1)
                    self.fields[row][3] = prop[1]
                    if prop[0] in field_props:
                        self.fields[row][1] = 'txt'
                        self.fields[row][4] = QtWidgets.QLineEdit()
                        self.fields[row][4].setText(self.fields[row][3])
                    else:
                        if prop[0] in dir_props:
                            self.fields[row][1] = 'dir'
                        else:
                            self.fields[row][1] = 'fil'
                        self.fields[row][4] = ClickableQLabel()
                        self.fields[row][4].setText(self.fields[row][3])
                        self.fields[row][4].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
                        self.fields[row][4].clicked.connect(self.itemClicked)
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
            else:
                for prop in props:
                    row += 1
                    self.fields.append([section, 'txt', prop[0], prop[1], QtWidgets.QLineEdit()])
                    if self.fields[row][0] != self.fields[row - 1][0]:
                        self.grid.addWidget(QtWidgets.QLabel(self.fields[row][0]), row, 0)
                    self.grid.addWidget(QtWidgets.QLabel(self.fields[row][2]), row, 1)
                    self.fields[row][4].setText(self.fields[row][3])
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
        row += 1
        quit = QtWidgets.QPushButton('Quit', self)
        self.grid.addWidget(quit, row, 0)
        quit.clicked.connect(self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        save = QtWidgets.QPushButton('Save', self)
        self.grid.addWidget(save, row, 1)
        save.clicked.connect(self.saveClicked)
        launch = QtWidgets.QPushButton('Save && Open', self)
        self.grid.addWidget(launch, row, 3)
        launch.clicked.connect(self.saveLaunch)
        wdth = save.fontMetrics().boundingRect(launch.text()).width() + 9
        launch.setMaximumWidth(wdth)
        edit = QtWidgets.QPushButton('Save && Edit', self)
        self.grid.addWidget(edit, row, 2)
        edit.clicked.connect(self.saveEdit)
        edit.setMaximumWidth(wdth)
        help = QtWidgets.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        self.grid.addWidget(help, row, 4)
        help.clicked.connect(self.helpClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        self.grid.setColumnStretch(2, 5)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        utilities = ['getmap', 'getmerra2', 'makeweatherfiles', 'sirenupd']
        utilicon = ['map.png', 'download.png', 'weather.png', 'download.png']
        spawns = []
        icons = []
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            for i in range(len(utilities)):
                if os.path.exists(utilities[i] + '.exe'):
                    spawns.append(utilities[i] + '.exe')
                    icons.append(utilicon[i])
                else:
                    if os.path.exists(utilities[i] + '.py'):
                        spawns.append(utilities[i] + '.py')
                        icons.append(utilicon[i])
        else:
            for i in range(len(utilities)):
                if os.path.exists(utilities[i] + '.py'):
                    spawns.append(utilities[i] + '.py')
                    icons.append(utilicon[i])
        if len(spawns) > 0:
            spawnitem = []
            menubar = QtWidgets.QMenuBar()
            spawnMenu = menubar.addMenu('&Tools')
            for i in range(len(spawns)):
                if type(spawns[i]) is list:
                    who = spawns[i][0][:spawns[i][0].find('.')]
                else:
                    who = spawns[i][:spawns[i].find('.')]
                spawnitem.append(QtWidgets.QAction(QtGui.QIcon(icons[i]), who, self))
                spawnitem[-1].triggered.connect(partial(self.spawn, spawns[i]))
                spawnMenu.addAction(spawnitem[-1])
            self.layout.setMenuBar(menubar)
        self.setWindowTitle('SIREN (' + fileVersion() + ') - Create Preferences file')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.show()

    def filenameChanged(self):
        if self.fields[0][4].text().lower() == 'siren_default.ini' or \
          self.fields[0][4].text().lower() == 'siren_default' or \
          self.fields[0][4].text().lower() == 'siren_windows_default.ini' or \
          self.fields[0][4].text().lower() == 'siren_windows_default' or \
          self.fields[0][4].text().lower() == 'getfiles' or \
          self.fields[0][4].text().lower() == 'getfiles.ini':
            self.msg.setText('Proposed file name not allowed.')
        else:
            self.msg.setText('')

    def helpClicked(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.help,
                 title='Help for SIREN Preferences file', section='prefs')
        dialog.exec_()

    def itemClicked(self):
        for i in range(len(self.fields)):
            if self.fields[i][4].hasFocus():
                upd_field = self.fields[i][4].text()
                for key, value in iter(self.parents.items()):
                    upd_field = upd_field.replace(key, value)
                if self.fields[i][1] == 'dir':
                    curdir = upd_field
                    newone = QtWidgets.QFileDialog.getExistingDirectory(self,
                             'Choose ' + self.fields[i][2] + ' Folder', curdir,
                             QtWidgets.QFileDialog.ShowDirsOnly)
                    if newone != '':
                        if self.fields[i][0] == '[Parents]':
                            self.parents[self.fields[i][2]] = newone
                        else:
                            longest = [0, '']
                            for key, value in iter(self.parents.items()):
                                if len(newone) > len(value) and len(value) > longest[0]:
                                    if newone[:len(value)] == value:
                                        longest = [len(value), key]
                            if longest[0] > 0:
                                newone = longest[1] + newone[longest[0]:]
                            if self.parents['$YEAR$'] in newone:
                                newone = newone.replace(self.parents['$YEAR$'], '$YEAR$')
                            if self.parents['$USER$'] in newone:
                                newone = newone.replace(self.parents['$USER$'], '$USER$')
                        self.fields[i][4].setText(newone)
                elif self.fields[i][1] == 'fil':
                    curfil = upd_field
                    newone = QtWidgets.QFileDialog.getOpenFileName(self,
                             'Choose ' + self.fields[i][2] + ' File', curfil)[0]
                    if newone != '':
                        newone = QtCore.QDir.toNativeSeparators(newone)
                        longest = [0, '']
                        for key, value in iter(self.parents.items()):
                            if len(newone) > len(value) and len(value) > longest[0]:
                                if newone[:len(value)] == value:
                                    longest = [len(value), key]
                        if longest[0] > 0:
                            newone = longest[1] + newone[longest[0]:]
                        if self.parents['$YEAR$'] in newone:
                            newone = newone.replace(self.parents['$YEAR$'], '$YEAR$')
                        if self.parents['$USER$'] in newone:
                            newone = newone.replace(self.parents['$USER$'], '$USER$')
                        self.fields[i][4].setText(newone)
                break

    def quitClicked(self):
        self.close()

    def saveClicked(self):
        if self.saveIni() >= 0:
            self.close()
        else:
            QtWidgets.QMessageBox.about(self, 'SIREN - Error', self.msg.text())

    def saveEdit(self):
        if self.saveIni() >= 0:
            dialr = EdtDialog(self.new_ini)
            dialr.exec_()
            self.close()
        else:
            QtWidgets.QMessageBox.about(self, 'SIREN - Error', self.msg.text())

    def saveLaunch(self):
        if self.saveIni() >= 0:
            self.ini_file = self.new_ini
            self.close()
        else:
            QtWidgets.QMessageBox.about(self, 'SIREN - Error', self.msg.text())

    def saveIni(self):
        updates = {}
        lines = []
        newfile = self.fields[0][4].text()
        if newfile == 'siren_default.ini' or newfile == 'siren_default' or \
          newfile == 'siren_windows_default.ini' or newfile == 'siren_windows_default':
            self.msg.setText('Proposed file name not allowed.')
            return -1
        if newfile[-4:].lower() != '.ini':
            newfile = newfile + '.ini'
        newfile = self.siren_dir + newfile
        self.new_ini = newfile
        if os.path.exists(newfile):
            if os.path.exists(newfile + '~'):
                os.remove(newfile + '~')
            os.rename(newfile, newfile + '~')
        my_dir = adj_dir = os.getcwd()
        if self.parents['$YEAR$'] in adj_dir:
            adj_dir = adj_dir.replace(self.parents['$YEAR$'], '$YEAR$')
        if self.parents['$USER$'] in my_dir:
            adj_dir = adj_dir.replace(self.parents['$USER$'], '$USER$')
        for field in self.fields:
            if field[0] == 'section':
                continue
            if field[0] == '[Parents]' and (field[2] == '$USER$' or field[2] == '$YEAR$'):
                continue
            if field[0] not in list(updates.keys()):
                updates[field[0]] = []
            fld = field[4].text()
            if self.parents['$YEAR$'] in fld:
                if field[0] == '[Base]' and (field[2] == 'year' or field[2] == 'years'):
                    pass
                else:
                    fld = fld.replace(self.parents['$YEAR$'], '$YEAR$')
            if self.parents['$USER$'] in fld:
                fld = fld.replace(self.parents['$USER$'], '$USER$')
            if len(fld) > len(adj_dir):
                if fld[:len(adj_dir)] == adj_dir:
                    if field[1] == 'dir' and fld[len(adj_dir)] == '/':
                        fld = fld[len(adj_dir) + 1:]
                    elif field[1] == 'fil':
                        if fld.find('/') >= 0:
                            that_len = len(commonprefix([adj_dir, fld]))
                            if that_len > 0:
                                bits = adj_dir[that_len:].split('/')
                                pfx = ('..' + '/') * (len(bits) - 1)
                                fld = pfx + fld[that_len + 1:]
            updates[field[0]].append(field[2] + '=' + fld)
        if '[Parents]' in list(updates.keys()):
            my_dir = os.getcwd()
            my_dir = my_dir.replace(self.parents['$YEAR$'], '$YEAR$')
            my_dir = my_dir.replace(self.parents['$USER$'], '$USER$')
            for p in range(len(updates['[Parents]'])):
                i = updates['[Parents]'][p].find('=')
                value = updates['[Parents]'][p][i + 1:]
                if value.find('/') >= 0:
                    that_len = len(commonprefix([my_dir, value]))
                    if that_len > 0:
                        bits = my_dir[that_len:].split('/')
                        pfx = ('..' + '/') * (len(bits) - 1)
                        updates['[Parents]'][p] = updates['[Parents]'][p][:i + 1] + pfx + value[that_len + 1:]
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            ini_file = 'siren_windows_default.ini'
        else:
            ini_file = 'siren_default.ini'
        if os.path.exists(self.siren_dir + ini_file):
            ini_file = self.siren_dir + ini_file
        else:
            if not os.path.exists(ini_file):
                return -1
        inf = open(ini_file, 'r')
        lines = inf.readlines()
        inf.close()
        for section in updates:
            in_section = False
            properties = updates[section]
            props = []
            for i in range(len(properties)):
                props.append(properties[i].split('=')[0])
            for i in range(len(lines)):
                if lines[i][:len(section)] == section:
                    in_section = True
                elif in_section:
                    if lines[i][0] == '[':
                        i -= 1
                        break
                    elif lines[i][0] != ';' and lines[i][0] != '#':
                        bits = lines[i].split('=')
                        for j in range(len(properties) - 1, -1, -1):
                            if bits[0] == props[j]:
                                lines[i] = properties[j] + '\n'
                                del properties[j]
                                del props[j]
            if len(properties) > 0:
                if not in_section:
                    lines.append(section + '\n')
                    i += 1
                for j in range(len(properties)):
                    lines.insert(i + 1, properties[j] + '\n')
                    i += 1
        sou = open(newfile, 'w')
        for i in range(len(lines)):
            sou.write(lines[i])
        sou.close()
        return 0

    def spawn(self, who):
        if type(who) is list:
            if os.path.exists(who[0]):
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if who[0][-3:] == '.py':
                        pid = subprocess.Popen([who], shell=True).pid
                    else:
                        pid = subprocess.Popen([who]).pid
                else:
                    pid = subprocess.Popen(['python3', who[0], who[1]]).pid
        else:
            if os.path.exists(who):
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if who[-3:] == '.py':
                        pid = subprocess.Popen([who], shell=True).pid
                    else:
                        pid = subprocess.Popen([who]).pid
                else:
                    pid = subprocess.Popen(['python3', who]).pid
        return


if '__main__' == __name__:
    app = QtWidgets.QApplication(sys.argv)
    tabdialog = TabDialog()
    tabdialog.show()
    sys.exit(app.exec_())
