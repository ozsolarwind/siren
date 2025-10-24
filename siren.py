#!/usr/bin/python3
#
#  Copyright (C) 2016-2025 Sustainable Energy Now Inc., Angus King
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
import shutil
import subprocess
import sys
import time
import webbrowser

import displayobject
from credits import fileVersion
from editini import EdtDialog, EditFileSections
from getmodels import getModelFile, commonprefix
from senutils import ClickableQLabel, getUser


class TabDialog(QtWidgets.QDialog):

    def check_file(self, mdir, fil, errors=''):
        ok = True
        try:
            self.config.read(mdir + fil)
        except configparser.DuplicateOptionError as err:
            errors += 'DuplicateOptionError ' + str(err) + '\n'
            i = str(err).find('[line')
            if i >= 0:
                j = str(err).find(']', i)
                txt = ' ' + str(err)[i:j + 1]
            else:
                txt = ''
            model_name = 'Duplicate Option Error' + txt
            ok = False
        except configparser.DuplicateSectionError as err:
            errors += 'DuplicateSectionError ' + str(err) + '\n'
            i = str(err).find('[line')
            if i >= 0:
                j = str(err).find(']', i)
                txt = ' ' + str(err)[i:j + 1]
            else:
                txt = ''
            model_name = 'Duplicate Section Error' + txt
            ok = False
        except:
            err = sys.exc_info()[0]
            errors += 'Error: ' + str(err) + ' While reading from ' + mdir + fil + '\n'
            model_name = 'Error reading file'
            ok = False
        if ok:
            try:
                model_name = self.config.get('Base', 'name')
            except:
                model_name = ''
        return ok, model_name, errors

    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent)
        models_dir = '.'
        self.models_dirs = []
        if len(sys.argv) > 1:
            if sys.argv[1][-4:] == '.ini':
                self.invoke('powermap', sys.argv[1])
                sys.exit()
            elif os.path.isdir(sys.argv[1]):
                models_dir = sys.argv[1]
            else:
                ini_dir = sys.argv[1].replace('$USER$', getUser())
                if os.path.isdir(ini_dir):
                    models_dir = ini_dir
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                if models_dir[-1] != '\\' and models_dir[-1] != '/':
                    models_dir += '\\'
            elif models_dir[-1] != '/':
                models_dir += '/'
        else:
            self.models_dirs = getModelFile()
        if len(self.models_dirs) == 0:
            self.models_dirs = [models_dir]
        elif len(self.models_dirs) > 1:
            dups = []
            for i in range(len(self.models_dirs)):
                for j in range(i + 1, len(self.models_dirs)):
                    if os.path.samefile(self.models_dirs[i], self.models_dirs[j]):
                        if j not in dups:
                            dups.append(j)
            dups.sort(reverse=True)
            for dup in dups:
                del self.models_dirs[dup]
        self.entries = []
        for models_dir in self.models_dirs:
            fils = os.listdir(models_dir)
            self.help = 'help.html'
            self.about = 'about.html'
            self.config = configparser.RawConfigParser()
            ignore = ['siren_default.ini']
            errors = ''
            for fil in sorted(fils):
                if fil[-4:] == '.ini':
                    if fil in ignore:
                        continue
                    mod_time = time.strftime('%Y-%m-%d %H:%M:%S',
                               time.localtime(os.path.getmtime(models_dir + fil)))
                    ok, model_name, errors = self.check_file(models_dir, fil, errors)
                    if len(self.models_dirs) > 1:
                        self.entries.append([fil, model_name, mod_time, models_dir, ok])
                    else:
                        self.entries.append([fil, model_name, mod_time, ok])
        if len(errors) > 0:
            dialog = displayobject.AnObject(QtWidgets.QDialog(), errors,
                     title='SIREN (' + fileVersion() + ') - Preferences file errors')
            dialog.exec_()
        if len(self.entries) == 0:
            self.new_model()
        self.setWindowTitle('SIREN (' + fileVersion() + ') - Select SIREN Model')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        buttonLayout = QtWidgets.QHBoxLayout()
        self.quitButton = QtWidgets.QPushButton(self.tr('&Quit'))
        buttonLayout.addWidget(self.quitButton)
        self.quitButton.clicked.connect(self.quit)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
        self.newButton = QtWidgets.QPushButton(self.tr('&New Model'))
        buttonLayout.addWidget(self.newButton)
        self.newButton.clicked.connect(self.new_model)
        buttons = QtWidgets.QFrame()
        buttons.setLayout(buttonLayout)
        layout = QtWidgets.QGridLayout()
        self.table = QtWidgets.QTableWidget()
        self.table.setRowCount(len(self.entries))
        hdr_labels = ['Preference File', 'SIREN Model', 'Date modified']
        if len(self.models_dirs) > 1:
            hdr_labels.append('Folder')
        self.table.setColumnCount(len(hdr_labels))
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
            for cl in range(len(hdr_labels)):
                self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(self.entries[rw][cl]))
                ln += len(self.entries[rw][cl])
            if ln > max_row:
                max_row = ln
        self.sort_desc = False # start in date descending order
        self.sort_col = 2
        self.order(2)
        self.table.resizeColumnsToContents()
        self.table.viewport().installEventFilter(self)
        fnt = self.table.fontMetrics()
        ln = max_row * max(9, fnt.averageCharWidth())
        ln2 = (len(self.entries) + 8) * (fnt.xHeight() + fnt.lineSpacing())
        screen = QtWidgets.QDesktopWidget().screenGeometry()
        if ln > screen.width() * .9:
            ln = int(screen.width() * .9)
        if ln2 > screen.height() * .9:
            ln2 = int(screen.height() * .9)
        layout.addWidget(QtWidgets.QLabel('Left click on row for Desired Model or right click for Tools; Right click column header to sort'), 0, 0)
        layout.addWidget(self.table, 1, 0)
        layout.addWidget(buttons, 2, 0)
        menubar = QtWidgets.QMenuBar()
        self.utilities = ['flexiplot', 'getera5', 'getmap', 'getmerra2', 'makeweatherfiles', 'powerplot', 'sirenupd']
        self.utilicon = ['line.png', 'download.png', 'map.png', 'download.png', 'weather.png', 'line.png', 'download.png']
        spawns = []
        icons = []
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            for i in range(len(self.utilities)):
                if os.path.exists(self.utilities[i] + '.exe'):
                    spawns.append(self.utilities[i] + '.exe')
                    icons.append(self.utilicon[i])
                else:
                    if os.path.exists(self.utilities[i] + '.py'):
                        spawns.append(self.utilities[i] + '.py')
                        icons.append(self.utilicon[i])
        else:
            for i in range(len(self.utilities)):
                if os.path.exists(self.utilities[i] + '.py'):
                    spawns.append(self.utilities[i] + '.py')
                    icons.append(self.utilicon[i])
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
        self.model_tool = ['flexiplot', 'getmap', 'indexweather', 'makegrid', 'powermap',
                           'powermatch', 'powerplot', 'updateswis']
        self.model_icon = ['line.png', 'map.png', 'list.png', 'grid.png', 'sen_icon32.png',
                           'power.png', 'line.png', 'list.png']
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

    def eventFilter(self, source, event):
        if self.table.selectedIndexes() != []:
            if event.type() == QtCore.QEvent.MouseButtonRelease and \
              event.button() == QtCore.Qt.LeftButton:
                ent = self.table.item(self.table.currentRow(), 0).text()
                if ent in ['getfiles.ini', 'flexiplot.ini', 'powerplot.ini']:
                    return QtCore.QObject.event(source, event)
                self.table.viewport().removeEventFilter(self)
                if len(self.models_dirs) > 1:
                    ent_dir = self.table.item(self.table.currentRow(), 3).text()
                else:
                    ent_dir = self.models_dirs[0]
                self.invoke('powermap', ent_dir + ent)
                self.quit()
            if (event.type() == QtCore.QEvent.MouseButtonPress or event.type() == QtCore.QEvent.MouseButtonRelease) and \
              event.button() == QtCore.Qt.RightButton:
                ent = self.table.item(self.table.currentRow(), 0).text()
                index = self.table.indexAt(event.pos())
                selectionModel = self.table.selectionModel()
                selectionModel.select(self.table.model().index(self.table.currentRow(), 0),
                    selectionModel.Deselect|selectionModel.Rows)
                selectionModel.select(self.table.model().index(index.row(), 0),
                    selectionModel.Rows)
                menu = QtWidgets.QMenu()
                actions =  []
                if ent in ['flexiplot.ini', 'powerplot.ini']:
                    tool = ent[:ent.find('.')]
                    try:
                        i = self.model_tool.index(tool)
                        actions.append(menu.addAction(QtGui.QIcon(self.model_icon[i]),
                                                     'Execute ' + self.model_tool[i]))
                        actions[-1].setIconVisibleInMenu(True)
                    except:
                        pass
                    actions.append(menu.addAction(QtGui.QIcon('edit.png'), 'Edit Preferences'))
                    actions[-1].setIconVisibleInMenu(True)
                elif ent == 'getfiles.ini':
                    for i in range(len(self.utilities)):
                        if self.utilities[i] in ['flexiplot', 'powerplot']:
                            continue
                        actions.append(menu.addAction(QtGui.QIcon(self.utilicon[i]),
                                                     'Execute ' + self.utilities[i]))
                        actions[-1].setIconVisibleInMenu(True)
                    actions.append(menu.addAction(QtGui.QIcon('edit.png'), 'Edit Preferences'))
                    actions[-1].setIconVisibleInMenu(True)
                else:
                    for i in range(len(self.model_tool)):
                        if self.model_tool[i] == 'updateswis':
                            mdl = self.table.item(self.table.currentRow(), 1).text()
                            if mdl.find('SWIS') < 0:
                                continue
                        actions.append(menu.addAction(QtGui.QIcon(self.model_icon[i]),
                                                     'Execute ' + self.model_tool[i]))
                        actions[-1].setIconVisibleInMenu(True)
                    actions.append(menu.addAction(QtGui.QIcon('edit.png'), 'Edit Preferences'))
                    actions[-1].setIconVisibleInMenu(True)
                    actions.append(menu.addAction(QtGui.QIcon('edit.png'), 'Edit File Preferences'))
                    actions[-1].setIconVisibleInMenu(True)
                    actions.append(menu.addAction(QtGui.QIcon('copy.png'), 'Copy Preferences'))
                    actions[-1].setIconVisibleInMenu(True)
                    actions.append(menu.addAction(QtGui.QIcon('delete.png'), 'Delete Preferences'))
                    actions[-1].setIconVisibleInMenu(True)
                action = menu.exec_(self.mapToGlobal(event.pos()))
                if action is not None:
                    if len(self.models_dirs) > 1:
                        ent_dir = self.table.item(self.table.currentRow(), 3).text()
                    else:
                        ent_dir = self.models_dirs[0]
                    if action.text()[:8] == 'Execute ':
                        if not self.entries[self.table.currentRow()][-3]:
                            ok, model_name, errors = self.check_file(ent_dir, ent)
                            if len(errors) > 0:
                                dialog = displayobject.AnObject(QtWidgets.QDialog(), errors,
                                title='SIREN (' + fileVersion() + ') - Preferences file errors')
                                dialog.exec_()
                                return QtCore.QObject.event(source, event)
                        self.invoke(action.text()[8:], ent_dir + ent)
                    elif action.text()[-11:] == 'Preferences':
                        i = self.table.item(self.table.currentRow(), 1).text().find('[line ')
                        if i >= 0:
                            j = self.table.item(self.table.currentRow(), 1).text().find(']', i)
                            line = int(self.table.item(self.table.currentRow(), 1).text()[i + 5:j].strip()) - 1
                        else:
                            line = None
                        if action.text()[-16:] == 'File Preferences':
                            self.editIniFileSects(ent_dir + ent)
                        elif action.text()[:4] == 'Copy':
                            newfile = QtWidgets.QFileDialog.getSaveFileName(None, 'Copy Preferences file',
                                      ent_dir + ent, 'Preference files (*.ini)')[0]
                            if newfile != '':
                                if newfile.find('/') >= 0:
                                    my_dir = os.getcwd()
                                    if sys.platform == 'win32' or sys.platform == 'cygwin':
                                        my_dir = my_dir.replace('\\', '/')
                                    that_len = len(commonprefix([my_dir, newfile]))
                                    if that_len > 0:
                                        new_ent = newfile[that_len + 1:]
                                    else:
                                        new_ent = newfile
                                try:
                                    shutil.copy2(ent_dir + ent, newfile)
                                except:
                                    return QtCore.QObject.event(source, event)
                                i = new_ent.rfind('/')
                                if i >= 0:
                                    ent_dir = new_ent[:i + 1]
                                    ent = new_ent[i + 1:]
                                else:
                                    ent_dir = ''
                                    ent = new_ent
                                mod_time = time.strftime('%Y-%m-%d %H:%M:%S',
                                           time.localtime(os.path.getmtime(newfile)))
                                rw = self.table.currentRow() + 1
                                self.entries.insert(rw, [])
                                self.entries[rw].append(ent)
                                self.entries[rw].append(self.entries[self.table.currentRow()][1])
                                self.entries[rw].append(mod_time)
                                if self.table.columnCount() == 4:
                                    self.entries[rw].append(ent_dir)
                                self.entries[rw].append(True)
                                self.table.insertRow(rw)
                                for cl in range(self.table.columnCount()):
                                    self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(self.entries[rw][cl]))
                        elif action.text()[:6] == 'Delete':
                            reply = QtWidgets.QMessageBox.question(self, 'SIREN - Delete Preferences',
                                    "Is '" + ent_dir + ent + "' the one to delete?",
                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                    QtWidgets.QMessageBox.No)
                            if reply == QtWidgets.QMessageBox.Yes:
                                os.remove(ent_dir + ent)
                                del self.entries[self.table.currentRow()]
                                self.table.removeRow(self.table.currentRow())
                                return QtCore.QObject.event(source, event)
                        else:
                            mod_b4 = self.entries[self.table.currentRow()][2]
                            self.editIniFile(ent_dir + ent, line=line)
                            mod_time = time.strftime('%Y-%m-%d %H:%M:%S',
                               time.localtime(os.path.getmtime(ent_dir + ent)))
                            if mod_time != mod_b4:
                                self.entries[self.table.currentRow()][2] = mod_time
                                self.table.setItem(self.table.currentRow(), 2, QtWidgets.QTableWidgetItem(mod_time))
                        ok, model_name, errors = self.check_file(ent_dir, ent)
                        if model_name != self.entries[self.table.currentRow()][1]:
                            self.entries[self.table.currentRow()][1] = model_name
                            self.table.setItem(self.table.currentRow(), 1, QtWidgets.QTableWidgetItem(model_name))
                        if len(errors) > 0:
                            self.entries[self.table.currentRow()][-1] = False
                            dialog = displayobject.AnObject(QtWidgets.QDialog(), errors,
                            title='SIREN (' + fileVersion() + ') - Preferences file errors')
                            dialog.exec_()
                            return QtCore.QObject.event(source, event)
                        else:
                            self.entries[self.table.currentRow()][-1] = True
        return QtCore.QObject.event(source, event)

    def invoke(self, program, ent):
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            if os.path.exists(program + '.exe'):
                pid = subprocess.Popen([program + '.exe', ent]).pid
            elif os.path.exists(program + '.py'):
                pid = subprocess.Popen([program + '.py', ent], shell=True).pid
        else:
            pid = subprocess.Popen(['python3', program + '.py', ent]).pid # python3
        return

    def new_model(self):
        do_new = makeNew(self.models_dirs)
        do_new.exec_()
        if do_new.ini_file != '':
            if do_new.launch:
                self.invoke('powermap', do_new.ini_file)
                return
              #  self.quit()
            mod_time = time.strftime('%Y-%m-%d %H:%M:%S',
                       time.localtime(os.path.getmtime(do_new.ini_file)))
            i = do_new.ini_file.rfind('/')
            if i >= 0:
                models_dir = do_new.ini_file[:i + 1]
                fil = do_new.ini_file[i + 1:]
            else:
                model_dir = '/'
                fil = do_new.ini_file
            ok, model_name, errors = self.check_file(models_dir, fil)
            if len(self.models_dirs) > 1:
                self.entries.append([fil, model_name, mod_time, models_dir, ok])
            else:
                self.entries.append([fil, model_name, mod_time, ok])
            self.sort_desc = not self.sort_desc
            self.order(self.sort_col)

    def editIniFile(self, ini=None, line=None):
        dialr = EdtDialog(ini, line=line)
        dialr.exec_()
        return

    def editIniFileSects(self, ini=None):
        dialr = EditFileSections(ini)
        dialr.exec_()
        return

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
                        pid = subprocess.Popen(who, shell=True).pid
                    else:
                        pid = subprocess.Popen(who).pid
                else:
                    who.insert(0, 'python3')
                    pid = subprocess.Popen(who).pid
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
        if col == self.sort_col:
            if self.sort_desc:
                self.sort_desc = False
            else:
                self.sort_desc = True
        else:
            self.sort_desc = False
            self.sort_col = col
        self.entries = sorted(self.entries, key=lambda x: x[col], reverse=self.sort_desc)
        for item in self.entries:
            for cl in range(len(self.entries[0]) - 1):
                self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(item[cl]))
            rw += 1

    def quit(self):
        self.close()


class makeNew(QtWidgets.QDialog):

    def __init__(self, models_dirs=None, help='help.html'):
        super(makeNew, self).__init__()
        self.models_dirs = models_dirs
        self.help = help
        self.ini_file = ''
        self.launch = False
        ini_file = 'siren_default.ini'
        for models_dir in self.models_dirs:
            if os.path.exists(models_dir + ini_file):
                ini_file = models_dir + ini_file
                break
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
        row = -1
        self.fields = []
        self.fields.append(['section', 'typ', 'name', 'value', QtWidgets.QLineEdit()])
        self.grid = QtWidgets.QGridLayout()
        self.grid.addWidget(QtWidgets.QLabel('Model folder:'), row + 1, 0)
        self.fldrcombo = QtWidgets.QComboBox(self)
        for models_dir in self.models_dirs:
            self.fldrcombo.addItem(models_dir)
        self.fldrcombo.setCurrentIndex(0)
        self.grid.addWidget(self.fldrcombo, row + 1, 1, 1, 2)
        row += 1
        self.grid.addWidget(QtWidgets.QLabel('New file name:'), row + 1, 0)
        self.fields[row][4].setText('siren_new.ini')
        self.grid.addWidget(self.fields[row][4], row + 1, 1)
        self.fields[row][4].textChanged.connect(self.filenameChanged)
        self.msg = QtWidgets.QLabel('')
        self.grid.addWidget(self.msg, row + 1, 2, 1, 3)
        now = datetime.datetime.now()
        if '[Base]' in list(sections.keys()):
            for key, value in sections['[Base]']:
                row += 1
                self.fields.append(['[Base]', 'txt', key, value, QtWidgets.QLineEdit()])
                if key == 'year':
                    self.fields[-1][3] = str((now.year - 1))
                if self.fields[row][0] != self.fields[row - 1][0]:
                    self.grid.addWidget(QtWidgets.QLabel(self.fields[row][0]), row + 1, 0)
                self.grid.addWidget(QtWidgets.QLabel(self.fields[row][2]), row + 1, 1)
                self.fields[row][4].setText(self.fields[row][3])
                self.grid.addWidget(self.fields[row][4], row + 1, 2, 1, 3)
        self.parents = {}
        sections['[Parents]'].append(['$USER$', getUser()])
        sections['[Parents]'].append(['$YEAR$', str(now.year - 1)])
        if '[Parents]' in list(sections.keys()):
            for key, value in sections['[Parents]']:
                self.parents[key] = value
                row += 1
                self.fields.append(['[Parents]', '?', key, value, '?'])
                if self.fields[row][0] != self.fields[row - 1][0]:
                    self.grid.addWidget(QtWidgets.QLabel(self.fields[row][0]), row + 1, 0)
                self.grid.addWidget(QtWidgets.QLabel(self.fields[row][2]), row + 1, 1)
                if key == '$USER$' or key == '$YEAR$':
                    self.fields[row][1] = 'txt'
                    self.fields[row][4] = QtWidgets.QLabel(self.fields[row][3])
                    self.grid.addWidget(self.fields[row][4], row + 1, 2, 1, 3)
                else:
                    self.fields[row][1] = 'dir'
                    self.fields[row][4] = ClickableQLabel()
                    self.fields[row][4].setText(self.fields[row][3])
                    self.fields[row][4].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
                    self.grid.addWidget(self.fields[row][4], row + 1, 2, 1, 3)
                    self.fields[row][4].clicked.connect(self.itemClicked)
        for section, props in sections.items():
            if section == '[Base]' or section == '[Parents]':
                continue
            elif section == '[Map]':
                for prop in props:
                    row += 1
                    self.fields.append([section, '?', prop[0], prop[1], '?'])
                    if self.fields[row][0] != self.fields[row - 1][0]:
                        self.grid.addWidget(QtWidgets.QLabel(self.fields[row][0]), row + 1, 0)
                    self.grid.addWidget(QtWidgets.QLabel(self.fields[row][2]), row + 1, 1)
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
                    self.grid.addWidget(self.fields[row][4], row + 1, 2, 1, 3)
            elif section in file_sects:
                for prop in props:
                    row += 1
                    self.fields.append([section, '?', prop[0], prop[1], '?'])
                    if self.fields[row][0] != self.fields[row - 1][0]:
                        self.grid.addWidget(QtWidgets.QLabel(self.fields[row][0]), row + 1, 0)
                    self.grid.addWidget(QtWidgets.QLabel(self.fields[row][2]), row + 1, 1)
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
                    self.grid.addWidget(self.fields[row][4], row + 1, 2, 1, 3)
            else:
                for prop in props:
                    row += 1
                    self.fields.append([section, 'txt', prop[0], prop[1], QtWidgets.QLineEdit()])
                    if self.fields[row][0] != self.fields[row - 1][0]:
                        self.grid.addWidget(QtWidgets.QLabel(self.fields[row][0]), row + 1, 0)
                    self.grid.addWidget(QtWidgets.QLabel(self.fields[row][2]), row + 1, 1)
                    self.fields[row][4].setText(self.fields[row][3])
                    self.grid.addWidget(self.fields[row][4], row + 1, 2, 1, 3)
        row += 1
        quit = QtWidgets.QPushButton('Quit', self)
        self.grid.addWidget(quit, row + 1, 0)
        quit.clicked.connect(self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        save = QtWidgets.QPushButton('Save', self)
        self.grid.addWidget(save, row + 1, 1)
        save.clicked.connect(self.saveClicked)
        launch = QtWidgets.QPushButton('Save && Open', self)
        self.grid.addWidget(launch, row + 1, 3)
        launch.clicked.connect(self.saveLaunch)
        wdth = save.fontMetrics().boundingRect(launch.text()).width() + 9
        launch.setMaximumWidth(wdth)
        edit = QtWidgets.QPushButton('Save && Edit', self)
        self.grid.addWidget(edit, row + 1, 2)
        edit.clicked.connect(self.saveEdit)
        edit.setMaximumWidth(wdth)
        help = QtWidgets.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        self.grid.addWidget(help, row + 1, 4)
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
        self.utilities = ['getera5', 'getmap', 'getmerra2', 'makeweatherfiles', 'sirenupd']
        self.utilicon = ['download.png', 'map.png', 'download.png', 'weather.png', 'download.png']
        spawns = []
        icons = []
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            for i in range(len(self.utilities)):
                if os.path.exists(self.utilities[i] + '.exe'):
                    spawns.append(self.utilities[i] + '.exe')
                    icons.append(self.utilicon[i])
                else:
                    if os.path.exists(self.utilities[i] + '.py'):
                        spawns.append(self.utilities[i] + '.py')
                        icons.append(self.utilicon[i])
        else:
            for i in range(len(self.utilities)):
                if os.path.exists(self.utilities[i] + '.py'):
                    spawns.append(self.utilities[i] + '.py')
                    icons.append(self.utilicon[i])
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
        if self.fields[0][4].text().lower() in ['getfiles', 'getfiles.ini',
                                                'siren_default', 'siren_default.ini']:
            # and maybe flexiplot and powerplot
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
                for key, value in self.parents.items():
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
                            for key, value in self.parents.items():
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
                        for key, value in self.parents.items():
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
            self.ini_file = self.new_ini
            self.close()
        else:
            QtWidgets.QMessageBox.about(self, 'SIREN - Error', self.msg.text())

    def saveEdit(self):
        if self.saveIni() >= 0:
            dialr = EdtDialog(self.new_ini)
            dialr.exec_()
            self.ini_file = self.new_ini
            self.close()
        else:
            QtWidgets.QMessageBox.about(self, 'SIREN - Error', self.msg.text())

    def saveLaunch(self):
        if self.saveIni() >= 0:
            self.ini_file = self.new_ini
            self.launch = True
            self.close()
        else:
            QtWidgets.QMessageBox.about(self, 'SIREN - Error', self.msg.text())

    def saveIni(self):
        updates = {}
        lines = []
        newfile = self.fields[0][4].text()
        if newfile == 'siren_default.ini' or newfile == 'siren_default':
            self.msg.setText('Proposed file name not allowed.')
            return -1
        if newfile[-4:].lower() != '.ini':
            newfile = newfile + '.ini'
        newfile = self.fldrcombo.currentText() + newfile
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
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            adj_dir = adj_dir.replace('\\', '/')
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
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                my_dir = my_dir.replace('\\', '/')
            for p in range(len(updates['[Parents]'])):
                i = updates['[Parents]'][p].find('=')
                value = updates['[Parents]'][p][i + 1:]
                if value.find('/') >= 0:
                    that_len = len(commonprefix([my_dir, value]))
                    if that_len > 0:
                        bits = my_dir[that_len:].split('/')
                        pfx = ('..' + '/') * (len(bits) - 1)
                        updates['[Parents]'][p] = updates['[Parents]'][p][:i + 1] + pfx + value[that_len + 1:]
        ini_file = 'siren_default.ini'
        if os.path.exists(self.fldrcombo.currentText() + ini_file):
            ini_file = self.fldrcombo.currentText() + ini_file
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
    try:
        QtGui.QGuiApplication.setDesktopFileName('siren')
    except:
        pass
    tabdialog = TabDialog()
    tabdialog.show()
    sys.exit(app.exec_())
