#!/usr/bin/python
#
#  Copyright (C) 2016 Sustainable Energy Now Inc., Angus King
#
#  sirens.py - This file is part of SIREN.
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

import ConfigParser   # decode .ini file
import datetime
from functools import partial
import os
from PyQt4 import QtCore, QtGui
from PyQt4.QtCore import SIGNAL
import subprocess
import sys
import webbrowser

from credits import fileVersion
import displayobject
from editini import EdtDialog
from senuser import getUser

def commonprefix(args, sep='/'):
    arg2 = []
    for arg in args:
        arg2.append(arg)
        if arg[-1] != sep:
            arg2[-1] += sep
    return os.path.commonprefix(arg2).rpartition(sep)[0]

class TabDialog(QtGui.QDialog):
    def __init__(self, parent=None):
        QtGui.QDialog.__init__(self, parent)
        if len(sys.argv) > 1:
            if sys.argv[1][-4:] == '.ini':
                self.invoke(sys.argv[1])
                sys.exit()
        entries = []
        fils = sorted(os.listdir('.'))
        self.help = ''
        self.about = ''
        config = ConfigParser.RawConfigParser()
        for fil in fils:
            if fil[-4:] == '.ini':
                if fil == 'siren_default.ini' or fil == 'siren_windows_default.ini':
                    continue
                try:
                    config.read(fil)
                except:
                    continue
                try:
                    model_name = config.get('Base', 'name')
                except:
                    model_name = ''
                entries.append([fil, model_name])
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
        if len(entries) == 0:
            self.new()
     #    if len(entries) == 1:
     #        self.invoke(entries[0][0])
     #        sys.exit()
        self.setWindowTitle('Select SIREN Model')
        buttonLayout = QtGui.QHBoxLayout()
        self.quitButton = QtGui.QPushButton(self.tr('&Quit'))
        buttonLayout.addWidget(self.quitButton)
        self.connect(self.quitButton, QtCore.SIGNAL('clicked()'), self.quit)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
        self.newButton = QtGui.QPushButton(self.tr('&New Model'))
        buttonLayout.addWidget(self.newButton)
        self.connect(self.newButton, QtCore.SIGNAL('clicked()'), self.new)
        buttons = QtGui.QFrame()
        buttons.setLayout(buttonLayout)
        layout = QtGui.QGridLayout()
        self.table = QtGui.QTableWidget()
        self.table.setRowCount(len(entries))
        self.table.setColumnCount(2)
        hdr_labels = ['Preference File', 'SIREN Model']
        self.table.setHorizontalHeaderLabels(hdr_labels)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
        max_row = 30
        for rw in range(len(entries)):
            ln = 0
            for cl in range(2):
                self.table.setItem(rw, cl, QtGui.QTableWidgetItem(entries[rw][cl]))
                ln += len(entries[rw][cl])
            if ln > max_row:
                max_row = ln
        self.table.resizeColumnsToContents()
        self.table.itemClicked.connect(self.Clicked)
        fnt = self.table.fontMetrics()
        ln = max_row * max(9, fnt.averageCharWidth())
        ln2 = (len(entries) + 8) * (fnt.xHeight() + fnt.lineSpacing())
        screen = QtGui.QDesktopWidget().screenGeometry()
        if ln > screen.width() * .9:
            ln = int(screen.width() * .9)
        if ln2 > screen.height() * .9:
            ln2 = int(screen.height() * .9)
        layout.addWidget(QtGui.QLabel('Click on row for Desired Model'), 0, 0)
        layout.addWidget(self.table, 1, 0)
        layout.addWidget(buttons, 2, 0)
        menubar = QtGui.QMenuBar()
        utilities = ['getmap', 'makeweather2']
        utilicon = ['map.png', 'weather.png']
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
                spawnitem.append(QtGui.QAction(QtGui.QIcon(icons[i]), who, self))
                spawnitem[-1].triggered.connect(partial(self.spawn, spawns[i]))
                spawnMenu.addAction(spawnitem[-1])
            layout.setMenuBar(menubar)
        help = QtGui.QAction(QtGui.QIcon('help.png'), 'Help', self)
        help.setShortcut('F1')
        help.setStatusTip('Help')
        help.triggered.connect(self.showHelp)
        about = QtGui.QAction(QtGui.QIcon('about.png'), 'About', self)
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
            ent = str(self.table.item(row.row(), 0).text())
            self.invoke(ent)
        self.quit()

    def invoke(self, ent):
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            if os.path.exists('sirenm.exe'):
                pid = subprocess.Popen(['sirenm.exe', ent]).pid
            elif os.path.exists('sirenm.py'):
                pid = subprocess.Popen(['sirenm.py', ent], shell=True).pid
        else:
            pid = subprocess.Popen(['python', 'sirenm.py', ent]).pid

    def new(self):
        do_new = makeNew()
        do_new.exec_()
        if do_new.ini_file != '':
            self.invoke(do_new.ini_file)
            self.quit()

    def showAbout(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.about, title='About SENs SAM Model')
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
                    pid = subprocess.Popen(['python', who[0], who[1]]).pid
        else:
            if os.path.exists(who):
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if who[-3:] == '.py':
                        pid = subprocess.Popen([who], shell=True).pid
                    else:
                        pid = subprocess.Popen([who]).pid
                else:
                    pid = subprocess.Popen(['python', who]).pid
        return

    def quit(self):
        self.close()


class ClickableQLabel(QtGui.QLabel):
    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtGui.QApplication.widgetAt(event.globalPos()).setFocus()
        self.emit(QtCore.SIGNAL('clicked()'))


class makeNew(QtGui.QDialog):

    def __init__(self, help='help.html'):
        super(makeNew, self).__init__()
        self.help = help
        self.ini_file = ''
        self.initUI()

    def initUI(self):
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            ini_file = 'siren_windows_default.ini'
        else:
            ini_file = 'siren_default.ini'
        if not os.path.exists(ini_file):
            return
        file_sects = ['[Parents]', '[Files]', '[SAM Modules]']
        dir_props = ['pow_files', 'sam_sdk', 'solar_files', 'variable_files', 'wind_files']
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
        self.fields.append(['section', 'typ', 'name', 'value', QtGui.QLineEdit()])
        self.grid = QtGui.QGridLayout()
        self.grid.addWidget(QtGui.QLabel('New file name:'), row, 0)
        self.fields[row][4].setText('siren_new.ini')
        self.grid.addWidget(self.fields[row][4], row, 1)
        self.fields[row][4].textChanged.connect(self.filenameChanged)
        self.msg = QtGui.QLabel('')
        self.grid.addWidget(self.msg, row, 2, 1, 3)
        now = datetime.datetime.now()
        if '[Base]' in sections.keys():
            for key, value in sections['[Base]']:
                row += 1
                self.fields.append(['[Base]', 'txt', key, value, QtGui.QLineEdit()])
                if key == 'year':
                    self.fields[-1][3] = str((now.year - 1))
                if self.fields[row][0] != self.fields[row - 1][0]:
                    self.grid.addWidget(QtGui.QLabel(self.fields[row][0]), row, 0)
                self.grid.addWidget(QtGui.QLabel(self.fields[row][2]), row, 1)
                self.fields[row][4].setText(self.fields[row][3])
                self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
        self.parents = {}
        sections['[Parents]'].append(['$USER$', getUser()])
        sections['[Parents]'].append(['$YEAR$', str((now.year - 1))])
        if '[Parents]' in sections.keys():
            for key, value in sections['[Parents]']:
                self.parents[key] = value
                row += 1
                self.fields.append(['[Parents]', '?', key, value, '?'])
                if self.fields[row][0] != self.fields[row - 1][0]:
                    self.grid.addWidget(QtGui.QLabel(self.fields[row][0]), row, 0)
                self.grid.addWidget(QtGui.QLabel(self.fields[row][2]), row, 1)
                if key == '$USER$' or key == '$YEAR$':
                    self.fields[row][1] = 'txt'
                    self.fields[row][4] = QtGui.QLabel(self.fields[row][3])
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
                else:
                    self.fields[row][1] = 'dir'
                    self.fields[row][4] = ClickableQLabel()
                    self.fields[row][4].setText(self.fields[row][3])
                    self.fields[row][4].setFrameStyle(6)
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
                    self.connect(self.fields[row][4], QtCore.SIGNAL('clicked()'), self.itemClicked)
        for section, props in iter(sections.iteritems()):
            if section == '[Base]' or section == '[Parents]':
                continue
            elif section == '[Map]':
                for prop in props:
                    row += 1
                    self.fields.append([section, '?', prop[0], prop[1], '?'])
                    if self.fields[row][0] != self.fields[row - 1][0]:
                        self.grid.addWidget(QtGui.QLabel(self.fields[row][0]), row, 0)
                    self.grid.addWidget(QtGui.QLabel(self.fields[row][2]), row, 1)
                    if prop[0] == 'map' or (prop[0][:3] == 'map' and prop[0][3] != '_'):
                        self.fields[row][3] = prop[1]
                        self.fields[row][1] = 'fil'
                        self.fields[row][4] = ClickableQLabel()
                        self.fields[row][4].setFrameStyle(6)
                        self.connect(self.fields[row][4], QtCore.SIGNAL('clicked()'), self.itemClicked)
                    else:
                        self.fields[row][1] = 'txt'
                        self.fields[row][4] = QtGui.QLineEdit()
                    self.fields[row][4].setText(self.fields[row][3])
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
            elif section in file_sects:
                for prop in props:
                    row += 1
                    self.fields.append([section, '?', prop[0], prop[1], '?'])
                    if self.fields[row][0] != self.fields[row - 1][0]:
                        self.grid.addWidget(QtGui.QLabel(self.fields[row][0]), row, 0)
                    self.grid.addWidget(QtGui.QLabel(self.fields[row][2]), row, 1)
                    self.fields[row][3] = prop[1]
                    if prop[0] in dir_props:
                        self.fields[row][1] = 'dir'
                    else:
                        self.fields[row][1] = 'fil'
                    self.fields[row][4] = ClickableQLabel()
                    self.fields[row][4].setText(self.fields[row][3])
                    self.fields[row][4].setFrameStyle(6)
                    self.connect(self.fields[row][4], QtCore.SIGNAL('clicked()'), self.itemClicked)
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
            else:
                for prop in props:
                    row += 1
                    self.fields.append([section, 'txt', prop[0], prop[1], QtGui.QLineEdit()])
                    if self.fields[row][0] != self.fields[row - 1][0]:
                        self.grid.addWidget(QtGui.QLabel(self.fields[row][0]), row, 0)
                    self.grid.addWidget(QtGui.QLabel(self.fields[row][2]), row, 1)
                    self.fields[row][4].setText(self.fields[row][3])
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
        row += 1
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, row, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        save = QtGui.QPushButton('Save', self)
        self.grid.addWidget(save, row, 1)
        save.clicked.connect(self.saveClicked)
        launch = QtGui.QPushButton('Save && Open', self)
        self.grid.addWidget(launch, row, 3)
        launch.clicked.connect(self.saveLaunch)
        wdth = save.fontMetrics().boundingRect(launch.text()).width() + 9
        launch.setMaximumWidth(wdth)
        edit = QtGui.QPushButton('Save && Edit', self)
        self.grid.addWidget(edit, row, 2)
        edit.clicked.connect(self.saveEdit)
        edit.setMaximumWidth(wdth)
        help = QtGui.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        self.grid.addWidget(help, row, 4)
        help.clicked.connect(self.helpClicked)
        QtGui.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        self.grid.setColumnStretch(2, 5)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        utilities = ['getmap', 'makeweather2']
        utilicon = ['map.png', 'weather.png']
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
            menubar = QtGui.QMenuBar()
            spawnMenu = menubar.addMenu('&Tools')
            for i in range(len(spawns)):
                if type(spawns[i]) is list:
                    who = spawns[i][0][:spawns[i][0].find('.')]
                else:
                    who = spawns[i][:spawns[i].find('.')]
                spawnitem.append(QtGui.QAction(QtGui.QIcon(icons[i]), who, self))
                spawnitem[-1].triggered.connect(partial(self.spawn, spawns[i]))
                spawnMenu.addAction(spawnitem[-1])
            self.layout.setMenuBar(menubar)
        self.setWindowTitle('SIREN (' + fileVersion() + ') - Create Preferences file')
        self.show()

    def filenameChanged(self):
        if str(self.fields[0][4].text()).lower() == 'siren_default.ini' or \
          str(self.fields[0][4].text()).lower() == 'siren_default' or \
          str(self.fields[0][4].text()).lower() == 'siren_windows_default.ini' or \
          str(self.fields[0][4].text()).lower() == 'siren_windows_default':
            self.msg.setText('File name not allowed.')
        else:
            self.msg.setText('')

    def helpClicked(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.help, \
                 title='Help for SIREN Preferences file (' + fileVersion() + ')', section='prefs')
        dialog.exec_()

    def itemClicked(self):
        for i in range(len(self.fields)):
            if self.fields[i][4].hasFocus():
                raw_field = upd_field = self.fields[i][4].text()
                for key, value in iter(self.parents.iteritems()):
                    upd_field = upd_field.replace(key, value)
                if self.fields[i][1] == 'dir':
                    curdir = upd_field
                    newone = str(QtGui.QFileDialog.getExistingDirectory(self,
                             'Choose ' + self.fields[i][2] + ' Folder', curdir,
                             QtGui.QFileDialog.ShowDirsOnly))
                    if newone != '':
                        newone = newone.replace('\\', '/')
                        if self.fields[i][0] == '[Parents]':
                            self.parents[self.fields[i][2]] = newone
                        else:
                            longest = [0, '']
                            for key, value in iter(self.parents.iteritems()):
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
                    newone = str(QtGui.QFileDialog.getOpenFileName(self,
                             'Choose ' + self.fields[i][2] + ' File', curfil))
                    if newone != '':
                        newone = newone.replace('\\', '/')
                        longest = [0, '']
                        for key, value in iter(self.parents.iteritems()):
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
        self.saveIni()
        self.close()

    def saveEdit(self):
        self.saveIni()
        dialr = EdtDialog(self.new_ini)
        dialr.exec_()
        self.close()

    def saveLaunch(self):
        self.saveIni()
        self.ini_file = self.new_ini
        self.close()

    def saveIni(self):
        updates = {}
        lines = []
        newfile = str(self.fields[0][4].text())
        if newfile == 'siren_default.ini' or newfile == 'siren_default' or \
          newfile == 'siren_windows_default.ini' or newfile == 'siren_windows_default':
            self.msg.setText('File name not allowed.')
            return
        if newfile[-4:].lower() != '.ini':
            newfile = newfile + '.ini'
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
            if field[0] not in updates.keys():
                updates[field[0]] = []
            fld = str(field[4].text())
            if self.parents['$YEAR$'] in fld:
                if field[0] == '[Base]' and (field[2] == 'year' or field[2] == 'years'):
                    pass
                else:
                    fld = fld.replace(self.parents['$YEAR$'], '$YEAR$')
            if self.parents['$USER$'] in fld:
                fld = fld.replace(self.parents['$USER$'], '$USER$')
#
            if len(fld) > len(adj_dir):
                if fld[:len(adj_dir)] == adj_dir:
                    fld = fld[len(adj_dir) + 1:]
            updates[field[0]].append(field[2] + '=' + fld)
        if '[Parents]' in updates.keys():
            my_dir = os.getcwd()
            my_dir = my_dir.replace('\\', '/')
            my_dir = my_dir.replace(self.parents['$YEAR$'], '$YEAR$')
            my_dir = my_dir.replace(self.parents['$USER$'], '$USER$')
            for p in range(len(updates['[Parents]'])):
                i = updates['[Parents]'][p].find('=')
                value = updates['[Parents]'][p][i + 1:]
                if value.find('/') >= 0:
                    that_len = len(commonprefix([my_dir, value]))
                    if that_len > 0:
                        bits = my_dir[that_len:].split('/')
                        pfx = '../' * (len(bits) - 1)
                        updates['[Parents]'][p] = updates['[Parents]'][p][:i + 1] + pfx + value[that_len + 1:]
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            ini_file = 'siren_windows_default.ini'
        else:
            ini_file = 'siren_default.ini'
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

    def spawn(self, who):
        if type(who) is list:
            if os.path.exists(who[0]):
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if who[0][-3:] == '.py':
                        pid = subprocess.Popen([who], shell=True).pid
                    else:
                        pid = subprocess.Popen([who]).pid
                else:
                    pid = subprocess.Popen(['python', who[0], who[1]]).pid
        else:
            if os.path.exists(who):
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if who[-3:] == '.py':
                        pid = subprocess.Popen([who], shell=True).pid
                    else:
                        pid = subprocess.Popen([who]).pid
                else:
                    pid = subprocess.Popen(['python', who]).pid
        return


if '__main__' == __name__:
    app = QtGui.QApplication(sys.argv)
    tabdialog = TabDialog()
    tabdialog.show()
    sys.exit(app.exec_())
