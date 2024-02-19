#!/usr/bin/python3
#
#  Copyright (C) 2015-2023 Sustainable Energy Now Inc., Angus King
#
#  editini.py - This file is part of SIREN.
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

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QDesktopWidget
import configparser   # decode .ini file
import datetime
import os
import sys

from displaytable import Table
from getmodels import getModelFile, commonprefix
import inisyntax
from senutils import ClickableQLabel, getParents, getUser, techClean


class EdtDialog(QtWidgets.QDialog):
    def __init__(self, in_file, parent=None, line=None, section=None, save_as=False):
        self.in_file = in_file
        try:
            s = open(self.in_file, 'r').read()
            bits = s.split('\n')
            ln = 0
            for l in range(len(bits)):
                if len(bits[l]) > ln:
                    ln = len(bits[l])
                if section is not None:
                    if bits[l] == section:
                        line = l
            ln2 = len(bits)
        except:
            if self.in_file[self.in_file.rfind('.'):] == '.ini':
                s = ''
            else:
                s = ''
            ln = 36
            ln2 = 5
        QtWidgets.QDialog.__init__(self, parent)
        self.findbwd = QtWidgets.QPushButton('<')
        self.search = QtWidgets.QLineEdit('')
        self.findfwd = QtWidgets.QPushButton('>')
        self.saveButton = QtWidgets.QPushButton(self.tr('&Save'))
        self.cancelButton = QtWidgets.QPushButton(self.tr('Cancel'))
        buttonLayout = QtWidgets.QHBoxLayout()
        buttonLayout.addWidget(QtWidgets.QLabel('Find:'))
        buttonLayout.addWidget(self.findbwd)
        buttonLayout.addWidget(self.search)
        buttonLayout.addWidget(self.findfwd)
        self.findmsg = QtWidgets.QLabel('')
        buttonLayout.addWidget(self.findmsg)
        buttonLayout.addStretch(1)
        buttonLayout.addWidget(self.saveButton)
        buttonLayout.addWidget(self.cancelButton)
        self.findbwd.clicked.connect(self.findBwd)
        self.findfwd.clicked.connect(self.findText)
        self.saveButton.clicked.connect(self.accept)
        self.cancelButton.clicked.connect(self.reject)
        if save_as:
            self.saveasButton = QtWidgets.QPushButton(self.tr('Save as'))
            buttonLayout.addWidget(self.saveasButton)
            self.saveasButton.clicked.connect(self.saveas)
        self.widget = QtWidgets.QPlainTextEdit()
        self.widget.setStyleSheet('selection-background-color: #06A9D6; selection-color: black')
        highlight = inisyntax.IniHighlighter(self.widget.document(), line=line)
        if sys.platform == 'linux' or sys.platform == 'linux2':
            self.widget.setFont(QtGui.QFont('Ubuntu Mono 13', 12))
        else:
            self.widget.setFont(QtGui.QFont('Courier New', 12))
        fnt = self.widget.fontMetrics()
        ln = (ln + 5) * fnt.maxWidth()
        ln2 = (ln2 + 4) * fnt.height()
        screen = QDesktopWidget().availableGeometry()
        if ln > screen.width() * .67:
            ln = int(screen.width() * .67)
        if ln2 > screen.height() * .67:
            ln2 = int(screen.height() * .67)
        self.widget.resize(ln, ln2)
        self.widget.setPlainText(s)
        if section is not None:
            where = self.widget.find(section)
        elif line is not None:
            for l in range(line):
                self.widget.moveCursor(QtGui.QTextCursor.NextBlock)
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.widget)
        layout.addLayout(buttonLayout)
        self.setLayout(layout)
        self.setWindowTitle('SIREN - Edit - ' + self.in_file[self.in_file.rfind('/') + 1:])
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        size = self.geometry()
        self.setGeometry(1, 1, ln + 10, ln2 + 35)
        size = self.geometry()
        self.move(int((screen.width() - size.width()) / 2), int((screen.height() - size.height()) / 2))
        self.widget.show()

    def findBwd(self):
        self.findmsg.setText('')
        ft = self.search.text()
        if self.widget.find(ft, QtGui.QTextDocument.FindBackward):
            return
        else:
            self.findmsg.setText('Wrapped to bottom')
            self.widget.moveCursor(QtGui.QTextCursor.End,
                                   QtGui.QTextCursor.MoveAnchor)
        #    self.widget.moveCursor(-1)
       #     if self.widget.find(ft, QtGui.QTextDocument.FindBackward):
        #        self.widget.moveCursor(QtGui.QTextCursor.End,
         #                              QtGui.QTextCursor.MoveAnchor)

    def findText(self):
        self.findmsg.setText('')
        ft = self.search.text()
        if self.widget.find(ft):
            return
        else:
            self.findmsg.setText('Wrapped to top')
        #    self.widget.moveCursor(1)
         #   if self.widget.find(ft):
            self.widget.moveCursor(QtGui.QTextCursor.Start,
                                   QtGui.QTextCursor.MoveAnchor)

    def accept(self):
        t = self.widget.toPlainText()
        bits = t.split('\n')
        if self.in_file[self.in_file.rfind('.'):] != '.ini':
            bits.sort()
            while len(bits[0]) == 0:
                del bits[0]
        s = open(self.in_file, 'w')
        for lin in bits:
            if len(lin) > 0:
                s.write(lin + '\n')
        s.close()
        self.close()

    def saveas(self):
        newfile = QtWidgets.QFileDialog.getSaveFileName(None, 'Save Preferences file',
                  self.in_file, 'Preference files (*.ini)')[0]
        if newfile != '':
            self.in_file = newfile
            self.accept()


class EditFileSections(QtWidgets.QDialog):
    def __init__(self, ini_file=None, parent=None):
        QtWidgets.QDialog.__init__(self, parent)
        file_sects = ['Parents', 'Files', 'SAM Modules']
        dir_props = ['pow_files', 'sam_sdk', 'scenarios', 'solar_files', 'variable_files', 'wind_files']
        field_props = ['check', 'scenario_prefix']
        config = configparser.RawConfigParser()
        if ini_file is not None:
            self.config_file = ini_file
        elif len(sys.argv) > 1:
            self.config_file = sys.argv[1]
        else:
            self.config_file = getModelFile('SIREN.ini')
        config.read(self.config_file)
        sections = {}
        for section in file_sects:
            props = []
            try:
                items = config.items(section)
            except:
                continue
            for key, value in items:
                props.append([key, value])
            if len(props) > 0:
                sections[section] = props
        bgd = 'rgba({}, {}, {}, {})'.format(*self.palette().color(QtGui.QPalette.Window).getRgb())
        bgd_style = 'background-color: ' + bgd + '; border: 1px inset grey; min-height: 22px; border-radius: 4px;'
        bold = QtGui.QFont()
        bold.setBold(True)
        width0 = 0
        width1 = 0
        width2 = 0
        row = 0
        self.fields = []
        self.fields.append(['section', 'typ', 'name', 'value', QtWidgets.QLineEdit()])
        self.grid = QtWidgets.QGridLayout()
        label = QtWidgets.QLabel('Working directory')
        label.setFont(bold)
        width0 = label.fontMetrics().boundingRect(label.text()).width() * 1.1
        self.grid.addWidget(label, row, 0)
        self.fields[row][4].setText(os.getcwd())
        self.fields[row][4].setReadOnly(True)
        self.fields[row][4].setStyleSheet(bgd_style)
        width2 = self.fields[row][4].fontMetrics().boundingRect(self.fields[row][4].text()).width() * 1.1
        self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
        self.parents = {}
        sections['Parents'].append(['$USER$', getUser()])
        now = datetime.datetime.now()
        sections['Parents'].append(['$YEAR$', str(now.year - 1)])
        if 'Parents' in list(sections.keys()):
            for key, value in sections['Parents']:
                self.parents[key] = value
                row += 1
                self.fields.append(['Parents', '?', key, value, '?'])
                if self.fields[row][0] != self.fields[row - 1][0]:
                    label = QtWidgets.QLabel(self.fields[row][0])
                    label.setFont(bold)
                    width0 = max(width0, label.fontMetrics().boundingRect(label.text()).width() * 1.1)
                    self.grid.addWidget(label, row, 0)
                label = QtWidgets.QLabel(self.fields[row][2])
                width1 = max(width1, label.fontMetrics().boundingRect(label.text()).width() * 1.1)
                self.grid.addWidget(label, row, 1)
                if key == '$USER$' or key == '$YEAR$':
                    self.fields[row][1] = 'txt'
                    self.fields[row][4] = QtWidgets.QLineEdit()
                    self.fields[row][4].setText(self.fields[row][3])
                    width2 = max(width2, self.fields[row][4].fontMetrics().boundingRect(self.fields[row][4].text()).width() * 1.1)
                    self.fields[row][4].setReadOnly(True)
                    self.fields[row][4].setStyleSheet(bgd_style)
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
                else:
                    self.fields[row][1] = 'dir'
                    self.fields[row][4] = ClickableQLabel()
                    self.fields[row][4].setText(self.fields[row][3])
                    width2 = max(width2, self.fields[row][4].fontMetrics().boundingRect(self.fields[row][4].text()).width() * 1.1)
                    self.fields[row][4].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
                    self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
                    self.fields[row][4].clicked.connect(self.itemClicked)
        for section, props in sections.items():
            if section == 'Parents':
                continue
            for prop in props:
                row += 1
                self.fields.append([section, '?', prop[0], prop[1], '?'])
                if self.fields[row][0] != self.fields[row - 1][0]:
                    label = QtWidgets.QLabel(self.fields[row][0])
                    label.setFont(bold)
                    width0 = max(width0, label.fontMetrics().boundingRect(label.text()).width() * 1.1)
                    self.grid.addWidget(label, row, 0)
                label = QtWidgets.QLabel(self.fields[row][2])
                width1 = max(width1, label.fontMetrics().boundingRect(label.text()).width() * 1.1)
                self.grid.addWidget(label, row, 1)
                self.fields[row][3] = prop[1]
                if prop[0] in field_props:
                    self.fields[row][1] = 'txt'
                    self.fields[row][4] = QtWidgets.QLineEdit()
                    self.fields[row][4].setText(self.fields[row][3])
                    width2 = max(width2, self.fields[row][4].fontMetrics().boundingRect(self.fields[row][4].text()).width() * 1.1)
                else:
                    if prop[0] in dir_props:
                        self.fields[row][1] = 'dir'
                    else:
                        self.fields[row][1] = 'fil'
                    self.fields[row][4] = ClickableQLabel()
                    self.fields[row][4].setText(self.fields[row][3])
                    width2 = max(width2, self.fields[row][4].fontMetrics().boundingRect(self.fields[row][4].text()).width() * 1.1)
                    self.fields[row][4].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
                    self.fields[row][4].clicked.connect(self.itemClicked)
                self.grid.addWidget(self.fields[row][4], row, 2, 1, 3)
        row += 1
        quit = QtWidgets.QPushButton('Quit', self)
        self.grid.addWidget(quit, row, 0)
        quit.clicked.connect(self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        save = QtWidgets.QPushButton('Save & Exit', self)
        self.grid.addWidget(save, row, 1)
        save.clicked.connect(self.saveClicked)
        self.grid.setColumnStretch(2, 5)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Edit File Sections - ' + ini_file[ini_file.rfind('/') + 1:])
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.resize(int(width0 + width1 + width2), int(self.sizeHint().height() * 1.1))
        self.show()

    def itemClicked(self):
        for i in range(len(self.fields)):
            if self.fields[i][4].hasFocus():
                suffix = ''
                has_user = False
                has_year = False
                upd_field = self.fields[i][4].text()
                if upd_field[-1] == '*':
                    j = upd_field.rfind('/')
                    suffix = upd_field[j:]
                    upd_field = upd_field[:j]
                else:
                    suffix = ''
                if upd_field.find('$USER$') >= 0:
                    has_user = True
                if upd_field.find('$YEAR$') >= 0:
                    has_year = True
                parents = getParents(list(self.parents.items()))
                for key, value in parents:
                    upd_field = upd_field.replace(key, value)
                if self.fields[i][1] == 'dir':
                    curdir = upd_field
                    newone = QtWidgets.QFileDialog.getExistingDirectory(self,
                             'Choose ' + self.fields[i][2] + ' Folder', curdir,
                             QtWidgets.QFileDialog.ShowDirsOnly)
                    if newone != '':
                        if self.fields[i][0] == 'Parents':
                            self.parents[self.fields[i][2]] = newone
                        else:
                            longest = [0, '']
                            for key, value in self.parents.items():
                                if len(newone) > len(value) and len(value) > longest[0]:
                                    if newone[:len(value)] == value:
                                        longest = [len(value), key]
                            if longest[0] > 0:
                                newone = longest[1] + newone[longest[0]:]
                            if has_year and self.parents['$YEAR$'] in newone:
                                newone = newone.replace(self.parents['$YEAR$'], '$YEAR$')
                            if has_user and self.parents['$USER$'] in newone:
                                newone = newone.replace(self.parents['$USER$'], '$USER$')
                        self.fields[i][4].setText(newone + suffix)
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
                        if has_year and self.parents['$YEAR$'] in newone:
                            newone = newone.replace(self.parents['$YEAR$'], '$YEAR$')
                        if has_user and self.parents['$USER$'] in newone:
                            newone = newone.replace(self.parents['$USER$'], '$USER$')
                        self.fields[i][4].setText(newone + suffix)
                break

    def quitClicked(self):
        self.close()

    def saveClicked(self):
        updates = {}
        lines = []
        my_dir = adj_dir = os.getcwd()
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            adj_dir = adj_dir.replace('\\', '/')
        for field in self.fields:
            if field[0] == 'section':
                continue
            if field[0] == 'Parents' and (field[2] == '$USER$' or field[2] == '$YEAR$'):
                continue
            if field[0] not in list(updates.keys()):
                updates[field[0]] = []
            fld = field[4].text()
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
        if 'Parents' in list(updates.keys()):
            my_dir = os.getcwd()
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                my_dir = my_dir.replace('\\', '/')
            for p in range(len(updates['Parents'])):
                i = updates['Parents'][p].find('=')
                value = updates['Parents'][p][i + 1:]
                if value.find('/') >= 0:
                    that_len = len(commonprefix([my_dir, value]))
                    if that_len > 0:
                        bits = my_dir[that_len:].split('/')
                        pfx = ('../') * (len(bits) - 1)
                       # pfx = pfx[:-1]
                        updates['Parents'][p] = updates['Parents'][p][:i + 1] + pfx + value[that_len + 1:]
        SaveIni(updates, ini_file=self.config_file)
        self.close()

class EditSect():
    def __init__(self, section, save_folder, ini_file=None, txt_ok=None):
        self.section = section
        config = configparser.RawConfigParser()
        if ini_file is not None:
            config_file = ini_file
        elif len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = getModelFile('SIREN.ini')
        config.read(config_file)
        section_items = config.items(self.section)
        section_dict = {}
        for key, value in section_items:
            section_dict[key] = value
        dialog = Table(section_dict, fields=['property', 'value'], title=self.section + ' Parameters',
                 save_folder=save_folder, edit=True, txt_ok=txt_ok)
        dialog.exec_()
        values = dialog.getValues()
        if values is None:
            return
        section_dict = {}
        section_items = []
        for key in values:
            try:
                section_items.append(key + '=' + values[key][0][6:])
            except:
                section_items.append(key + '=')
        section_dict[self.section] = section_items
        SaveIni(section_dict, ini_file=config_file)


class EditTech():
    def __init__(self, save_folder, ini_file=None):
        config = configparser.RawConfigParser()
        if ini_file is not None:
            config_file = ini_file
        elif len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = getModelFile('SIREN.ini')
        config.read(config_file)
        technologies = config.get('Power', 'technologies')
        technologies += ' ' + config.get('Power', 'fossil_technologies')
        technologies = technologies.split()
        for i in range(len(technologies)):
            technologies[i] = techClean(technologies[i])
        technologies = sorted(technologies)
        tech_dict = {}
        for technology in technologies:
            area = 0.
            capital_cost = ''
            o_m_cost = ''
            try:
                area = float(config.get(technology, 'area'))
            except:
                pass
            try:
                capital_cost = config.get(technology, 'capital_cost')
                if capital_cost[-1] == 'K':
                    capital_cost = float(capital_cost[:-1]) * pow(10, 3)
                elif capital_cost[-1] == 'M':
                    capital_cost = float(capital_cost[:-1]) * pow(10, 6)
                else:
                    capital_cost = float(capital_cost)
            except:
                pass
            try:
                o_m_cost = config.get(technology, 'o_m_cost')
                if o_m_cost[-1] == 'K':
                    o_m_cost = float(o_m_cost[:-1]) * pow(10, 3)
                elif o_m_cost[-1] == 'M':
                    o_m_cost = float(o_m_cost[:-1]) * pow(10, 6)
                else:
                    o_m_cost = float(o_m_cost)
            except:
                pass
            tech_dict[technology] = [area, capital_cost, o_m_cost]
        dialog = Table(tech_dict, fields=['technology', 'area', 'capital_cost', 'o_m_cost'],
                 save_folder=save_folder, title='Technologies', edit=True)
        dialog.exec_()
        values = dialog.getValues()
        if values is None:
            return
        SaveIni(values)


class SaveIni():
    def __init__(self, values, ini_file=None):
        if ini_file is not None:
            config_file = ini_file
        elif len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = getModelFile('SIREN.ini')
        try:
            inf = open(config_file, 'r')
            lines = inf.readlines()
            inf.close()
        except:
            lines = []
        del_lines = []
        for section in values:
            in_section = False
            properties = values[section]
            props = []
            for i in range(len(properties)):
                props.append(properties[i].split('=')[0])
            for i in range(len(lines)):
                if lines[i][:len(section) + 2] == '[' + section + ']':
                    in_section = True
                elif in_section:
                    if lines[i][0] == '[':
                        i -= 1
                        break
                    elif lines[i][0] != ';' and lines[i][0] != '#':
                        bits = lines[i].split('=')
                        for j in range(len(properties) - 1, -1, -1):
                            if bits[0] == props[j]:
                                if properties[j] == props[j] + '=' \
                                  or properties[j] == props[j]: # delete empty values
                                    del_lines.append(i)
                                else:
                                    lines[i] = properties[j] + '\n'
                                del properties[j]
                                del props[j]
            if len(properties) > 0:
                if not in_section:
                    lines.append('[' + section + ']\n')
                    i += 1
                for j in range(len(properties)):
                    k = properties[j].find('=')
                    if k > 0 and k != len(properties[j]) - 1:
                        lines.insert(i + 1, properties[j] + '\n')
                        i += 1
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            try:
                sou = open(config_file, 'w')
            except PermissionError as err:
                print(f"Permission error: {err}")
                return
            except:
                return
        if os.path.exists(config_file + '~'):
            os.remove(config_file + '~')
        try:
            os.rename(config_file, config_file + '~')
        except:
            pass
        sou = open(config_file, 'w')
        for i in range(len(lines)):
            if i in del_lines:
                pass
            else:
                sou.write(lines[i])
        sou.close()
