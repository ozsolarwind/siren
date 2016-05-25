#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
#
#  Editini.py - This file is part of SIREN.
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

from PyQt4 import QtCore, QtGui
from PyQt4.QtGui import QDesktopWidget
import ConfigParser   # decode .ini file
import os
import sys

import displaytable
import inisyntax


class EdtDialog(QtGui.QDialog):
    def __init__(self, in_file, parent=None):
        self.in_file = in_file
        try:
            s = open(self.in_file, 'r').read()
            bits = s.split('\n')
            ln = 0
            for lin in bits:
                if len(lin) > ln:
                    ln = len(lin)
            ln2 = len(bits)
        except:
            if self.in_file[self.in_file.rfind('.'):] == '.ini':
                s = ''
            else:
                s = ''
            ln = 36
            ln2 = 5
        QtGui.QDialog.__init__(self, parent)
        self.saveButton = QtGui.QPushButton(self.tr('&Save'))
        self.cancelButton = QtGui.QPushButton(self.tr('Cancel'))
        buttonLayout = QtGui.QHBoxLayout()
        buttonLayout.addStretch(1)
        buttonLayout.addWidget(self.saveButton)
        buttonLayout.addWidget(self.cancelButton)
        self.connect(self.saveButton, QtCore.SIGNAL('clicked()'), self,
                     QtCore.SLOT('accept()'))
        self.connect(self.cancelButton, QtCore.SIGNAL('clicked()'),
                     self, QtCore.SLOT('reject()'))
        self.widget = QtGui.QPlainTextEdit()
        highlight = inisyntax.IniHighlighter(self.widget.document())
        if sys.platform == 'linux2':
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
        layout = QtGui.QVBoxLayout()
        layout.addWidget(self.widget)
        layout.addLayout(buttonLayout)
        self.setLayout(layout)
        self.setWindowTitle('SIREN - Edit - ' + self.in_file[self.in_file.rfind(os.sep) + 1:])
        size = self.geometry()
        self.setGeometry(1, 1, ln + 10, ln2 + 35)
        size = self.geometry()
        self.move((screen.width() - size.width()) / 2,
            (screen.height() - size.height()) / 2)
        self.widget.show()

    def accept(self):
        t = str(self.widget.toPlainText())
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


class EditSect():
    def __init__(self, section, save_folder):
        self.section = section
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        section_items = config.items(self.section)
        section_dict = {}
        for key, value in section_items:
            section_dict[key] = value
        dialog = displaytable.Table(section_dict, fields=['property', 'value'], title=self.section + ' Parameters',
                 save_folder=save_folder, edit=True)
        dialog.exec_()
        values = dialog.getValues()
        if values is None:
            return
        section_dict = {}
        section_items = []
        for key in values:
            section_items.append(key + '=' + values[key][0][6:])
        section_dict[self.section] = section_items
        SaveIni(section_dict)


class EditTech():
    def __init__(self, save_folder):
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        technologies = config.get('Power', 'technologies')
        technologies += ' ' + config.get('Power', 'fossil_technologies')
        technologies = technologies.split(' ')
        for i in range(len(technologies)):
            technologies[i] = technologies[i].replace('_', ' ').title()
            technologies[i] = technologies[i].replace('Pv', 'PV')
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
        dialog = displaytable.Table(tech_dict, fields=['technology', 'area', 'capital_cost', 'o_m_cost'],
                 save_folder=save_folder, title='Technologies', edit=True)
        dialog.exec_()
        values = dialog.getValues()
        if values is None:
            return
        SaveIni(values)


class SaveIni():
    def __init__(self, values):
        if len(sys.argv) > 1:
            ini_file = sys.argv[1]
        else:
            ini_file = 'SIREN.ini'
        inf = open(ini_file, 'r')
        lines = inf.readlines()
        inf.close()
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
                                lines[i] = properties[j] + '\n'
                                del properties[j]
                                del props[j]
            if len(properties) > 0:
                if not in_section:
                    lines.append('[' + section + ']\n')
                    i += 1
                for j in range(len(properties)):
                    lines.insert(i + 1, properties[j] + '\n')
                    i += 1
        if os.path.exists(ini_file + '~'):
            os.remove(ini_file + '~')
        os.rename(ini_file, ini_file + '~')
        sou = open(ini_file, 'w')
        for i in range(len(lines)):
            sou.write(lines[i])
        sou.close()
