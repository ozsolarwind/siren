#!/usr/bin/python3
#
#  Copyright (C) 2015-2019 Sustainable Energy Now Inc., Angus King
#
#  displayobject.py - This file is part of SIREN.
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

import os
import sys
from PyQt4 import QtCore
from PyQt4 import QtGui

from credits import fileVersion
from turbine import Turbine


class AnObject(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, dialog, anobject, readonly=True, title=None, section=None, textedit=True):
        super(AnObject, self).__init__()
        self.anobject = anobject
        self.readonly = readonly
        self.title = title
        self.section = section
        self.textedit = textedit
        dialog.setObjectName('Dialog')
        self.initUI()

    def set_stuff(self, grid, widths, heights, i):
        if widths[1] > 0:
            grid.setColumnMinimumWidth(0, widths[0] + 10)
            grid.setColumnMinimumWidth(1, widths[1] + 10)
        i += 1
        if isinstance(self.anobject, str):
            quit = QtGui.QPushButton('Close', self)
            width = quit.fontMetrics().boundingRect('Close').width() + 10
            quit.setMaximumWidth(width)
        else:
            quit = QtGui.QPushButton('Quit', self)
        grid.addWidget(quit, i + 1, 0)
        quit.clicked.connect(self.quitClicked)
        if not self.readonly:
            save = QtGui.QPushButton("Save", self)
            grid.addWidget(save, i + 1, 1)
            save.clicked.connect(self.saveClicked)
        self.setLayout(grid)
        screen = QtGui.QDesktopWidget().availableGeometry()
        h = heights * i
        if h > screen.height():
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                pct = 0.85
            else:
                pct = 0.90
            h = int(screen.height() * pct)
        self.resize(widths[0] + widths[1] + 40, h)
        if self.title is not None:
            self.setWindowTitle('SIREN - ' + self.title)
        elif isinstance(self.anobject, str) or isinstance(self.anobject, dict):
            self.setWindowTitle('SIREN - ?')
        else:
            self.setWindowTitle('SIREN - Review ' + getattr(self.anobject, '__module__'))
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))

    def initUI(self):
        label = []
        self.edit = []
        self.field_type = []
        metrics = []
        widths = [0, 0]
        heights = 0
        i = -1
        grid = QtGui.QGridLayout()
        if isinstance(self.anobject, str):
            self.web = QtGui.QTextEdit()
            if os.path.exists(self.anobject):
                htf = open(self.anobject, 'r')
                html = htf.read()
                htf.close()
                if self.anobject[-5:].lower() == '.html' or \
                   self.anobject[-4:].lower() == '.htm' or \
                   html[:5] == '<html':
                    html = html.replace('[VERSION]', fileVersion())
                    if self.section is not None:
                        line = html.split('\n')
                        html = ''
                        for i in range(len(line)):
                            html += line[i] + '\n'
                            if line[i].strip() == '<body>':
                               break
                        for i in range(i, len(line)):
                            if line[i][:2] == '<h':
                                if line[i].find('id="' + self.section + '"') > 0:
                                    break
                        for i in range(i, len(line)):
                            if line[i].find('Back to top<') > 0:
                                break
                            j = line[i].find(' (see <a href=')
                            if j > 0:
                                k = line[i].find('</a>)', j)
                                line[i] = line[i][:j] + line[i][k + 5:]
                            html += line[i] + '\n'
                        for i in range(i, len(line)):
                            if line[i].strip() == '</body>':
                                break
                        for i in range(i, len(line)):
                            html += line[i] + '\n'
                    self.web.setHtml(html)
                else:
                    self.web.setPlainText(html)
            else:
                html = self.anobject
                if self.anobject[:5] == '<html':
                    self.anobject = self.anobject.replace('[VERSION]', fileVersion())
                    self.web.setHtml(self.anobject)
                else:
                    self.web.setPlainText(self.anobject)
            metrics.append(self.web.fontMetrics())
            try:
                widths[0] = metrics[0].boundingRect(self.web.text()).width()
                heights = metrics[0].boundingRect(self.web.text()).height()
            except:
                bits = html.split('\n')
                for lin in bits:
                    if len(lin) > widths[0]:
                        widths[0] = len(lin)
                heights = len(bits)
                fnt = self.web.fontMetrics()
                widths[0] = (widths[0]) * fnt.maxWidth()
                heights = (heights) * fnt.height()
                screen = QtGui.QDesktopWidget().availableGeometry()
                if widths[0] > screen.width() * .67:
                    heights = int(heights / .67)
                    widths[0] = int(screen.width() * .67)
            if self.readonly:
                self.web.setReadOnly(True)
            i = 1
            grid.addWidget(self.web, 0, 0)
            self.set_stuff(grid, widths, heights, i)
        elif isinstance(self.anobject, dict):
            if self.textedit:
                self.keys = []
                for key, value in self.anobject.items():
                    self.field_type.append('str')
                    label.append(QtGui.QLabel(key + ':'))
                    self.keys.append(key)
                    self.edit.append(QtGui.QTextEdit())
                    self.edit[-1].setPlainText(value)
                 #   print '(160)', key, self.edit[-1].document().blockCount()
                    if i < 0:
                        metrics.append(label[-1].fontMetrics())
                        metrics.append(self.edit[-1].fontMetrics())
                    bits = value.split('\n')
                    ln = 0
                    for lin in bits:
                        if len(lin) > ln:
                            ln = len(lin)
                    ln2 = len(bits)
                    ln = (ln + 5) * metrics[0].maxWidth()
                    ln2 = (ln2 + 4) * metrics[0].height()
                    self.edit[-1].resize(ln, ln2)
                    if metrics[0].boundingRect(label[-1].text()).width() > widths[0]:
                        widths[0] = metrics[0].boundingRect(label[-1].text()).width()
                    try:
                        if metrics[1].boundingRect(self.edit[-1].text()).width() > widths[1]:
                            widths[1] = metrics[1].boundingRect(self.edit[-1].text()).width()
                    except:
                        widths[1] = ln
                    for j in range(2):
                        try:
                            if metrics[j].boundingRect(label[-1].text()).height() > heights:
                                heights = metrics[j].boundingRect(label[-1].text()).height()
                        except:
                            heights = ln2
                    if self.readonly:
                        self.edit[-1].setReadOnly(True)
                    i += 1
                    grid.addWidget(label[-1], i + 1, 0)
                    grid.addWidget(self.edit[-1], i + 1, 1)
            else:
                self.keys = []
                for key, value in self.anobject.items():
                    self.field_type.append('str')
                    label.append(QtGui.QLabel(key + ':'))
                    self.keys.append(key)
                    self.edit.append(QtGui.QLineEdit())
                    self.edit[-1].setText(value)
                    if i < 0:
                        metrics.append(label[-1].fontMetrics())
                        metrics.append(self.edit[-1].fontMetrics())
                    ln = (len(value) + 5) * metrics[0].maxWidth()
                    ln2 = metrics[0].height()
                    self.edit[-1].resize(ln, ln2)
                    if metrics[0].boundingRect(label[-1].text()).width() > widths[0]:
                        widths[0] = metrics[0].boundingRect(label[-1].text()).width()
                    try:
                        if metrics[1].boundingRect(self.edit[-1].text()).width() > widths[1]:
                            widths[1] = metrics[1].boundingRect(self.edit[-1].text()).width()
                    except:
                        widths[1] = ln
                    if self.readonly:
                        self.edit[-1].setReadOnly(True)
                    i += 1
                    grid.addWidget(label[-1], i, 0)
                    grid.addWidget(self.edit[-1], i, 1)
            self.set_stuff(grid, widths, heights, i)
        else:
            units = {'area': 'sq. Km', 'capacity': 'MW', 'rotor': 'm', 'generation': 'MWh', 'grid_len': 'Km',
                     'grid_path_len': 'Km'}
            for prop in dir(self.anobject):
                if prop[:2] != '__' and prop[-2:] != '__':
                    attr = getattr(self.anobject, prop)
                    if isinstance(attr, int):
                         self.field_type.append('int')
                    elif isinstance(attr, float):
                         self.field_type.append('float')
                    else:
                         self.field_type.append('str')
                    label.append(QtGui.QLabel(prop.title() + ':'))
                    if self.field_type[-1] != "str":
                        self.edit.append(QtGui.QLineEdit(str(attr)))
                    else:
                        self.edit.append(QtGui.QLineEdit(attr))
                    if i < 0:
                        metrics.append(label[-1].fontMetrics())
                        metrics.append(self.edit[-1].fontMetrics())
                    if metrics[0].boundingRect(label[-1].text()).width() > widths[0]:
                        widths[0] = metrics[0].boundingRect(label[-1].text()).width()
                    if metrics[1].boundingRect(self.edit[-1].text()).width() > widths[1]:
                        widths[1] = metrics[1].boundingRect(self.edit[-1].text()).width()
                    for j in range(2):
                        if metrics[j].boundingRect(label[-1].text()).height() > heights:
                            heights = metrics[j].boundingRect(label[-1].text()).height()
                    if self.readonly:
                        self.edit[-1].setReadOnly(True)
                    i += 1
                    grid.addWidget(label[-1], i + 1, 0)
                    grid.addWidget(self.edit[-1], i + 1, 1)
                    if prop in list(units.keys()):
                        grid.addWidget(QtGui.QLabel(units[prop]), i + 1, 2)
                    if prop == 'turbine':
                        i += 1
                        curve = QtGui.QPushButton('Show Power Curve', self)
                        grid.addWidget(curve, i + 1, 1)
                        curve.clicked.connect(self.curveClicked)
                        self.turbine = attr
            self.set_stuff(grid, widths, heights, i)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)

    def curveClicked(self):
        Turbine(self.turbine).PowerCurve()
        return

    def quitClicked(self):
        self.close()

    def saveClicked(self):
        if isinstance(self.anobject, dict):
            if self.textedit:
                for i in range(len(self.keys)):
                    self.anobject[self.keys[i]] = str(self.edit[i].toPlainText())
            else:
                for i in range(len(self.keys)):
                    self.anobject[self.keys[i]] = str(self.edit[i].text())
        else:
            i = -1
            for prop in dir(self.anobject):
                if prop[:2] != '__' and prop[-2:] != '__':
                    i += 1
                    if self.field_type[i] == 'int':
                        setattr(self.anobject, prop, int(self.edit[i].text()))
                    elif self.field_type[i] == 'float':
                        setattr(self.anobject, prop, float(self.edit[i].text()))
                    else:
                        setattr(self.anobject, prop, str(self.edit[i].text()))
        self.close()

    def getValues(self):
        return self.anobject
