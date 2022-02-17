#!/usr/bin/python3
#
#  Copyright (C) 2015-2022 Sustainable Energy Now Inc., Angus King
#
#  colours.py - This file is part of SIREN.
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
from PyQt5 import QtCore, QtGui, QtWidgets
import configparser  # decode .ini file
from editini import SaveIni
from getmodels import getModelFile
from senuser import techClean

class Colours(QtWidgets.QDialog):

    def __init__(self, ini_file=None, section='Colors', add_colour=False, palette=None, underscore=False):
        super(Colours, self).__init__()
        config = configparser.RawConfigParser()
        if ini_file is not None:
            self.config_file = ini_file
        elif len(sys.argv) > 1:
            self.config_file = sys.argv[1]
        else:
            self.config_file = getModelFile('SIREN.ini')
        self.section = section
        self.underscore = underscore
        config.read(self.config_file)
        groups = ['Fossil Technologies', 'Grid', 'Map', 'Plot', 'Resource', 'Technologies', 'The Rest']
        map_colours = ['background', 'border', 'fossil', 'fossil_name', 'station', 'station_name',
                       'town', 'town_name']
        plot_colours = ['cumulative', 'gross_load', 'load', 'shortfall']
        resource_colours = ['dhi_high', 'dhi_low', 'dni_high', 'dni_low', 'ghi_high', 'ghi_low',
                            'rain_high', 'rain_low', 'temp_high', 'temp_low', 'wind50_high',
                            'wind50_low', 'wind_high', 'wind_low']
        colour_groups = {}
        for group in groups:
            colour_groups[group] = []
        try:
            technologies = config.get('Power', 'technologies')
            technologies = technologies.split()
        except:
            technologies = []
        try:
            fossil_technologies = config.get('Power', 'fossil_technologies')
            fossil_technologies = fossil_technologies.split()
        except:
            fossil_technologies = []
        self.map = ''
        if self.section == 'Colors':
            try:
                self.map = config.get('Map', 'map_choice')
            except:
                pass
        self.colours = {}
        try:
            colours0 = config.items(self.section)
            for it, col in colours0:
                if it == 'ruler':
                    continue
                self.colours[it] = ['', col]
        except:
            pass
        self.default_col = 1
        if self.map != '':
            self.default_col = 2
            try:
                colours1 = config.items(self.section + self.map)
                for it, col in colours1:
                    if it == 'ruler':
                        continue
                    if it in self.colours:
                        self.colours[it] = [col, self.colours[it][1]]
                    else:
                        self.colours[it] = [col, '']
            except:
                pass
        if palette is not None and len(palette) > 0: # set palette of colours
            col = ['', '']
            col[0] = QtWidgets.QColorDialog.getColor(QtCore.Qt.white, None, 'Select colour for item 1')
            if len(palette) > 1:
                col[1] = QtWidgets.QColorDialog.getColor(QtCore.Qt.white, None,
                         'Select colour for item ' + str(len(palette)))
                inc = []
                for c in range(3):
                    inc.append((col[1].getRgb()[c] - col[0].getRgb()[c]) / (len(palette) - 1))
                for i in range(len(palette)):
                    colr = []
                    for c in range(3):
                        colr.append(int(col[0].getRgb()[c] + inc[c] * i))
                    QtGui.QColor.setRgb(col[1], colr[0], colr[1], colr[2])
                    if self.underscore:
                        self.colours[palette[i].lower()] = ['', col[1].name()]
                    else:
                        self.colours[palette[i].lower().replace(' ', '_')] = ['', col[1].name()]
            else:
                if self.underscore:
                    self.colours[palette[0].lower()] = ['', col[0].name()]
                else:
                    self.colours[palette[0].lower().replace(' ', '_')] = ['', col[0].name()]
        group_colours = False
        try:
            gc = config.get('View', 'group_colours')
            if gc.lower() in ['true', 'on', 'yes']:
                group_colours = True
        except:
            pass
        self.old_ruler = ''
        for key, value in self.colours.items():
            if key in fossil_technologies:
                colour_groups['Fossil Technologies'].append(key)
            elif key.find('grid') >= 0:
                colour_groups['Grid'].append(key)
            elif key in map_colours:
                colour_groups['Map'].append(key)
            elif key in plot_colours:
                colour_groups['Plot'].append(key)
            elif key in resource_colours:
                colour_groups['Resource'].append(key)
            elif key in technologies:
                colour_groups['Technologies'].append(key)
            else:
                colour_groups['The Rest'].append(key)
            for i in range(2):
                if value[i] != '':
                    value[i] = QtGui.QColor(value[i])
            self.colours[key] = value
        self.width = [0, 0]
        self.height = [0, 0]
        self.item = []
        self.btn = []
        self.grid = QtWidgets.QGridLayout()
        self.grid.addWidget(QtWidgets.QLabel('Item'), 0, 0)
        if self.map != '':
            self.grid.addWidget(QtWidgets.QLabel('Colour for ' + self.map), 0, 1)
        self.grid.addWidget(QtWidgets.QLabel('Default Colour'), 0, self.default_col)
        i = 1
        if group_colours:
            bold = QtGui.QFont()
            bold.setBold(True)
            for gkey, gvalue in iter(sorted(colour_groups.items())):
                label = QtWidgets.QLabel(gkey)
                label.setFont(bold)
                self.grid.addWidget(label, i, 0)
                i += 1
                for key in sorted(gvalue):
                    value = self.colours[key]
                    self.add_item(key, value, i)
                    i += 1
        else:
            for key, value in iter(sorted(self.colours.items())):
                self.add_item(key, value, i)
                i += 1
        if add_colour:
            if self.underscore:
                key = add_colour.lower()
            else:
                key = add_colour.lower().replace(' ', '_')
            self.colours[key] = ['', '']
            self.add_item(key, ['', ''], -1)
            self.showDialog(colour=key)
        buttonLayout = QtWidgets.QHBoxLayout()
        quit = QtWidgets.QPushButton('Quit', self)
        quit.setMaximumWidth(70)
        buttonLayout.addWidget(quit)
        quit.clicked.connect(self.quitClicked)
        save = QtWidgets.QPushButton('Save && Exit', self)
        buttonLayout.addWidget(save)
        save.clicked.connect(self.saveClicked)
        if self.section != 'Colors':
            add = QtWidgets.QPushButton('Add', self)
            buttonLayout.addWidget(add)
            add.clicked.connect(self.addClicked)
        buttons = QtWidgets.QFrame()
        buttons.setLayout(buttonLayout)
        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(self.grid)
        layout.addWidget(buttons)
        self.setLayout(layout)
        self.grid.setColumnMinimumWidth(0, self.width[0])
        self.grid.setColumnMinimumWidth(1, self.width[0])
        if self.map != '':
            self.grid.setColumnMinimumWidth(2, self.width[0])
        frame = QtWidgets.QFrame()
        frame.setLayout(layout)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        screen = QtWidgets.QDesktopWidget().availableGeometry()
       # if self.height[0] == 0:
        metrics = quit.fontMetrics()
        self.height[0] = metrics.boundingRect(quit.text()).height() + self.grid.verticalSpacing() * 3
        height = (self.height[0] + self.grid.verticalSpacing() * 2) * self.grid.rowCount() + buttons.sizeHint().height()
        if height > int(screen.height() * .9):
            self.height[1] = int(screen.height() * .9)
        else:
            self.height[1] = height
        if self.width[0] == 0:
            self.width[1] = buttons.sizeHint().width() + 80
        else:
            self.width[1] = self.width[0] * 3 + 80
        self.resize(self.width[1], self.height[1])
     #   self.resize(640, 480)
        self.setWindowTitle('SIREN - Color dialog')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)

    def add_item(self, key, value, i):
        if i < 0:
            i = self.grid.rowCount() # could possibly use this always just not sure about colour groups
            screen = QtWidgets.QDesktopWidget().availableGeometry()
            if self.height[1] + self.height[0] < int(screen.height() * .9):
                width = max(self.width[1], self.frameGeometry().width())
                self.height[1] += self.height[0]
                self.resize(width, self.height[1])
        wht = techClean(key, full=True)
        self.grid.addWidget(QtWidgets.QLabel(wht), i, 0)
        if self.map != '':
            self.btn.append(QtWidgets.QPushButton(key, self))
            self.btn[-1].clicked.connect(self.showDialog)
            if value[0] != '':
                self.btn[-1].setStyleSheet('QPushButton {background-color: %s; color: %s;}' %
                             (value[0].name(), value[0].name()))
            self.grid.addWidget(self.btn[-1], i, 1)
            metrics = self.btn[-1].fontMetrics()
            if metrics.boundingRect(self.btn[-1].text()).width() > self.width[0]:
                self.width[0] = metrics.boundingRect(self.btn[-1].text()).width()
            if metrics.boundingRect(self.btn[-1].text()).height() > self.height[0]:
                self.height[0] = metrics.boundingRect(self.btn[-1].text()).height() + self.grid.verticalSpacing() * 2
        self.btn.append(QtWidgets.QPushButton(key + '_base', self))
        metrics = self.btn[-1].fontMetrics()
        if i < 2:
            self.default_style = self.btn[-1].styleSheet()
        self.btn[-1].clicked.connect(self.showDialog)
        if value[1] != '':
            self.btn[-1].setStyleSheet('QPushButton {background-color: %s; color: %s;}' %
                         (value[1].name(), value[1].name()))
        self.grid.addWidget(self.btn[-1], i, self.default_col)
        if metrics.boundingRect(self.btn[-1].text()).width() > self.width[0]:
            self.width[0] = metrics.boundingRect(self.btn[-1].text()).width()
        if metrics.boundingRect(self.btn[-1].text()).height() > self.height[0]:
            self.height[0] = metrics.boundingRect(self.btn[-1].text()).height() + self.grid.verticalSpacing() * 2

    def mousePressEvent(self, event):
        if QtCore.Qt.RightButton == event.button():
            for btn in self.btn:
                if btn.hasFocus():
                    if btn.text()[-5:] != '_base':
                        key = btn.text()
                        if self.colours[key][0] != '':
                            self.colours[key] = ['delete', self.colours[key][1]]
                        btn.setStyleSheet(self.default_style)
                    elif self.section != 'Colors':
                        if btn.text()[-5:] == '_base':
                            key = btn.text()[:-5]
                        else:
                            key = btn.text()
                        self.colours[key] = ['', '']
                        btn.setStyleSheet(self.default_style)
                        break

    def showDialog(self, colour=False):
        sender = self.sender()
        if not colour:
            if sender.text()[-5:] == '_base':
                key = sender.text()[:-5]
                ndx = 1
            else:
                key = sender.text()
                ndx = 0
        else:
            key = colour
            ndx = 1
        if self.colours[key][ndx] != '' and self.colours[key][ndx] != 'delete':
            col = QtWidgets.QColorDialog.getColor(self.colours[key][ndx], None, 'Select colour for ' + key.title())
        else:
            col = QtWidgets.QColorDialog.getColor(QtGui.QColor('white'), None, 'Select colour for ' + key.title())
        if col.isValid():
            if not colour:
                if ndx == 0:
                    self.colours[key] = [col, self.colours[key][1]]
                else:
                    self.colours[key] = [self.colours[key][0], col]
                for i in range(len(self.btn)):
                    if self.btn[i] == sender:
                        self.btn[i].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % (col.name(), col.name()))
                        break
            else:
                for i in range(len(self.btn) -1, -1, -1):
                    self.colours[key] = [self.colours[key][0], col]
                    if self.btn[i].text() == key + '_base':
                        self.btn[i].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % (col.name(), col.name()))
                        break

    def quitClicked(self):
        self.close()

    def addClicked(self):
        text, ok = QtWidgets.QInputDialog.getText(self, 'Add Colour Item', 'Enter name for colour item:')
        if ok:
            if self.underscore:
                key = text.lower()
            else:
                key = text.lower().replace(' ', '_')
            self.colours[key] = ['', '']
            self.add_item(key, ['', ''], -1)

    def saveClicked(self):
        updates = {}
        lines = [[], []]
        for key, value in self.colours.items():
            for i in range(2):
                if value[i] == 'delete':
                    lines[i].append(key + '=')
                elif value[i] != '':
                    lines[i].append(key + '=' + value[i].name())
                elif self.section != 'Colors':
                    lines[i].append(key + '=')
        if len(lines[0]) > 0:
            updates[self.section + self.map] = lines[0]
        updates[self.section] = lines[1]
        SaveIni(updates, ini_file=self.config_file)
        self.close()

    def getValues(self):
        return self.anobject
