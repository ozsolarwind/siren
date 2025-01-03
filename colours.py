#!/usr/bin/python3
#
#  Copyright (C) 2015-2023 Sustainable Energy Now Inc., Angus King
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
import random
from senutils import techClean

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
        self.cancelled = False
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
            if len(palette) > 1:
                col = ['', '']
                col[0] = QtWidgets.QColorDialog.getColor(QtCore.Qt.white, None, 'Select colour for item 1')
                if not col[0].isValid():
                    self.cancelled = True
                    self.close()
                col[1] = QtWidgets.QColorDialog.getColor(QtCore.Qt.white, None, 'Select colour for item ' + str(len(palette)))
                if not col[1].isValid():
                    self.cancelled = True
                    self.close()
                inc = []
                for c in range(3):
                    inc.append((col[1].getRgb()[c] - col[0].getRgb()[c]) / (len(palette) - 1))
                for i in range(len(palette)):
                    colr = []
                    for c in range(3):
                        colr.append(int(col[0].getRgb()[c] + inc[c] * i))
                    QtGui.QColor.setRgb(col[1], colr[0], colr[1], colr[2])
                    if self.underscore:
                        self.colours[palette[i].lower().replace(' ', '_')] = ['', col[1].name()]
                    else:
                        self.colours[palette[i].lower()] = ['', col[1].name()]
            else:
                if self.underscore:
                    key = palette[0].lower().replace(' ', '_')
                else:
                    key = palette[0].lower()
                if key not in self.colours:
                    self.colours[key] = ['', '']
                col = QtWidgets.QColorDialog.getColor(QtGui.QColor(self.colours[key][1]), None, 'Select colour for ' + key.title())
                if col.isValid():
                    self.colours[key] = ['', col.name()]
                else:
                    self.cancelled = True
                    if not add_colour:
                        self.close()
        if add_colour:
            if self.underscore:
                key = add_colour.lower()
            else:
                key = add_colour.lower().replace(' ', '_')
            if key not in self.colours:
                self.colours[key] = ['', '']
            col = QtWidgets.QColorDialog.getColor(QtGui.QColor(self.colours[key][1]), None, 'Select colour for ' + key.title())
            if col.isValid():
                self.colours[key] = ['', col.name()]
            else:
                self.cancelled = True
                self.close()
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
        layout.addWidget(buttons)
        layout.addLayout(self.grid)
   #     layout.addWidget(buttons)
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
            col = QtWidgets.QColorDialog.getColor(QtGui.QColor(self.colours[key][ndx]), None, 'Select colour for ' + key.title())
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
                key = text.lower().replace(' ', '_')
            else:
                key = text.lower()
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

# create a colour blindness palette
def PlotPalette(keys, palette=15, black=False, lower=True):
    # https://mk.bcgsc.ca/colorblind/palettes.mhtml#15-color-palette-for-colorbliness
    colour_dict = {}
    if palette == 15 or palette == '15':
        colours = [[0,0,0],
                   [0,73,73],
                   [0,146,146],
                   [255,109,182],
                   [255,182,219],
                   [73,0,146],
                   [0,109,219],
                   [182,109,255],
                   [109,182,255],
                   [182,219,255],
                   [146,0,0],
                   [146,73,0],
                   [219,109,0],
                   [36,255,36],
                   [255,255,109]
                  ]
        # colour order to "separate" similar colours
        order = [3, 1, 14, 2, 10, 13, 4, 7, 5, 11, 9, 8, 6, 12]
        if black:
            order.insert(0, 0)
    elif palette == 24 or palette == '24':
        colours = [[0,61,48],
                   [0,87,69],
                   [0,115,92],
                   [0,145,117],
                   [0,175,142],
                   [0,203,167],
                   [0,235,193],
                   [134,255,222],
                   [0,48,111],
                   [0,72,158],
                   [0,95,204],
                   [0,121,250],
                   [0,159,250],
                   [0,194,249],
                   [0,229,248],
                   [124,255,250],
                   [255,213,253],
                   [0,64,2],
                   [95,9,20],
                   [0,90,1],
                   [134,8,28],
                   [0,119,2],
                   [178,7,37],
                   [0,149,3],
                   [222,13,46],
                   [0,180,8],
                   [255,66,53],
                   [0,211,2],
                   [255,135,53],
                   [0,244,7],
                   [255,185,53],
                   [255,226,57]]
        order = []
        for o in range(len(colours)):
            order.append(o)
    elif palette[0].lower() == 'r': #random
        for key in keys:
            r = lambda: random.randint(0,255)
            new_colr = '#%02X%02X%02X' % (r(),r(),r())
            colour_dict[key] = new_colr
        return colour_dict
    else: # later version of 15 palette
        colours = [[104,2,63],
                   [0,129,105],
                   [239,0,150],
                   [0,220,181],
                   [255,207,226],
                   [0,60,134],
                   [148,0,230],
                   [0,159,250],
                   [255,113,253],
                   [124,255,250],
                   [106,2,19],
                   [0,134,7],
                   [246,2,57],
                   [0,227,7],
                   [255,220,61]
              ]
        # colour order to "separate" similar colours
        order = [0, 5, 9, 13, 14, 1, 2, 6, 3, 12, 11, 7, 4, 8, 10]
    o = -1
    for key in keys:
        o += 1
        if o >= len(order):
            o = 0
        colr = '#'
        for j in range(3):
            colr += str(hex(colours[order[o]][j]))[2:].upper().zfill(2)
        if lower:
            colour_dict[key.lower()] = colr
        else:
            colour_dict[key] = colr
    return colour_dict
