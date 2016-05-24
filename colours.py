#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
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
from PyQt4 import QtGui
from PyQt4.QtCore import Qt
import ConfigParser  # decode .ini file


class Colours(QtGui.QDialog):

    def __init__(self):
        super(Colours, self).__init__()
        self.initUI()

    def initUI(self):
        def add_item(key, value, i):
            wht = key.replace('_', ' ').title()
            wht = wht.replace('Pv', 'PV')
            wht = wht.replace('hi ', 'HI ')
            wht = wht.replace('ni ', 'NI ')
            self.grid.addWidget(QtGui.QLabel(wht), i, 0)
            if self.map != '':
                self.btn.append(QtGui.QPushButton(key, self))
                self.btn[-1].clicked.connect(self.showDialog)
                if value[0] != '':
                    self.btn[-1].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % \
                                 (value[0].name(), value[0].name()))
                self.grid.addWidget(self.btn[-1], i, 1)
                metrics = self.btn[-1].fontMetrics()
                if metrics.boundingRect(self.btn[-1].text()).width() > self.width:
                    self.width = metrics.boundingRect(self.btn[-1].text()).width()
            self.btn.append(QtGui.QPushButton(key + '_base', self))
            metrics = self.btn[-1].fontMetrics()
            if i < 2:
                self.default_style = self.btn[-1].styleSheet()
                height = metrics.boundingRect(self.btn[-1].text()).height()
            self.btn[-1].clicked.connect(self.showDialog)
            if value[1] != '':
                self.btn[-1].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % \
                             (value[1].name(), value[1].name()))
            self.grid.addWidget(self.btn[-1], i, default_col)
            if metrics.boundingRect(self.btn[-1].text()).width() > self.width:
                self.width = metrics.boundingRect(self.btn[-1].text()).width()

        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
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
            technologies = technologies.split(' ')
        except:
            technologies = []
        try:
            fossil_technologies = config.get('Power', 'fossil_technologies')
            fossil_technologies = fossil_technologies.split(' ')
        except:
            fossil_technologies = []

        try:
            self.map = config.get('Map', 'map_choice')
        except:
            self.map = ''
        self.colours = {}
        try:
            colours0 = config.items('Colors')
            for it, col in colours0:
                if it == 'ruler':
                    continue
                self.colours[it] = ['', col]
        except:
            pass
        default_col = 1
        if self.map != '':
            default_col = 2
            try:
                colours1 = config.items('Colors' + self.map)
                for it, col in colours1:
                    if it == 'ruler':
                        continue
                    if it in self.colours:
                        self.colours[it] = [col, self.colours[it][1]]
                    else:
                        self.colours[it] = [col, '']
            except:
                pass
        group_colours = False
        try:
            gc = config.get('View', 'group_colours')
            if gc.lower() in ['true', 'on', 'yes']:
                group_colours = True
        except:
            pass
        self.old_ruler = ''
        for key, value in self.colours.iteritems():
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
        self.width = 0
        height = 0
        self.item = []
        self.btn = []
        self.grid = QtGui.QGridLayout()
        self.grid.addWidget(QtGui.QLabel('Item'), 0, 0)
        if self.map != '':
            self.grid.addWidget(QtGui.QLabel('Colour for ' + self.map), 0, 1)
        self.grid.addWidget(QtGui.QLabel('Default Colour'), 0, default_col)
        i = 1
        if group_colours:
            bold = QtGui.QFont()
            bold.setBold(True)
            for gkey, gvalue in iter(sorted(colour_groups.iteritems())):
                label = QtGui.QLabel(gkey)
                label.setFont(bold)
                self.grid.addWidget(label, i, 0)
                i += 1
                for key in sorted(gvalue):
                    value = self.colours[key]
                    add_item(key, value, i)
                    i += 1
        else:
            for key, value in iter(sorted(self.colours.iteritems())):
                add_item(key, value, i)
                i += 1
        quit = QtGui.QPushButton('Quit', self)
        quit.setMaximumWidth(70)
        self.grid.addWidget(quit, i + 1, 0)
        quit.clicked.connect(self.quitClicked)
        save = QtGui.QPushButton('Save && Exit', self)
        self.grid.addWidget(save, i + 1, 1)
        save.clicked.connect(self.saveClicked)
        self.setLayout(self.grid)
        self.grid.setColumnMinimumWidth(0, self.width)
        self.grid.setColumnMinimumWidth(1, self.width)
        if self.map != '':
            self.grid.setColumnMinimumWidth(2, self.width)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        screen = QtGui.QDesktopWidget().availableGeometry()
        self.resize(self.width * 3 + 80, int(screen.height() * .9))
        self.setWindowTitle('SIREN - Color dialog')
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)

    def mousePressEvent(self, event):
        if Qt.RightButton == event.button():
            for btn in self.btn:
                if btn.hasFocus():
                    if btn.text()[-5:] != '_base':
                        key = str(btn.text())
                        self.colours[key] = ['', self.colours[key][1]]
                        btn.setStyleSheet(self.default_style)
                    break

    def showDialog(self):
        sender = self.sender()
        if sender.text()[-5:] == '_base':
            key = str(sender.text())[:-5]
            ndx = 1
        else:
            key = str(sender.text())
            ndx = 0
        if self.colours[key][ndx] != '':
            col = QtGui.QColorDialog.getColor(self.colours[key][ndx])
        else:
            col = QtGui.QColorDialog.getColor(QtGui.QColor(''))
        if col.isValid():
            if ndx == 0:
                self.colours[key] = [col, self.colours[key][1]]
            else:
                self.colours[key] = [self.colours[key][0], col]
            for i in range(len(self.btn)):
                if self.btn[i] == sender:
                    self.btn[i].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % (col.name(), col.name()))
                    break

    def quitClicked(self):
        self.close()

    def saveClicked(self):
        if len(sys.argv) > 1:
            ini_file = sys.argv[1]
        else:
            ini_file = 'SIREN.ini'
        inf = open(ini_file, 'r')
        lines = inf.readlines()
        inf.close()
        in_color = False
        for i in range(len(lines)):
            if lines[i][:8] == '[Colors]':
                in_color = True
            elif in_color:
                if lines[i][0] == '[':
                    in_color = False
                    add_in_here = i
                elif lines[i][0] != ';' and lines[i][0] != '#':
                    bits = lines[i].split('=')

                    if bits[0] in self.colours:
                        lines[i] = bits[0] + '=' + self.colours[bits[0]][1].name() + '\n'
        if self.map != '':
            got_him = False
            del_lines = []
            for i in range(len(lines)):
                if lines[i][:8 + len(self.map)] == '[Colors' + self.map + ']':
                    in_color = True
                    got_him = True
                elif in_color:
                    if lines[i][0] == '[':
                        in_color = False
                        add_in_here = i
                    elif lines[i][0] != ';' and lines[i][0] != '#':
                        bits = lines[i].split('=')
                        if bits[0] in self.colours:
                            if self.colours[bits[0]][0] == '':
                                del_lines.append(i)
                            else:
                                lines[i] = bits[0] + '=' + str(self.colours[bits[0]][0].name()) + '\n'
                            del self.colours[bits[0]]
            more_lines = []
            for key, value in self.colours.iteritems():
                if value[0] != '':
                    more_lines.append(key + '=' + str(value[0].name()) + '\n')
            if len(more_lines) > 0:
                if not got_him:
                    more_lines.insert(0, '[Colors' + self.map + ']\n')
        if os.path.exists(ini_file + '~'):
            os.remove(ini_file + '~')
        os.rename(ini_file, ini_file + '~')
        sou = open(ini_file, 'w')
        for i in range(len(lines)):
            if i in del_lines:
                pass
            else:
                sou.write(lines[i])
            if add_in_here == (i + 1):
                for j in range(len(more_lines)):
                    sou.write(more_lines[j])
        sou.close()
        self.close()

    def getValues(self):
        return self.anobject
