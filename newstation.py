#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
#
#  newstation.py - This file is part of SIREN.
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

import csv
import os
import sys

import ConfigParser  # decode .ini file
from PyQt4 import QtGui, QtCore

from senuser import getUser
from station import Station
from turbine import Turbine


class AnObject(QtGui.QDialog):

    def get_config(self):
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        try:
            self.base_year = config.get('Base', 'year')
        except:
            self.base_year = '2012'
        parents = []
        try:
            parents = config.items('Parents')
        except:
            pass
        try:
            self.sam_file = config.get('Files', 'sam_turbines')
            for key, value in parents:
                self.sam_file = self.sam_file.replace(key, value)
            self.sam_file = self.sam_file.replace('$USER$', getUser())
            self.sam_file = self.sam_file.replace('$YEAR$', self.base_year)
        except:
            self.sam_file = ''
        try:
            self.pow_dir = config.get('Files', 'pow_files')
            for key, value in parents:
                self.pow_dir = self.pow_dir.replace(key, value)
            self.pow_dir = self.pow_dir.replace('$USER$', getUser())
            self.pow_dir = self.pow_dir.replace('$YEAR$', self.base_year)
        except:
            self.pow_dir = ''
        try:
            self.map = config.get('Map', 'map_choice')
        except:
            self.map = ''
        self.upper_left = [0., 0.]
        self.lower_right = [-90., 180.]
        try:
             upper_left = config.get('Map', 'upper_left' + self.map).split(',')
             self.upper_left[0] = float(upper_left[0].strip())
             self.upper_left[1] = float(upper_left[1].strip())
             lower_right = config.get('Map', 'lower_right' + self.map).split(',')
             self.lower_right[0] = float(lower_right[0].strip())
             self.lower_right[1] = float(lower_right[1].strip())
        except:
             try:
                 lower_left = config.get('Map', 'lower_left' + self.map).split(',')
                 upper_right = config.get('Map', 'upper_right' + self.map).split(',')
                 self.upper_left[0] = float(upper_right[0].strip())
                 self.upper_left[1] = float(lower_left[1].strip())
                 self.lower_right[0] = float(lower_left[0].strip())
                 self.lower_right[1] = float(upper_right[1].strip())
             except:
                 pass
        self.technologies = ['']
        self.areas = {}
        try:
            technologies = config.get('Power', 'technologies')
            for item in technologies.split(' '):
                itm = item.replace('_', ' ').title()
                itm = itm.replace('Pv', 'PV')
                self.technologies.append(itm)
                try:
                    self.areas[itm] = float(config.get(itm, 'area'))
                except:
                    self.areas[itm] = 0.
        except:
            pass
        self.tshours = 0
        try:
            self.tshours = float(config.get('Solar Thermal', 'tshours'))
        except:
            pass

    def __init__(self, dialog, anobject, scenarios=None):
        super(AnObject, self).__init__()
        self.get_config()
        self.anobject = anobject
        if scenarios is not None:
            self.scenarios = []
            for i in range(len(scenarios)):
                if scenarios[i][0] != 'Existing':
                    self.scenarios.append(scenarios[i][0])
            if len(self.scenarios) < 2:
                self.scenarios = None
        else:
            self.scenarios = None
        dialog.setObjectName('Dialog')
        self.initUI()

    def initUI(self):
        self.save = False
        self.field = ['name', 'technology', 'lat', 'lon', 'capacity', 'turbine', 'rotor',
                  'no_turbines', 'area', 'scenario', 'power_file', 'grid_line', 'storage_hours', 'direction']
        self.label = []
        self.edit = []
        self.field_type = []
        metrics = []
        widths = [0, 0]
        heights = 0
        self.turbines = [['', '', 0.]]
       #  self.turbine_class = ['']
        grid = QtGui.QGridLayout()
        turbcombo = QtGui.QComboBox(self)
        sam = open(self.sam_file)
        sam_turbines = csv.DictReader(sam)
        for turb in sam_turbines:
            if turb['Name'] == 'Units' or turb['Name'] == '[0]':
                pass
            else:
                self.turbines.append([turb['Name'], '', turb['Rotor Diameter']])
                if turb['IEC Wind Speed Class'] in ['', ' ', '0', 'Unknown', 'unknown', 'not listed']:
                   pass
                else:
                   cls = turb['IEC Wind Speed Class'].replace('|', ', ')
                   cls.replace('Class ', '')
                   self.turbines[-1][1] = 'Class: ' + cls
        sam.close()
        pow_files = os.listdir(self.pow_dir)
        for pow in pow_files:
            if pow[-4:] == '.pow':
                turb = Turbine(pow[:-4])
                self.turbines.append([pow[:-4], '', str(turb.rotor)])
              #   self.turbine_class.append('')
        self.turbines.sort()
        j = -1
        for i in range(len(self.turbines)):
            if self.turbines[i][0] == 'Vestas V90-2.0':
                j = i
                break
        for i in range(len(self.turbines)):
            turbcombo.addItem(self.turbines[i][0])
            if self.turbines[i][0] == self.anobject.turbine:
                turbcombo.setCurrentIndex(i)
                if self.turbines[i][0] != '':
                    j = -1
        if j >= 0:
            turbcombo.setCurrentIndex(j)
        techcombo = QtGui.QComboBox(self)
        for i in range(len(self.technologies)):
            techcombo.addItem(self.technologies[i])
            if self.technologies[i] == self.anobject.technology:
                techcombo.setCurrentIndex(i)
        if self.scenarios is not None:
            scencombo = QtGui.QComboBox(self)
            for i in range(len(self.scenarios)):
                scencombo.addItem(self.scenarios[i])
                if self.scenarios[i] == self.anobject.scenario:
                    scencombo.setCurrentIndex(i)
        for i in range(len(self.field)):
            self.label.append(QtGui.QLabel(self.field[i].title() + ':'))
            if i == 0:
                metrics.append(self.label[-1].fontMetrics())
                if metrics[0].boundingRect(self.label[-1].text()).width() > widths[0]:
                    widths[0] = metrics[0].boundingRect(self.label[-1].text()).width()
            grid.addWidget(self.label[-1], i + 1, 0)
        self.capacity_was = 0.0
        self.no_turbines_was = 0
        self.turbine_was = ''
        self.turbine_classd = QtGui.QLabel('')
        self.show_hide = {}
        units = {'area': 'sq. Km', 'capacity': 'MW'}
        for i in range(len(self.field)):
            attr = getattr(self.anobject, self.field[i])
            if isinstance(attr, int):
                self.field_type.append('int')
            elif isinstance(attr, float):
                self.field_type.append('float')
            else:
                self.field_type.append('str')
            if self.field[i] == 'name':
                self.name = attr
                self.edit.append(QtGui.QLineEdit(self.name))
                metrics.append(self.label[-1].fontMetrics())
            elif self.field[i] == 'technology':
                self.techcomb = techcombo
                self.edit.append(self.techcomb)
                self.techcomb.currentIndexChanged.connect(self.technologyChanged)
            elif self.field[i] == 'lat':
                self.lat = attr
                self.edit.append(QtGui.QLineEdit(str(self.lat)))
            elif self.field[i] == 'lon':
                self.lon = attr
                self.edit.append(QtGui.QLineEdit(str(self.lon)))
            elif self.field[i] == 'capacity':
                self.capacity = attr
                self.capacity_was = attr
                self.edit.append(QtGui.QDoubleSpinBox())  # QtGui.QLineEdit(str(self.capacity)))
                self.edit[-1].setRange(0, 10000)
                self.edit[-1].setValue(self.capacity)
                self.edit[-1].setDecimals(1)
            elif self.field[i] == 'turbine':
                if self.techcomb.currentText() != 'Wind' and self.techcomb.currentText() != '':
                    turbcombo.setCurrentIndex(0)
                self.turbine = turbcombo
                self.turbines_was = turbcombo
                self.show_hide['turbine'] = len(self.edit)
                self.edit.append(self.turbine)
                self.turbine_classd.setText(self.turbines[turbcombo.currentIndex()][1])
                self.turbine.currentIndexChanged.connect(self.turbineChanged)
                grid.addWidget(self.turbine_classd, i + 1, 2)
            elif self.field[i] == 'rotor':
                self.rotor = attr
                self.show_hide['rotor'] = len(self.edit)
                self.edit.append(QtGui.QLineEdit(str(self.rotor)))
                self.edit[-1].setEnabled(False)
                self.curve = QtGui.QPushButton('Show Power Curve', self)
                grid.addWidget(self.curve, i + 1, 2)
                self.curve.clicked.connect(self.curveClicked)
            elif self.field[i] == 'no_turbines':
                self.no_turbines = attr
                self.no_turbines_was = attr
                self.show_hide['no_turbines'] = len(self.edit)
                self.edit.append(QtGui.QSpinBox())  # QtGui.QLineEdit(str(self.no_turbines)))
                self.edit[-1].setRange(0, 299)
                if self.no_turbines == '':
                    self.edit[-1].setValue(0)
                else:
                    self.edit[-1].setValue(int(self.no_turbines))
            elif self.field[i] == 'area':
                self.area = attr
                self.edit.append(QtGui.QLineEdit(str(self.area)))
                self.edit[-1].setEnabled(False)
            elif self.field[i] == 'scenario':
                self.scenario = attr
                if self.scenarios is None:
                    self.edit.append(QtGui.QLineEdit(self.scenario))
                    self.edit[-1].setEnabled(False)
                else:
                    self.scencomb = scencombo
                    self.scencomb.currentIndexChanged.connect(self.scenarioChanged)
                    self.edit.append(self.scencomb)
            elif self.field[i] == 'power_file':
                self.power_file = attr
                self.edit.append(QtGui.QLineEdit(self.power_file))
            elif self.field[i] == 'grid_line':
                self.grid_line = attr
                if attr is not None:
                    self.edit.append(QtGui.QLineEdit(self.grid_line))
                else:
                    self.edit.append(QtGui.QLineEdit(''))
            elif self.field[i] == 'storage_hours':
                self.storage_hours = attr
                self.show_hide['storage_hours'] = len(self.edit)
                if attr is not None:
                    self.edit.append(QtGui.QLineEdit(str(self.storage_hours)))
                else:
                    self.edit.append(QtGui.QLineEdit(str(self.tshours)))
            elif self.field[i] == 'direction':
                self.direction = attr
                self.show_hide['direction'] = len(self.edit)
                if attr is not None:
                    self.edit.append(QtGui.QLineEdit(str(self.direction)))
                else:
                    self.edit.append(QtGui.QLineEdit(''))
            try:
                if metrics[1].boundingRect(self.edit[-1].text()).width() > widths[1]:
                    widths[1] = metrics[1].boundingRect(self.edit[-1].text()).width()
            except:
                pass
            grid.addWidget(self.edit[-1], i + 1, 1)
            if self.field[i] in units.keys():
                grid.addWidget(QtGui.QLabel(units[self.field[i]]), i + 1, 2)
        self.technologyChanged(self.techcomb.currentIndex)
        grid.setColumnMinimumWidth(0, widths[0] + 10)
        grid.setColumnMinimumWidth(1, widths[1] + 10)
        grid.setColumnMinimumWidth(2, 30)
        i += 1
        self.message = QtGui.QLabel('')
        msg_font = self.message.font()
        msg_font.setBold(True)
        self.message.setFont(msg_font)
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.message.setPalette(msg_palette)
        grid.addWidget(self.message, i + 1, 0, 1, 2)
        i += 1
        quit = QtGui.QPushButton("Quit", self)
        grid.addWidget(quit, i + 1, 0)
        quit.clicked.connect(self.quitClicked)
        save = QtGui.QPushButton("Save && Exit", self)
        grid.addWidget(save, i + 1, 1)
        save.clicked.connect(self.saveClicked)
        self.setLayout(grid)
        self.resize(widths[0] + widths[1] + 40, heights * i)
        self.setWindowTitle('SIREN - Edit ' + getattr(self.anobject, '__module__'))
        QtGui.QShortcut(QtGui.QKeySequence("q"), self, self.quitClicked)
       #  self.connect(self, SIGNAL('status_text'), self.setStatusText)
        # self.emit(SIGNAL('status_text'), 'Well Hello!')

    def technologyChanged(self, val):
        wind_fields = ['turbine', 'rotor', 'no_turbines']
        cst_fields = ['storage_hours']
        pv_fields = ['direction']
        show_fields = []
        hide_fields = []
        if str(self.techcomb.currentText()) == 'Wind':
            hide_fields.append(cst_fields)
            hide_fields.append(pv_fields)
            show_fields.append(wind_fields)
            self.curve.show()
            self.turbine_classd.show()
        else:
            self.curve.hide()
            self.turbine_classd.hide()
            if str(self.techcomb.currentText()) == 'Solar Thermal':
                show_fields.append(cst_fields)
                hide_fields.append(pv_fields)
                hide_fields.append(wind_fields)
            elif 'PV' in str(self.techcomb.currentText()):
                hide_fields.append(cst_fields)
                show_fields.append(pv_fields)
                hide_fields.append(wind_fields)
            else:
                hide_fields.append(cst_fields)
                hide_fields.append(pv_fields)
                hide_fields.append(wind_fields)
        for i in range(len(hide_fields)):
            for j in range(len(hide_fields[i])):
                k = self.field.index(hide_fields[i][j])
                self.label[k].hide()
                self.edit[self.show_hide[hide_fields[i][j]]].hide()
                self.edit[self.show_hide[hide_fields[i][j]]].setEnabled(False)
        for i in range(len(show_fields)):
            for j in range(len(show_fields[i])):
                k = self.field.index(show_fields[i][j])
                self.label[k].show()
                self.edit[self.show_hide[show_fields[i][j]]].show()
                if show_fields[i][j] != 'rotor':
                    self.edit[self.show_hide[show_fields[i][j]]].setEnabled(True)
        self.technology = str(self.techcomb.currentText())

    def scenarioChanged(self, val):
        self.scenario = self.scencomb.currentText()

    def turbineChanged(self, val):
        self.turbine_classd.setText(self.turbines[val][1])

    def setStatusText(self, text):
        if text == self.statusBar().currentMessage():
            return
        self.statusBar().clearMessage()
        self.statusBar().showMessage(text)

    def curveClicked(self):
        if str(self.turbine.currentText()) != '':
            Turbine(str(self.turbine.currentText())).PowerCurve()
        return

    def quitClicked(self):
        self.close()

    def saveClicked(self):
        for i in range(len(self.field)):
            try:
                if not self.edit[i].isEnabled():
                    continue
            except:
                continue
            self.edit[i].setFocus()
            if isinstance(self.edit[i], QtGui.QLineEdit):
                if self.field[i] == 'direction' and 'PV' in self.technology:
                    try:
                        setattr(self, self.field[i], int(self.edit[i].text()))
                        if self.direction < 0:
                            self.direction = 360 + self.direction
                        if self.direction > 360:
                            self.message.setText('Error with ' + self.field[i].title() + ' field')
                            return
                    except:
                        setattr(self, self.field[i], str(self.edit[i].text()))
                        if self.direction.upper() in ['', 'N', 'NNE', 'NE', 'ENE', 'E', 'ESE',
                                                      'SE', 'SSE', 'S', 'SSW', 'SW',
                                                      'WSW', 'W', 'WNW', 'NW', 'NNW']:
                            self.direction = self.direction.upper()
                        else:
                            self.message.setText('Error with ' + self.field[i].title() + ' field')
                            return
                elif self.field_type[i] == 'int':
                    try:
                        setattr(self, self.field[i], int(self.edit[i].text()))
                    except:
                        self.message.setText('Error with ' + self.field[i].title() + ' field')
                        return
                elif self.field_type[i] == 'float':
                    try:
                        setattr(self, self.field[i], float(self.edit[i].text()))
                        if self.field[i] == 'lon':
                            if self.lon < self.upper_left[1]:
                                self.message.setText('Error with ' + self.field[i].title() + ' field. Too far west')
                                return
                            elif self.lon > self.lower_right[1]:
                                self.message.setText('Error with ' + self.field[i].title() + ' field. Too far east')
                                return
                        elif self.field[i] == 'lat':
                            if self.lat > self.upper_left[0]:
                                self.message.setText('Error with ' + self.field[i].title() + ' field. Too far north')
                                return
                            elif self.lat < self.lower_right[0]:
                                self.message.setText('Error with ' + self.field[i].title() + ' field. Too far south')
                                return
                    except:
                        self.message.setText('Error with ' + self.field[i].title() + ' field')
                        return
                elif self.field[i] == 'grid_line':
                    if self.edit[i].text() == '':
                        setattr(self, self.field[i], None)
                    else:
                        setattr(self, self.field[i], self.edit[i].text())
                else:
                    setattr(self, self.field[i], self.edit[i].text())
            elif isinstance(self.edit[i], QtGui.QComboBox):
                setattr(self, self.field[i], str(self.edit[i].currentText()))
                if self.field[i] == 'technology' and str(self.edit[i].currentText()) == '':
                    self.message.setText('Error with ' + self.field[i].title() + '. Choose technology')
                    return
            elif isinstance(self.edit[i], QtGui.QDoubleSpinBox):
                setattr(self, self.field[i], self.edit[i].value())
            elif isinstance(self.edit[i], QtGui.QSpinBox):
                setattr(self, self.field[i], self.edit[i].value())
        if self.technology != 'Wind':
            self.rotor = 0.0
            self.turbine == ''
        if self.technology == 'Biomass':
            self.area = self.areas[self.technology] * float(self.capacity)
        elif 'PV' in self.technology:
            self.area = self.areas[self.technology] * float(self.capacity)
        elif self.technology == 'Wind':
            if self.turbine == '':
                self.edit[5].setFocus()
                self.message.setText('Error with ' + self.field[5].title() + '. Choose turbine')
                return
            turbine = Turbine(self.turbine)
            if self.capacity != self.capacity_was or \
              (self.turbine != self.turbine_was and self.no_turbines == self.no_turbines_was):
                self.no_turbines = int(round((self.capacity * 1000.) / turbine.capacity))
            self.capacity = self.no_turbines * turbine.capacity / 1000.  # reduce from kW to MW
            self.rotor = turbine.rotor
            self.area = self.areas[self.technology] * float(self.no_turbines) * pow((self.rotor * .001), 2)
        elif self.technology == 'Solar Thermal':
            self.area = self.areas[self.technology] * float(self.capacity)  # temp calc. Should be 3.83 x collector area
        elif self.technology == 'Geothermal':
            self.area = self.areas[self.technology] * float(self.capacity)
        elif self.technology == 'Wave':
            self.area = self.areas[self.technology] * float(self.capacity)
        elif self.technology == 'Hydro':
            self.area = self.areas[self.technology] * float(self.capacity)
        elif self.technology[:5] == 'Other':
            self.area = self.areas[self.technology] * float(self.capacity)
        else:
            self.message.setText('This technology not yet implemented. Choose another.')
            self.edit[1].setFocus()
            return
        self.save = True
        self.close()

    def getValues(self):
        if self.save:
            station = Station(str(self.name), self.technology, self.lat, self.lon, self.capacity, self.turbine, self.rotor,
                      self.no_turbines, self.area, self.scenario, power_file=self.power_file)
            if self.grid_line is not None:
                station.grid_line = self.grid_line
            if self.direction is not None and 'PV' in self.technology:
                station.direction = self.direction
            if self.storage_hours is not None and self.technology == 'Solar Thermal':
                try:
                    if self.storage_hours != self.tshours:
                        station.storage_hours = float(self.storage_hours)
                except:
                    pass
            return station
        else:
            return None
