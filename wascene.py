#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
#
#  wascene.py - This file is part of SIREN.
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
import datetime
from math import sin, cos, pi, sqrt, degrees, radians, asin, atan2
import os
import sys
import xlrd

import ConfigParser   # decode .ini file
from PyQt4 import QtCore, QtGui
from PyQt4.QtCore import Qt
import mpl_toolkits.basemap.pyproj as pyproj   # Import the pyproj module

from towns import Towns
from grid import Grid, Grid_Boundary, Line
from senuser import getUser
from station import Station, Stations
from dijkstra_4 import Shortest


class WAScene(QtGui.QGraphicsScene):

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
        self.existing = True
        try:
            existing = config.get('Base', 'existing')
            if existing.lower() in ['false', 'no', 'off']:
                self.existing = False
        except:
            pass
        try:
            self.model_name = config.get('Base', 'name')
        except:
            self.model_name = ''
        parents = []
        try:
            parents = config.items('Parents')
        except:
            pass
        try:
            self.map = config.get('Map', 'map_choice')
        except:
            self.map = ''
        try:
            self.map_file = config.get('Files', 'map' + self.map)
            for key, value in parents:
                self.map_file = self.map_file.replace(key, value)
            self.map_file = self.map_file.replace('$USER$', getUser())
            self.map_file = self.map_file.replace('$YEAR$', self.base_year)
        except:
            try:
                self.map_file = config.get('Map', 'map' + self.map)
                for key, value in parents:
                    self.map_file = self.map_file.replace(key, value)
                self.map_file = self.map_file.replace('$USER$', getUser())
                self.map_file = self.map_file.replace('$YEAR$', self.base_year)
            except:
                self.map_file = ''
        try:
            self.projection = config.get('Map', 'projection' + self.map)
        except:
            try:
                self.projection = config.get('Map', 'projection')
            except:
                self.projection = 'EPSG:3857'
        self.map_upper_left = [0., 0.]
        self.map_lower_right = [-90., 180.]
        try:
             upper_left = config.get('Map', 'upper_left' + self.map).split(',')
             self.map_upper_left[0] = float(upper_left[0].strip())
             self.map_upper_left[1] = float(upper_left[1].strip())
             lower_right = config.get('Map', 'lower_right' + self.map).split(',')
             self.map_lower_right[0] = float(lower_right[0].strip())
             self.map_lower_right[1] = float(lower_right[1].strip())
        except:
             try:
                 lower_left = config.get('Map', 'lower_left' + self.map).split(',')
                 upper_right = config.get('Map', 'upper_right' + self.map).split(',')
                 self.map_upper_left[0] = float(upper_right[0].strip())
                 self.map_upper_left[1] = float(lower_left[1].strip())
                 self.map_lower_right[0] = float(lower_left[0].strip())
                 self.map_lower_right[1] = float(upper_right[1].strip())
             except:
                 pass
        self.map_polygon = [[self.map_upper_left[0], self.map_upper_left[1]],
                            [self.map_upper_left[0], self.map_lower_right[1]],
                            [self.map_lower_right[0], self.map_lower_right[1]],
                            [self.map_lower_right[0], self.map_upper_left[1]],
                            [self.map_upper_left[0], self.map_upper_left[1]]]
        self.scale = False
        try:
            scale = config.get('Map', 'scale' + self.map)
        except:
            scale = config.get('Map', 'scale')
        if scale.lower() in ['true', 'yes', 'on']:
            self.scale = True
        try:
            self.scenarios = config.get('Files', 'scenarios')
            for key, value in parents:
                self.scenarios = self.scenarios.replace(key, value)
            self.scenarios = self.scenarios.replace('$USER$', getUser())
            self.scenarios = self.scenarios.replace('$YEAR$', self.base_year)
            i = self.scenarios.rfind('/')
            self.scenarios = self.scenarios[:i + 1]
        except:
            self.scenarios = ''
        try:
            self.scenario = config.get('Files', 'scenario')
            for key, value in parents:
                self.scenario = self.scenario.replace(key, value)
            self.scenario = self.scenario.replace('$USER$', getUser())
            self.scenario = self.scenario.replace('$YEAR$', self.base_year)
        except:
            self.scenario = ''
        try:
            self.resource_grid = config.get('Files', 'resource_grid')
            for key, value in parents:
                self.resource_grid = self.resource_grid.replace(key, value)
            self.resource_grid = self.resource_grid.replace('$USER$', getUser())
        except:
            self.resource_grid = ''
        self.colors = {}
        self.colors['background'] = 'darkBlue'
        self.colors['border'] = ''
        self.colors['grid_boundary'] = 'blue'
        self.colors['grid_trace'] = 'white'
        self.colors['ruler'] = 'white'
        self.colors['station'] = '#00FF00'
        self.colors['station_name'] = 'white'
        self.colors['town'] = 'red'
        self.colors['town_name'] = 'lightGray'
        technologies = config.get('Power', 'technologies')
        technologies = technologies.split(' ')
        try:
            colours = config.items('Colors')
            for item, colour in colours:
                if item in technologies or (item[:6] == 'fossil' and item != 'fossil_name'):
                    itm = item.replace('_', ' ').title()
                    itm = itm.replace('Pv', 'PV')
                    self.colors[itm] = colour
                else:
                    self.colors[item] = colour
        except:
            pass
        self.areas = {}
        try:
            for item in technologies:
                itm = item.replace('_', ' ').title()
                itm = itm.replace('Pv', 'PV')
                try:
                    self.areas[itm] = float(config.get(itm, 'area'))
                except:
                    self.areas[itm] = 0.
        except:
            pass
        try:
            technologies = config.get('Power', 'fossil_technologies')
            technologies = technologies.split(' ')
            for item in technologies:
                itm = item.replace('_', ' ').title()
                itm = itm.replace('Pv', 'PV')
                try:
                    self.areas[itm] = float(config.get(itm, 'area'))
                except:
                    self.areas[itm] = 0.
        except:
            pass
        if self.map != '':
            try:
                colours = config.items('Colors' + self.map)
                for item, colour in colours:
                    if item in technologies or (item[:6] == 'fossil' and item != 'fossil_name'):
                        itm = item.replace('_', ' ').title()
                        itm = itm.replace('Pv', 'PV')
                        self.colors[itm] = colour
                    else:
                        self.colors[item] = colour
            except:
                pass
        self.show_generation = False
        self.show_capacity = False
        try:
            show_capacity = config.get('View', 'capacity')
            if show_capacity.lower() in ['true', 'yes', 'on']:
                self.show_capacity = True
        except:
            pass
        self.show_capacity_fill = False
        self.capacity_opacity = 1.
        try:
            show_capacity_fill = config.get('View', 'capacity_fill')
            if show_capacity_fill.lower() in ['true', 'yes', 'on']:
                self.show_capacity_fill = True
            else:
                try:
                    self.capacity_opacity = float(show_capacity_fill)
                    self.show_capacity_fill = True
                    if self.capacity_opacity < 0. or self.capacity_opacity > 1.:
                        self.capacity_opacity = 1.
                except:
                    self.colors['border'] = ''
        except:
            pass
        try:
            capacity_area = config.get('View', 'capacity_area')
            self.capacity_area = float(capacity_area)
        except:
            self.capacity_area = 10.
        self.show_fossil = False
        try:
            show_fossil = config.get('View', 'fossil')
            if show_fossil.lower() in ['true', 'yes', 'on']:
                self.show_fossil = True
        except:
            pass
        self.show_station_name = True
        try:
            show_station_name = config.get('View', 'station_name')
            if show_station_name.lower() in ['false', 'no', 'off']:
                self.show_station_name = False
        except:
            pass
        self.show_legend = False
        self.show_ruler = False
        try:
            show_ruler = config.get('View', 'show_ruler')
            if show_ruler in ['true', 'yes', 'on']:
                self.show_ruler = True
        except:
            pass
        self.ruler = 100.
        self.ruler_ticks = 10.
        try:
            ruler = config.get('View', 'ruler')
            rule = ruler.split(',')
            self.ruler = float(rule[0])
            if len(rule) > 1:
                self.ruler_ticks = float(rule[1].strip())
        except:
            pass
        self.show_towns = True
        try:
            show_towns = config.get('View', 'show_towns')
            if show_towns.lower() in ['false', 'no', 'off']:
                self.show_towns = False
        except:
            pass
        self.center_on_click = False
        try:
            center_on_click = config.get('View', 'center_on_click')
            if center_on_click in ['true', 'yes', 'on']:
                self.center_on_click = True
        except:
            pass
        self.new_grid = False
        try:
            new_grid = config.get('View', 'new_grid')
            if new_grid.lower() in ['true', 'yes', 'on']:
                self.new_grid = True
        except:
            try:
                new_grid = config.get('Grid', 'new_grid')
                if new_grid.lower() in ['true', 'yes', 'on']:
                   self.new_grid = True
            except:
                pass
            pass
        self.existing_grid = True
        try:
            existing_grid = config.get('View', 'existing_grid')
            if existing_grid.lower() in ['false', 'no', 'off']:
                self.existing_grid = False
        except:
            try:
                existing_grid = config.get('Grid', 'existing_grid')
                if existing_grid.lower() in ['false', 'no', 'off']:
                   self.existing_grid = False
            except:
                pass
            pass
        try:
            grid = config.get('Files', 'grid_network2')
            self.existing_grid2 = True
        except:
            self.existing_grid2 = False
        self.trace_grid = True
        try:
            trace_grid = config.get('View', 'trace_grid')
            if trace_grid.lower() in ['false', 'no', 'off']:
                self.trace_grid = False
            #     self.existing_grid = False
        except:
            try:
                trace_grid = config.get('Grid', 'trace_grid')
                if trace_grid.lower() in ['false', 'no', 'off']:
                   self.trace_grid = False
                #    self.existing_grid = False
            except:
                pass
            pass
        self.cost_existing = False
        try:
            cost_existing = config.get('Grid', 'cost_existing')
            if cost_existing.lower() in ['true', 'yes', 'on']:
                self.cost_existing = True
        except:
            pass
        self.trace_existing = False
        try:
            trace_existing = config.get('View', 'trace_existing')
            if trace_existing.lower() in ['true', 'yes', 'on']:
                self.trace_existing = True
        except:
            try:
                trace_existing = config.get('Grid', 'trace_existing')
                if trace_existing.lower() in ['true', 'yes', 'on']:
                   self.trace_existing = True
            except:
                pass
        self.load_centre = None
        try:
            load_centre = config.get('Grid', 'load_centre')
            load_centre = load_centre.replace(' ', '')
            load_centre = load_centre.replace('(', '')
            load_centre = load_centre.replace('))', '')
            load_centre = load_centre.replace('),', ',')
            load_centre = load_centre.replace(')', '')
            load_centre = load_centre.split(',')
            try:
                float(load_centre[0])
                self.load_centre = [['1', float(load_centre[0]), float(load_centre[1])]]
                for j in range(2, len(load_centre), 2):
                    self.load_centre.append([str(j), float(load_centre[j]), float(load_centre[j + 1])])
            except:
                self.load_centre = [[load_centre[0], float(load_centre[1]), float(load_centre[2])]]
                for j in range(3, len(load_centre), 3):
                    self.load_centre.append([load_centre[j], float(load_centre[j + 1]),
                        float(load_centre[j + 2])])
        except:
            pass
        self.station_square = False
        try:
            square = config.get('View', 'station_shape')
            if square[0].lower() == 's':
                self.station_square = True
        except:
            pass
        self.dispatchable = None
        self.line_loss = 0.
       #  self.subs_cost=0.
      #   self.subs_loss=0.
        try:
            itm = config.get('Grid', 'dispatchable')
            itm = itm.replace('_', ' ').title()
            self.dispatchable = itm.replace('Pv', 'PV')
            line_loss = config.get('Grid', 'line_loss')
            if line_loss[-1] == '%':
                self.line_loss = float(line_loss[:-1]) / 100000.
            else:
                self.line_loss = float(line_loss) / 1000.
       #      line_cost = config.get('Grid', 'substation_cost')
        #     if line_cost[-1] == 'K':
         #        self.subs_cost = float(line_cost[:-1]) * pow(10, 3)
          #   elif line_cost[-1] == 'M':
           #      self.subs_cost = float(line_cost[:-1]) * pow(10, 6)
            # else:
             #    self.subs_cost = float(line_cost)
        #     line_loss = config.get('Grid', 'substation_loss')
        #     if line_loss[-1] == '%':
         #        self.subs_loss = float(line_loss[:-1]) / 100.
         #    else:
         #        self.subs_loss = float(line_loss)
        except:
            pass
        self.line_width = 0
        try:
            line_width = config.get('View', 'line_width')
            try:
                self.line_width = int(line_width)
            except:
                self.line_width = float(line_width)
        except:
            pass
        self.tshours = 0
        try:
            self.tshours = float(config.get('Solar Thermal', 'tshours'))
        except:
            pass

    def destinationxy(self, lon1, lat1, bearing, distance):
        """
        Given a start point, initial bearing, and distance, calculate
        the destination point and final bearing travelling along a
        (shortest distance) great circle arc
        """
        radius = 6367.   # km is the radius of the Earth

     # convert decimal degrees to radians
        ln1, lt1, baring = map(radians, [lon1, lat1, bearing])

     # "reverse" haversine formula
        lat2 = asin(sin(lt1) * cos(distance / radius) +
                                cos(lt1) * sin(distance / radius) * cos(baring))
        lon2 = ln1 + atan2(sin(baring) * sin(distance / radius) * cos(lt1),
                                            cos(distance / radius) - sin(lt1) * sin(lat2))
        return degrees(lon2), degrees(lat2)

    def __init__(self):
        QtGui.QGraphicsScene.__init__(self)
        self.get_config()
        self.setBackgroundBrush(QtGui.QColor(self.colors['background']))
        if os.path.exists(self.map_file):
            war = QtGui.QImageReader(self.map_file)
            wa = war.read()
            pixMap = QtGui.QPixmap.fromImage(wa)
        else:
            ratio = abs(self.map_upper_left[1] - self.map_lower_right[1]) / \
                     abs(self.map_upper_left[0] - self.map_lower_right[0])
            pixMap = QtGui.QPixmap(200 * ratio, 200)
            if self.map != '':
                painter = QtGui.QPainter(pixMap)
                brush = QtGui.QBrush(QtGui.QColor('white'))
                painter.fillRect(QtCore.QRectF(0, 0, 200 * ratio, 200), brush)
                painter.setPen(QtGui.QColor('lightgray'))
                painter.drawText(QtCore.QRectF(0, 0, 200 * ratio, 200), QtCore.Qt.AlignCenter,
                                 'Map not found.')
                painter.end()
        self.addPixmap(pixMap)
        w = self.width()
        h = self.height()
        if isinstance(self.line_width, float):
            if self.line_width < 1:
                self.line_width = w * self.line_width
        self.upper_left = [0, 0, self.map_upper_left[1], self.map_upper_left[0]]
        self.lower_right = [w, h, self.map_lower_right[1], self.map_lower_right[0]]
        self.setSceneRect(-w * 0.05, -h * 0.05, w * 1.1, h * 1.1)
        self._positions = {}
        self._setupCoordTransform()
        self._gridGroup = QtGui.QGraphicsItemGroup()
        self._gridGroup2 = QtGui.QGraphicsItemGroup()
        self._setupGrid()
        self.addItem(self._gridGroup)
        if not self.existing_grid:
            self._gridGroup.setVisible(False)
        if self.existing_grid2:
            self.addItem(self._gridGroup2)
        self._townGroup = QtGui.QGraphicsItemGroup()
        self._setupTowns()
        self.addItem(self._townGroup)
        if not self.show_towns:
            self._townGroup.setVisible(False)
        self._scenarios = []
        self._capacityGroup = QtGui.QGraphicsItemGroup()
        self._generationGroup = QtGui.QGraphicsItemGroup()
        self._nameGroup = QtGui.QGraphicsItemGroup()
        self._fossilGroup = QtGui.QGraphicsItemGroup()
        self._fcapacityGroup = QtGui.QGraphicsItemGroup()
        self._fnameGroup = QtGui.QGraphicsItemGroup()
        self._setupStations()
        self.addItem(self._capacityGroup)
        if not self.show_capacity:
            self._capacityGroup.setVisible(False)
            self._fcapacityGroup.setVisible(False)
        self.addItem(self._generationGroup)
        if not self.show_generation:
            self._generationGroup.setVisible(False)
        self.addItem(self._fossilGroup)
        self.addItem(self._fcapacityGroup)
        self.addItem(self._fnameGroup)
        if not self.show_fossil:
            self._fossilGroup.setVisible(False)
            self._fcapacityGroup.setVisible(False)
            self._fnameGroup.setVisible(False)
        self.addItem(self._nameGroup)
        if not self.show_station_name:
            self._nameGroup.setVisible(False)
            self._fnameGroup.setVisible(False)
        self._plot_cache = {}

    def _setupCoordTransform(self):
        self._proj = pyproj.Proj('+init=' + self.projection)   # LatLon with WGS84 datum used by GPS units and Google Earth
        x1, y1, lon1, lat1 = self.upper_left
        x2, y2, lon2, lat2 = self.lower_right
        ul = self._proj(lon1, lat1)
        lr = self._proj(lon2, lat2)
        self._lat_scale = y2 / (lr[1] - ul[1])
        self._lon_scale = x2 / (lr[0] - ul[0])
        self._orig_lat = ul[1]
        self._orig_lon = ul[0]

    def _setupTowns(self):
        self._towns = {}
        self._towns = Towns(ul_lat=self.upper_left[3], ul_lon=self.upper_left[2],
                      lr_lat=self.lower_right[3], lr_lon=self.lower_right[2])
        for st in self._towns.towns:
            p = self.mapFromLonLat(QtCore.QPointF(st.lon, st.lat))
            el = QtGui.QGraphicsEllipseItem(p.x() - 1, p.y() - 1, 2, 2)   # here to adjust town circles
            el.setBrush(QtGui.QColor(self.colors['town']))
            el.setPen(QtGui.QColor(self.colors['town']))
            el.setZValue(0)
            self._townGroup.addToGroup(el)
            txt = QtGui.QGraphicsSimpleTextItem(st.name)
            new_font = txt.font()
            new_font.setPointSizeF(self.width() / 20)
            txt.setFont(new_font)
            txt.setPos(p + QtCore.QPointF(1.5, -0.5))
            txt.scale(0.1, 0.1)
            txt.setBrush(QtGui.QColor(self.colors['town_name']))
            txt.setZValue(0)
            self._townGroup.addToGroup(txt)
        return

    def _setupStations(self):
        self._current_name = QtGui.QGraphicsSimpleTextItem('')
        new_font = self._current_name.font()
        new_font.setPointSizeF(self.width() / 10)
        self._current_name.setFont(new_font)
        self._current_name.scale(0.1, 0.1)
        self._current_name.setZValue(2)
        self.addItem(self._current_name)
        self._stations = {}
        self._station_positions = {}
        self._stationGroups = {}
        self._stationLabels = []
        self._stationCircles = {}
        for key, value in self.colors.iteritems():
            self._stationCircles[key] = []
        if self.existing:
            self._stations = Stations()
            for st in self._stations.stations:
                self.addStation(st)
            self._scenarios.append(['Existing', False, 'Existing stations'])
        else:
            self._stations = Stations(existing=False)
        if self.scenario != '':
            self._setupScenario(self.scenario)

    def _setupScenario(self, scenario):
        i = scenario.rfind('/')
        if i > 0:
            scen_file = scenario
            scen_filter = scenario[i + 1:]
        else:
            scen_file = self.scenarios + scenario
            scen_filter = scenario
        if os.path.exists(scen_file):
            description = ''
            if scen_file[-4:] == '.xls' or scen_file[-5:] == '.xlsx' or \
             scen_file[-5:] == '.xls~' or scen_file[-6:] == '.xlsx~':
                var = {}
                workbook = xlrd.open_workbook(scen_file)
                worksheet = workbook.sheet_by_index(0)
                num_rows = worksheet.nrows - 1
                num_cols = worksheet.ncols - 1
                if worksheet.cell_value(0, 0) == 'Description:' or worksheet.cell_value(0, 0) == 'Comment:':
                    curr_row = 1
                    description = worksheet.cell_value(0, 1)
                else:
                    curr_row = 0
#               get column names
                curr_col = -1
                while curr_col < num_cols:
                    curr_col += 1
                    var[worksheet.cell_value(curr_row, curr_col)] = curr_col
                while curr_row < num_rows:
                    curr_row += 1
                    try:
                        new_st = Station(str(worksheet.cell_value(curr_row, var['Station Name'])),
                                                 str(worksheet.cell_value(curr_row, var['Technology'])),
                                                 worksheet.cell_value(curr_row, var['Latitude']),
                                                 worksheet.cell_value(curr_row, var['Longitude']),
                                                 worksheet.cell_value(curr_row, var['Maximum Capacity (MW)']),
                                                 str(worksheet.cell_value(curr_row, var['Turbine'])),
                                                 worksheet.cell_value(curr_row, var['Rotor Diam']),
                                                 worksheet.cell_value(curr_row, var['No. turbines']),
                                                 worksheet.cell_value(curr_row, var['Area']),
                                                 scen_filter)
                        name_ok = False
                        new_name = new_st.name
                        ctr = 0
                        while not name_ok:
                            for i in range(len(self._stations.stations)):
                                if self._stations.stations[i].name == new_name:
                                    ctr += 1
                                    new_name = new_st.name + ' ' + str(ctr)
                                    break
                            else:
                                name_ok = True
                        if new_name != new_st.name:
                            new_st.name = new_name
                        if new_st.area == 0 or new_st.area == '':
                            if new_st.technology == 'Wind':
                                new_st.area = self.areas[new_st.technology] * float(new_st.no_turbines) * \
                                              pow((new_st.rotor * .001), 2)
                            else:
                                new_st.area = self.areas[new_st.technology] * float(new_st.capacity)
                        try:
                            power_file = worksheet.cell_value(curr_row, var['Power File'])
                            if power_file != '':
                                new_st.power_file = power_file
                        except:
                            pass
                        try:
                            grid_line = worksheet.cell_value(curr_row, var['Grid Line'])
                            if grid_line != '':
                                new_st.grid_line = grid_line
                        except:
                            pass
                        try:
                            direction = worksheet.cell_value(curr_row, var['Direction'])
                            if direction != '':
                                new_st.direction = direction
                        except:
                            pass
                        try:
                            storage_hours = worksheet.cell_value(curr_row, var['Storage Hours'])
                            if storage_hours != '':
                                new_st.storage_hours = storage_hours
                        except:
                            pass
                        self._stations.stations.append(new_st)
                        self.addStation(self._stations.stations[-1])
                    except:
                        break
            else:
                scene = open(scen_file)
                line = scene.readline()
                if len(line) > 13 and line[:13] == 'Description:,':
                        description = line[13:]
                        if description[0] == '"':
                            i = description.rfind('"')
                            description = description[1:i - 1]
                        else:
                            bits = line.split(',')
                            description = bits[1]
                else:
                    scene.seek(0)
                new_stations = csv.DictReader(scene)
                for st in new_stations:
                    new_st = Station(st['Station Name'], st['Technology'], float(st['Latitude']),
                             float(st['Longitude']), float(st['Maximum Capacity (MW)']),
                             st['Turbine'], float(st['Rotor Diam']), int(st['No. turbines']), float(st['Area']), scen_filter)
                    if new_st.area == 0:
                        if new_st.technology == 'Wind':
                            new_st.area = self.areas[new_st.technology] * float(new_st.no_turbines) * \
                                          pow((new_st.rotor * .001), 2)
                        else:
                            new_st.area = self.areas[new_st.technology] * float(new_st.capacity)
                    self._stations.stations.append(new_st)
                    self.addStation(self._stations.stations[-1])
                scene.close()
            self._scenarios.append([scen_filter, False, description])

    def _setupGrid(self):
        def do_them(lines, width=self.line_width, grid2=False):
            for line in lines:
                color = QtGui.QColor()
                color.setNamedColor(line.style)
                pen = QtGui.QPen(color, width)
                pen.setJoinStyle(QtCore.Qt.RoundJoin)
                pen.setCapStyle(QtCore.Qt.RoundCap)
                start = self.mapFromLonLat(QtCore.QPointF(line.coordinates[0][1], line.coordinates[0][0]))
                for pt in range(1, len(line.coordinates)):
                    end = self.mapFromLonLat(QtCore.QPointF(line.coordinates[pt][1], line.coordinates[pt][0]))
                    ln = QtGui.QGraphicsLineItem(QtCore.QLineF(start, end))
                    ln.setPen(pen)
                    ln.setZValue(0)
                    self.addItem(ln)
                    if grid2:
                        self._gridGroup2.addToGroup(ln)
                    else:
                        self._gridGroup.addToGroup(ln)
                    start = end
            return
        self.lines = Grid()
        do_them(self.lines.lines)
        self.grid_lines = len(self.lines.lines)
        lines = Grid_Boundary()
        if len(lines.lines) > 0:
            lines.lines[0].style = self.colors['grid_boundary']
            do_them(lines.lines, width=0)
        if self.existing_grid2:
            lines2 = Grid(grid2=True)
            do_them(lines2.lines, grid2=True)

    def addStation(self, st):
        self._stationGroups[st.name] = []
        p = self.mapFromLonLat(QtCore.QPointF(st.lon, st.lat))
        size = -1
        if self.station_square:
            if self.scale:
                try:
                    size = sqrt(st.area) / 2.
                except:
                    if st.technology == 'Wind':
                        size = self.areas[st.technology] * float(st.no_turbines) * pow((float(st.rotor) * .001), 2)
                    else:
                        size = self.areas[st.technology] * float(st.capacity)
                e = self.destinationxy(st.lon, st.lat, 90., size)
                w = self.destinationxy(st.lon, st.lat, 270., size)
                n = self.destinationxy(st.lon, st.lat, 0., size)
                s = self.destinationxy(st.lon, st.lat, 180., size)
                p2 = self.mapFromLonLat(QtCore.QPointF(w[0], n[1]))
                pe = self.mapFromLonLat(QtCore.QPointF(e[0], n[1]))
                ps = self.mapFromLonLat(QtCore.QPointF(w[0], s[1]))
                x_d = pe.x() - p2.x()
                y_d = ps.y() - p2.y()
                el = QtGui.QGraphicsRectItem(p2.x(), p2.y(), x_d, y_d)
            else:
                el = QtGui.QGraphicsRectItem(p.x() - 1.5, p.y() - 1.5, 3, 3)   # here to adjust station squares when not scaling
        else:
            if self.scale:
                try:
                    size = sqrt(st.area / pi) * 2.   # need diameter
                except:
                    if st.technology == 'Wind':
                        size = sqrt(self.areas[st.technology] * float(st.no_turbines) * pow((float(st.rotor) * .001), 2)
                               / pi) * 2.
                    else:
                        size = sqrt(self.areas[st.technology] * float(st.capacity) / pi) * 2.
                east = self.destinationxy(st.lon, st.lat, 90., size)
                pe = self.mapFromLonLat(QtCore.QPointF(east[0], st.lat))
                north = self.destinationxy(st.lon, st.lat, 0., size)
                pn = self.mapFromLonLat(QtCore.QPointF(st.lon, north[1]))
                x_d = p.x() - pe.x()
                y_d = pn.y() - p.y()
                el = QtGui.QGraphicsEllipseItem(p.x() - x_d / 2, p.y() - y_d / 2, x_d, y_d)
            else:
                el = QtGui.QGraphicsEllipseItem(p.x() - 1.5, p.y() - 1.5, 3, 3)   # here to adjust station circles when not scaling
        el.setBrush(QtGui.QColor(self.colors[st.technology]))
        if self.colors['border'] != '':
            el.setPen(QtGui.QColor(self.colors['border']))
        else:
            el.setPen(QtGui.QColor(self.colors[st.technology]))
        el.setZValue(1)
        self.addItem(el)
        self._stationGroups[st.name].append(el)
        self._stationCircles[st.technology].append(el)
        if st.technology[:6] == 'Fossil':
            self._fossilGroup.addToGroup(el)
        size = sqrt(float(st.capacity) * self.capacity_area / pi)
        east = self.destinationxy(st.lon, st.lat, 90., size)
        pe = self.mapFromLonLat(QtCore.QPointF(east[0], st.lat))
        north = self.destinationxy(st.lon, st.lat, 0., size)
        pn = self.mapFromLonLat(QtCore.QPointF(st.lon, north[1]))
        x_d = p.x() - pe.x()
        y_d = pn.y() - p.y()
        el = QtGui.QGraphicsEllipseItem(p.x() - x_d / 2, p.y() - y_d / 2, x_d, y_d)
        if self.show_capacity_fill:
            el.setBrush(QtGui.QColor(self.colors[st.technology]))
            el.setOpacity(self.capacity_opacity)
        if self.colors['border'] != '':
            el.setPen(QtGui.QColor(self.colors['border']))
        else:
            el.setPen(QtGui.QColor(self.colors[st.technology]))
        el.setZValue(1)
        self.addItem(el)
        self._stationGroups[st.name].append(el)
        self._stationCircles[st.technology].append(el)
        if st.technology[:6] == 'Fossil':
            self._fcapacityGroup.addToGroup(el)
        else:
            self._capacityGroup.addToGroup(el)
        self.addLine(st)
        txt = QtGui.QGraphicsSimpleTextItem(st.name)
        new_font = txt.font()
        new_font.setPointSizeF(self.width() / 10)
        txt.setFont(new_font)
        txt.setPos(p + QtCore.QPointF(1.5, -0.5))
        txt.scale(0.1, 0.1)
        if st.technology[:6] == 'Fossil':
            txt.setBrush(QtGui.QColor(self.colors['fossil_name']))
        else:
            txt.setBrush(QtGui.QColor(self.colors['station_name']))
        txt.setZValue(2)
        self.addItem(txt)
        self._stationGroups[st.name].append(txt)
        self._stationLabels.append(txt)
        if st.technology[:6] == 'Fossil':
            self._fnameGroup.addToGroup(txt)
        else:
            self._nameGroup.addToGroup(txt)
        return

    def addGeneration(self, st):
        if (st.technology[:6] != 'Fossil' or self.show_fossil) and st.generation > 0:
            p = self.mapFromLonLat(QtCore.QPointF(st.lon, st.lat))
            size = sqrt(float(st.generation) * (self.capacity_area / 2000) / pi)
            east = self.destinationxy(st.lon, st.lat, 90., size)
            pe = self.mapFromLonLat(QtCore.QPointF(east[0], st.lat))
            north = self.destinationxy(st.lon, st.lat, 0., size)
            pn = self.mapFromLonLat(QtCore.QPointF(st.lon, north[1]))
            x_d = p.x() - pe.x()
            y_d = pn.y() - p.y()
            el = QtGui.QGraphicsEllipseItem(p.x() - x_d / 2, p.y() - y_d / 2, x_d, y_d)
            if self.show_capacity_fill:
                brush = QtGui.QBrush()
                brush.setColor(QtGui.QColor(self.colors[st.technology]))
            #     brush.setStyle(Qt.Dense1Pattern)
                brush.setStyle(Qt.SolidPattern)
                el.setBrush(brush)
                el.setOpacity(self.capacity_opacity)
            if self.colors['border'] != '':
                el.setPen(QtGui.QColor(self.colors['border']))
            else:
                el.setPen(QtGui.QColor(self.colors[st.technology]))
            el.setZValue(1)
            self.addItem(el)
            self._stationGroups[st.name].append(el)
            self._stationCircles[st.technology].append(el)
            self._generationGroup.addToGroup(el)

    def addLine(self, st):
        if st.technology == 'Rooftop PV':
            return
        if self.new_grid and (st.scenario != 'Existing' or self.trace_existing):
            dims = [[st.lat, st.lon]]
            grid_len = 0.
            grid_path_len = 0.
            gridn = ''
            if st.grid_line is not None:
                gridn = st.grid_line.replace('(', '')
                gridn = gridn.replace('))', '')
                gridn = gridn.replace('),', ',')
                gridn = gridn.replace(')', '')
                gridn = gridn.split(',')
                for j in range(2, len(gridn), 2):
                    dims.append([float(gridn[j]), float(gridn[j + 1])])
                    grid_len += self.lines.actualDistance(dims[-2][0], dims[-2][1], dims[-1][0], dims[-1][1])
            grid_point = self.lines.gridConnect(dims[-1][0], dims[-1][1])
            if grid_point[0] == -1:
                if self.load_centre is not None:   # start with a load centre
                    nearest = 99999
                    j = -1
                    for i in range(len(self.load_centre)):
                        thisone = self.lines.actualDistance(self.load_centre[i][1], self.load_centre[i][2],
                                  dims[-1][0], dims[-1][1])
                        if thisone < nearest:
                            nearest = thisone
                            j = i
                    grid_point[0] = nearest
                    grid_point[1] = self.load_centre[j][1]
                    grid_point[2] = self.load_centre[j][2]
                    grid_point[3] = -1
                else:
                    try:
                        nearest, dist = self._stations.Nearest(dims[-1][0], dims[-1][1], distance=True, fossil=True,
                                        ignore=st.name)
                    except:
                        return
                    grid_point[0] = dist
                    grid_point[1] = nearest.lat
                    grid_point[2] = nearest.lon
                    grid_point[3] = -1
            if grid_point[0] > 0 or len(gridn) > 1:
                grid_len += grid_point[0]
                dims.append([grid_point[1], grid_point[2]])
                if st.technology in self.dispatchable:
                    dispatchable = 'Y'
                else:
                    dispatchable = ''
                l = Line(st.name, self.colors['new_grid'], dims, round(grid_len, 1), grid_point[3], dispatchable, 0,
                    0.)
            #     l.peak_loss = round(st.capacity * self.line_loss * l.length + self.subs_loss * st.capacity, 3)
              #   subs_cost = self.subs_cost * st.capacity
                if l.dispatchable == 'Y':
                    cost, l.line_table = self.lines.Line_Cost(st.capacity, st.capacity)
                else:
                    cost, l.line_table = self.lines.Line_Cost(st.capacity, 0.)
                l.line_cost = round(l.length * cost, 0)
                next = grid_point[3]
                self.lines.lines.append(l)
                if self.trace_grid and self.load_centre is not None:
                    nearest = 99999
                    j = -1
                    for i in range(len(self.load_centre)):
                        thisone = self.lines.actualDistance(self.load_centre[i][1], self.load_centre[i][2],
                                  dims[0][0], dims[0][1])
                        if thisone < nearest:
                            nearest = thisone
                            j = i
                    path = Shortest(self.lines.lines, dims[0], [self.load_centre[j][1],
                           self.load_centre[j][2]], self.grid_lines)
                    line = path.getLines()
                    for li in line:
                        if self.lines.lines[li].peak_load is None:
                            self.lines.lines[li].peak_load = 0.
                        if self.lines.lines[li].peak_loss is None:
                            self.lines.lines[li].peak_loss = 0.
                        self.lines.lines[li].peak_loss += round((st.capacity * self.line_loss * self.lines.lines[li].length), 3)
                        self.lines.lines[li].peak_load += st.capacity
                        if dispatchable == 'Y':
                            self.lines.lines[li].dispatchable = 'Y'
                            if self.lines.lines[li].peak_dispatchable is None:
                                self.lines.lines[li].peak_dispatchable = 0.
                            self.lines.lines[li].peak_dispatchable += st.capacity
                    route = path.getPath()
#                   check we don't go through another load_centre
                    if len(self.load_centre) > 1:
                        for co in range(len(route) - 1, 0, -1):
                            for i in range(len(self.load_centre)):
                                if route[co][0] == self.load_centre[i][1] and \
                                  route[co][1] == self.load_centre[i][2]:
                                    route = route[i:]
                                    break
                    for i in range(1, len(route)):
                        grid_path_len += self.lines.actualDistance(route[i - 1][0], route[i - 1][1],
                                         route[i][0], route[i][1])
                else:
                    if self.lines.lines[next].peak_load is None:
                        self.lines.lines[next].peak_load = 0.
                        self.lines.lines[next].peak_loss = 0.
                    self.lines.lines[next].peak_load += st.capacity
                    if dispatchable == 'Y':
                        if self.lines.lines[next].peak_dispatchable is None:
                            self.lines.lines[next].peak_dispatchable = 0.
                        self.lines.lines[next].peak_dispatchable += st.capacity
                    try:
                        self.lines.lines[next].peak_loss += round((st.capacity * self.line_loss * l.length), 3)
                    except:
                        pass
                color = QtGui.QColor()
                color.setNamedColor(self.lines.lines[-1].style)
                pen = QtGui.QPen(color, self.line_width)
                pen.setJoinStyle(QtCore.Qt.RoundJoin)
                pen.setCapStyle(QtCore.Qt.RoundCap)
                start = self.mapFromLonLat(QtCore.QPointF(l.coordinates[0][1],
                  l.coordinates[0][0]))
                for i in range(1, len(l.coordinates)):
                    end = self.mapFromLonLat(QtCore.QPointF(l.coordinates[i][1],
                      l.coordinates[i][0]))
                    ln = QtGui.QGraphicsLineItem(QtCore.QLineF(start, end))
                    ln.setPen(pen)
                    ln.setZValue(0)
                    self.addItem(ln)
                    self._stationGroups[st.name].append(ln)
                    start = end
            if grid_len > 0:
                st.grid_len = grid_len
            if grid_path_len > 0:
                st.grid_path_len = grid_path_len

    def refreshGrid(self):
        for i in range(self.grid_lines):
            if self.lines.lines[i].peak_load is not None:
                self.lines.lines[i].peak_dispatchable = 0.
                self.lines.lines[i].peak_load = 0.
                self.lines.lines[i].peak_loss = 0.
        if self.new_grid:
            kept_line = []
            for i in range(len(self.lines.lines) - 1, self.grid_lines - 1, -1):
                if self.lines.lines[i].peak_load is not None:
                    self.lines.lines[i].peak_dispatchable = 0.
                    self.lines.lines[i].peak_load = 0.
                    self.lines.lines[i].peak_loss = 0.
                if self.lines.lines[i].length > 0:
                    try:
                        for j in range(len(self._stationGroups[self.lines.lines[i].name])):
                            if isinstance(self._stationGroups[self.lines.lines[i].name][j], QtGui.QGraphicsLineItem):
                                self.removeItem(self._stationGroups[self.lines.lines[i].name][j])
                                del self._stationGroups[self.lines.lines[i].name][j]
                                break
                    except:
                        pass
                    del self.lines.lines[i]
            for st in self._stations.stations:
                self.addLine(st)

    def changeDate(self, d):
        return
        d = datetime.date(d.year(), d.month(), d.day())
        for st in self._stations.values():
            st.changeDate(d)
        self._power_tot.changeDate(d)

    def mapToLonLat(self, p):
        x = p.x() / self._lon_scale + self._orig_lon
        y = p.y() / self._lat_scale + self._orig_lat
        lon, lat = self._proj(x, y, inverse=True)
        return QtCore.QPointF(round(lon, 4), round(lat, 4))

    def mapFromLonLat(self, p):
        lon, lat = p.x(), p.y()
        x, y = self._proj(lon, lat)
        x = (x - self._orig_lon) * self._lon_scale
        y = (y - self._orig_lat) * self._lat_scale
        return QtCore.QPointF(x, y)

    def positions(self):
        try:
            return self._positions
        except:
            return

    def stationPositions(self):
        return self._station_positions

    def toggleTotal(self, start):
        if self._power_tot.infoVisible():
            self._power_tot.hideInfo()
        else:
            self._power_tot.showInfo(0, start)

    def powerPlotImage(self, name):
        return self._stations[name].powerPlotImage()
