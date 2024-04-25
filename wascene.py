#!/usr/bin/python3
#
#  Copyright (C) 2015-2023 Sustainable Energy Now Inc., Angus King
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

import datetime
from math import sin, cos, pi, sqrt, degrees, radians, asin, atan2
import os
import sys
import configparser   # decode .ini file
from PyQt5 import QtCore, QtGui, QtWidgets
try:
    import mpl_toolkits.basemap.pyproj as pyproj   # Import the pyproj module
except:
    import pyproj
from towns import Towns
from getmodels import getModelFile
from grid import Grid, Grid_Area, Grid_Boundary, Grid_Zones, Line
from senutils import getParents, getUser, techClean, WorkBook
from station import Station, Stations
from dijkstra_4 import Shortest


class WAScene(QtWidgets.QGraphicsScene):

    def get_config(self):
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
            if config_file.rfind('/') >= 0:
                self.config_file = config_file[config_file.rfind('/') + 1:]
            elif config_file.rfind('\\') >= 0:
                self.config_file = config_file[config_file.rfind('\\') + 1:]
            else:
                self.config_file = config_file
        else:
            config_file = getModelFile('SIREN.ini')
            self.config_file = 'SIREN.ini'
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
            parents = getParents(config.items('Parents'))
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
            try:
                scale = config.get('Map', 'scale')
            except:
                scale = 'false'
        if scale.lower() in ['true', 'yes', 'on']:
            self.scale = True
        try:
            scenario_prefix = config.get('Files', 'scenario_prefix')
        except:
            scenario_prefix = ''
        try:
            self.scenarios = config.get('Files', 'scenarios')
            if scenario_prefix != '' :
                self.scenarios += '/' + scenario_prefix
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
        self.colors['grid_areas'] = '#00FF00'
        self.colors['grid_zones'] = '#FFFF00'
        self.colors['grid_trace'] = 'white'
        self.colors['ruler'] = 'white'
        self.colors['station'] = '#00FF00'
        self.colors['station_name'] = 'white'
        self.colors['town'] = 'red'
        self.colors['town_name'] = 'lightGray'
        try:
            technologies = config.get('Power', 'technologies')
        except:
            technologies = ''
        technologies = technologies.split()
        try:
            colours = config.items('Colors')
            for item, colour in colours:
                if item in technologies or (item[:6] == 'fossil' and item != 'fossil_name'):
                    itm = techClean(item)
                    self.colors[itm] = colour
                else:
                    self.colors[item] = colour
        except:
            pass
        self.areas = {}
        try:
            for item in technologies:
                itm = techClean(item)
                try:
                    self.areas[itm] = float(config.get(itm, 'area'))
                except:
                    self.areas[itm] = 0.
        except:
            pass
        try:
            fossil_technologies = config.get('Power', 'fossil_technologies')
            fossil_technologies = technologies.split()
            for item in fossil_technologies:
                itm = techClean(item)
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
                        itm = techClean(item)
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
            if show_ruler.lower() in ['true', 'yes', 'on']:
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
            if center_on_click.lower() in ['true', 'yes', 'on']:
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
            config.get('Files', 'grid2_network')
            self.existing_grid2 = True
        except configparser.NoOptionError:
            try:
                config.get('Files', 'grid_network2')
                self.existing_grid2 = True
            except:
                self.existing_grid2 = False
        except:
            self.existing_grid2 = False
        try:
            grid_zones = config.get('Files', 'grid_zones')
            self.grid_zones = True
            try:
                grid_zones = config.get('View', 'grid_zones')
                if grid_zones.lower() in ['false', 'no', 'off']:
                    self.grid_zones = False
            except:
                pass
        except:
            self.grid_zones = False
        self.zone_opacity = 0.
        try:
            self.zone_opacity = float(config.get('View', 'zone_opacity'))
            if self.zone_opacity < 0. or self.zone_opacity > 1.:
                self.zone_opacity = 0.
        except:
            pass
        self.grid_areas = False
        for s in ['', '1', '2', '3', '4', '5']:
            try:
                grid_areas = config.get('Files', 'grid_areas' + s)
                self.grid_areas = True
                try:
                    grid_areas = config.get('View', 'grid_areas')
                    if grid_areas.lower() in ['false', 'no', 'off']:
                        self.grid_areas = False
                except:
                    pass
                break
            except:
                pass
        self.area_opacity = 0.1
        try:
            self.area_opacity = float(config.get('View', 'area_opacity'))
            if self.area_opacity < 0. or self.area_opacity > 1.:
                self.area_opacity = 0.1
        except:
            pass
        self.line_group = True
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
        self.hide_map = False
        try:
            hide_map = config.get('View', 'hide_map')
            if hide_map.lower() in ['true', 'yes', 'on']:
                self.hide_map = True
        except:
            pass
        self.show_coord = False
        try:
            show_coord = config.get('View', 'show_coord')
            if show_coord.lower() in ['true', 'yes', 'on']:
                self.show_coord = True
        except:
            pass
        self.coord_grid = [0, 0, 'c']
        try:
            coord_grid = config.get('View', 'coord_grid')
            if coord_grid.lower()[0] == 'm': # merra-2
                self.coord_grid = [.5, .625, 'c']
            elif coord_grid.lower()[0] == 'e': # era5
                self.coord_grid = [.25, .25, 'c'] # set to 't' to have grid at top left
            else: # lat,lon
                try:
                    bits = coord_grid.split(',')
                    self.coord_grid[0] = float(bits[0])
                    self.coord_grid[1] = float(bits[0])
                except:
                    pass
        except:
            pass
        try:
            self.txt_ratio = float(config.get('View', 'txt_ratio')) / 10.
        except:
            self.txt_ratio = 0.01
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
                try:
                    self.load_centre = [[load_centre[0], float(load_centre[1]), float(load_centre[2])]]
                    for j in range(3, len(load_centre), 3):
                        self.load_centre.append([load_centre[j], float(load_centre[j + 1]),
                            float(load_centre[j + 2])])
                except:
                    pass
        except:
            pass
        self.station_square = False
        try:
            square = config.get('View', 'station_shape')
            if square[0].lower() == 's':
                self.station_square = True
        except:
            pass
        self.station_opacity = 1.
        try:
            self.station_opacity = float(config.get('View', 'station_opacity'))
            if self.station_opacity < 0. or self.station_opacity > 1.:
                self.station_opacity = 1.
        except:
            pass
        self.dispatchable = None
        self.line_loss = 0.
       #  self.subs_cost=0.
      #   self.subs_loss=0.
        try:
            itm = config.get('Grid', 'dispatchable')
            self.dispatchable = techClean(itm)
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
        self.cst_tshours = 0
        try:
            self.cst_tshours = float(config.get('CST', 'tshours'))
        except:
            pass
        self.st_tshours = 0
        try:
            self.st_tshours = float(config.get('Solar Thermal', 'tshours'))
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
        ln1, lt1, baring = list(map(radians, [lon1, lat1, bearing]))
     # "reverse" haversine formula
        lat2 = asin(sin(lt1) * cos(distance / radius) +
                                cos(lt1) * sin(distance / radius) * cos(baring))
        lon2 = ln1 + atan2(sin(baring) * sin(distance / radius) * cos(lt1),
                                            cos(distance / radius) - sin(lt1) * sin(lat2))
        return degrees(lon2), degrees(lat2)

    def __init__(self):
        QtWidgets.QGraphicsScene.__init__(self)
        self.exitLoop = False
        self.loopMax = 0
        self.get_config()
        self.last_locn = None
        self.setBackgroundBrush(QtGui.QColor(self.colors['background']))
        if os.path.exists(self.map_file):
            war = QtGui.QImageReader(self.map_file)
            wa = war.read()
            pixMap = QtGui.QPixmap.fromImage(wa)
            if self.hide_map:
                pixMap.fill()
        else:
            ratio = abs(self.map_upper_left[1] - self.map_lower_right[1]) / \
                     abs(self.map_upper_left[0] - self.map_lower_right[0])
            pixMap = QtGui.QPixmap(int(200 * ratio), 200)
            if self.map != '':
                painter = QtGui.QPainter(pixMap)
                brush = QtGui.QBrush(QtGui.QColor('white'))
                painter.fillRect(QtCore.QRectF(0, 0, 200 * ratio, 200), brush)
                painter.setPen(QtGui.QColor('lightgray'))
                painter.drawText(QtCore.QRectF(0, 0, 200 * ratio, 200), QtCore.Qt.AlignCenter,
                                 'Map not found.')
                painter.end()
        self.pixmap = pixMap
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
        self._gridGroup = QtWidgets.QGraphicsItemGroup() # normal grid group
        self._gridGroup2 = QtWidgets.QGraphicsItemGroup() # extra grid group
        self._gridGroupz = QtWidgets.QGraphicsItemGroup() # zone group
        self._gridGroupa = QtWidgets.QGraphicsItemGroup() # areas group
        self._lineGroup = QtWidgets.QGraphicsItemGroup() # stations lines group
        try:
            self._setupGrid()
        except Exception as e:
            msgbox = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Critical,
                                           'Error setting up grid.',
                                            e.msg + '\n\nMay need to check map coordinates\n(upper_left' + self.map \
                                            + ' and lower_right' + self.map + ' in [Map]) or' \
                                            + '\nthe status of grid-related (KML) files.' \
                                            + '\nExecution aborted.',
                                           QtWidgets.QMessageBox.Ok)
            reply = msgbox.exec_()
            return
        self.addItem(self._gridGroup)
        self.addItem(self._lineGroup)
        if not self.existing_grid:
            self._gridGroup.setVisible(False)
        if self.existing_grid2:
            self.addItem(self._gridGroup2)
        if self.grid_zones:
            self.addItem(self._gridGroupz)
        if self.grid_areas:
            self.addItem(self._gridGroupa)
        self._coordGroup = QtWidgets.QGraphicsItemGroup()
        self._setupCoordGrid()
        self.addItem(self._coordGroup)
        if not self.show_coord:
            self._coordGroup.setVisible(False)
        self._townGroup = QtWidgets.QGraphicsItemGroup()
        self._setupTowns()
        self.addItem(self._townGroup)
        if not self.show_towns:
            self._townGroup.setVisible(False)
        self._scenarios = []
        self._capacityGroup = QtWidgets.QGraphicsItemGroup()
        self._generationGroup = QtWidgets.QGraphicsItemGroup()
        self._nameGroup = QtWidgets.QGraphicsItemGroup()
        self._fossilGroup = QtWidgets.QGraphicsItemGroup()
        self._fcapacityGroup = QtWidgets.QGraphicsItemGroup()
        self._fnameGroup = QtWidgets.QGraphicsItemGroup()
        self._setupStations()
        if len(self._stations.tech_missing) > 0:
            msg = 'Preferences file needs checking for -\n'
            for tech in self._stations.tech_missing:
                msg += " '" + tech + "',"
            msgbox = QtWidgets.QMessageBox()
            msgbox.setWindowTitle('SIREN - ' + self.config_file + ' setup stations')
            msgbox.setText("Error encountered in setting up stations.\n" + \
                           msg + '\nExecution will continue but need to check stations.')
            msgbox.setIcon(QtWidgets.QMessageBox.Warning)
            msgbox.setStandardButtons(QtWidgets.QMessageBox.Ok)
            reply = msgbox.exec_()
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
        self._proj = pyproj.Proj(self.projection)   # LatLon with WGS84 datum used by GPS units and Google Earth
        x1, y1, lon1, lat1 = self.upper_left
        x2, y2, lon2, lat2 = self.lower_right
        ul = self._proj(lon1, lat1)
        lr = self._proj(lon2, lat2)
        self._lat_scale = y2 / (lr[1] - ul[1])
        self._lon_scale = x2 / (lr[0] - ul[0])
        self._orig_lat = ul[1]
        self._orig_lon = ul[0]

    def _setupCoordGrid(self):
        color = QtGui.QColor()
        color.setNamedColor((self.colors['town_name']))
        pen = QtGui.QPen(color, self.line_width)
        pen.setJoinStyle(QtCore.Qt.RoundJoin)
        pen.setCapStyle(QtCore.Qt.RoundCap)
        if self.coord_grid[0] > 0 and self.coord_grid[1] > 0:
            latn = self.coord_grid[0] * round((self.map_upper_left[0]) / self.coord_grid[0])
            lats = self.coord_grid[0] * round((self.map_lower_right[0]) / self.coord_grid[0])
            lonw = self.coord_grid[1] * round((self.map_upper_left[1]) / self.coord_grid[1])
            lone = self.coord_grid[1] * round((self.map_lower_right[1]) / self.coord_grid[1])
            lat_step = self.coord_grid[0]
            lon_step = self.coord_grid[1]
            if self.coord_grid[2] == 'c':
                lat = lats + self.coord_grid[0] / 2
                lon = lonw + self.coord_grid[1] / 2
            else:
                lat = lats
                lon = lonw
        else:
            bnds = [45., 90., 180.]
            degs = [2.5, 5., 10.]
            mlt = 1
            step1 = abs((self.map_upper_left[0] - self.map_lower_right[0]))
            step2 = abs((self.map_upper_left[1] - self.map_lower_right[1]))
            step = min(step1, step2)
            while step < bnds[0] * mlt:
                mlt = mlt / 10.
            if step > bnds[2] * mlt:
                step = degs[2] * mlt
            elif step > bnds[1] * mlt:
                step = degs[1] * mlt
            else:
                step = degs[0] * mlt
            lat_step = lon_step = step
            lat = step * round((self.map_lower_right[0] + step / 2.) / step)
            lon = step * round((self.map_upper_left[1] + step / 2.) / step)
        while lat <= self.map_upper_left[0]:
            fromm = self.mapFromLonLat(QtCore.QPointF(self.map_upper_left[1], lat))
            too = self.mapFromLonLat(QtCore.QPointF(self.map_lower_right[1], lat))
            item = QtWidgets.QGraphicsLineItem(fromm.x(), fromm.y(), too.x(), too.y())
            item.setPen(pen)
            item.setZValue(3)
            self._coordGroup.addToGroup(item)
            lat += lat_step
        while lon <= self.map_lower_right[1]:
            fromm = self.mapFromLonLat(QtCore.QPointF(lon, self.map_upper_left[0]))
            too = self.mapFromLonLat(QtCore.QPointF(lon, self.map_lower_right[0]))
            item = QtWidgets.QGraphicsLineItem(fromm.x(), fromm.y(), too.x(), too.y())
            item.setPen(pen)
            item.setZValue(3)
            self._coordGroup.addToGroup(item)
            lon += lon_step
        return

    def _setupTowns(self):
        self._towns = {}
        self._towns = Towns(ul_lat=self.upper_left[3], ul_lon=self.upper_left[2],
                      lr_lat=self.lower_right[3], lr_lon=self.lower_right[2])
        for st in self._towns.towns:
            p = self.mapFromLonLat(QtCore.QPointF(st.lon, st.lat))
            el = QtWidgets.QGraphicsEllipseItem(p.x() - 1, p.y() - 1, 2, 2)   # here to adjust town circles
            el.setBrush(QtGui.QColor(self.colors['town']))
            el.setPen(QtGui.QColor(self.colors['town']))
            el.setZValue(0)
            self._townGroup.addToGroup(el)
            txt = QtWidgets.QGraphicsSimpleTextItem(st.name)
            new_font = txt.font()
            new_font.setPointSizeF(self.width() * (self.txt_ratio / 2.))
            txt.setFont(new_font)
            txt.setPos(p + QtCore.QPointF(1.5, -0.5))
            txt.setBrush(QtGui.QColor(self.colors['town_name']))
            txt.setZValue(0)
            self._townGroup.addToGroup(txt)
        return

    def _setupStations(self):
        self._current_name = QtWidgets.QGraphicsSimpleTextItem('')
        new_font = self._current_name.font()
        new_font.setPointSizeF(self.width() * self.txt_ratio)
        self._current_name.setFont(new_font)
        self._current_name.setZValue(2)
        self.addItem(self._current_name)
        self._stations = []
        self._station_positions = {}
        self._stationGroups = {}
        self._stationLabels = []
        self._stationCircles = {}
        for key, value in self.colors.items():
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
            var = {}
            workbook = WorkBook()
            workbook.open_workbook(scen_file)
            worksheet = workbook.sheet_by_index(0)
            num_rows = worksheet.nrows - 1
            num_cols = worksheet.ncols - 1
            if worksheet.cell_value(0, 0) == 'Description:' or worksheet.cell_value(0, 0) == 'Comment:':
                curr_row = 1
                description = worksheet.cell_value(0, 1)
            else:
                curr_row = 0
#           get column names
            curr_col = -1
            while curr_col < num_cols:
                curr_col += 1
                var[worksheet.cell_value(curr_row, curr_col)] = curr_col
            while curr_row < num_rows:
                curr_row += 1
                try:
                    try:
                        area = worksheet.cell_value(curr_row, var['Area'])
                    except:
                        area = 0
                    new_st = Station(str(worksheet.cell_value(curr_row, var['Station Name'])),
                                     str(worksheet.cell_value(curr_row, var['Technology'])),
                                     worksheet.cell_value(curr_row, var['Latitude']),
                                     worksheet.cell_value(curr_row, var['Longitude']),
                                     worksheet.cell_value(curr_row, var['Maximum Capacity (MW)']),
                                     str(worksheet.cell_value(curr_row, var['Turbine'])),
                                     worksheet.cell_value(curr_row, var['Rotor Diam']),
                                     worksheet.cell_value(curr_row, var['No. turbines']),
                                     area,
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
                        hub_height = worksheet.cell_value(curr_row, var['Hub Height'])
                        if hub_height != '':
                            setattr(new_st, 'hub_height', hub_height)
                    except:
                        pass
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
                        tilt = worksheet.cell_value(curr_row, var['Tilt'])
                        if tilt != '':
                            setattr(new_st, 'tilt', tilt)
                    except:
                        pass
                    try:
                        storage_hours = worksheet.cell_value(curr_row, var['Storage Hours'])
                        if storage_hours != '':
                            setattr(new_st, 'storage_hours', storage_hours)
                    except:
                        pass
                    self._stations.stations.append(new_st)
                    self.addStation(self._stations.stations[-1])
                except Exception as error:
                    print('wascene error:', error)
                    pass
            self._scenarios.append([scen_filter, False, description])

    def _setupGrid(self):
        def do_them(lines, width=self.line_width, grid_lines='', opacity=0.):
            for line in lines:
                color = QtGui.QColor()
                color.setNamedColor(line.style)
                pen = QtGui.QPen(color, width)
                pen.setJoinStyle(QtCore.Qt.RoundJoin)
                pen.setCapStyle(QtCore.Qt.RoundCap)
                if grid_lines in ['a', 'b', 'z']:
                    bnds = []
                    for pt in range(len(line.coordinates)):
                        bnds.append(self.mapFromLonLat(QtCore.QPointF(line.coordinates[pt][1], line.coordinates[pt][0])))
                    path = QtGui.QPainterPath()
                    path.addPolygon(QtGui.QPolygonF(bnds))
                    poly = QtWidgets.QGraphicsPathItem(path)
                    brush = QtGui.QColor(color)
                    brush.setAlphaF(opacity)
                    poly.setBrush(brush)
                    poly.setPen(pen)
                    if grid_lines == 'a':
                        self._gridGroupa.addToGroup(poly)
                    elif grid_lines == 'b':
                        self._gridGroup.addToGroup(poly)
                    elif grid_lines == 'z':
                        self._gridGroupz.addToGroup(poly)
                else:
                    start = self.mapFromLonLat(QtCore.QPointF(line.coordinates[0][1], line.coordinates[0][0]))
                    for pt in range(1, len(line.coordinates)):
                        end = self.mapFromLonLat(QtCore.QPointF(line.coordinates[pt][1], line.coordinates[pt][0]))
                        # FIXME Can't identify the QGraphicsScene in the arguments of the QGraphicsItem
                        ln = QtWidgets.QGraphicsLineItem(QtCore.QLineF(start, end))
                        ln.setPen(pen)
                        ln.setZValue(0)
                        self.addItem(ln)
                        if grid_lines == '':
                            self._gridGroup.addToGroup(ln)
                        elif grid_lines == '2':
                            self._gridGroup2.addToGroup(ln)
                        start = end
            return
        self.lines = Grid()
        do_them(self.lines.lines)
        self.grid_lines = len(self.lines.lines)
        lines = Grid_Boundary()
        if len(lines.lines) > 0:
            lines.lines[0].style = self.colors['grid_boundary']
            do_them(lines.lines, width=0, grid_lines='b')
        if self.existing_grid2:
            lines2 = Grid(grid2=True)
            do_them(lines2.lines, grid_lines='2')
        if self.grid_zones:
            self.linesz = Grid_Zones()
            if len(self.linesz.lines) > 0:
                do_them(self.linesz.lines, grid_lines='z', opacity=self.zone_opacity)
        if self.grid_areas:
            linesa = Grid_Area('grid_areas')
            if len(linesa.lines) > 0:
                do_them(linesa.lines, grid_lines='a', opacity=self.area_opacity)

    def addStation(self, st):
        self._stationGroups[st.name] = []
        p = self.mapFromLonLat(QtCore.QPointF(st.lon, st.lat))
        try:
            if len(self.linesz.lines) > 0:
                st.zone = self.linesz.getZone(st.lat, st.lon)
        except:
            pass
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
                el = QtWidgets.QGraphicsRectItem(p2.x(), p2.y(), x_d, y_d)
            else:
                el = QtWidgets.QGraphicsRectItem(p.x() - 1.5, p.y() - 1.5, 3, 3)   # here to adjust station squares when not scaling
        else:
            if self.scale:
                try:
                    size = sqrt(st.area / pi) * 2.   # need diameter
                except:
                    if st.technology == 'Wind':
                        size = sqrt(self.areas[st.technology] *
                                    float(st.no_turbines) * pow((float(st.rotor) *
                                                                 .001), 2) / pi) * 2.
                    else:
                        size = sqrt(self.areas[st.technology] * float(st.capacity) / pi) * 2.
                east = self.destinationxy(st.lon, st.lat, 90., size)
                pe = self.mapFromLonLat(QtCore.QPointF(east[0], st.lat))
                north = self.destinationxy(st.lon, st.lat, 0., size)
                pn = self.mapFromLonLat(QtCore.QPointF(st.lon, north[1]))
                x_d = p.x() - pe.x()
                y_d = pn.y() - p.y()
                el = QtWidgets.QGraphicsEllipseItem(p.x() - x_d / 2, p.y() - y_d / 2, x_d, y_d)
            else:
                el = QtWidgets.QGraphicsEllipseItem(p.x() - 1.5, p.y() - 1.5, 3, 3)   # here to adjust station circles when not scaling
        try:
            el.setBrush(QtGui.QColor(self.colors[st.technology]))
        except:
            self.colors[st.technology] = 'gray'
            el.setBrush(QtGui.QColor('gray'))
        if self.station_opacity < 1.:
            el.setOpacity(self.station_opacity)
        if self.colors['border'] != '':
            el.setPen(QtGui.QColor(self.colors['border']))
        else:
            el.setPen(QtGui.QColor(self.colors[st.technology]))
        el.setZValue(1)
        self.addItem(el)
        self._stationGroups[st.name].append(el)
        try:
            self._stationCircles[st.technology].append(el)
        except:
            self._stationCircles[st.technology] = [(el)]
        if st.technology[:6] == 'Fossil':
            self._fossilGroup.addToGroup(el)
        size = sqrt(float(st.capacity) * self.capacity_area / pi)
        east = self.destinationxy(st.lon, st.lat, 90., size)
        pe = self.mapFromLonLat(QtCore.QPointF(east[0], st.lat))
        north = self.destinationxy(st.lon, st.lat, 0., size)
        pn = self.mapFromLonLat(QtCore.QPointF(st.lon, north[1]))
        x_d = p.x() - pe.x()
        y_d = pn.y() - p.y()
        el = QtWidgets.QGraphicsEllipseItem(p.x() - x_d / 2, p.y() - y_d / 2, x_d, y_d)
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
        txt = QtWidgets.QGraphicsSimpleTextItem(st.name)
        new_font = txt.font()
        new_font.setPointSizeF(self.width() * self.txt_ratio)
        txt.setFont(new_font)
        txt.setPos(p + QtCore.QPointF(1.5, -0.5))
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
            el = QtWidgets.QGraphicsEllipseItem(p.x() - x_d / 2, p.y() - y_d / 2, x_d, y_d)
            if self.show_capacity_fill:
                brush = QtGui.QBrush()
                brush.setColor(QtGui.QColor(self.colors[st.technology]))
            #     brush.setStyle(QtCore.Qt.Dense1Pattern)
                brush.setStyle(QtCore.Qt.SolidPattern)
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
                    # FIXME Can't identify the QGraphicsScene in the arguments of the QGraphicsItem
                    ln = QtWidgets.QGraphicsLineItem(QtCore.QLineF(start, end))
                    ln.setPen(pen)
                    ln.setZValue(0)
                    self.addItem(ln)
                    self._stationGroups[st.name].append(ln)
                    self._lineGroup.addToGroup(ln)
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
            for i in range(len(self.lines.lines) - 1, self.grid_lines - 1, -1):
                if self.lines.lines[i].peak_load is not None:
                    self.lines.lines[i].peak_dispatchable = 0.
                    self.lines.lines[i].peak_load = 0.
                    self.lines.lines[i].peak_loss = 0.
                if self.lines.lines[i].length > 0:
                    try:
                        for j in range(len(self._stationGroups[self.lines.lines[i].name])):
                            if isinstance(self._stationGroups[self.lines.lines[i].name][j], QtWidgets.QGraphicsLineItem):
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
        for st in list(self._stations.values()):
            st.changeDate(d)
        self._power_tot.changeDate(d)

    def mapToLonLat(self, p, decpts=4):
        x = p.x() / self._lon_scale + self._orig_lon
        y = p.y() / self._lat_scale + self._orig_lat
        lon, lat = self._proj(x, y, inverse=True)
        return QtCore.QPointF(round(lon, decpts), round(lat, decpts))

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
