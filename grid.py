#!/usr/bin/python
#
#  Copyright (C) 2015-2019 Sustainable Energy Now Inc., Angus King
#
#  grid.py - This file is part of SIREN.
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
import StringIO
import sys
from math import sin, cos, radians, asin, acos, atan2, sqrt, degrees
import zipfile

import ConfigParser   # decode .ini file
from xml.etree.ElementTree import ElementTree, fromstring

from parents import getParents
from senuser import getUser

RADIUS = 6367.   # radius of earth in km


def within_map(x, y, poly):
    n = len(poly)
    inside = False
    p1x, p1y = poly[0]
    for i in range(n + 1):
        p2x, p2y = poly[i % n]
        if y > min(p1y, p2y):
            if y <= max(p1y, p2y):
                if x <= max(p1x, p2x):
                    if p1y != p2y:
                        xints = (y - p1y) * (p2x - p1x) / (p2y - p1y) + p1x
                    if p1x == p2x or x <= xints:
                        inside = not inside
        p1x, p1y = p2x, p2y
    return inside


def dust(pyd, pxd, y1d, x1d, y2d, x2d):   # debug
    px = radians(pxd)
    py = radians(pyd)
    x1 = radians(x1d)
    y1 = radians(y1d)
    x2 = radians(x2d)
    y2 = radians(y2d)
    p_x = x2 - x1
    p_y = y2 - y1
    something = p_x * p_x + p_y * p_y
    u = ((px - x1) * p_x + (py - y1) * p_y) / float(something)
    if u > 1:
        u = 1
    elif u < 0:
        u = 0
    x = x1 + u * p_x
    y = y1 + u * p_y
    dx = x - px
    dy = y - py
    dist = sqrt(dx * dx + dy * dy)
    return [round(abs(dist) * RADIUS, 2), round(degrees(y), 6), round(degrees(x), 6)]


class Line:
    def __init__(self, name, style, coordinates, length=0., connector=-1, dispatchable=None, line_cost=None, peak_load=None,
                 peak_dispatchable=None, peak_loss=None, line_table=None, substation_cost=None):
        self.name = name
        self.style = style
        self.coordinates = []
        for i in range(len(coordinates)):
            self.coordinates.append([])
            for j in range(len(coordinates[i])):
                self.coordinates[-1].append(round(coordinates[i][j], 6))
        self.connector = connector
        self.length = length
        self.dispatchable = dispatchable
        self.line_cost = line_cost
        self.peak_load = peak_load
        self.peak_dispatchable = peak_dispatchable
        self.peak_loss = peak_loss
        self.line_table = line_table
        self.substation_cost = substation_cost


class Grid:
    def decode2(self, proces, substation=False):
        operands = ['+', '-', '*', '/', 'x']
        proces = proces.replace(' ', '')
        proces = proces.split('=')
        numbr = [0., 0.]
        for p in range(len(proces)):
            stack = []
            i = 0
            j = 0
            while j < len(proces[p]):
                if proces[p][j] in operands:
                    stack.append(proces[p][i:j])
                    stack.append(proces[p][j])
                    i = j + 1
                j += 1
            if j > i:
                stack.append(proces[p][i:])
            if p == 1:
                i = 0
                while i < len(stack):
                    try:
                        if substation:
                            stack[i] = self.substation_costs[stack[i]]
                        else:
                            stack[i] = self.line_costs[stack[i]]
                    except:
                        pass
                    i += 2
            numbr[p] = float(stack[0])
            for i in range(1, len(stack), 2):
                if stack[i] == '*' or stack[i] == 'x':
                    numbr[p] = numbr[p] * float(stack[i + 1])
                elif stack[i] == '/':
                    numbr[p] = numbr[p] / float(stack[i + 1])
                if stack[i] == '+':
                    numbr[p] = numbr[p] + float(stack[i + 1])
                if stack[i] == '-':
                    numbr[p] = numbr[p] - float(stack[i + 1])
        return [numbr[0], numbr[1], proces[1]]

    def decode(self, s_lines, the_lines=None):
        if the_lines is None:
            new_lines = []
        else:
            new_lines = the_lines
        l = 0
        while True:
            l = s_lines.find('(', l)
            if l < 0:
                break
            r = s_lines.find(')', l)
            l2 = s_lines.find('(', l + 1)
            if l2 > r:   # normal
                proces = s_lines[l + 1:r]
                l += 1
            else:
                b = 1
                while l2 < r:
                    b += 1
                    l2 = s_lines.find('(', l2 + 1)
                    if l2 < 0:
                        break
                while b > 0:
                    b -= 1
                    r = s_lines.find(')', r + 1)
                proces = s_lines[l + 1:r]
                l = r
            if proces[:4] == 'for(':
                i = 4
                while proces[i] != '=':
                    i += 1
                subs = proces[4:i]
                j = i + 1
                while proces[j] != ',':
                    j += 1
                frm = int(proces[i + 1:j])
                i = j + 1
                while proces[i] != ',':
                    i += 1
                too = int(proces[j + 1:i])
                for j in range(frm, too + 1):
                    this_bit = proces[i + 1:-1].replace(subs, str(j))
                    if j == 1:
                        this_bit = this_bit.replace('1*', '')
                        this_bit = this_bit.replace('*1', '')
                    self.decode(this_bit, new_lines)
            else:
                new_lines.append(self.decode2(proces))
        return new_lines

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
            parents = getParents(aparents = config.items('Parents'))
        except:
            pass
        try:
            self.kml_file = config.get('Files', 'grid_network')
            for key, value in parents:
                self.kml_file = self.kml_file.replace(key, value)
            self.kml_file = self.kml_file.replace('$USER$', getUser())
            self.kml_file = self.kml_file.replace('$YEAR$', self.base_year)
        except:
            self.kml_file = ''
        try:
            self.kml_file2 = config.get('Files', 'grid_network2')
            for key, value in parents:
                self.kml_file2 = self.kml_file2.replace(key, value)
            self.kml_file2 = self.kml_file2.replace('$USER$', getUser())
            self.kml_file2 = self.kml_file2.replace('$YEAR$', self.base_year)
        except:
            self.kml_file2 = ''
        try:
            mapc = config.get('Map', 'map_choice')
        except:
            mapc = ''
        self.colors = {}
        self.colors['grid_boundary'] = 'blue'
        self.grid2_colors = {'500': '#%02x%02x%02x' % (255, 237, 0),
                             '400': '#%02x%02x%02x' % (242, 170, 45),
                             '330': '#%02x%02x%02x' % (242, 170, 45),
                             '275': '#%02x%02x%02x' % (249, 13, 227),
                             '220': '#%02x%02x%02x' % (0, 0, 255),
                             '132': '#%02x%02x%02x' % (228, 2, 45),
                             '110': '#%02x%02x%02x' % (228, 2, 45),
                             '88': '#%02x%02x%02x' % (119, 65, 16),
                             '66': '#%02x%02x%02x' % (119, 65, 16),
                             'dc': '#%02x%02x%02x' % (86, 154, 105)}
        try:
            colours = config.items('Colors')
            for item, colour in colours:
                if item in ['grid_boundary', 'grid_trace']:
                    self.colors[item] = colour
                elif item[:5] == 'grid_':
                    self.colors[item[5:]] = colour
                elif item[:6] == 'grid2_':
                    self.grid2_colors[item[6:]] = colour
        except:
            pass
        if mapc != '':
            try:
                colours = config.items('Colors' + mapc)
                for item, colour in colours:
                    if item in ['grid_boundary', 'grid_trace']:
                        self.colors[item] = colour
                    elif item[:5] == 'grid_':
                        self.colors[item[5:]] = colour
                    elif item[:6] == 'grid2_':
                        self.grid2_colors[item[6:]] = colour
            except:
                pass
        upper_left = [0., 0.]
        lower_right = [-90., 180.]
        try:
             upper_left = config.get('Map', 'upper_left' + mapc).split(',')
             upper_left[0] = float(upper_left[0].strip())
             upper_left[1] = float(upper_left[1].strip())
             lower_right = config.get('Map', 'lower_right' + mapc).split(',')
             lower_right[0] = float(lower_right[0].strip())
             lower_right[1] = float(lower_right[1].strip())
        except:
             try:
                 lower_left = config.get('Map', 'lower_left' + mapc).split(',')
                 upper_right = config.get('Map', 'upper_right' + mapc).split(',')
                 upper_left[0] = float(upper_right[0].strip())
                 upper_left[1] = float(lower_left[1].strip())
                 lower_right[0] = float(lower_left[0].strip())
                 lower_right[1] = float(upper_right[1].strip())
             except:
                 pass
        self.map_polygon = [upper_left, [upper_left[0], lower_right[1]], lower_right,
           [lower_right[0], upper_left[1]], upper_left]
        self.default_length = -1
        try:
            trace_existing = config.get('View', 'trace_existing')
            if trace_existing.lower() in ['true', 'yes', 'on']:
                self.default_length = 0
        except:
            try:
                trace_existing = config.get('Grid', 'trace_existing')
                if trace_existing.lower() in ['true', 'yes', 'on']:
                   self.default_length = 0
            except:
                pass
        self.line_costs = {}
        try:
            line_costs = config.get('Grid', 'line_costs')
            line_costs = line_costs.replace('(', '')
            line_costs = line_costs.replace(')', '')
            line_costs = line_costs.split(',')
            for i in range(len(line_costs)):
                line = line_costs[i].split('=')
                if line[1].strip()[-1] == 'K':
                    self.line_costs[line[0].strip()] = float(line[1].strip()[:-1]) * pow(10, 3)
                elif line[1].strip()[-1] == 'M':
                    self.line_costs[line[0].strip()] = float(line[1].strip()[:-1]) * pow(10, 6)
                else:
                    self.line_costs[line[0].strip()] = float(line[1].strip()[:-1])
        except:
            pass
        self.substation_costs = {}
        try:
            substation_costs = config.get('Grid', 'substation_costs')
            substation_costs = substation_costs.replace('(', '')
            substation_costs = substation_costs.replace(')', '')
            substation_costs = substation_costs.split(',')
            for i in range(len(substation_costs)):
                line = substation_costs[i].split('=')
                if line[1].strip()[-1] == 'K':
                    self.substation_costs[line[0].strip()] = float(line[1].strip()[:-1]) * pow(10, 3)
                elif line[1].strip()[-1] == 'M':
                    self.substation_costs[line[0].strip()] = float(line[1].strip()[:-1]) * pow(10, 6)
                else:
                    self.substation_costs[line[0].strip()] = float(line[1].strip()[:-1])
        except:
            pass
        try:
            s_lines = config.get('Grid', 's_lines')
            s_lines = s_lines.strip()
            self.s_line_table = self.decode(s_lines)
        except:
            self.s_line_table = []
        try:
            d_lines = config.get('Grid', 'd_lines')
            d_lines = d_lines.strip()
            self.d_line_table = self.decode(d_lines)
        except:
            self.d_line_table = []
        self.dummy_fix = False
        try:
            dummy_fix = config.get('Grid', 'dummy_fix')
            if dummy_fix.lower() in ['true', 'yes', 'on']:
                self.dummy_fix = True
        except:
            pass

    def __init__(self, grid2=False):
        self.get_config()
        self.lines = []
        if grid2:
            kml_file = self.kml_file2
        else:
            kml_file = self.kml_file
        if not os.path.exists(kml_file):
            if grid2:
                self.kml_file = ''
            else:
                self.kml_file2 = ''
            return
        style = {}
        styl = ''
        zipped = False
        if kml_file[-4:] == '.kmz': # zipped file?
            zipped = True
            zf = zipfile.ZipFile(kml_file, 'r')
            inner_file = ''
            for name in zf.namelist():
                if name[-4:] == '.kml':
                    inner_file = name
                    break
            if inner_file == '':
                return
            memory_file = StringIO.StringIO()
            memory_file.write(zf.open(inner_file).read())
            root = ElementTree(fromstring(memory_file.getvalue()))
        else:
            kml_data = open(kml_file, 'rb')
            root = ElementTree(fromstring(kml_data.read()))
         # Create an iterator
        iterat = root.getiterator()
        placemark_id = ''
        line_names = []
        stylm = ''
        for element in iterat:
            elem = element.tag[element.tag.find('}') + 1:]
            if elem == 'Style':
                for name, value in element.items():
                    if name == 'id':
                        styl = value
            elif elem == 'StyleMap':
                for name, value in element.items():
                    if name == 'id':
                        stylm = value
            elif elem == 'color':
                if styl in self.colors:
                    style[styl] = self.colors[styl]
                else:
                    style[styl] = '#' + element.text[-2:] + element.text[-4:-2] + element.text[-6:-4]
                if stylm != '':
                    if stylm in self.colors:
                        style[stylm] = self.colors[stylm]
                    else:
                        style[stylm] = '#' + element.text[-2:] + element.text[-4:-2] + element.text[-6:-4]
            elif elem == 'name':
                line_name = element.text
                if placemark_id != '':
                    line_name += placemark_id
                    placemark_id = ''
            elif elem == 'Placemark' and grid2:
                for key, value in element.items():
                    if key == 'id':
                        if value[:4] == 'kml_':
                            placemark_id = value[3:]
                        else:
                            placemark_id = value
            elif elem == 'SimpleData' and grid2:
                for key, value in element.items():
                    if key == 'name' and (value == 'CAPACITY_kV' or value == 'CAPACITYKV'):
                        try:
                            styl = self.grid2_colors[element.text]
                        except:
                            styl = self.grid2_colors['66']
            elif elem == 'styleUrl':
                styl = element.text[1:]
            elif elem == 'coordinates':
                coords = []
                coordinates = ' '.join(element.text.split()).split(' ')
                for i in range(len(coordinates)):
                    coords.append([float(coordinates[i].split(',')[1]), float(coordinates[i].split(',')[0])])
                inmap = False
                for coord in coords:
                    if within_map(coord[0], coord[1], self.map_polygon):
                        inmap = True
                        break
                if inmap:
                    if self.default_length >= 0:
                        grid_len = 0.
                        for j in range(1, len(coords)):
                            grid_len += self.actualDistance(coords[j - 1][0], coords[j - 1][1],
                                        coords[j][0], coords[j][1])
                    else:
                        grid_len = self.default_length
                    if line_name in line_names:
                        i = 2
                        while line_name + '#' + str(i) in line_names:
                            i += 1
                        line_name += '#' + str(i)
                    line_names.append(line_name)
                    if grid2:
                        self.lines.append(Line(line_name, styl, coords, length=grid_len))
                    else:
                        try:
                            self.lines.append(Line(line_name, style[styl], coords, length=grid_len))
                        except:
                            style[styl] = '#FFFFFF'
                            self.lines.append(Line(line_name, style[styl], coords, length=grid_len))
        if zipped:
            memory_file.close()
            zf.close()
        else:
            kml_data.close()
     # connect together
     # if load_centres connect closest end to closest load centre
     #    for i in range(len(self.lines)):
     #        connect = []
     #        con = -1
     #        connect.append(self.gridConnect(self.lines[i].coordinates[0][0], self.lines[i].coordinates[0][1], \
     #                       ignore=[i]))
     #        connect.append(self.gridConnect(self.lines[i].coordinates[-1][0], self.lines[i].coordinates[-1][1], \
     #                       ignore=[i]))
     #        if connect[0][0] > 0:
     #            if connect[1][0] < connect[0][0]:
     #                con = connect[1][2]
     #                self.lines[i].coordinates.append([connect[1][1], connect[1][2]])
     #            else:
     #                con = connect[0][2]
     #                self.lines[i].coordinates.insert(0, [connect[0][1], connect[0][2]])
     #        self.lines[i].connector = con

    def gridConnect(self, lat, lon, ignore=[]):
        shortest = [99999, -1., -1., -1]
        for l in range(len(self.lines)):
            if l in ignore:
                continue
            for i in range(len(self.lines[l].coordinates) - 1):
                if self.dummy_fix:
                    dist = dust(lat, lon, self.lines[l].coordinates[i][0], self.lines[l].coordinates[i][1],
                           self.lines[l].coordinates[i + 1][0], self.lines[l].coordinates[i + 1][1])
                elif self.kml_file == '' and self.dummy_fix:
                    dist = dust(lat, lon, self.lines[l].coordinates[i][0], self.lines[l].coordinates[i][1],
                           self.lines[l].coordinates[i + 1][0], self.lines[l].coordinates[i + 1][1])
                else:
                    dist = self.DistancePointLine(lat, lon, self.lines[l].coordinates[i][0], self.lines[l].coordinates[i][1],
                           self.lines[l].coordinates[i + 1][0], self.lines[l].coordinates[i + 1][1])
                if dist[0] >= 0 and dist[0] < shortest[0]:
                    shortest = dist[:]
                    shortest.append(l)
        if shortest[0] == 99999:
             shortest[0] = -1
        return shortest   # length, lat, lon, line#

    def DistancePointLine(self, pyd, pxd, y1d, x1d, y2d, x2d):
# px,py is the point to test.
# x1,y1,x2,y2 is the line to check distance.
# Returns distance from the line, or if the intersecting point on the line nearest
# the point tested is outside the endpoints of the line, the distance to the
# nearest endpoint.
# Returns -1 on 0 denominator conditions.
        px = radians(pxd)
        py = radians(pyd)
        x1 = radians(x1d)
        y1 = radians(y1d)
        x2 = radians(x2d)
        y2 = radians(y2d)
        b13 = self.Bearing(y1, x1, py, px)
        b12 = self.Bearing(y1, x1, y2, x2)
        d13 = self.Distance(y1, x1, py, px)
        d23 = self.Distance(y2, x2, py, px)
        dxt = asin(sin(d13) * sin(b13 - b12))
        dat = acos(cos(d13) / cos(dxt))
        iy = asin(sin(y1) * cos(dat) + cos(y1) * sin(dat) * cos(b12))
        ix = x1 + atan2(sin(b12) * sin(dat) * cos(y1), cos(dat) - sin(y1) * sin(iy))
        if abs(ix - x1) > abs(x1 - x2) or abs(iy - y1) > abs(y1 - y2):
            dst = d13
            ix = x1
            iy = y1
        else:
            dst = self.Distance(iy, ix, py, px)
            if d13 < dst:   # must be another way but this'll do for now
                dst = d13
                ix = x1
                iy = y1
        if d23 < dst:   # must be another way but this'll do for now
            dst = d23
            ix = x2
            iy = y2
        return [round(abs(dst) * RADIUS, 2), round(degrees(iy), 6), round(degrees(ix), 6)]

    def Bearing(self, y1, x1, y2, x2):
# find the bearing between the coordinates
        return atan2(sin(x2 - x1) * cos(y2), cos(y1) * sin(y2) - sin(y1) * cos(y2) * cos(x2 - x1))

    def Distance(self, y1, x1, y2, x2):
# find the differences between the coordinates
        dy = y2 - y1
        dx = x2 - x1
        ra13 = pow(sin(dy / 2.), 2) + cos(y1) * cos(y2) * pow(sin(dx / 2.), 2)
        return 2 * asin(min(1, sqrt(ra13)))

    def actualDistance(self, y1d, x1d, y2d, x2d):
        x1 = radians(x1d)
        y1 = radians(y1d)
        x2 = radians(x2d)
        y2 = radians(y2d)
        dst = self.Distance(y1, x1, y2, x2)
        return round(abs(dst) * RADIUS, 2)

    def Line_Cost(self, peak_load, peak_dispatchable):
        if peak_dispatchable is None or peak_dispatchable == 0:
            s = 0
            while self.s_line_table[s][0] < peak_load and (s + 1) < len(self.s_line_table):
                s += 1
            return self.s_line_table[s][1], self.s_line_table[s][2]
        elif peak_load == peak_dispatchable:
            p = 0
            while self.d_line_table[p][0] < peak_load and (p + 1) < len(self.d_line_table):
                p += 1
            return self.d_line_table[p][1], self.d_line_table[p][2]
        else:
            if peak_load > peak_dispatchable + peak_dispatchable:
                p = 0
                while self.d_line_table[p][0] < peak_load and (p + 1) < len(self.d_line_table):
                    p += 1
                return self.d_line_table[p][1], self.d_line_table[p][2]
            p = 0
            while self.d_line_table[p][0] < peak_dispatchable and (p + 1) < len(self.d_line_table):
                p += 1
            s = 0
            while self.s_line_table[s][0] < peak_load - peak_dispatchable and (s + 1) < len(self.s_line_table):
                s += 1
            if self.d_line_table[p][0] >= self.s_line_table[s][0]:
                return self.d_line_table[p][1], self.d_line_table[p][2]
            else:
                return self.d_line_table[p][1], self.d_line_table[p][2]

    def Substation_Cost(self, line):
        if line in self.substation_costs:
            return self.substation_costs[line]
        return 0.


class Grid_Boundary:
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
            parents = getParents(config.items('Parents'))
        except:
            pass
        try:
            self.kml_file = config.get('Files', 'grid_boundary')
            for key, value in parents:
                self.kml_file = self.kml_file.replace(key, value)
            self.kml_file = self.kml_file.replace('$USER$', getUser())
            self.kml_file = self.kml_file.replace('$YEAR$', self.base_year)
        except:
            self.kml_file = ''
        self.colour = '#0000FF'
        try:
            mapc = config.get('Map', 'map_choice')
        except:
            mapc = ''
        try:
            self.colour = config.get('Colors', 'grid_boundary')
        except:
            pass
        if mapc != '':
            try:
                self.colour = config.get('Colors' + mapc, 'grid_boundary')
            except:
                pass
        upper_left = [0., 0.]
        lower_right = [-90., 180.]
        try:
             upper_left = config.get('Map', 'upper_left' + mapc).split(',')
             upper_left[0] = float(upper_left[0].strip())
             upper_left[1] = float(upper_left[1].strip())
             lower_right = config.get('Map', 'lower_right' + mapc).split(',')
             lower_right[0] = float(lower_right[0].strip())
             lower_right[1] = float(lower_right[1].strip())
        except:
             try:
                 lower_left = config.get('Map', 'lower_left' + mapc).split(',')
                 upper_right = config.get('Map', 'upper_right' + mapc).split(',')
                 upper_left[0] = float(upper_right[0].strip())
                 upper_left[1] = float(lower_left[1].strip())
                 lower_right[0] = float(lower_left[0].strip())
                 lower_right[1] = float(upper_right[1].strip())
             except:
                 pass
        self.map_polygon = [upper_left, [upper_left[0], lower_right[1]], lower_right,
           [lower_right[0], upper_left[1]], upper_left]

    def __init__(self):
        self.get_config()
        self.lines = []
        if not os.path.exists(self.kml_file):
            return
        style = {}
        styl = ''
        zipped = False
        if self.kml_file[-4:] == '.kmz': # zipped file?
            zipped = True
            zf = zipfile.ZipFile(kml_file, 'r')
            inner_file = ''
            for name in zf.namelist():
                if name[-4:] == '.kml':
                    inner_file = name
                    break
            if inner_file == '':
                return
            memory_file = StringIO.StringIO()
            memory_file.write(zf.open(inner_file).read())
            root = ElementTree(fromstring(memory_file.getvalue()))
        else:
            kml_data = open(self.kml_file, 'rb')
            root = ElementTree(fromstring(kml_data.read()))
         # Create an iterator
        iterat = root.getiterator()
        for element in iterat:
            elem = element.tag[element.tag.find('}') + 1:]
            if elem == 'Style':
                for name, value in element.items():
                    if name == 'id':
                        styl = value
            elif elem == 'color':
                style[styl] = self.colour
            elif elem == 'name':
                line_name = element.text
            elif elem == 'styleUrl':
                styl = element.text[1:]
            elif elem == 'coordinates':
                coords = []
                coordinates = ' '.join(element.text.split()).split(' ')
                for i in range(len(coordinates)):
                    coords.append([round(float(coordinates[i].split(',')[1]), 6),
                      round(float(coordinates[i].split(',')[0]), 6)])
                i = int(len(coords) / 2)
                if within_map(coords[0][0], coords[0][1], self.map_polygon) and \
                   within_map(coords[i][0], coords[i][1], self.map_polygon):
                    try:
                        self.lines.append(Line(line_name, style[styl], coords))
                    except:
                        self.lines.append(Line(line_name, self.colour, coords))
        if zipped:
            memory_file.close()
            zf.close()
        else:
            kml_data.close()
