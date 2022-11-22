#!/usr/bin/python3
#
#  Copyright (C) 2015-2022 Sustainable Energy Now Inc., Angus King
#
#  station.py - This file is part of SIREN.
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
from math import radians, cos, sin, asin, sqrt, pow

import configparser   # decode .ini file

from getmodels import getModelFile
from senutils import getParents, getUser, techClean, WorkBook

def within_map(y, x, poly):
    n = len(poly)
    inside = False
    p1y, p1x = poly[0]
    for i in range(n + 1):
        p2y, p2x = poly[i % n]
        if x > min(p1x, p2x):
            if x <= max(p1x, p2x):
                if y <= max(p1y, p2y):
                    if p1x != p2x:
                        yints = (x - p1x) * (p2y - p1y) / (p2x - p1x) + p1y
                    if p1y == p2y or y <= yints:
                        inside = not inside
        p1y, p1x = p2y, p2x
    return inside


class Station:
    def __init__(self, name, technology, lat, lon, capacity, turbine, rotor, no_turbines, area, scenario, generation=None,
                 power_file=None, grid_line=None, grid_len=None, grid_path_len=None, direction=None, tilt=None,
                 storage_hours=None, zone=None):
        self.name = name
        self.technology = technology
        self.lat = lat
        self.lon = lon
        self.capacity = capacity
        self.turbine = turbine
        self.rotor = rotor
        self.no_turbines = no_turbines
        self.area = area
        self.scenario = scenario
        self.generation = generation
        self.power_file = power_file
        self.grid_line = grid_line
        self.grid_len = grid_len
        self.grid_path_len = grid_path_len
        self.direction = direction
        self.storage_hours = storage_hours
        if tilt is not None:
            self.tilt = tilt
        self.zone = zone


class Stations:
    def get_config(self):
        config = configparser.RawConfigParser()
        if __name__ == '__main__':
            for i in range(1, len(sys.argv)):
                if sys.argv[i][-4:] == '.ini':
                    config_file = sys.argv[i]
                    break
        else:
            if len(sys.argv) > 1:
                config_file = sys.argv[1]
            else:
                config_file = getModelFile('SIREN.ini')
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
        self.fac_files = []
        try:
            fac_file = config.get('Files', 'grid_stations')
            for key, value in parents:
                fac_file = fac_file.replace(key, value)
            fac_file = fac_file.replace('$USER$', getUser())
            fac_file = fac_file.replace('$YEAR$', self.base_year)
            self.fac_files.append(fac_file)
        except:
            pass
        if self.stations2:
            try:
                fac_file = config.get('Files', 'grid_stations2')
                for key, value in parents:
                    fac_file = fac_file.replace(key, value)
                fac_file = fac_file.replace('$USER$', getUser())
                fac_file = fac_file.replace('$YEAR$', self.base_year)
                self.fac_files.append(fac_file)
            except:
                pass
        self.ignore_deleted = True
        try:
            if config.get('Grid', 'ignore_deleted_existing').lower() in ['false', 'off', 'no']:
                self.ignore_deleted = False
        except:
            pass
        self.technologies = ['']
        self.areas = {}
        try:
            technologies = config.get('Power', 'technologies')
            for item in technologies.split():
                itm = techClean(item)
                self.technologies.append(itm)
                try:
                    self.areas[itm] = float(config.get(itm, 'area'))
                except:
                    self.areas[itm] = 0.
        except:
            pass
        try:
            technologies = config.get('Power', 'fossil_technologies')
            technologies = technologies.split()
            for item in technologies:
                itm = techClean(item)
                try:
                    self.areas[itm] = float(config.get(itm, 'area'))
                except:
                    self.areas[itm] = 0.
        except:
            pass
        try:
            mapc = config.get('Map', 'map_choice')
        except:
            mapc = ''
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

    def haversine(self, lat1, lon1, lat2, lon2):
        """
        Calculate the great circle distance between two points
        on the earth (specified in decimal degrees)
        """
     # convert decimal degrees to radians
        lon1, lat1, lon2, lat2 = list(map(radians, [lon1, lat1, lon2, lat2]))

     # haversine formula
        dlon = lon2 - lon1
        dlat = lat2 - lat1
        a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
        c = 2 * asin(sqrt(a))

     # 6367 km is the radius of the Earth
        km = 6367 * c
        return km

    def __init__(self, code=False, fossil=False, existing=True, stations2=True):
        self.stations2 = stations2
        self.get_config()
        try:
            self.stations in locals()
        except:
            self.stations = []
        if not existing:
            return
        if not os.path.exists(self.sam_file):
            return
        sam = open(self.sam_file)
        sam_turbines = csv.DictReader(sam)
        for fac_file in self.fac_files:
            if os.path.exists(fac_file):
                if fac_file[-4:] == '.csv':
                    facile = open(fac_file)
                    facilities = csv.DictReader(facile)
                    for facility in facilities:
                        if self.ignore_deleted and facility['Balancing Status'] == 'Deleted':
                            continue
                        if facility['Longitude'] != '':
                            if not within_map(float(facility['Latitude']),
                              float(facility['Longitude']), self.map_polygon):
                                continue
                            if 'Facility Code' in facilities.fieldnames:   # IMO format
                                bit = facility['Facility Code'].split('_')
                                rotor = 0.
                                turbine = ''
                                no_turbines = 0
                                if bit[-1][:2] == 'WW' or bit[-1][:2] == 'WF':
                                    tech = 'Wind'
                                    turbine = facility['Turbine']
                                    if turbine[:7] == 'Enercon':
                                        bit = turbine[9:].split(' _')
                                        if len(bit) == 1:
                                            rotor = bit[0].split('_')[0]
                                        else:
                                            bit = bit[0].split('_')
                                            if len(bit) == 1:
                                                rotor = bit[0]
                                            else:
                                                if bit[1][-1] == 'm':
                                                    rotor = bit[1][:-1]
                                                else:
                                                    rotor = bit[0]
                                    else:
                                        turb = turbine.split(';')
                                        if len(turb) > 1:
                                            turbine = turb[1]
                                        else:
                                            turbine = turb[0]
                                        sam.seek(0)
                                        for turb in sam_turbines:
                                            if turb['Name'] == turbine:
                                                rotor = turb['Rotor Diameter']
                                                break
                                        else: # try and find .pow file
                                            pow_file = self.pow_dir + '/' + turbine + '.pow'
                                            if os.path.exists(pow_file):
                                                tf = open(pow_file, 'r')
                                                lines = tf.readlines()
                                                tf.close()
                                                rotor = float(lines[1].strip('" \t\n\r'))
                                                del lines
                                    no_turbines = int(facility['No. turbines'])
                                    try:
                                        rotor = float(rotor)
                                    except:
                                        pass
                                    area = self.areas[tech] * float(no_turbines) * pow((rotor * .001), 2)
                                elif bit[0] == 'MERSOLAR':
                                    tech = 'Single Axis PV'
                                    area = self.areas[tech] * float(facility['Maximum Capacity (MW)'])
                                elif bit[-1][:2] == 'PV':
                                    tech = 'Fixed PV'
                                    area = self.areas[tech] * float(facility['Maximum Capacity (MW)'])
                                elif bit[-1] == 'PLANT' and bit[-2] == 'BIOMASS':
                                    tech = 'Biomass'
                                    area = self.areas[tech] * float(facility['Maximum Capacity (MW)'])
                                elif facility['Participant Code'] in ['CTE', 'LNDFLLGP', 'PERTHNRGY', 'WGRES']:
                                    tech = 'Biomass'
                                    area = self.areas[tech] * float(facility['Maximum Capacity (MW)'])
                                else:
                                    tech = 'Fossil '
                                    if 'Fossil' in facilities.fieldnames:
                                        tech += facility['Fossil']
                                    else:
                                        if bit[-1][:2] == 'GT' or bit[-1][0] == 'U':
                                            tech += 'OCGT'
                                        elif bit[-1][:2] == 'CC':
                                            tech += 'CCGT'
                                        elif bit[-1][:3] == 'COG':
                                            tech += 'Cogen'
                                        elif bit[0] == 'TESLA' or bit[-1] == 'WGP':
                                            tech += 'Distillate'
                                        elif bit[-2] == 'WGP':
                                            tech += 'Mixed'
                                        elif bit[-1][0] == 'G' and (bit[-1][1] >= '1' and bit[-1][1] <= '9'):
                                            tech += 'Coal'
                                        else:
                                            tech += 'Mixed'
                                           #  tech += 'Distillate'
                                    try:
                                        area = self.areas[tech] * float(facility['Maximum Capacity (MW)'])
                                    except:
                                        area = 0
                                if code:
                                    nice_name = facility['Facility Code']
                                else:
                                    nice_name = facility['Facility Name']
                                    if nice_name == '':
                                        name_split = facility['Facility Code'].split('_')
                                        if len(name_split) > 1:
                                            nice_name = ''
                                            for i in range(len(name_split) - 1):
                                                nice_name += name_split[i].title() + '_'
                                            nice_name += name_split[-1]
                                        else:
                                            nice_name = facility['Facility Code']
                                stn = self.Get_Station(nice_name)
                                if stn is None:   # new station?
                                    self.stations.append(Station(nice_name, tech,
                                        float(facility['Latitude']), float(facility['Longitude']),
                                        float(facility['Maximum Capacity (MW)']), turbine, rotor, no_turbines, area, 'Existing'))
                                    if tech == 'Fixed PV':
                                        try:
                                            if facility['Tilt'] != '':
                                                self.stations[-1].tilt = float(facility['Tilt'])
                                        except:
                                            pass
                                else:   # additional generator in existing station
                                    if stn.technology != tech:
                                        if stn.technology[:6] == 'Fossil' and tech[:6] == 'Fossil':
                                            stn.technology = 'Fossil Mixed'
                                    stn.capacity = stn.capacity + float(facility['Maximum Capacity (MW)'])
                                    stn.area += area
                                    stn.no_turbines = stn.no_turbines + no_turbines
                            else:   # SIREN format
                                try:
                                    turbs = int(facility['No. turbines'])
                                except:
                                    turbs = 0
                                self.stations.append(Station(facility['Station Name'],
                                             facility['Technology'],
                                             float(facility['Latitude']),
                                             float(facility['Longitude']),
                                             float(facility['Maximum Capacity (MW)']),
                                             facility['Turbine'],
                                             0.,
                                             turbs,
                                             float(facility['Area']),
                                             'Existing'))
                                if 'Wind' in self.stations[-1].technology:
                                    if self.stations[-1].rotor == 0 or self.stations[-1].rotor == '':
                                        rotor = 0
                                        if self.stations[-1].turbine[:7] == 'Enercon':
                                            bit = self.stations[-1].turbine[9:].split(' _')
                                            if len(bit) == 1:
                                                rotor = bit[0].split('_')[0]
                                            else:
                                                bit = bit[0].split('_')
                                                if len(bit) == 1:
                                                    rotor = bit[0]
                                                else:
                                                    if bit[1][-1] == 'm':
                                                        rotor = bit[1][:-1]
                                                    else:
                                                        rotor = bit[0]
                                        else:
                                            turb = self.stations[-1].turbine.split(';')
                                            if len(turb) > 1:
                                                sam.seek(0)
                                                turbine = turb[1]
                                                for turb in sam_turbines:
                                                    if turb['Name'] == turbine:
                                                        rotor = turb['Rotor Diameter']
                                    try:
                                        self.stations[-1].rotor = float(rotor)
                                    except:
                                        pass
                                if self.stations[-1].area == 0 or self.stations[-1].area == '':
                                    if 'Wind' in self.stations[-1].technology:
                                        self.stations[-1].area = self.areas[self.stations[-1].technology] * \
                                                                 float(self.stations[-1].no_turbines) * \
                                                                 pow((self.stations[-1].rotor * .001), 2)
                                    else:
                                        self.stations[-1].area = self.areas[self.stations[-1].technology] * \
                                                                 float(self.stations[-1].capacity)
                                try:
                                    if facility['Power File'] != '':
                                        self.stations[-1].power_file = facility['Power File']
                                except:
                                    pass
                                try:
                                    if facility['Grid Line'] != '':
                                        self.stations[-1].grid_line = facility['Grid Line']
                                except:
                                    pass
                                if 'PV' in self.stations[-1].technology:
                                    try:
                                        if facility['Direction'] != '':
                                            self.stations[-1].direction = facility['Direction']
                                    except:
                                        pass
                                    try:
                                        if facility['Tilt'] != '':
                                            self.stations[-1].tilt = float(facility['Tilt'])
                                    except:
                                        pass
                                if self.stations[-1].technology in ['CST', 'Solar Thermal']:
                                    try:
                                        if facility['Storage Hours'] != '':
                                            self.stations[-1].storage_hours = float(facility['Storage Hours'])
                                    except:
                                        pass
                    facile.close()
                else:   # assume excel and in our format
                    var = {}
                    workbook = WorkBook()
                    workbook.open_workbook(fac_file)
                    worksheet = workbook.sheet_by_index(0)
                    num_rows = worksheet.nrows - 1
                    num_cols = worksheet.ncols - 1
                    if worksheet.cell_value(0, 0) == 'Description:' or worksheet.cell_value(0, 0) == 'Comment:':
                        curr_row = 1
                        self.description = worksheet.cell_value(0, 1)
                    else:
                        curr_row = 0
                        self.description = ''
#                   get column names
                    curr_col = -1
                    while curr_col < num_cols:
                        curr_col += 1
                        var[worksheet.cell_value(curr_row, curr_col)] = curr_col
                    while curr_row < num_rows:
                        curr_row += 1
                        if not within_map(worksheet.cell_value(curr_row, var['Latitude']),
                          worksheet.cell_value(curr_row, var['Longitude']), self.map_polygon):
                            continue
                        self.stations.append(Station(str(worksheet.cell_value(curr_row, var['Station Name'])),
                                             str(worksheet.cell_value(curr_row, var['Technology'])),
                                             worksheet.cell_value(curr_row, var['Latitude']),
                                             worksheet.cell_value(curr_row, var['Longitude']),
                                             worksheet.cell_value(curr_row, var['Maximum Capacity (MW)']),
                                             str(worksheet.cell_value(curr_row, var['Turbine'])),
                                             worksheet.cell_value(curr_row, var['Rotor Diam']),
                                             worksheet.cell_value(curr_row, var['No. turbines']),
                                             worksheet.cell_value(curr_row, var['Area']),
                                             'Existing'))
                        if 'Wind' in self.stations[-1].technology:
                            if self.stations[-1].rotor == 0 or self.stations[-1].rotor == '':
                                rotor = 0
                                if self.stations[-1].turbine[:7] == 'Enercon':
                                    bit = self.stations[-1].turbine[9:].split(' _')
                                    if len(bit) == 1:
                                        rotor = bit[0].split('_')[0]
                                    else:
                                        bit = bit[0].split('_')
                                        if len(bit) == 1:
                                            rotor = bit[0]
                                        else:
                                            if bit[1][-1] == 'm':
                                                rotor = bit[1][:-1]
                                            else:
                                                rotor = bit[0]
                                else:
                                    turb = self.stations[-1].turbine.split(';')
                                    if len(turb) > 1:
                                        sam.seek(0)
                                        turbine = turb[1]
                                        for turb in sam_turbines:
                                            if turb['Name'] == turbine:
                                                rotor = turb['Rotor Diameter']
                                try:
                                    self.stations[-1].rotor = float(rotor)
                                except:
                                    pass
                        try:
                            if self.stations[-1].area == 0 or self.stations[-1].area == '':
                                if 'Wind' in self.stations[-1].technology:
                                    self.stations[-1].area = self.areas[self.stations[-1].technology] * \
                                                             float(self.stations[-1].no_turbines) * \
                                                             pow((self.stations[-1].rotor * .001), 2)
                                else:
                                    self.stations[-1].area = self.areas[self.stations[-1].technology] * \
                                                             float(self.stations[-1].capacity)
                        except:
                            self.stations[-1].area = 0.
                        try:
                            power_file = worksheet.cell_value(curr_row, var['Power File'])
                            if power_file != '':
                                self.stations[-1].power_file = power_file
                        except:
                            pass
                        try:
                            grid_line = worksheet.cell_value(curr_row, var['Grid Line'])
                            if grid_line != '':
                                self.stations[-1].grid_line = grid_line
                        except:
                            pass
                        if 'PV' in self.stations[-1].technology:
                            try:
                                direction = worksheet.cell_value(curr_row, var['Direction'])
                                if direction != '':
                                    self.stations[-1].direction = direction
                            except:
                                pass
                            try:
                                tilt = worksheet.cell_value(curr_row, var['Tilt'])
                                if tilt != '':
                                    self.stations[-1].tilt = tilt
                            except:
                                pass
                        if self.stations[-1].technology in ['CST', 'Solar Thermal']:
                            try:
                                storage_hours = worksheet.cell_value(curr_row, var['Storage Hours'])
                                if storage_hours != '':
                                    self.stations[-1].storage_hours = storage_hours
                            except:
                                pass
        sam.close()

    def Nearest(self, lat, lon, distance=False, fossil=False, ignore=None):
        hdr = ''
        distnce = 999999
        for station in self.stations:
            if station.technology[:6] == 'Fossil' and not fossil:
                continue
            dist = self.haversine(lat, lon, station.lat, station.lon)
            if dist < distnce:
                if ignore is not None and ignore == station.name:
                    continue
                hdr = station.name
                distnce = dist
        for station in self.stations:
            if station.name == hdr:
                if distance:
                    return station, distnce
                else:
                    return station
        return None

    def Stn_Location(self, name):
        for station in self.stations:
            if station.name == name:
                return str(station.lat) + ' ' + str(station.lon)
        return ''

    def Stn_Turbine(self, name):
        for station in self.stations:
            if station.name == name:
                return station.turbine
        return ''

    def Get_Station(self, name):
        for station in self.stations:
            if station.name == name:
                return station
        return None

    def Description(self):
        return self.description
