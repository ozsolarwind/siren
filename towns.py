#!/usr/bin/python
#
#  Copyright (C) 2015 Sustainable Energy Now Inc., Angus King
#
#  towns.py - This file is part of SIREN.
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

import ConfigParser    # decode .ini file
import csv
import os
import sys
import xlrd
from math import radians, cos, sin, asin, sqrt

from senuser import getUser


class Town:
    def __init__(self, name, lat, lon, lid='??', state='WA', country='AUS', elev=0, zone=8):
        self.name = name.strip().title()
        self.lat = float(lat)
        self.lon = float(lon)
        if isinstance(lid, float):
            self.lid = str(int(lid))
        else:
            self.lid = str(lid)
        self.state = state.strip()
        self.country = country
        self.elev = elev
        self.zone = zone


class Towns:
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
            self.bom_file = config.get('Files', 'bom')
            for key, value in parents:
                self.bom_file = self.bom_file.replace(key, value)
            self.bom_file = self.bom_file.replace('$USER$', getUser())
            self.bom_file = self.bom_file.replace('$YEAR$', self.base_year)
        except:
            self.bom_file = ''
        try:
            self.town_file = config.get('Files', 'towns')
            for key, value in parents:
                self.town_file = self.town_file.replace(key, value)
            self.town_file = self.town_file.replace('$USER$', getUser())
            self.town_file = self.town_file.replace('$YEAR$', self.base_year)
        except:
            self.town_file = ''

    def haversine(self, lat1, lon1, lat2, lon2):
        """
        Calculate the great circle distance between two points
        on the earth (specified in decimal degrees)
        """
     # convert decimal degrees to radians
        lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])

     # haversine formula
        dlon = lon2 - lon1
        dlat = lat2 - lat1
        a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
        c = 2 * asin(sqrt(a))

     # 6367 km is the radius of the Earth
        km = 6367 * c
        return km

#   might extend capability to restrict towns to network boundary
    def town_in_boundary(self, x, y, poly):
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

    def __init__(self, ul_lat=None, ul_lon=None, lr_lat=None, lr_lon=None, remove_duplicates=True):
        """Initializes the data."""

        def read_town_csv(town_file):
            townfile = open(town_file)
            twns = csv.DictReader(townfile)
            if 'Town' in twns.fieldnames:
                name_field = 'Town'
            else:
                name_field = 'Site name'
            if remove_duplicates:
                for twn in twns:
                    for town in self.towns:
                        if twn[name_field] == town.name:
                            break
                        if len(twn[name_field]) < len(town.name) and town.name.find(' ') > 0:
                            if twn[name_field].title() == town.name[:len(twn[name_field])]:
                                break
                            if town.name[:6].title() == 'North ' and twn[name_field] == town.name[6:]:
                                break
                        if len(twn[name_field]) > len(town.name) and twn[name_field].find(' ') > 0:
                            if twn[name_field].title()[:len(town.name)] == town.name:
                                break
                    else:
                        if ul_lat is None or \
                          (float(twn['Latitude']) >= lr_lat and
                           float(twn['Latitude']) <= ul_lat and
                           float(twn['Longitude']) >= ul_lon and
                           float(twn['Longitude']) <= lr_lon):
                            if 'Lid' in twns.fieldnames and twn['Lid'] != '':
                                self.towns.append(Town(twn[name_field], twn['Latitude'], twn['Longitude'], twn['Lid'],
                                             twn['State'], twn['Country'], twn['Elev'], twn['Zone']))
                            else:
                                self.towns.append(Town(twn[name_field], twn['Latitude'], twn['Longitude']))
            else:
                for twn in twns:
                    if ul_lat is None or \
                      (float(twn['Latitude']) >= lr_lat and
                       float(twn['Latitude']) <= ul_lat and
                       float(twn['Longitude']) >= ul_lon and
                       float(twn['Longitude']) <= lr_lon):
                        if 'Lid' in twns.fieldnames and twn['Lid'] != '':
                            self.towns.append(Town(twn[name_field], twn['Latitude'], twn['Longitude'], twn['Lid'],
                                        twn['State'], twn['Country'], twn['Elev'], twn['Zone']))
                        else:
                            self.towns.append(Town(twn[name_field], twn['Latitude'], twn['Longitude']))
            townfile.close()

        self.get_config()
        self.towns = []
#   Process BOM stations first
        if os.path.exists(self.bom_file):
            if self.bom_file[-4:] == '.csv':
                read_town_csv(self.bom_file)
            else:
                c_closed = -1
                c_id = -1
                c_nme = -1
                c_state = -1
                c_lat = -1
                c_lon = -1
                c_elev = -1
                c_ctry = -1
                workbook = xlrd.open_workbook(self.bom_file)
                worksheet = workbook.sheet_by_index(0)
                num_rows = worksheet.nrows - 1
                num_cols = worksheet.ncols - 1
#     get column names
                curr_col = -1
                while curr_col < num_cols:
                    curr_col += 1
                    if worksheet.cell_value(0, curr_col) == 'Bureau of Meteorology Station Number' or \
                       worksheet.cell_value(0, curr_col) == 'Lid':
                        c_id = curr_col
                    elif worksheet.cell_value(0, curr_col) == 'Station Name' or \
                       worksheet.cell_value(0, curr_col) == 'Site Name' or \
                       worksheet.cell_value(0, curr_col) == 'Town':
                        c_nme = curr_col
                    elif worksheet.cell_value(0, curr_col) == 'Latitude to 4 decimal places, in decimal degrees' or \
                       worksheet.cell_value(0, curr_col) == 'Latitude':
                        c_lat = curr_col
                    elif worksheet.cell_value(0, curr_col) == 'Longitude to 4 decimal places, in decimal degrees' or \
                       worksheet.cell_value(0, curr_col) == 'Longitude':
                        c_lon = curr_col
                    elif worksheet.cell_value(0, curr_col) == 'State':
                        c_state = curr_col
                    elif worksheet.cell_value(0, curr_col) == 'Height of station above mean sea level in metres' or \
                       worksheet.cell_value(0, curr_col) == 'Elev':
                        c_elev = curr_col
                    elif worksheet.cell_value(0, curr_col) == 'Month/Year site closed. (MM/YYYY)':
                        c_closed = curr_col
                    elif worksheet.cell_value(0, curr_col) == 'Country':
                        c_ctry = curr_col
#     WMO (World Meteorological Organisation) Index Number
                curr_row = 0
                while curr_row < num_rows:
                    curr_row += 1
                    if c_closed >= 0:
                        if worksheet.cell_value(curr_row, c_nme) != '' \
                        and worksheet.cell_value(curr_row, c_closed).strip() == '':
                            if ul_lat is None or \
                              (worksheet.cell_value(curr_row, c_lat) >= lr_lat and
                               worksheet.cell_value(curr_row, c_lat) <= ul_lat and
                               worksheet.cell_value(curr_row, c_lon) >= ul_lon and
                               worksheet.cell_value(curr_row, c_lon) <= lr_lon):
                                self.towns.append(Town(worksheet.cell_value(curr_row, c_nme),
                                                       worksheet.cell_value(curr_row, c_lat),
                                                       worksheet.cell_value(curr_row, c_lon),
                                                       worksheet.cell_value(curr_row, c_id),
                                                       worksheet.cell_value(curr_row, c_state),
                                                       elev=worksheet.cell_value(curr_row, c_elev)))
                    else:
                        if worksheet.cell_value(curr_row, c_nme) != '':
                            if ul_lat is None or \
                              (worksheet.cell_value(curr_row, c_lat) >= lr_lat and
                               worksheet.cell_value(curr_row, c_lat) <= ul_lat and
                               worksheet.cell_value(curr_row, c_lon) >= ul_lon and
                               worksheet.cell_value(curr_row, c_lon) <= lr_lon):
                                self.towns.append(Town(worksheet.cell_value(curr_row, c_nme),
                                                       worksheet.cell_value(curr_row, c_lat),
                                                       worksheet.cell_value(curr_row, c_lon)))
                                if c_id >= 0:
                                    self.towns[-1].lid = worksheet.cell_value(curr_row, c_id)
                                if c_state >= 0:
                                    self.towns[-1].state = worksheet.cell_value(curr_row, c_state)
                                if c_elev >= 0:
                                    self.towns[-1].elev = worksheet.cell_value(curr_row, c_elev)
                                if c_ctry >= 0:
                                    self.towns[-1].country = worksheet.cell_value(curr_row, c_ctry)
# Process list of towns second
        if os.path.exists(self.town_file):
            read_town_csv(self.town_file)

    def SAM_Header(self, lat, lon):
#   SAM CSV (Solar)
#   Location ID,City,State,Country,Latitude,Longitude,Elevation,Time Zone,Source
        hdr = ''
        distance = 999999
        for town in self.towns:
            dist = self.haversine(lat, lon, town.lat, town.lon)
            if dist < distance:
                hdr = 'Location ID,City,State,Country,Latitude,Longitude,Time Zone,Elevation,Source\n' + \
                    '%s,%s,%s,%s,%s,%s,%s,%s,IWEC' % (town.lid, town.name, town.state, town.country,
                    lat, lon, town.zone, town.elev)
                distance = dist
        return hdr

    def SMW_Header(self, lat, lon, year=2014):
#   SAM SMW (Solar)
#   Location ID,Name,State,Time Zone,Latitude,Longitude,Elevation,Time Step,Start year,Start time
        hdr = ''
        distance = 999999
        for town in self.towns:
            dist = self.haversine(lat, lon, town.lat, town.lon)
            if dist < distance:
                hdr = '%s,"%s",%s,%s,%s,%s,%s,3600.0,%s,0:30:00' % (town.lid, town.name, town.state, town.zone,
                    lat, lon, town.elev, year)
                distance = dist
        return hdr

    def SRW_Header(self, lat, lon, year=2014):
#   SRW (Wind)
#   <location id>,<city>,<state>,<country>,<year>,<latitude>,<longitude>,<elevation>,<time step in hours>,<number of rows>
        hdr = ''
        distance = 999999
        for town in self.towns:
            dist = self.haversine(lat, lon, town.lon, town.lat)
            if dist < distance:
                hdr = '%s,%s,%s,%s,%s,%s,%s,%s,1,8760' % (town.lid, town.name, town.state, town.country,
                    year, round(lat, 4), round(lon, 4), town.elev)
                distance = dist
        return hdr

    def Nearest(self, lat, lon, distance=False):
        the_town = ''
        distnce = 999999
        for twn in self.towns:
            dist = self.haversine(lat, lon, twn.lat, twn.lon)
            if dist < distnce:
                the_town = twn.name
                distnce = dist
        for twn in self.towns:
            if twn.name == the_town:
                if distance:
                    return twn, distnce
                else:
                    return twn
        return

    def Stn_Location(self, lid):
        for town in self.towns:
            if town.lid == lid.lstrip('0'):
                return str(town.lat) + ' ' + str(town.lon)
        return ''

    def Get_Town(self, name):
        for town in self.towns:
            if town.name == name:
                return town
        return ''
