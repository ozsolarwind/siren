#!/usr/bin/python3
#
#  Copyright (C) 2015-2022 Sustainable Energy Now Inc., Angus King
#
#  superpower.py - This file is part of SIREN.
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

from math import asin, ceil, cos, fabs, pow, radians, sin, sqrt, floor
import csv
import os
import sys
import ssc
import time

import configparser  # decode .ini file
from PyQt5 import Qt, QtCore, QtGui, QtWidgets

from getmodels import getModelFile
from senutils import getParents, getUser, techClean, extrapolateWind, WorkBook
from powerclasses import *
# import Station
from turbine import Turbine

import tempfile # for wind extrapolate

the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

class SuperPower():
    log = QtCore.pyqtSignal()
    log2 = QtCore.pyqtSignal()
    barProgress = QtCore.pyqtSignal(int, str)
    barRange = QtCore.pyqtSignal(int, int)

    def haversine(self, lat1, lon1, lat2, lon2):
        """
        Calculate the great circle distance between two points
        on the earth (specified in decimal degrees)
        """
   #     convert decimal degrees to radians
        lon1, lat1, lon2, lat2 = list(map(radians, [lon1, lat1, lon2, lat2]))

   #     haversine formula
        dlon = lon2 - lon1
        dlat = lat2 - lat1
        a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
        c = 2 * asin(sqrt(a))

   #     6367 km is the radius of the Earth
        km = 6367 * c
        return km

    def find_closest(self, latitude, longitude, wind=False):
        dist = 99999
        closest = ''
        if wind:
            filetype = ['.srw']
            technology = 'wind_index'
            index_file = self.wind_index
            folder = self.wind_files
        else:
            filetype = ['.csv', '.smw']
            technology = 'solar_index'
            index_file = self.solar_index
            folder = self.solar_files
        if index_file == '':
            fils = os.listdir(folder)
            for fil in fils:
                if fil[-4:] in filetype:
                    bit = fil.split('_')
                    if bit[-1][:4] == self.base_year:
                        dist1 = self.haversine(float(bit[-3]), float(bit[-2]), latitude, longitude)
                        if dist1 < dist:
                            closest = fil
                            dist = dist1
        else:
            fils = []
            if self.default_files[technology] is None:
                dft_file = index_file
                if os.path.exists(dft_file):
                    pass
                else:
                    dft_file = folder + '/' + index_file
                if os.path.exists(dft_file):
                    self.default_files[technology] = WorkBook()
                    self.default_files[technology].open_workbook(dft_file, )
                else:
                    return closest
            var = {}
            worksheet = self.default_files[technology].sheet_by_index(0)
            num_rows = worksheet.nrows - 1
            num_cols = worksheet.ncols - 1
#           get column names
            curr_col = -1
            while curr_col < num_cols:
                curr_col += 1
                var[worksheet.cell_value(0, curr_col)] = curr_col
            curr_row = 0
            while curr_row < num_rows:
                curr_row += 1
                lat = worksheet.cell_value(curr_row, var['Latitude'])
                lon = worksheet.cell_value(curr_row, var['Longitude'])
                fil = worksheet.cell_value(curr_row, var['Filename'])
                fils.append([lat, lon, fil])
            for fil in fils:
                dist1 = self.haversine(fil[0], fil[1], latitude, longitude)
                if dist1 < dist:
                    closest = fil[2]
                    dist = dist1
        if __name__ == '__main__':
            print(closest)
        return closest

    def do_defaults(self, station):
        if 'PV' in station.technology:
            technology = 'PV'
        elif 'Wind' in station.technology:
            technology = 'Wind'
        else:
            technology = station.technology
        if self.default_files[technology] is None:
            dft_file = self.variable_files + '/' + self.defaults[technology]
            if os.path.exists(dft_file):
                self.default_files[technology] = WorkBook()
                self.default_files[technology].open_workbook(dft_file)
            else:
                return
        var = {}
        worksheet = self.default_files[technology].sheet_by_index(0)
        num_rows = worksheet.nrows - 1
        num_cols = worksheet.ncols - 1
#       get column names
        curr_col = -1
        while curr_col < num_cols:
            curr_col += 1
            var[worksheet.cell_value(0, curr_col)] = curr_col
        curr_row = 0
        while curr_row < num_rows:
            curr_row += 1
            if (worksheet.cell_value(curr_row, var['TYPE']) == 'SSC_INPUT' or \
              worksheet.cell_value(curr_row, var['TYPE']) == 'SSC_INOUT') and \
              worksheet.cell_value(curr_row, var['DEFAULT']) != '' and \
              str(worksheet.cell_value(curr_row, var['DEFAULT'])).lower() != 'input':
                if worksheet.cell_value(curr_row, var['DATA']) == 'SSC_STRING':
                    self.data.set_string(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'),
                      worksheet.cell_value(curr_row, var['DEFAULT']).encode('utf-8'))
                elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_ARRAY':
                    arry = split_array(worksheet.cell_value(curr_row, var['DEFAULT']))
                    self.data.set_array(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'), arry)
                elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_NUMBER':
                    if isinstance(worksheet.cell_value(curr_row, var['DEFAULT']), float):
                        self.data.set_number(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'),
                          float(worksheet.cell_value(curr_row, var['DEFAULT'])))
                    else:
                        self.data.set_number(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'),
                          worksheet.cell_value(curr_row, int(var['DEFAULT'])))
                elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_MATRIX':
                    mtrx = split_matrix(worksheet.cell_value(curr_row, var['DEFAULT']))
                    self.data.set_matrix(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'), mtrx)

    def debug_sam(self, name, tech, module, data, status):
        data_typs = ['invalid', 'string', 'number', 'array', 'matrix', 'table']
        var_typs = ['?', 'input', 'output', 'inout']
        var_names = []
        ssc_info = ssc.Info(module)
        while ssc_info.get():
            var_names.append([ssc_info.name(), ssc_info.var_type(), ssc_info.data_type()])
        if status:
            status.log('SAM Variable list for ' + tech + ' - ' + name)
        else:
            print('SAM Variable list for ' + tech + ' - ' + name)
        var_names = sorted(var_names, key=lambda s: s[0].lower())
        info = []
        for fld in var_names:
            msg = fld[0].decode() + ',' + var_typs[fld[1]] + ',' + data_typs[fld[2]] + ','
            if fld[2] == 1:
                msg += data.get_string(fld[0]).decode()
            elif fld[2] == 2:
                msg += str(data.get_number(fld[0]))
            elif fld[2] == 3:
                dat = data.get_array(fld[0])
                if len(dat) == 0:
                    msg += '"[]"'
                else:
                    msg += '"['
                    stop = len(dat)
                    if stop > 12:
                        fin = '...] ' + str(stop) + ' items in total"'
                        stop = 13
                    else:
                        fin = ']"'
                    for i in range(stop):
                        msg += str(dat[i]) + ','
                    msg = msg[:-1]
                    msg += fin
            elif fld[2] == 4:
                msg += str(len(data.get_matrix(fld[0]))) + ' entries'
            info.append(msg)
        if status:
            for msg in info:
                status.log2(msg)
            status.log('Variable list complete for ' + tech + ' - ' + name)
        else:
            for msg in info:
                print(msg)
            print('Variable list complete for ' + tech + ' - ' + name)
        return

    def __init__(self, stations, plots, parent=None, year=None, selected=None, status=None, progress=None):
        self.stations = stations
        self.plots = plots
        self.power_summary = []
        self.selected = selected
        self.status = status
        self.progress = progress
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = getModelFile('SIREN.ini')
        config.read(config_file)
        self.expert = False
        try:
            expert = config.get('Base', 'expert_mode')
            if expert.lower() in ['true', 'on', 'yes']:
                self.expert = True
        except:
            pass
        if year is None:
            try:
                self.base_year = config.get('Base', 'year')
            except:
                self.base_year = '2012'
        else:
            self.base_year = year
        parents = []
        try:
            parents = getParents(config.items('Parents'))
        except:
            pass
        try:
            self.biomass_multiplier = float(config.get('Biomass', 'multiplier'))
        except:
            self.biomass_multiplier = 8.25
        try:
            resource = config.get('Geothermal', 'resource')
            if resource.lower()[0:1] == 'hy':
                self.geo_res = 0
            else:
                self.geo_res = 1
        except:
            self.geo_res = 0
        self.pv_dc_ac_ratio = [1.1] * 5
        try:
            self.pv_dc_ac_ratio = [float(config.get('PV', 'dc_ac_ratio'))] * 5
        except:
            pass
        try:
            self.pv_dc_ac_ratio[0] = float(config.get('Fixed PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.pv_dc_ac_ratio[1] = float(config.get('Rooftop PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.pv_dc_ac_ratio[2] = float(config.get('Single Axis PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.pv_dc_ac_ratio[3] = float(config.get('Backtrack PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.pv_dc_ac_ratio[4] = float(config.get('Tracking PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.pv_dc_ac_ratio[4] = float(config.get('Dual Axis PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.pv_losses = float(config.get('PV', 'losses'))
        except:
            self.pv_losses = 5
        try:
            self.wave_cutout = float(config.get('Wave', 'cutout'))
        except:
            self.wave_cutout = 0
        try:
            wf = config.get('Wave', 'efficiency')
            if wf[-1] == '%':
                self.wave_efficiency = float(wf[-1]) / 100.
            else:
                self.wave_efficiency = float(wf)
            if self.wave_efficiency > 1:
                self.wave_efficiency = self.wave_efficiency / 100.
        except:
            self.wave_efficiency = 0
        self.wind_turbine_spacing = [8, 8] # onshore and offshore winds
        self.wind_row_spacing = [8, 8]
        self.wind_offset_spacing = [4, 4]
        self.wind_farm_losses_percent = [2, 2]
        self.wind_hub_formula = [None, None]
        self.wind_law = ['l', 'l']
        try:
            self.wind_turbine_spacing[0] = int(config.get('Wind', 'turbine_spacing'))
        except:
            try:
                self.wind_turbine_spacing[0] = int(config.get('Onshore Wind', 'turbine_spacing'))
            except:
                pass
        try:
            self.wind_row_spacing[0] = int(config.get('Wind', 'row_spacing'))
        except:
            try:
                self.wind_row_spacing[0] = int(config.get('Onshore Wind', 'row_spacing'))
            except:
                pass
        try:
            self.wind_offset_spacing[0] = int(config.get('Wind', 'offset_spacing'))
        except:
            try:
                self.wind_offset_spacing[0] = int(config.get('Onshore Wind', 'offset_spacing'))
            except:
                pass
        try:
            self.wind_farm_losses_percent[0] = int(config.get('Wind', 'wind_farm_losses_percent').strip('%'))
        except:
            try:
                self.wind_farm_losses_percent[0] = int(config.get('Onshore Wind', 'wind_farm_losses_percent').strip('%'))
            except:
                pass
        try:
            self.wind_law[0] = config.get('Wind', 'extrapolate')
        except:
            try:
                self.wind_law[0] = config.get('Onshore Wind', 'extrapolate')
            except:
                pass
        try:
            self.wind_hub_formula[0] = config.get('Wind', 'hub_formula')
        except:
            try:
                self.wind_hub_formula[0] = config.get('Onshore Wind', 'hub_formula')
            except:
                pass
        try:
            self.wind_turbine_spacing[1] = int(config.get('Offshore Wind', 'turbine_spacing'))
        except:
            pass
        try:
            self.wind_row_spacing[1] = int(config.get('Offshore Wind', 'row_spacing'))
        except:
            pass
        try:
            self.wind_offset_spacing[1] = int(config.get('Offshore Wind', 'offset_spacing'))
        except:
            pass
        try:
            self.wind_farm_losses_percent[1] = int(config.get('Offshore Wind', 'wind_farm_losses_percent').strip('%'))
        except:
            pass
        try:
            self.wind_law[1] = config.get('Offshore Wind', 'extrapolate')
        except:
            pass
        try:
            self.wind_hub_formula[1] = config.get('Offshore Wind', 'hub_formula')
        except:
            pass
        self.st_gross_net = 0.87
        try:
            self.st_gross_net = float(config.get('Solar Thermal', 'gross_net'))
        except:
            pass
        self.st_tshours = 0
        try:
            self.st_tshours = float(config.get('Solar Thermal', 'tshours'))
        except:
            pass
        self.st_volume = 12.9858
        try:
            self.st_volume = float(config.get('Solar Thermal', 'volume'))
        except:
            pass
        self.cst_gross_net = 0.87
        try:
            self.cst_gross_net = float(config.get('CST', 'gross_net'))
        except:
            pass
        self.cst_tshours = 0
        try:
            self.cst_tshours = float(config.get('CST', 'tshours'))
        except:
            pass
        try:
            self.hydro_cf = float(config.get('Hydro', 'cf'))
        except:
            self.hydro_cf = 0.33
        try:
            self.actual_power = config.get('Files', 'actual_power')
            for key, value in parents:
                self.actual_power = self.actual_power.replace(key, value)
            self.actual_power = self.actual_power.replace('$USER$', getUser())
            self.actual_power = self.actual_power.replace('$YEAR$', self.base_year)
        except:
            self.actual_power = ''
        try:
            self.solar_files = config.get('Files', 'solar_files')
            for key, value in parents:
                self.solar_files = self.solar_files.replace(key, value)
            self.solar_files = self.solar_files.replace('$USER$', getUser())
            self.solar_files = self.solar_files.replace('$YEAR$', self.base_year)
        except:
            self.solar_files = ''
        try:
            self.solar_index = config.get('Files', 'solar_index')
            for key, value in parents:
                self.solar_index = self.solar_index.replace(key, value)
            self.solar_index = self.solar_index.replace('$USER$', getUser())
            self.solar_index = self.solar_index.replace('$YEAR$', self.base_year)
        except:
            self.solar_index = ''
        try:
            self.wind_files = config.get('Files', 'wind_files')
            for key, value in parents:
                self.wind_files = self.wind_files.replace(key, value)
            self.wind_files = self.wind_files.replace('$USER$', getUser())
            self.wind_files = self.wind_files.replace('$YEAR$', self.base_year)
        except:
            self.wind_files = ''
        try:
            self.wind_index = config.get('Files', 'wind_index')
            for key, value in parents:
                self.wind_index = self.wind_index.replace(key, value)
            self.wind_index = self.wind_index.replace('$USER$', getUser())
            self.wind_index = self.wind_index.replace('$YEAR$', self.base_year)
        except:
            self.wind_index = ''
        try:
            self.variable_files = config.get('Files', 'variable_files')
            for key, value in parents:
                self.variable_files = self.variable_files.replace(key, value)
            self.variable_files = self.variable_files.replace('$USER$', getUser())
            self.defaults = {}
            self.default_files = {}
            defaults = config.items('SAM Modules')
            for tech, default in defaults:
                if '_variables' in tech:
                    tec = tech.replace('_variables', '')
                    tec = techClean(tec)
                    self.defaults[tec] = default
                    self.default_files[tec] = None
                elif tech == 'helio_positions':
                    self.default_files['helio_positions'] = default
                    self.defaults['helio_positions'] = None
                elif tech == 'optical_table':
                    self.default_files['optical_table'] = default
                    self.defaults['optical_table'] = None
            self.default_files['solar_index'] = None
            self.default_files['wind_index'] = None
            self.default_files['actual'] = None
        except:
            pass
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
        if not os.path.exists(self.scenarios + self.actual_power):
            self.actual_power = ''
        try:
            self.helpfile = config.get('Files', 'help')
            for key, value in parents:
                self.helpfile = self.helpfile.replace(key, value)
            self.helpfile = self.helpfile.replace('$USER$', getUser())
            self.helpfile = self.helpfile.replace('$YEAR$', self.base_year)
        except:
            self.helpfile = ''
        try:
            subs_loss = config.get('Grid', 'substation_loss')
            if subs_loss[-1] == '%':
                self.grid_subs_loss = float(subs_loss[:-1]) / 100.
            else:
                self.grid_subs_loss = float(subs_loss) / 10.
            line_loss = config.get('Grid', 'line_loss')
            if line_loss[-1] == '%':
                self.grid_line_loss = float(line_loss[:-1]) / 100000.
            else:
                self.grid_line_loss = float(line_loss) / 1000.
        except:
            self.grid_subs_loss = 0.
            self.grid_line_loss = 0.
        if self.progress is None:
            self.show_progress = None
        else:
            progress_bar = True
            try:
                progress_bar = config.get('View', 'progress_bar')
                if progress_bar in ['false', 'no', 'off']:
                    self.show_progress = None
                else:
                    self.show_progress = True
                    try:
                        self.progress_bar = int(progress_bar)
                    except:
                        self.progress_bar = 0
            except:
                self.show_progress = True
                self.progress_bar = 0
        self.debug = False
        try:
            debug = config.get('Power', 'debug_sam')
            if debug.lower() in ['true', 'yes', 'on']:
                self.debug = True
        except:
            pass
        self.gen_pct = None
        ssc_api = ssc.API()
# to supress messages
        if not self.expert:
            ssc_api.set_print(0)

    def getPower(self):
        self.x = []
        self.stored = []
        self.ly = {}
        if self.plots['save_data'] or self.plots['financials'] or self.plots['save_detail']:
            self.stn_outs = []
            self.stn_tech = []
            self.stn_size = []
            self.stn_pows = []
            self.stn_grid = []
            self.stn_path = []
        elif self.plots['save_tech'] or self.plots['save_match']:
            self.stn_outs = []
            self.stn_tech = []
            if self.plots['visualise']:
                self.stn_pows = []
        elif self.plots['visualise']:
            self.stn_outs = []
            self.stn_pows = []
        if self.plots['save_zone']:
            self.stn_zone = []
        len_x = 8760
        for i in range(len_x):
            self.x.append(i)
        if self.plots['grid_losses']:
            self.ly['Generation'] = []
            for i in range(len_x):
                self.ly['Generation'].append(0.)
        self.getPowerLoop()

    def getPowerLoop(self):
        self.all_done = False
        to_do = []
        for st in range(len(self.stations)):
            if self.plots['by_station']:
                if self.stations[st].name not in self.selected:
                    continue
            if self.stations[st].technology == 'Rooftop PV' \
              and self.stations[st].scenario == 'Existing' \
              and not self.plots['gross_load']:
                continue
            if self.stations[st].technology[:6] == 'Fossil' \
              and not self.plots['actual']:
                continue
            to_do.append(st)
        show_progress = False
        if self.show_progress:
            if self.progress_bar == 0:
                show_progress = True
                self.progress.barRange(0, len(to_do))
            elif len(to_do) >= self.progress_bar:
                show_progress = True
                self.progress.barRange(0, len(to_do))
        for st in range(len(to_do)):
  #      for st in range(len(self.stations)):
            stn = self.stations[to_do[st]]
            if show_progress:
                try:
                    self.progress.barProgress(st, 'Processing ' + stn.name + ' (' + stn.technology + ')')
                    QtWidgets.QApplication.processEvents()
                    if not self.progress.be_open:
                        break
                except:
                    break
            if stn.technology[:6] == 'Fossil' and not self.plots['actual']:
                continue
            if self.plots['by_station']:
                if stn.name not in self.selected:
                    continue
            if stn.technology == 'Rooftop PV' and stn.scenario == 'Existing' \
              and not self.plots['gross_load']:
                continue
            if self.plots['by_station']:
                if stn.name not in self.selected:
                    continue
                key = stn.name
            else:
                if stn.technology == 'Rooftop PV' and (stn.scenario.find('Existing') >= 0):
                    key = 'Existing Rooftop PV'
                else:
                    key = stn.technology
            if self.plots['save_zone']:
                key = stn.zone + '.' + key
            if self.plots['save_data'] or self.plots['financials'] or self.plots['save_detail']:
                self.stn_outs.append(stn.name)
                self.stn_tech.append(stn.technology)
                self.stn_size.append(stn.capacity)
                self.stn_pows.append([])
                if stn.grid_len is not None:
                    self.stn_grid.append(stn.grid_len)
                else:
                    self.stn_grid.append(0.)
                if stn.grid_path_len is not None:
                    self.stn_path.append(stn.grid_path_len)
                else:
                    self.stn_path.append(0.)
            elif self.plots['save_tech'] or self.plots['save_match']:
                self.stn_outs.append(stn.name)
                self.stn_tech.append(stn.technology)
                if self.plots['visualise']:
                    self.stn_pows.append([])
            elif self.plots['visualise']:
                self.stn_outs.append(stn.name)
                self.stn_pows.append([])
            if self.plots['save_zone']:
                self.stn_zone.append(stn.zone)
            power = self.getStationPower(stn)
            total_power = 0.
            total_energy = 0.
            if power is None:
                pass
            else:
                if key in self.ly:
                    pass
                else:
                    self.ly[key] = []
                    for i in range(len(self.x)):
                        self.ly[key].append(0.)
                for i in range(len(power)):
                    if self.plots['grid_losses']:
                        if stn.grid_path_len is not None:
                            enrgy = power[i] * (1 - self.grid_line_loss * stn.grid_path_len - self.grid_subs_loss)
                        else:
                            enrgy = power[i] * (1 - self.grid_subs_loss)
                        self.ly[key][i] += enrgy / 1000.
                        total_energy += enrgy / 1000.
                        self.ly['Generation'][i] += power[i] / 1000.
                    else:
                        self.ly[key][i] += power[i] / 1000.
                    total_power += power[i] / 1000.
                    if self.plots['save_data'] or self.plots['financials'] or self.plots['save_detail'] \
                      or self.plots['visualise']:
                        self.stn_pows[-1].append(power[i] / 1000.)
            if total_energy > 0:
                pt = PowerSummary(stn.name, stn.technology, total_power, stn.capacity, total_energy)
            else:
                pt = PowerSummary(stn.name, stn.technology, total_power, stn.capacity)
            if self.plots['save_zone']:
                pt.zone = stn.zone
            self.power_summary.append(pt)
        if show_progress:
            self.progress.barProgress(-1)

    def getStationPower(self, station):
        def do_module(modname, station, field):
            if self.debug and self.status:
                do_time = True
                clock_start = time.time()
            else:
                do_time = False
            module = ssc.Module(modname.encode('utf-8'))
            if do_time:
                time2 = time.time() - clock_start
                self.status.log('Load (%.6f seconds)' % (time2))
            if (module.exec_(self.data)):
                if do_time:
                    time3 = time.time() - clock_start - time2
                    self.status.log('Execute (%.6f seconds)' % (time3))
                if self.debug:
                    self.debug_sam(station.name, station.technology, module, self.data, self.status)
                farmpwr = self.data.get_array(field.encode('utf-8'))
                if do_time:
                    time4 = time.time() - clock_start - time2 - time3
                    self.status.log('Get data (%.6f seconds)' % (time4))
                del module
                return farmpwr
            else:
                if self.status:
                   self.status.log('Errors encountered processing ' + station.name)
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    if self.status:
                       self.status.log(modname + ' error [' + str(idx) + ']: ' + msg.decode())
                    else:
                        print(modname + ' error [', idx, ' ]: ', msg.decode())
                    idx += 1
                    msg = module.log(idx)
                del module
                return None

        if self.plots['actual'] and self.actual_power != '':
            if self.default_files['actual'] is None:
                if os.path.exists(self.scenarios + self.actual_power):
                    self.default_files['actual'] = WorkBook()
                    self.default_files['actual'].open_workbook(self.scenarios + self.actual_power)
            worksheet = self.default_files['actual'].sheet_by_index(0)
            for col in range(worksheet.ncols):
                if worksheet.cell_value(0, col) == 'Hourly energy | (kWh)':
                    break
                if worksheet.cell_value(0, col) == 'Hourly Energy | (kW)':
                    break
                if worksheet.cell_value(0, col) == station.name:
                    break
            else:
                col += 1
            farmpwr = []
            if col < worksheet.ncols:
                for i in range(1, worksheet.nrows):
                    farmpwr.append(worksheet.cell_value(i, col) * 1000.)
            return farmpwr
        elif station.power_file is not None and station.power_file != '':
            if os.path.exists(self.scenarios + station.power_file):
                xl_file = WorkBook()
                xl_file.open_workbook(self.scenarios + station.power_file)
                worksheet = xl_file.sheet_by_index(0)
                for col in range(worksheet.ncols):
                    if worksheet.cell_value(0, col) == 'Hourly energy | (kWh)':
                        break
                    if worksheet.cell_value(0, col) == 'Hourly Energy | (kW)':
                        break
                    if worksheet.cell_value(0, col) == station.name:
                        break
                if col < worksheet.ncols:
                    for i in range(1, worksheet.nrows):
                        farmpwr.append(worksheet.cell_value(i, col) * 1000.)
                return farmpwr
        if station.capacity == 0:
            return None
        if self.status:
            self.status.log('Processing ' + station.name + ' (' + station.technology + ')')
            QtWidgets.QApplication.processEvents()
        self.data = None
        self.data = ssc.Data()
        farmpwr = [] # just in case
        if 'Wind' in station.technology:
            if 'Off' in station.technology: # offshore?
                wtyp = 1
            else:
                wtyp = 0
            closest = self.find_closest(station.lat, station.lon, wind=True)
            turbine = Turbine(station.turbine)
            if not hasattr(turbine, 'capacity'):
                return None
            wind_file = self.wind_files + '/' + closest
            hub_hght = 0
            if self.wind_hub_formula[wtyp] is not None: # if a hub height is specified
                formula = self.wind_hub_formula[wtyp].replace('rotor', str(turbine.rotor))
                try:
                    hub_hght = eval(formula)
                except:
                    pass
            temp_file = None
            if turbine.rotor > 85 and hub_hght > 0: # if a hub height is specified
                try:
                    temp_dir = tempfile.gettempdir()
                    temp_file = 'windfile.srw'
                    wind_data = extrapolateWind(self.wind_files + '/' + closest, hub_hght, law=self.wind_law[wtyp])
                    wf = open(temp_dir + '/' + temp_file, 'w')
                    for line in wind_data:
                        wf.write(line)
                    wf.close()
                    wind_file = temp_dir + '/' + temp_file
                except:
                    pass
            self.data.set_string(b'wind_resource_filename', wind_file.encode('utf-8'))
            no_turbines = int(station.no_turbines)
            if station.scenario == 'Existing' and (no_turbines * turbine.capacity) != (station.capacity * 1000):
                loss = round(1. - (station.capacity * 1000) / (no_turbines * turbine.capacity), 2)
                loss = loss * 100
                if loss < self.wind_farm_losses_percent[wtyp]:
                    loss = self.wind_farm_losses_percent[wtyp]
                self.data.set_number(b'system_capacity', station.capacity * 1000000)
                self.data.set_number(b'wind_farm_losses_percent', loss)
            else:
                self.data.set_number(b'system_capacity', no_turbines * turbine.capacity * 1000)
                self.data.set_number(b'wind_farm_losses_percent', self.wind_farm_losses_percent[wtyp])
            pc_wind = turbine.speeds
            pc_power = turbine.powers
            self.data.set_array(b'wind_turbine_powercurve_windspeeds', pc_wind)
            self.data.set_array(b'wind_turbine_powercurve_powerout', pc_power)
            t_rows = int(ceil(sqrt(no_turbines)))
            ctr = no_turbines
            wt_x = []
            wt_y = []
            for r in range(t_rows):
                for c in range(t_rows):
                    wt_x.append(r * self.wind_row_spacing[wtyp] * turbine.rotor)
                    wt_y.append(c * self.wind_turbine_spacing[wtyp] * turbine.rotor +
                                (r % 2) * self.wind_offset_spacing[wtyp] * turbine.rotor)
                    ctr -= 1
                    if ctr < 1:
                        break
                if ctr < 1:
                    break
            self.data.set_array(b'wind_farm_xCoordinates', wt_x)
            self.data.set_array(b'wind_farm_yCoordinates', wt_y)
            self.data.set_number(b'wind_turbine_rotor_diameter', turbine.rotor)
            self.data.set_number(b'wind_turbine_cutin', turbine.cutin)
            try:
                if station.hub_height > 0:
                    self.data.set_number(b'wind_turbine_hub_ht', station.hub_height)
            except:
                if hub_hght > 0:
                    self.data.set_number(b'wind_turbine_hub_ht', hub_hght)
            self.do_defaults(station)
            farmpwr = do_module('windpower', station, 'gen')
            if temp_file is not None: # if a hub height is specified
                os.remove(temp_dir + '/' + temp_file)
            return farmpwr
        elif station.technology == 'CST':
            closest = self.find_closest(station.lat, station.lon)
            base_capacity = 104.
            self.data.set_string(b'file_name', (self.solar_files + '/' + closest).encode('utf-8'))
            self.data.set_number(b'system_capacity', int(base_capacity * 1000))
            self.data.set_number(b'w_des', base_capacity / self.cst_gross_net)
            self.data.set_number(b'latitude', station.lat)
            self.data.set_number(b'longitude', station.lon)
            self.data.set_number(b'timezone', int(round(station.lon / 15.)))
            if station.storage_hours is None:
                tshours = self.cst_tshours
            else:
                tshours = station.storage_hours
            self.data.set_number(b'hrs_tes', tshours)
            sched = [[1] * 24] * 12
            self.data.set_matrix(b'weekday_schedule', sched[:])
            self.data.set_matrix(b'weekend_schedule', sched[:])
            if self.defaults['optical_table'] is None:
                optic_file = self.variable_files + '/' + self.default_files['optical_table']
                if not os.path.exists(optic_file):
                    if self.status:
                        self.status.log('optical_table file required for ' + station.name)
                    return
                self.defaults['optical_table'] = []
                hf = open(optic_file)
                lines = hf.readlines()
                hf.close()
                for line in lines:
                    row = []
                    bits = line.split(',')
                    for bit in bits:
                        row.append(float(bit))
                    self.defaults['optical_table'].append(row)
                del lines
            self.data.set_matrix(b'OpticalTable', self.defaults['optical_table'][:])
            self.do_defaults(station)
            farmpwr = do_module('tcsgeneric_solar', station, 'gen')
            if farmpwr is not None:
                if station.capacity != base_capacity:
                    for i in range(len(farmpwr)):
                        farmpwr[i] = farmpwr[i] * station.capacity / float(base_capacity)
            return farmpwr
        elif station.technology == 'Solar Thermal':
            closest = self.find_closest(station.lat, station.lon)
            base_capacity = 104
            self.data.set_string(b'solar_resource_file', (self.solar_files + '/' + closest).encode('utf-8'))
            self.data.set_number(b'system_capacity', base_capacity * 1000)
            self.data.set_number(b'P_ref', base_capacity / self.st_gross_net)
            if station.storage_hours is None:
                tshours = self.st_tshours
            else:
                tshours = station.storage_hours
            self.data.set_number(b'tshours', tshours)
            sched = [[1.] * 24] * 12
            self.data.set_matrix(b'weekday_schedule', sched[:])
            self.data.set_matrix(b'weekend_schedule', sched[:])
            if ssc.API().version() >= 159:
                self.data.set_matrix(b'dispatch_sched_weekday', sched[:])
                self.data.set_matrix(b'dispatch_sched_weekend', sched[:])
                if ssc.API().version() >= 206:
                    self.data.set_number(b'gross_net_conversion_factor', self.st_gross_net)
                    if self.defaults['helio_positions'] is None:
                        helio_file = self.variable_files + '/' + self.default_files['helio_positions']
                        if not os.path.exists(helio_file):
                            if self.status:
                               self.status.log('helio_positions file required for ' + station.name)
                            return
                        self.defaults['helio_positions'] = []
                        hf = open(helio_file)
                        lines = hf.readlines()
                        hf.close()
                        for line in lines:
                            row = []
                            bits = line.split(',')
                            for bit in bits:
                                row.append(float(bit))
                            self.defaults['helio_positions'].append(row)
                        del lines
                    self.data.set_matrix(b'helio_positions', self.defaults['helio_positions'][:])
            else:
                self.data.set_number(b'Design_power', base_capacity / self.st_gross_net)
                self.data.set_number(b'W_pb_design', base_capacity / self.st_gross_net)
                vol_tank = base_capacity * tshours * self.st_volume
                self.data.set_number(b'vol_tank', vol_tank)
                f_tc_cold = self.data.get_number(b'f_tc_cold')
                V_tank_hot_ini = vol_tank * (1. - f_tc_cold)
                self.data.set_number(b'V_tank_hot_ini', V_tank_hot_ini)
            self.do_defaults(station)
            farmpwr = do_module('tcsmolten_salt', station, 'gen')
            if farmpwr is not None:
                if station.capacity != base_capacity:
                    for i in range(len(farmpwr)):
                        farmpwr[i] = farmpwr[i] * station.capacity / float(base_capacity)
            return farmpwr
        elif 'PV' in station.technology:
            closest = self.find_closest(station.lat, station.lon)
            self.data.set_string(b'solar_resource_file', (self.solar_files + '/' + closest).encode('utf-8'))
            dc_ac_ratio = self.pv_dc_ac_ratio[0]
            if station.technology[:5] == 'Fixed':
                dc_ac_ratio = self.pv_dc_ac_ratio[0]
                self.data.set_number(b'array_type', 0)
            elif station.technology[:7] == 'Rooftop':
                dc_ac_ratio = self.pv_dc_ac_ratio[1]
                self.data.set_number(b'array_type', 1)
            elif station.technology[:11] == 'Single Axis':
                dc_ac_ratio = self.pv_dc_ac_ratio[2]
                self.data.set_number(b'array_type', 2)
            elif station.technology[:9] == 'Backtrack':
                dc_ac_ratio = self.pv_dc_ac_ratio[3]
                self.data.set_number(b'array_type', 3)
            elif station.technology[:8] == 'Tracking' or station.technology[:9] == 'Dual Axis':
                dc_ac_ratio = self.pv_dc_ac_ratio[4]
                self.data.set_number(b'array_type', 4)
            self.data.set_number(b'system_capacity', station.capacity * 1000 * dc_ac_ratio)
            self.data.set_number(b'dc_ac_ratio', dc_ac_ratio)
            try:
                self.data.set_number(b'tilt', fabs(station.tilt))
            except:
                self.data.set_number(b'tilt', fabs(station.lat))
            if float(station.lat) < 0:
                azi = 0
            else:
                azi = 180
            if station.direction is not None:
                if isinstance(station.direction, int):
                    if station.direction >= 0 and station.direction <= 360:
                        azi = station.direction
                else:
                    dirns = {'N': 0, 'NNE': 23, 'NE': 45, 'ENE': 68, 'E': 90, 'ESE': 113,
                             'SE': 135, 'SSE': 157, 'S': 180, 'SSW': 203, 'SW': 225,
                             'WSW': 248, 'W': 270, 'WNW': 293, 'NW': 315, 'NNW': 338}
                    try:
                        azi = dirns[station.direction]
                    except:
                        pass
            self.data.set_number(b'azimuth', azi)
            self.data.set_number(b'losses', self.pv_losses)
            self.do_defaults(station)
            farmpwr = do_module('pvwattsv5', station, 'gen')
            return farmpwr
        elif station.technology == 'Biomass':
            closest = self.find_closest(station.lat, station.lon)
            self.data.set_string(b'file_name', (self.solar_files + '/' + closest).encode('utf-8'))
            self.data.set_number(b'system_capacity', station.capacity * 1000)
            self.data.set_number(b'biopwr.plant.nameplate', station.capacity * 1000)
            feedstock = station.capacity * 1000 * self.biomass_multiplier
            self.data.set_number(b'biopwr.feedstock.total', feedstock)
            self.data.set_number(b'biopwr.feedstock.total_biomass', feedstock)
            carbon_pct = self.data.get_number(b'biopwr.feedstock.total_biomass_c')
            self.data.set_number(b'biopwr.feedstock.total_c', feedstock * carbon_pct / 100.)
            self.do_defaults(station)
            farmpwr = do_module('biomass', station, 'gen')
            return farmpwr
        elif station.technology == 'Geothermal':
            closest = self.find_closest(station.lat, station.lon)
            self.data.set_string(b'file_name', (self.solar_files + '/' + closest).encode('utf-8'))
            self.data.set_number(b'nameplate', station.capacity * 1000)
            self.data.set_number(b'resource_potential', station.capacity * 10.)
            self.data.set_number(b'resource_type', self.geo_res)
            self.data.set_string(b'hybrid_dispatch_schedule', ('1' * 24 * 12).encode('utf-8'))
            self.do_defaults(station)
            pwr = do_module('geothermal', station, 'monthly_energy')
            if pwr is not None:
                farmpwr = []
                for i in range(12):
                    for j in range(the_days[i] * 24):
                        farmpwr.append(pwr[i] / (the_days[i] * 24))
                return farmpwr
            else:
                return None
        elif station.technology == 'Hydro':   # fudge Hydro purely by Capacity Factor
            pwr = station.capacity * 1000 * self.hydro_cf
            for hr in range(8760):
                farmpwr.append(pwr)
            return farmpwr
        elif station.technology == 'Wave':   # fudge Wave using 10m wind speed
            closest = self.find_closest(station.lat, station.lon)
            tf = open(self.solar_files + '/' + closest, 'r')
            lines = tf.readlines()
            tf.close()
            fst_row = len(lines) - 8760
            wnd_col = 4
            for i in range(fst_row, len(lines)):
                bits = lines[i].split(',')
                if float(bits[wnd_col]) == 0:
                    farmpwr.append(0.)
                else:
                    wave_height = 0.0070104 * pow(float(bits[wnd_col]) * 1.94384, 2)   # 0.023 * 0.3048 = 0.0070104 ft to metres
                    if self.wave_cutout > 0 and wave_height > self.wave_cutout:
                        farmpwr.append(0.)
                    else:
                        wave_period = 0.45 * float(bits[wnd_col]) * 1.94384
                        wave_pwr = pow(wave_height, 2) * wave_period * self.wave_efficiency
                        if wave_pwr > 1.:
                            wave_pwr = 1.
                        pwr = station.capacity * 1000 * wave_pwr
                        farmpwr.append(pwr)
            return farmpwr
        elif station.technology[:5] == 'Other':
            config = configparser.RawConfigParser()
            if len(sys.argv) > 1:
                config_file = sys.argv[1]
            else:
                config_file = getModelFile('SIREN.ini')
            config.read(config_file)
            props = []
            propty = {}
            wnd50 = False
            try:
                props = config.items(station.technology)
                for key, value in props:
                    propty[key] = value
                closest = self.find_closest(station.lat, station.lon)
                tf = open(self.solar_files + '/' + closest, 'r')
                lines = tf.readlines()
                tf.close()
                fst_row = len(lines) - 8760
                if closest[-4:] == '.smw':
                    dhi_col = 9
                    dni_col = 8
                    ghi_col = 7
                    tmp_col = 0
                    wnd_col = 4
                elif closest[-10:] == '(TMY2).csv' or closest[-10:] == '(TMY3).csv' \
                  or closest[-10:] == '(INTL).csv' or closest[-4:] == '.csv':
                    cols = lines[fst_row - 1].strip().split(',')
                    for i in range(len(cols)):
                        if cols[i].lower() in ['df', 'dhi', 'diffuse', 'diffuse horizontal',
                                               'diffuse horizontal irradiance']:
                            dhi_col = i
                        elif cols[i].lower() in ['dn', 'dni', 'beam', 'direct normal',
                                                 'direct normal irradiance']:
                            dni_col = i
                        elif cols[i].lower() in ['gh', 'ghi', 'global', 'global horizontal',
                                                 'global horizontal irradiance']:
                            ghi_col = i
                        elif cols[i].lower() in ['temp', 'tdry']:
                            tmp_col = i
                        elif cols[i].lower() in ['wspd', 'wind speed']:
                            wnd_col = i
                formula = propty['formula']
                operators = '+-*/%'
                for i in range(len(operators)):
                    formula = formula.replace(operators[i], ' ' + operators[i] + ' ')
                formula.replace(' /  / ', ' // ')
                formula.replace(' *  * ', ' ** ')
                propty['formula'] = formula
                if formula.find('wind50') >= 0:
                    closest = self.find_closest(station.lat, station.lon, wind=True)
                    tf = open(self.wind_files + '/' + closest, 'r')
                    wlines = tf.readlines()
                    tf.close()
                    if closest[-4:] == '.srw':
                        units = wlines[3].strip().split(',')
                        heights = wlines[4].strip().split(',')
                        for j in range(len(units)):
                            if units[j] == 'm/s':
                               if heights[j] == '50':
                                   wnd50_col = j
                                   break
                    fst_wrow = len(wlines) - 8760
                    wnd50_row = fst_wrow - fst_row
                    wnd50 = True
                for i in range(fst_row, len(lines)):
                    bits = lines[i].strip().split(',')
                    if wnd50:
                        bits2 = wlines[i + wnd50_row].strip().split(',')
                    formulb = propty['formula'].lower().split()
                    formula = ''
                    for form in formulb:
                        if form == 'dhi':
                            formula += bits[dhi_col]
                        elif form == 'dni':
                            formula += bits[dni_col]
                        elif form == 'ghi':
                            formula += bits[ghi_col]
                        elif form == 'temp':
                            formula += bits[tmp_col]
                        elif form == 'wind':
                            formula += bits[wnd_col]
                        elif form == 'wind50':
                            formula += bits2[wnd50_col]
                        else:
                            for key in list(propty.keys()):
                                if form == key:
                                    formula += propty[key]
                                    break
                            else:
                                formula += form
                        formula += ''
                    try:
                        pwr = eval(formula)
                        if pwr > 1:
                            pwr = 1.
                    except:
                        pwr = 0.
                    pwr = station.capacity * 1000 * pwr
                    farmpwr.append(pwr)
            except:
                pass
            return farmpwr
        else:
            return None

    def getValues(self):
        return self.power_summary

    def getPct(self):
        return self.gen_pct

    def getLy(self):
        try:
            return self.ly, self.x
        except:
            return None, None

    def getStnOuts(self):
        return self.stn_outs, self.stn_tech, self.stn_size, self.stn_pows, self.stn_grid, self.stn_path

    def getStnTech(self):
        return self.stn_outs, self.stn_tech

    def getStnPows(self):
        return self.stn_outs, self.stn_pows

    def getStnZones(self):
        return self.stn_zone

    def getVisual(self):
        return self.model.getVisual()
