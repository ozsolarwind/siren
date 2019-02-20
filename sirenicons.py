#!/usr/bin/python
#
#  Copyright (C) 2016-2019 Sustainable Energy Now Inc., Angus King
#
#  sirenicons.py - This file is part of SIREN.
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

import ConfigParser   # decode .ini file
import sys
from senuser import techClean


class Icons:
    def __init__(self):
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        self.icons = {'Biomass': 'biomass_g.png', 'CST': 'solar_g.png', 'Fossil': 'fossil_g.png',
                      'Geothermal': 'hot_rocks_g.png', 'Hydro': 'hydro_g.png',
                      'Other': 'question.png', 'PV': 'solar_pv_g.png',
                      'Solar Thermal': 'solar_g.png', 'Wave': 'wave_g.png', 'Wind': 'wind_g.png'}
        try:
            technologies = config.get('Power', 'technologies').split(' ')
            for tec in technologies:
                tec = techClean(tec)
                try:
                    self.icons[tec] = config.get(tec, 'icon')
                except:
                    pass
            technologies = config.get('Power', 'fossil_technologies').split(' ')
            for tec in technologies:
                tec = techClean(tec)
                try:
                    self.icons[tec] = config.get(tec, 'icon')
                except:
                    pass
        except:
            pass

    def getIcon(self, technology):
        if technology in self.icons.keys():
            return self.icons[technology]
        tech = technology
        if technology == 'Biogas':
            tech = 'Biomass'
        elif 'PV' in technology:
            tech = 'PV'
        elif technology[:6] == 'Fossil':
            tech = 'Fossil'
        try:
            return self.icons[tech]
        except:
            return 'question.png'
