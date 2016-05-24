#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King     
#
#  turbine.py - This file is part of SIREN.
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

import pylab as plt
import csv
import math
import os
import sys

import ConfigParser # decode .ini file

from senuser import getUser

class Power_Curve:

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
            self.pow_dir = config.get('Files', 'pow_files')
            for key, value in parents:
                self.pow_dir = self.pow_dir.replace(key, value)
            self.pow_dir = self.pow_dir.replace('$USER$', getUser())
            self.pow_dir = self.pow_dir.replace('$YEAR$', self.base_year)
        except:
            self.pow_dir = ''
    """Represents a Power Curve for a Wind Turbine."""

    def __init__(self, name, poly):
        """Initializes the data."""
        self.get_config()
        self.name = name
        pow_file = self.pow_dir + '/' + self.name + '.pow'
        if os.path.exists(pow_file):
            tf = open(pow_file, 'r')
            lines = tf.readlines()
            tf.close()
            self.capacity = 0.
            self.cutin = float(lines[4].strip('" \t\n\r'))
            self.rotor = float(lines[1].strip('" \t\n\r'))
            x = []
            y = []
            last_valu = 0
            for ln in range(5,len(lines)):
                try:
                    valu = float(lines[ln].strip('" \t\n\r'))
                    if valu > 0 and self.cutin == 0:
                        self.cutin = float(ln-4)
                    if valu < last_valu:
                        if valu > 0 or ln > 5:
                            break
                    x.append(float(ln-4))
                    y.append(float(valu))
                    last_valu = valu
                except:
                    break
            self.cutout = ln - 5
            self.power_curve = plt.polyfit(x, y, poly)
            self.capacity = last_valu
        else:
            print 'No', pow_file
     
    def Power(self, wind_speed):
        """return power for wind speed."""
        if wind_speed < self.cutin or wind_speed > self.cutout:
            return 0.
        reslt = plt.polyval(self.power_curve, wind_speed)
        if reslt < 0:
            return 0.
        if reslt > self.capacity:
            return self.capacity
        return reslt


class Turbine:
    """Specifications for a Wind Turbine (Power Curve, ...)."""

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

    def __init__(self, name, poly=13):
        """Initializes the data."""
        self.get_config()
        self.name = name
        self.maxp = 0
        if os.path.exists(self.sam_file):
             turb_fil = open(self.sam_file)
             turbines = csv.DictReader(turb_fil)
             for turbine in turbines:
                 if turbine['Name'] == name:
                     self.capacity = float(turbine['KW Rating'])
                     self.cutin = 0
                     self.powers = turbine['Power Curve Array'].split('|')
                     self.rotor = float(turbine['Rotor Diameter'])
                     self.speeds = turbine['Wind Speed Array'].split('|')
                     for i in range(len(self.powers)):
                         self.speeds[i] = float(self.speeds[i])
                         self.powers[i] = float(self.powers[i])
                         if self.powers[i] > 0 and self.cutin == 0:
                             self.cutin = self.speeds[i]
                     self.wind_class = turbine['IEC Wind Speed Class']
                     break
             else:
                 pow_turbine = Power_Curve(name, poly)
                 self.capacity = pow_turbine.capacity
                 self.cutin = pow_turbine.cutin
                 self.rotor = pow_turbine.rotor
                 self.wind_class = '?'
                 self.powers = []
                 self.speeds = []
                 last_pow = 0
                 for ws in range(161): # 0 to 40 by 0.25
                     self.speeds.append(ws / 4.)
                     powr = round(pow_turbine.Power(ws / 4.), 2)
                     if powr > 0:
                         if powr > last_pow:
                             last_pow = powr
                         self.powers.append(last_pow)
                     else:
                         self.powers.append(powr)
             turb_fil.close()

    def Power(self):
        """return power curve values for wind speed."""
        return self.speeds, self.powers

    def PowerCurve(self):
        """plot power curve."""
        plt.plot(self.speeds, self.powers, linewidth=2.0)
        plt.title('Power Curve for ' + self.name)
        plt.grid(True)
        plt.xlabel('wind speed (m/s)')
        plt.ylabel('generation (kW)')
        plt.show(block=True)
