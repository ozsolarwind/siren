#!/usr/bin/python
#
#  Copyright (C) 2015-2017 Sustainable Energy Now Inc., Angus King
#
#  makeweatherfiles.py - Make weather files for SAM
#
#  makeweatherfiles.py is free software: you can redistribute it
#  and/or modify it under the terms of the GNU Affero General
#  Public License as #  published by the Free Software Foundation,
#  either version 3 of the License, or (at your option) any later
#  version.
#
#  makeweatherfiles.py is distributed in the hope that it will be
#  useful, but WITHOUT ANY WARRANTY; without even the implied
#  warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
#  See the GNU Affero General Public License for more details.
#
#  You should have received a copy of the GNU Affero General
#  Public License along with SIREN.  If not, see
#  <http://www.gnu.org/licenses/>.
#
#  The routines getDHI and getDNI in this program have been derived
#  from models #  developed by The National Renewable Energy
#  Laboratory (NREL) Center for Renewable Energy Resources.
#  Copyright for these routines remain with them.
#
#  getDNI is derived from the DISC DNI Model
#  <http://rredc.nrel.gov/solar/models/DISC/>
#
#  getDHI is derived from the NREL DNI-GHI to DHI Calculator
#  <https://sam.nrel.gov/sites/sam.nrel.gov/files/content/documents/xls/DNI-GHI_to_DHI_Calculator.xlsx>
#
import datetime
import gzip
import math
import os
import sys
if sys.platform == 'win32' or sys.platform == 'cygwin':
    from netCDF4 import Dataset
else:
    from Scientific.IO.NetCDF import *
import time
from PyQt4 import QtCore, QtGui


class AnObject(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, dialog, anobject, title=None, section=None):
        super(AnObject, self).__init__()
        self.anobject = anobject
        self.title = title
        self.section = section
        dialog.setObjectName('Dialog')
        self.initUI()

    def set_stuff(self, grid, widths, heights, i):
        if widths[1] > 0:
            grid.setColumnMinimumWidth(0, widths[0] + 10)
            grid.setColumnMinimumWidth(1, widths[1] + 10)
        i += 1
        if isinstance(self.anobject, str):
            quit = QtGui.QPushButton('Close', self)
            width = quit.fontMetrics().boundingRect('Close').width() + 10
            quit.setMaximumWidth(width)
        else:
            quit = QtGui.QPushButton('Quit', self)
        grid.addWidget(quit, i + 1, 0)
        quit.clicked.connect(self.quitClicked)
        self.setLayout(grid)
        screen = QtGui.QDesktopWidget().availableGeometry()
        h = heights * i
        if h > screen.height():
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                pct = 0.85
            else:
                pct = 0.90
            h = int(screen.height() * pct)
        self.resize(widths[0] + widths[1] + 40, h)
        if self.title is None:
            self.setWindowTitle('?')
        else:
            self.setWindowTitle(self.title)

    def initUI(self):
        fields = []
        label = []
        self.edit = []
        self.field_type = []
        metrics = []
        widths = [0, 0]
        heights = 0
        i = -1
        grid = QtGui.QGridLayout()
        self.web = QtGui.QTextEdit()
        if os.path.exists(self.anobject):
            htf = open(self.anobject, 'r')
            html = htf.read()
            htf.close()
            self.web.setHtml(QtCore.QString(html))
        else:
            html = self.anobject
            if self.anobject[:5] == '<html':
                self.anobject = self.anobject.replace('[VERSION]', credits.fileVersion())
                self.web.setHtml(QtCore.QString(self.anobject))
            else:
                self.web.setPlainText(QtCore.QString(self.anobject))
        metrics.append(self.web.fontMetrics())
        try:
            widths[0] = metrics[0].boundingRect(self.web.text()).width()
            heights = metrics[0].boundingRect(self.web.text()).height()
        except:
            bits = html.split('\n')
            for lin in bits:
                if len(lin) > widths[0]:
                    widths[0] = len(lin)
            heights = len(bits)
            fnt = self.web.fontMetrics()
            widths[0] = (widths[0]) * fnt.maxWidth()
            heights = (heights) * fnt.height()
            screen = QtGui.QDesktopWidget().availableGeometry()
            if widths[0] > screen.width() * .67:
                heights = int(heights / .67)
                widths[0] = int(screen.width() * .67)
        self.web.setReadOnly(True)
        i = 1
        grid.addWidget(self.web, 0, 0)
        self.set_stuff(grid, widths, heights, i)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
    def curveClicked(self):
        Turbine(self.turbine).PowerCurve()
        return

    def quitClicked(self):
        self.close()

    def saveClicked(self):
        self.close()

    def getValues(self):
        return self.anobject


def getDNI(ghi=0, hour=0, lat=0, lon=0, press=1013.25, zone=8):
    hour_of_year = hour
    day_of_year = int((hour_of_year - 1) / 24) + 1
    latitude = lat
    longitude = lon
    day_angle = 6.283185 * (day_of_year - 1) / 365.
    time_zone = zone
    pressure = press
    etr = 1370 * (1.00011 + 0.034221 * math.cos(day_angle) + 0.00128 * \
          math.sin(day_angle) + 0.000719 * math.cos(2 * day_angle) + \
          0.000077 * math.sin(2 * day_angle))
    dec = (0.006918 - 0.399912 * math.cos(day_angle) + 0.070257 * \
          math.sin(day_angle) - 0.006758 * math.cos(2 * day_angle) + \
          0.000907 * math.sin(2 * day_angle) - 0.002697 * math.cos(3 * \
          day_angle) + 0.00148 * math.sin(3 * day_angle)) * (180. / 3.14159)
    eqt = (0.000075 + 0.001868 * math.cos(day_angle) - 0.032077 * \
          math.sin(day_angle) - 0.014615 * math.cos(2 * day_angle) - \
          0.040849 * math.sin(2 * day_angle)) * (229.18)
    hour_angle = 15 * (hour_of_year - 12 - 0.5 + eqt / 60 + ((longitude - \
          time_zone * 15) * 4) / 60)
    zenith_angle = math.acos(math.cos(math.radians(dec)) * \
          math.cos(math.radians(latitude)) * \
          math.cos(math.radians(hour_angle)) + \
          math.sin(math.radians(dec)) * math.sin(math.radians(latitude))) \
          * (180. / 3.14159)
    if zenith_angle < 80:
        am = 1 / (math.cos(math.radians(zenith_angle)) + 0.15 / \
        pow(93.885 - zenith_angle, 1.253)) * (pressure / 1013.25)
    else:
        am = 0.
    if am > 0:
        kt = ghi / (math.cos(math.radians(zenith_angle)) * etr)
    else:
        kt = 0.
    a = 0.
    if kt > 0:
        if kt > 0.6:
            a = -5.743 + 21.77 * kt - 27.49 * pow(kt, 2) + 11.56 * pow(kt, 3)
        elif kt < 0.6:
            a = 0.512 - 1.56 * kt + 2.286 * pow(kt, 2) - 2.222 * pow(kt, 3)
    b = 0.
    if kt > 0:
        if kt > 0.6:
            b = 41.4 - 118.5 * kt + 66.05 * pow(kt, 2) + 31.9 * pow(kt, 3)
        elif kt < 0.6:
            b = 0.37 + 0.962 * kt
    c = 0.
    if kt > 0:
        if kt > 0.6:
            c = -47.01 + 184.2 * kt - 222 * pow(kt, 2) + 73.81 * pow(kt, 3)
        elif kt < 0.6:
            c = -0.28 + 0.932 * kt - 2.048 * pow(kt, 2)
    kn = 0.
    if kt > 0:
        kn = a + b * math.exp(c * am)
    knc = 0.
    if kt > 0:
        knc = 0.886 - 0.122 * am + 0.0121 * pow(am, 2) - 0.000653 * \
              pow(am, 3) + 0.000014 * pow(am, 4)
    if kt > 0:
        if etr * (knc - kn) >= 0:
            return etr * (knc - kn)
    return 0.

def getDHI(ghi=0, dni=0, hour=0, lat=0, azimuth=0., tilt=0., reflectance=0.2):
# Ic = Ib math.cos(i) + Idh cos^2(B/2) + p Ih sin^2(B/2)
# Ic = insolation on a collector;
# Ib = beam insolation; Idh = diffuse insolation on horizontal;
# Ih = total insolation on horizontal
# i = incidence angle; B = surface tilt; p = gnd reflection
# Idh = Ih - Ibh; Ibh = Ib math.sin(a)  where a = altitude angle

    hour_of_year = hour   # B
    day_of_year = int((hour_of_year - 1) / 24) + 1   # H
    declination_angle = math.asin(math.sin(-23.45 * math.pi / 180.) * \
                        math.cos(360. / 365. * (10.5 + day_of_year) * \
                        math.pi / 180.)) * 180. / math.pi
    latitude = lat
# I
    if hour_of_year % 24 != 0:
        hour_angle = (15 * (12 - hour_of_year % 24)) + 7.5
    else:
        hour_angle = -172.5
# L
    sun_rise_hour_angle = math.acos(-1 * math.tan(latitude * math.pi / \
                          180.) * math.tan(declination_angle * math.pi / \
                          180.)) * 180. / math.pi
# J
    if abs(hour_angle) - 7.5 < sun_rise_hour_angle and \
      abs(hour_angle) + 7.5 > sun_rise_hour_angle:
        hour_angle_2 = abs(hour_angle) - 7.5 + (1 - (((abs(hour_angle) + \
                       7.5) - sun_rise_hour_angle) / 15)) * 7.5
    else:
        hour_angle_2 = 15 * (12 - hour_of_year % 24) + 7.5
# K
    if hour_angle < 0:
        sun_rise_set_adjusted_hour_angle = abs(hour_angle_2) * -1
    else:
        sun_rise_set_adjusted_hour_angle = hour_angle_2
# M
    sun_rise_hr_am = 12 - (math.acos(-1 * math.tan(latitude * math.pi / \
                     180.) * math.tan(declination_angle * math.pi / \
                     180.)) * 180. / math.pi) / 15.
# N
    sun_set_hr_pm = math.acos(-1 * math.tan(latitude * math.pi / 180.) * \
                    math.tan(declination_angle * math.pi / 180.)) * 180. / \
                    math.pi / 15. + 12
# O
    if (math.acos(-math.tan(latitude * math.pi / 180.) * \
          math.tan(declination_angle * math.pi / 180.)) * 180. / math.pi) > \
          abs(sun_rise_set_adjusted_hour_angle):
        sun_rise_set_hsr = 'UP'
    else:
        sun_rise_set_hsr = 'DOWN'
# Q
    altitude_angle = math.asin((math.sin(latitude * math.pi / 180.) * \
                     math.sin(declination_angle * math.pi / 180.)) + \
                     (math.cos(latitude * math.pi / 180.) * \
                     math.cos(declination_angle * math.pi / 180.) * \
                     math.cos(sun_rise_set_adjusted_hour_angle * math.pi / \
                     180.))) * 180. / math.pi
# R
    if sun_rise_set_adjusted_hour_angle > 0:
        solar_azimuth = abs(math.acos(((math.cos(declination_angle * \
                        math.pi / 180.) * math.sin(latitude * math.pi / \
                        180.) * math.cos(sun_rise_set_adjusted_hour_angle * \
                        math.pi / 180.)) - (math.sin(declination_angle * \
                        math.pi / 180.) * math.cos(latitude * math.pi / \
                        180.))) / math.cos(altitude_angle * math.pi / 180.)) \
                        * 180. / math.pi)
    else:
        solar_azimuth = -1 * abs(math.acos(((math.cos(declination_angle * \
                        math.pi / 180.) * math.sin(latitude * math.pi / 180.) \
                        * math.cos(sun_rise_set_adjusted_hour_angle * \
                        math.pi / 180.)) - (math.sin(declination_angle * \
                        math.pi / 180.) * math.cos(latitude * math.pi / \
                        180.))) / math.cos(altitude_angle * math.pi / \
                        180.)) * 180. / math.pi)
# S
    incidence_angle = math.acos((math.cos(altitude_angle * math.pi / 180.) * \
                      math.cos((solar_azimuth - azimuth) * math.pi / 180.) * \
                      math.sin(tilt * math.pi / 180.)) + \
                      ((math.sin(altitude_angle * math.pi / 180.) * \
                      math.cos(tilt * math.pi / 180.)))) * 180. / math.pi
# T
    if incidence_angle < 90:
        beam_component = dni * math.cos(incidence_angle * math.pi / 180.)
    else:
        beam_component = 0.
# U, V
    if altitude_angle > 0:
        diffuse_component = (ghi - (dni * math.sin(altitude_angle * math.pi / \
                            180.))) * pow(math.cos((tilt / 2) * math.pi / \
                            180.), 2)
        reflected_component = reflectance * ghi * pow(math.sin((tilt / 2) * \
                              math.pi / 180.), 2)
    else:
        diffuse_component = 0.
        reflected_component = 0.
# W
    total_wh = beam_component + diffuse_component + reflected_component
# X
    check_total = total_wh - ghi
# Y
    if diffuse_component < 0:
        dhi_negative = 0.
    else:
        dhi_negative = diffuse_component
# Z
    dhi_rounded = round(dhi_negative, 1)
    return dhi_rounded


class makeWeather():

    def unZip(self, inp_file):
        if inp_file is None:
            self.log += 'Terminating as file not found - %s\n' % inp_file
            self.return_code = 12
            return None
        if inp_file[-3:] == '.gz':
            if not os.path.exists(inp_file):
                self.log += 'Terminating as file not found - %s\n' % inp_file
                self.return_code = 4
                return None
            out_file = inp_file + '.nc'
            if os.path.exists(out_file):
                pass
            else:
                fin = gzip.open(inp_file, 'rb')
                file_content = fin.read()
                fin.close()
                fou = open(out_file, 'wb')
                fou.write(file_content)
                fou.close()
            return out_file
        else:
            if not os.path.exists(inp_file):
                self.log += 'Terminating as file not found - %s\n' % inp_file
                self.return_code = 4
                return None
            return inp_file

    def getSpeed(self, vmi, umi):
        try:
            um = umi.data
            vm = vmi.data
        except:
            um = umi[:]
            vm = vmi[:]
        sm = []
        for hr in range(len(um)):   # 24 hours
            hm = []
            for lat in range(len(um[hr])):   # n latitudes
                lm = []
                for lon in range(len(um[hr][lat])):   # m longitudes
                    lm.append(round(math.sqrt(um[hr][lat][lon] * um[hr][lat][lon] +
                                    vm[hr][lat][lon] * vm[hr][lat][lon]), 4))
                hm.append(lm)
            sm.append(hm)
        return sm

    def getDirn(self, vmi, umi):
        try:
            um = umi.data
            vm = vmi.data
        except:
            um = umi[:]
            vm = vmi[:]
        dm = []
        for hr in range(len(um)):   # 24 hours
            hm = []
            for lat in range(len(um[hr])):   # n latitudes
                lm = []
                for lon in range(len(um[hr][lat])):   # m longitudes
     # Calculate the wind direction
                    if abs(vm[hr][lat][lon]) < 0.000001:   # No v-component of velocity
                        if vm[hr][lat][lon] >= 0:
                            lm.append(270)
                        else:
                            lm.append(90)
                    else:   # Calculate angle and convert to degrees
                        theta = math.atan(um[hr][lat][lon] / vm[hr][lat][lon])
                        theta = math.degrees(theta)
                        if vm[hr][lat][lon] > 0:
                            lm.append(int(theta + 180.0))
                        else:   # Make sure angle is positive
                            theta = theta + 360.0
                            lm.append(int(theta % 360.0))
                hm.append(lm)
            dm.append(hm)
        return dm

    def getTemp(self, tmi):
        try:
            tm = tmi.data
        except:
            tm = tmi[:]
        tmp = []
        for hr in range(len(tm)):   # 24 hours
            ht = []
            for lat in range(len(tm[hr])):   # n latitudes
                lt = []
                for lon in range(len(tm[hr][lat])):   # m longitudes
                    lt.append(round(tm[hr][lat][lon] - 273.15, 1))   # K to C
                ht.append(lt)
            tmp.append(ht)
        return tmp

    def getPress(self, pmi):
        if self.fmat == 'srw':
            div = 101325   # Pa to atm
            rnd = 6
        else:
            div = 100.   # Pa to mbar (hPa)
            rnd = 0
        try:
            pm = pmi.data
        except:
            pm = pmi[:]
        ps = []
        for hr in range(len(pm)):   # 24 hours
            hp = []
            for lat in range(len(pm[hr])):   # n latitudes
                lp = []
                for lon in range(len(pm[hr][lat])):   # m longitudes
                    lp.append(round(pm[hr][lat][lon] / div, rnd))   # Pa to mbar
                hp.append(lp)
            ps.append(hp)
        return ps

    def getGHI(self, pmi):
        try:
            pm = pmi.data
        except:
            pm = pmi[:]
        ps = []
        for hr in range(len(pm)):   # 24 hours
            hp = []
            for lat in range(len(pm[hr])):   # n latitudes
                lp = []
                for lon in range(len(pm[hr][lat])):   # m longitudes
                    lp.append(pm[hr][lat][lon])
                hp.append(lp)
            ps.append(hp)
        return ps

    def decodeError(self, inp_file):
        self.log += 'Terminating as error with - %s\n' % inp_file
        try:
            tf = open(inp_file, 'r')
            lines = tf.read()
            if lines[:22] == 'Content-type:text/html':
                lines = lines[22:]
                lines = lines.replace('\r', '')
                lines = lines.replace('\n', '')
                i = lines.find('</')
                while i >= 0:
                    j = lines.find('>', i)
                    tag = lines[i: j + 1]
                    lines = lines.replace(tag, '')
                    tag = tag.replace('/', '')
                    lines = lines.replace(tag, '')
                    i = lines.find('</')
                self.log += lines
        except:
            pass
        self.return_code = 5
        return

    def valu(self, data, lat1, lon1, lat_rat, lon_rat, rnd=4):
        if rnd > 0:
            return round(lat_rat * lon_rat * data[lat1][lon1] +
                (1.0 - lat_rat) * lon_rat * data[lat1 + 1][lon1] +
                lat_rat * (1.0 - lon_rat) * data[lat1][lon1 + 1] +
                (1.0 - lat_rat) * (1.0 - lon_rat) * data[lat1 + 1][lon1 + 1], rnd)
        else:
            return int(lat_rat * lon_rat * data[lat1][lon1] +
                (1.0 - lat_rat) * lon_rat * data[lat1 + 1][lon1] +
                lat_rat * (1.0 - lon_rat) * data[lat1][lon1 + 1] +
                (1.0 - lat_rat) * (1.0 - lon_rat) * data[lat1 + 1][lon1 + 1])

    def get_data(self, inp_file):
        unzip_file = self.unZip(inp_file)
        if self.return_code != 0:
            return
        try:
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                cdf_file = Dataset(unzip_file, 'r')
            else:
                cdf_file = NetCDFFile(unzip_file, 'r')
        except:
            self.decodeError(inp_file)
            return

     #   Variable Description                                          Units
     #   -------- ---------------------------------------------------- --------
     #   ps       Time averaged surface pressure                       Pa
     #   u10m     Eastward wind at 10 m above displacement height      m/s
     #   u2m      Eastward wind at 2 m above displacement height       m/s
     #   u50m     Eastward wind at 50 m above surface                  m/s
     #   v10m     Northward wind at 10 m above the displacement height m/s
     #   v2m      Northward wind at 2 m above the displacement height  m/s
     #   v50m     Northward wind at 50 m above surface                 m/s
     #   t10m     Temperature at 10 m above the displacement height    K
     #   t2m      Temperature at 2 m above the displacement height     K
        self.lati = cdf_file.variables[self.vars['latitude']][:]
        self.longi = cdf_file.variables[self.vars['longitude']][:]
        self.tims = cdf_file.variables['time'][:]
        if self.vars['u10m'] in cdf_file.variables:
            self.s10m += self.getSpeed(cdf_file.variables[self.vars['v10m']], cdf_file.variables[self.vars['u10m']])
            self.d10m += self.getDirn(cdf_file.variables[self.vars['v10m']], cdf_file.variables[self.vars['u10m']])
            self.t_10m += self.getTemp(cdf_file.variables[self.vars['t10m']])
        else:
            self.s10m += self.getSpeed(cdf_file.variables[self.vars['v2m']], cdf_file.variables[self.vars['u2m']])
            self.d10m += self.getDirn(cdf_file.variables[self.vars['v2m']], cdf_file.variables[self.vars['u2m']])
            self.t_10m += self.getTemp(cdf_file.variables[self.vars['t2m']])
        self.p_s += self.getPress(cdf_file.variables[self.vars['ps']])
        if self.make_wind:
            self.s2m += self.getSpeed(cdf_file.variables[self.vars['v2m']], cdf_file.variables[self.vars['u2m']])
            self.d2m += self.getDirn(cdf_file.variables[self.vars['v2m']], cdf_file.variables[self.vars['u2m']])
            self.t_2m += self.getTemp(cdf_file.variables[self.vars['t2m']])
            self.s50m += self.getSpeed(cdf_file.variables[self.vars['v50m']], cdf_file.variables[self.vars['u50m']])
            self.d50m += self.getDirn(cdf_file.variables[self.vars['v50m']], cdf_file.variables[self.vars['u50m']])
        cdf_file.close()

    def get_rad_data(self, inp_file):
        unzip_file = self.unZip(inp_file)
        if self.return_code != 0:
            return
        try:
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                cdf_file = Dataset(unzip_file, 'r')
            else:
                cdf_file = NetCDFFile(unzip_file, 'r')
        except:
            self.decodeError(inp_file)
            return
     #   Variable Description                         Units
     #   -------- ----------------------------------- --------
     #   swgnt    Surface net downward shortwave flux W m-2
        self.lati = cdf_file.variables[self.vars['latitude']][:]
        self.longi = cdf_file.variables[self.vars['longitude']][:]
        self.tims = cdf_file.variables['time'][:]
        self.swgnt += self.getGHI(cdf_file.variables[self.vars['swgnt']])
        cdf_file.close()

    def checkZone(self):
        self.return_code = 0
        if self.longrange[0] is not None:
            if int(round(self.longrange[0] / 15)) != self.src_zone:
                self.log += 'MERRA west longitude (%s) in different time zone: %s\n' % ('{:0.4f}'.format(self.longrange[0]),
                            int(round(self.longrange[0] / 15)))
                self.return_code = 1
        if self.longrange[1] is not None:
            if int(round(self.longrange[1] / 15)) != self.src_zone:
                self.log += 'MERRA east longitude (%s) in different time zone: %s\n' % ('{:0.4f}'.format(self.longrange[1]),
                            int(round(self.longrange[1] / 15)))
                self.return_code = 1
        return

    def close(self):
        return

    def getLog(self):
        return self.log

    def returnCode(self):
        return str(self.return_code)

    def findFile(self, inp_strt, wind=True):
        if wind:
            for p in range(len(self.src_w_pfx)):
                inp_file = self.src_dir_w + self.src_w_pfx[p] + inp_strt + self.src_w_sfx[p]
                if os.path.exists(inp_file):
                    break
                else:
                    if inp_file.find('MERRA300') >= 0:
                        inp_file = inp_file.replace('MERRA300', 'MERRA301')
                        if os.path.exists(inp_file):
                            break
            else:
                self.log += 'No Wind file found for ' + inp_strt + '\n'
                return None
        else:
            for p in range(len(self.src_s_pfx)):
                inp_file = self.src_dir_s + self.src_s_pfx[p] + inp_strt + self.src_s_sfx[p]
                if os.path.exists(inp_file):
                    break
                else:
                    if inp_file.find('MERRA300') >= 0:
                        inp_file = inp_file.replace('MERRA300', 'MERRA301')
                        if os.path.exists(inp_file):
                            break
            else:
                self.log += 'No Solar file found for ' + inp_strt + '\n'
                return None
        return inp_file

    def getInfo(self, inp_file):
        if not os.path.exists(inp_file):
            if inp_file.find('MERRA300') >= 0:
                inp_file = inp_file.replace('MERRA300', 'MERRA301')
            else:
                return
        if not os.path.exists(inp_file):
            return
        unzip_file = self.unZip(inp_file)
        if self.return_code != 0:
            return
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            cdf_file = Dataset(unzip_file, 'r')
        else:
            cdf_file = NetCDFFile(unzip_file, 'r')
        i = unzip_file.rfind('/')
        self.log += '\nFile:\n    '
        self.log += unzip_file[i + 1:] + '\n'
        self.log += ' Format:\n    '
        self.log += str(cdf_file.Format) + '\n'
        self.log += ' Dimensions:\n    '
        vals = ''
        keys = cdf_file.dimensions.keys()
        values = cdf_file.dimensions.values()
        if type(cdf_file.dimensions) is dict:
            for i in range(len(keys)):
                vals += keys[i] + ': ' + str(values[i]) + ', '
        else:
            for i in range(len(keys)):
                bits = str(values[i]).strip().split(' ')
                vals += keys[i] + ': ' + bits[-1] + ', '
        self.log += vals[:-2] + '\n'
        self.log += ' Variables:\n    '
        vals = ''
        for key in iter(sorted(cdf_file.variables.iterkeys())):
            vals += key + ', '
        self.log += vals[:-2] + '\n'
        latitude = cdf_file.variables[self.vars['latitude']][:]
        self.log += ' Latitudes:\n    '
        vals = ''
        lati = latitude[:]
        for val in lati:
            vals += '{:0.4f}'.format(val) + ', '
        self.log += vals[:-2] + '\n'
        longitude = cdf_file.variables[self.vars['longitude']][:]
        self.log += ' Longitudes:\n    '
        vals = ''
        longi = longitude[:]
        for val in longi:
            vals += '{:0.4f}'.format(val) + ', '
        self.log += vals[:-2] + '\n'
        cdf_file.close()
        return

    def __init__(self, src_year, src_zone, src_dir_s, src_dir_w, tgt_dir, fmat, wrap=None, src_lat_lon=None, info=False):
      #  self.last_time = datetime.datetime.now()
        self.log = ''
        self.return_code = 0
        self.src_year = int(src_year)
        self.src_zone = src_zone
        self.src_dir_s = src_dir_s
        self.src_dir_w = src_dir_w
        self.tgt_dir = tgt_dir
        self.fmat = fmat
        self.wrap = False
        if wrap is None or wrap == '':
            pass
        elif wrap[0].lower() == 'y' or wrap[0].lower() == 't' or wrap[:2].lower() == 'on':
            self.wrap = True
        self.src_lat_lon = src_lat_lon
        self.src_s_pfx = []
        self.src_s_sfx = []
        merra300 = False
        self.vars = {'latitude': 'lat', 'longitude': 'lon', 'ps': 'PS', 'swgnt': 'SWGNT',
                     'time': 'time', 't2m': 'T2M', 't10m': 'T10M', 't50m': 'T50M', 'u2m': 'U2M',
                     'u10m': 'U10M', 'u50m': 'U50M', 'v2m': 'V2M', 'v10m': 'V10M', 'v50m': 'V50M'}
        if self.src_dir_s != '':
            self.src_dir_s += '/'
            fils = os.listdir(self.src_dir_s)
            for fil in fils:
                if fil.find('MERRA') >= 0:
                    j = fil.find('.tavg1_2d_rad_Nx.')
                    if j > 0:
                        if fil[:j + 17] not in self.src_s_pfx:
                            self.src_s_pfx.append(fil[:j + 17])
                            if self.src_s_pfx[-1].find('MERRA3') > 0:
                                merra300 = True
                                self.src_s_pfx[-1] = self.src_s_pfx[-1].replace('MERRA301', 'MERRA300')
                            self.src_s_sfx.append(fil[j + 17 + 8:])
                     #   break
            del fils
        if merra300:
            for key in self.vars.keys():
                self.vars[key] = self.vars[key].lower()
            self.vars['latitude'] = 'latitude'
            self.vars['longitude'] = 'longitude'
        self.src_w_pfx = []
        self.src_w_sfx = []
        merra300 = False
        if self.src_dir_w != '':
            self.src_dir_w += '/'
            fils = os.listdir(self.src_dir_w)
            for fil in fils:
                if fil.find('MERRA') >= 0:
                    j = fil.find('.tavg1_2d_slv_Nx.')
                    if j > 0:
                        if fil[:j + 17] not in self.src_w_pfx:
                            self.src_w_pfx.append(fil[:j + 17])
                            if self.src_w_pfx[-1].find('MERRA3') > 0:
                                merra300 = True
                                self.src_w_pfx[-1] = self.src_w_pfx[-1].replace('MERRA301', 'MERRA300')
                            self.src_w_sfx.append(fil[j + 17 + 8:])
                     #   break
            del fils
        if merra300:
            for key in self.vars.keys():
                self.vars[key] = self.vars[key].lower()
            self.vars['latitude'] = 'latitude'
            self.vars['longitude'] = 'longitude'
        if self.tgt_dir != '':
            self.tgt_dir += '/'
        if info:
            inp_strt = '{0:04d}'.format(self.src_year) + '0101'
            self.log += '\nWind file for: ' + inp_strt + '\n'
             # get variables from "wind" file
            self.getInfo(self.findFile(inp_strt, True))
            self.log += '\nSolar file for: ' + inp_strt + '\n'
             # get variables from "solar" file
            self.getInfo(self.findFile(inp_strt, False))
            return
        if self.src_zone.lower() == 'auto':
            self.auto_zone = True
            inp_strt = '{0:04d}'.format(self.src_year) + '0101'
             # get longitude from "wind" file
            unzip_file = self.unZip(self.findFile(inp_strt, True))
            if self.return_code != 0:
                return
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                cdf_file = Dataset(unzip_file, 'r')
            else:
                cdf_file = NetCDFFile(unzip_file, 'r')
            longitude = cdf_file.variables[self.vars['longitude']][:]
            self.src_zone = int(round(longitude[0] / 15))
            self.log += 'Time zone: %s based on MERRA (west) longitude (%s to %s)\n' % (str(self.src_zone),
                        '{:0.4f}'.format(longitude[0]), '{:0.4f}'.format(longitude[-1]))
            cdf_file.close()
        else:
            self.src_zone = int(self.src_zone)
            self.auto_zone = False
        if self.src_lat_lon != '':
            self.src_lat = []
            self.src_lon = []
            latlon = self.src_lat_lon.replace(' ','').split(',')
            try:
                for j in range(0, len(latlon), 2):
                    self.src_lat.append(float(latlon[j]))
                    self.src_lon.append(float(latlon[j + 1]))
            except:
                self.log += 'Error with Coordinates field'
                self.return_code = 2
                return
            self.log += 'Processing %s, %s\n' % (self.src_lat, self.src_lon)
        else:
            self.src_lat = None
            self.src_lon = None
        if self.fmat not in ['csv', 'smw', 'srw', 'wind']:
            self.log += 'Invalid output file format specified - %s\n' % self.fmat
            self.return_code = 3
            return
        if self.fmat == 'wind':
            self.fmat = 'srw'
        if self.fmat == 'srw':
            self.make_wind = True
        else:
            self.make_wind = False
            if self.src_dir_s[0] == self.src_dir_s[-1] and (self.src_dir_s[0] == '"' or
               self.src_dir_s[0] == "'"):
                self.src_dir_s = self.src_dir_s[1:-1]
        if self.src_dir_w[0] == self.src_dir_w[-1] and (self.src_dir_w[0] == '"' or
           self.src_dir_w[0] == "'"):
            self.src_dir_w = self.src_dir_w[1:-1]
        if self.tgt_dir[0] == self.tgt_dir[-1] and (self.tgt_dir[0] == '"' or
           self.tgt_dir[0] == "'"):
            self.tgt_dir = self.tgt_dir[1:-1]
        self.lati = []
        self.longi = []
        self.longrange = [None, None]
        self.tims = []
        self.s10m = []
        self.d10m = []
        self.t_10m = []
        self.p_s = []
        self.swgnt = []
        self.s2m = []
        self.s50m = []
        self.d2m = []
        self.d50m = []
        self.t_2m = []
        dys = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if self.src_zone > 0:
            inp_strt = '{0:04d}'.format(self.src_year - 1) + '1231'
        elif self.src_zone <= 0:
            inp_strt = '{0:04d}'.format(self.src_year) + '0101'
        # get variables from "wind" files
        self.get_data(self.findFile(inp_strt, True))   # get wind data
        if self.return_code != 0:
            return
        if self.src_zone != 0:
         # last self.src_zone hours of (self.src_year -1) GMT = first self.src_zone of self.src_year
         # first self.src_zone hours of self.src_year GMT = last self.src_zone of (self.src_year - 1)
            self.p_s = self.p_s[-self.src_zone:]
            if len(self.s10m) > 0:
                self.s10m = self.s10m[-self.src_zone:]
                self.d10m = self.d10m[-self.src_zone:]
                self.t_10m = self.t_10m[-self.src_zone:]
            if self.make_wind:
                self.s2m = self.s2m[-self.src_zone:]
                self.s50m = self.s50m[-self.src_zone:]
                self.d2m = self.d2m[-self.src_zone:]
                self.d50m = self.d50m[-self.src_zone:]
                self.t_2m = self.t_2m[-self.src_zone:]
        elif self.src_zone < 0:
         # first self.src_zone hours of self.src_year GMT = last self.src_zone of (self.src_year - 1)
            self.p_s = self.p_s[:self.src_zone]
            if len(self.s10m) > 0:
                self.s10m = self.s10m[:self.src_zone]
                self.d10m = self.d10m[:self.src_zone]
                self.t_10m = self.t_10m[:self.src_zone]
            if self.make_wind:
                self.s2m = self.s2m[:self.src_zone]
                self.s50m = self.s50m[:self.src_zone]
                self.d2m = self.d2m[:self.src_zone]
                self.d50m = self.d50m[:self.src_zone]
                self.t_2m = self.t_2m[:self.src_zone]
        self.the_year = self.src_year  # start with their year
        if self.wrap:
            yrs = 2
        else:
            yrs = 1
        for mt in range(len(dys)):
            self.log += 'Processing month %s wind\n' % str(mt + 1)
            for dy in range(1, dys[mt] + 1):
                if self.src_zone <= 0:
                    if mt == 0 and dy == 1:
                        continue
                for yr in range(yrs):
                    inp_strt = '{0:04d}'.format(self.the_year) + '{0:02d}'.format(mt + 1) + \
                               '{0:02d}'.format(dy)
                    inp_file = self.findFile(inp_strt, True)
                    if inp_file is None:
                        if self.wrap and self.the_year == self.src_year: # check if need to go back a year
                            self.the_year = self.src_year - 1
                            self.log += 'Wrapping to prior year - %.4d-%.2d-%.2d\n' % (self.the_year, (mt + 1), dy)
                            yrs = 1
                self.get_data(inp_file)   # get wind data
                if self.return_code != 0:
                    return
        if self.src_zone > 0:
            for i in range(self.src_zone):   # delete last n hours
                del self.p_s[-1]
            if len(self.s10m) > 0:
                for i in range(self.src_zone):
                    del self.s10m[-1]
                    del self.d10m[-1]
                    del self.t_10m[-1]
            if self.make_wind:
                for i in range(self.src_zone):
                    del self.s2m[-1]
                    del self.s50m[-1]
                    del self.d2m[-1]
                    del self.d50m[-1]
                    del self.t_2m[-1]
        elif self.src_zone < 0:
            inp_strt = '{0:04d}'.format(self.the_year + 1) + '0101'
            self.get_data(self.findFile(inp_strt, True))
            if self.return_code != 0:
                return
            for i in range(24 + self.src_zone):   # delete last n hours
                del self.p_s[-1]
            if len(self.s10m) > 0:
                for i in range(24 + self.src_zone):
                    del self.s10m[-1]
                    del self.d10m[-1]
                    del self.t_10m[-1]
            if self.make_wind:
                for i in range(24 + self.src_zone):   # delete last n hours
                    del self.s2m[-1]
                    del self.s50m[-1]
                    del self.d2m[-1]
                    del self.d50m[-1]
                    del self.t_2m[-1]
        if self.make_wind:
            target_dir = self.tgt_dir
            self.log += 'Target directory is %s\n' % target_dir
            if not os.path.exists(target_dir):
                self.log += 'mkdir %s\n' % target_dir
                os.makedirs(target_dir)
            if self.src_lat is not None:   # specific location(s)
                for i in range(len(self.src_lat)):
                    for lat2 in range(len(self.lati)):
                        if self.src_lat[i] <= self.lati[lat2]:
                            break
                    for lon2 in range(len(self.longi)):
                        if self.src_lon[i] <= self.longi[lon2]:
                            break
                    if self.longrange[0] is None:
                        self.longrange[0] = self.longi[lon2]
                    else:
                        if self.longi[lon2] < self.longrange[0]:
                            self.longrange[0] = self.longi[lon2]
                    if self.longrange[1] is None:
                        self.longrange[1] = self.longi[lon2]
                    else:
                        if self.longi[lon2] > self.longrange[1]:
                            self.longrange[1] = self.longi[lon2]
                    lat1 = lat2 - 1
                    lat_rat = (self.lati[lat2] - self.src_lat[i]) / (self.lati[lat2] - self.lati[lat1])
                    lon1 = lon2 - 1
                    lon_rat = (self.longi[lon2] - self.src_lon[i]) / (self.longi[lon2] - self.longi[lon1])
                    out_file = self.tgt_dir + 'wind_weather_' + str(self.src_lat[i]) + '_' + \
                               str(self.src_lon[i]) + '_' + str(self.src_year) + '.' + self.fmat
                    tf = open(out_file, 'w')
                    hdr = 'id,<city>,<state>,<country>,%s,%s,%s,0,1,8760\n' % (str(self.src_year),
                          round(self.src_lat[i], 4), round(self.src_lon[i], 4))
                    tf.write(hdr)
                    tf.write('Wind data derived from MERRA tavg1_2d_slv_Nx' + '\n')
                    if len(self.s10m) > 0:
                        tf.write('Temperature,Pressure,Direction,Speed,Temperature,Direction,Speed,' +
                                 'Direction,Speed' + '\n')
                        tf.write('C,atm,degrees,m/s,C,degrees,m/s,degrees,m/s' + '\n')
                        tf.write('2,0,2,2,10,10,10,50,50' + '\n')
                        for hr in range(len(self.s50m)):
                            tf.write(str(self.valu(self.t_2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                                str(self.valu(self.p_s[hr], lat1, lon1, lat_rat, lon_rat, rnd=6)) + ',' +
                                str(self.valu(self.d2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                str(self.valu(self.s2m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                                str(self.valu(self.t_10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                                str(self.valu(self.d10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                str(self.valu(self.s10m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                                str(self.valu(self.d50m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                str(self.valu(self.s50m[hr], lat1, lon1, lat_rat, lon_rat)) + '\n')
                    else:
                        tf.write('Temperature,Pressure,Direction,Speed,Direction,Speed' + '\n')
                        tf.write('C,atm,degrees,m/s,degrees,m/s' + '\n')
                        tf.write('2,0,2,2,50,50' + '\n')
                        for hr in range(len(self.s50m)):
                            tf.write(str(self.valu(self.t_2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                                str(self.valu(self.p_s[hr], lat1, lon1, lat_rat, lon_rat, rnd=6)) + ',' +
                                str(self.valu(self.d2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                str(self.valu(self.s2m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                                str(self.valu(self.d50m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                str(self.valu(self.s50m[hr], lat1, lon1, lat_rat, lon_rat)) + '\n')
                    tf.close()
                    self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
            else:
                for lat in range(len(self.s50m[0])):
                    for lon in range(len(self.s50m[0][0])):
                        if self.longrange[0] is None:
                            self.longrange[0] = self.longi[lon]
                        else:
                            if self.longi[lon] < self.longrange[0]:
                                self.longrange[0] = self.longi[lon]
                        if self.longrange[1] is None:
                            self.longrange[1] = self.longi[lon]
                        else:
                            if self.longi[lon] > self.longrange[1]:
                                self.longrange[1] = self.longi[lon]
                        out_file = self.tgt_dir + 'wind_weather_' + '{:0.4f}'.format(self.lati[lat]) + \
                                   '_' + '{:0.4f}'.format(self.longi[lon]) + '_' + str(self.src_year) + '.srw'
                        tf = open(out_file, 'w')
                        hdr = 'id,<city>,<state>,<country>,%s,%s,%s,0,1,8760\n' % (str(self.src_year),
                              round(self.lati[lat], 4), round(self.longi[lon], 4))
                        tf.write(hdr)
                        tf.write('Wind data derived from MERRA tavg1_2d_slv_Nx' + '\n')
                        if len(self.s10m) > 0:
                            tf.write('Temperature,Pressure,Direction,Speed,Temperature,Direction,' +
                                     'Speed,Direction,Speed' + '\n')
                            tf.write('C,atm,degrees,m/s,C,degrees,m/s,degrees,m/s' + '\n')
                            tf.write('2,0,2,2,10,10,10,50,50' + '\n')
                            for hr in range(len(self.s50m)):
                                tf.write(str(self.t_2m[hr][lat][lon]) + ',' + str(self.p_s[hr][lat][lon]) + ',' +
                                        str(self.d2m[hr][lat][lon]) + ',' + str(self.s2m[hr][lat][lon]) + ',' +
                                        str(self.t_10m[hr][lat][lon]) + ',' + str(self.d10m[hr][lat][lon]) + ',' +
                                        str(self.s10m[hr][lat][lon]) + ',' +
                                        str(self.d50m[hr][lat][lon]) + ',' + str(self.s50m[hr][lat][lon]) + '\n')
                        else:
                            tf.write('Temperature,Pressure,Direction,Speed,Direction,Speed' + '\n')
                            tf.write('C,atm,degrees,m/s,degrees,m/s' + '\n')
                            tf.write('2,0,2,2,50,50' + '\n')
                            for hr in range(len(self.s50m)):
                                tf.write(str(self.t_2m[hr][lat][lon]) + ',' + str(self.p_s[hr][lat][lon]) + ',' +
                                        str(self.d2m[hr][lat][lon]) + ',' + str(self.s2m[hr][lat][lon]) + ',' +
                                        str(self.d50m[hr][lat][lon]) + ',' + str(self.s50m[hr][lat][lon]) + '\n')
                        tf.close()
                        self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
                        self.checkZone()
            return  # that's it for wind
         # get variable from solar files
        if self.src_zone > 0:
            inp_strt = '{0:04d}'.format(self.src_year - 1) + '1231'
        elif self.src_zone <= 0:
            inp_strt = '{0:04d}'.format(self.src_year) + '0101'
        self.get_rad_data(self.findFile(inp_strt, False))  # get solar data
        if self.return_code != 0:
            return
        if self.src_zone != 0:
         # last self.src_zone hours of (self.src_year -1) GMT = first self.src_zone of self.src_year
         # first self.src_zone hours of self.src_year GMT = last self.src_zone of (self.src_year - 1)
            self.swgnt = self.swgnt[-self.src_zone:]
        elif self.src_zone < 0:
         # first self.src_zone hours of self.src_year GMT = last self.src_zone of (self.src_year - 1)
            self.swgnt = self.swgnt[:self.src_zone]
        self.the_year = self.src_year  # start with their year
        if self.wrap:
            yrs = 2
        else:
            yrs = 1
        for mt in range(len(dys)):
            self.log += 'Processing month %s solar\n' % str(mt + 1)
            for dy in range(1, dys[mt] + 1):
                if self.src_zone <= 0:
                    if mt == 0 and dy == 1:
                        continue
                found = False
                for yr in range(yrs):
                    inp_strt = '{0:04d}'.format(self.the_year) + '{0:02d}'.format(mt + 1) + \
                               '{0:02d}'.format(dy)
                    inp_file = self.findFile(inp_strt, False)
                    if inp_file is None:
                        if self.wrap and self.the_year == self.src_year:
                            self.the_year = self.src_year - 1
                            self.log += 'Wrapping to prior year - %.4d-%.2d-%.2d\n' % (self.the_year, (mt + 1), dy)
                            yrs = 1
                    if found:
                        break
                self.get_rad_data(inp_file)  # get solar data
                if self.return_code != 0:
                    return
        if self.src_zone > 0:
            for i in range(self.src_zone):  # delete last n hours
                del self.swgnt[-1]
        elif self.src_zone < 0:
            inp_strt = '{0:04d}'.format(self.the_year + 1) + '0101'
            self.get_rad_data(self.findFile(inp_strt, False))
            for i in range(24 + self.src_zone):  # delete last n hours
                del self.swgnt[-1]
        target_dir = self.tgt_dir
        self.log += 'Target directory is %s\n' % target_dir
        if not os.path.exists(target_dir):
            self.log += 'mkdir %s\n' % target_dir
            os.makedirs(target_dir)
        if self.src_lat is not None:  # specific location(s)
            for i in range(len(self.src_lat)):
                for lat2 in range(len(self.lati)):
                    if self.src_lat[i] <= self.lati[lat2]:
                        break
                for lon2 in range(len(self.longi)):
                    if self.src_lon[i] <= self.longi[lon2]:
                        break
                if self.longrange[0] is None:
                    self.longrange[0] = self.longi[lon2]
                else:
                    if self.longi[lon2] < self.longrange[0]:
                        self.longrange[0] = self.longi[lon2]
                if self.longrange[1] is None:
                    self.longrange[1] = self.longi[lon2]
                else:
                    if self.longi[lon2] > self.longrange[1]:
                        self.longrange[1] = self.longi[lon2]
                lat1 = lat2 - 1
                lat_rat = (self.lati[lat2] - self.src_lat[i]) / (self.lati[lat2] - self.lati[lat1])
                lon1 = lon2 - 1
                lon_rat = (self.longi[lon2] - self.src_lon[i]) / (self.longi[lon2] - self.longi[lon1])
                out_file = self.tgt_dir + 'solar_weather_' + \
                           str(self.src_lat[i]) + '_' + str(self.src_lon[i]) + '_' + str(self.src_year) + '.' + self.fmat
                tf = open(out_file, 'w')
                if self.fmat == 'csv':
                    hdr = 'Location,City,Region,Country,Latitude,Longitude,Time Zone,Elevation,Source\n' + \
                          'id,<city>,<state>,<country>,%s,%s,%s,0,IWEC\n' % (round(self.src_lat[i], 4),
                          round(self.src_lon[i], 4), str(self.src_zone))
                    tf.write(hdr)
                    tf.write('Year,Month,Day,Hour,GHI,DNI,DHI,Tdry,Pres,Wspd,Wdir' + '\n')
                    mth = 0
                    day = 1
                    hour = 1
                    for hr in range(len(self.s10m)):
                        ghi = self.valu(self.swgnt[hr], lat1, lon1, lat_rat, lon_rat)
                        dni = getDNI(ghi, hour=hr + 1, lat=self.src_lat[i], lon=self.src_lon[i],
                              press=self.valu(self.p_s[hr], lat1,
                              lon1, lat_rat, lon_rat, rnd=0), zone=self.src_zone)
                        dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.src_lat[i])
                        tf.write(str(self.src_year) + ',' + str(mth + 1) + ',' +
                        str(day) + ',' + str(hour) + ',' +
                        str(int(ghi)) + ',' + str(int(dni)) + ',' + str(int(dhi)) + ',' +
                        str(self.valu(self.t_10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +

                        str(self.valu(self.p_s[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                        str(self.valu(self.s10m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                        str(self.valu(self.d10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + '\n')
                        hour += 1
                        if hour > 24:
                            hour = 1
                            day += 1
                            if day > dys[mth]:
                                mth += 1
                                day = 1
                                hour = 1
                else:
                    hdr = 'id,<city>,<state>,%s,%s,%s,0,3600.0,%s,0:30:00\n' % (str(self.src_zone),
                          round(self.src_lat[i], 4),
                          round(self.src_lon[i], 4), str(self.src_year))
                    tf.write(hdr)
                    for hr in range(len(self.s10m)):
                        ghi = self.valu(self.swgnt[hr], lat1, lon1, lat_rat, lon_rat)
                        dni = getDNI(ghi, hour=hr + 1, lat=self.src_lat[i], lon=self.src_lon[i],
                              press=self.valu(self.p_s[hr], lat1, lon1, lat_rat,
                              lon_rat, rnd=0), zone=self.src_zone)
                        dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.src_lat[i])
                        tf.write(str(self.valu(self.t_10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) +
                        ',-999,-999,-999,' +
                        str(self.valu(self.s10m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                        str(self.valu(self.d10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                        str(self.valu(self.p_s[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                        str(int(ghi)) + ',' + str(int(dni)) + ',' + str(int(dhi)) +
                        ',-999,-999,\n')
                tf.close()
                self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
        else:
            for lat in range(len(self.s10m[0])):
                for lon in range(len(self.s10m[0][0])):
                    if self.longrange[0] is None:
                        self.longrange[0] = self.longi[lon]
                    else:
                        if self.longi[lon] < self.longrange[0]:
                            self.longrange[0] = self.longi[lon]
                    if self.longrange[1] is None:
                        self.longrange[1] = self.longi[lon]
                    else:
                        if self.longi[lon] > self.longrange[1]:
                            self.longrange[1] = self.longi[lon]
                    out_file = self.tgt_dir + 'solar_weather_' + '{:0.4f}'.format(self.lati[lat]) + \
                            '_' + '{:0.4f}'.format(self.longi[lon]) + '_' + str(self.src_year) + '.' + self.fmat
                    tf = open(out_file, 'w')
                    if self.fmat == 'csv':
                        hdr = 'Location,City,Region,Country,Latitude,Longitude,Time Zone,Elevation,Source\n' + \
                              'id,<city>,<state>,<country>,%s,%s,%s,0,IWEC\n' % (round(self.lati[lat], 4),
                              round(self.longi[lon], 4), str(self.src_zone))
                        tf.write(hdr)
                        tf.write('Year,Month,Day,Hour,GHI,DNI,DHI,Tdry,Pres,Wspd,Wdir' + '\n')
                        mth = 0
                        day = 1
                        hour = 1
                        for hr in range(len(self.s10m)):
                            ghi = self.swgnt[hr][lat][lon]
                            dni = getDNI(ghi, hour=hr + 1, lat=self.lati[lat], lon=self.longi[lon],
                                  press=self.p_s[hr][lat][lon], zone=self.src_zone)
                            dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.lati[lat])
                            tf.write(str(self.src_year) + ',' + '{0:02d}'.format(mth + 1) + ',' +
                                '{0:02d}'.format(day) + ',' + '{0:02d}'.format(hour) + ',' +
                                '{:0.1f}'.format(ghi) + ',' + '{:0.1f}'.format(dni) + ',' +
                                '{:0.1f}'.format(dhi) + ',' +
                                str(self.t_10m[hr][lat][lon]) + ',' +
                                str(self.p_s[hr][lat][lon]) + ',' +
                                str(self.s10m[hr][lat][lon]) + ',' +
                                str(self.d10m[hr][lat][lon]) + '\n')
                            hour += 1
                            if hour > 24:
                                hour = 1
                                day += 1
                                if day > dys[mth]:
                                    mth += 1
                                    day = 1
                                    hour = 1
                    else:
                        hdr = 'id,<city>,<state>,%s,%s,%s,0,3600.0,%s,0:30:00\n' % (str(self.src_zone),
                              round(self.lati[lat], 4),
                              round(self.longi[lon], 4), str(self.src_year))
                        tf.write(hdr)
                        for hr in range(len(self.s10m)):
                            ghi = self.swgnt[hr][lat][lon]
                            dni = getDNI(ghi, hour=hr + 1, lat=self.lati[lat], lon=self.longi[lon],
                                  press=self.p_s[hr][lat][lon], zone=self.src_zone)
                            dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.lati[lat])
                            tf.write(str(self.t_10m[hr][lat][lon]) +
                                ',-999,-999,-999,' +
                                str(self.s10m[hr][lat][lon]) + ',' +
                                str(self.d10m[hr][lat][lon]) + ',' +
                                str(self.p_s[hr][lat][lon]) + ',' +
                                str(int(ghi)) + ',' + str(int(dni)) + ',' + str(int(dhi)) +
                                ',-999,-999,\n')
                    tf.close()
                    self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
                    self.checkZone()

class ClickableQLabel(QtGui.QLabel):
    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtGui.QApplication.widgetAt(event.globalPos()).setFocus()
        self.emit(QtCore.SIGNAL('clicked()'))


class getParms(QtGui.QWidget):

    def __init__(self, help='makeweatherfiles.html'):
        super(getParms, self).__init__()
        self.help = help
        self.initUI()

    def initUI(self):
        self.grid = QtGui.QGridLayout()
        self.grid.addWidget(QtGui.QLabel('Year:'), 0, 0)
        self.yearSpin = QtGui.QSpinBox()
        now = datetime.datetime.now()
        self.yearSpin.setRange(1979, now.year)
        self.yearSpin.setValue(now.year - 1)
        self.grid.addWidget(self.yearSpin, 0, 1)
        self.grid.addWidget(QtGui.QLabel('Wrap to prior year:'), 1, 0)
        self.wrapbox = QtGui.QCheckBox()
        self.wrapbox.setCheckState(QtCore.Qt.Checked)
        self.grid.addWidget(self.wrapbox, 1, 1)
        self.grid.addWidget(QtGui.QLabel('If checked will wrap back to prior year'), 1, 2, 1, 3)
        self.grid.addWidget(QtGui.QLabel('Time Zone:'), 2, 0)
        self.zoneCombo = QtGui.QComboBox()
        self.zoneCombo.addItem('Auto')
        for i in range(-12, 13):
            self.zoneCombo.addItem(str(i))
        self.zoneCombo.currentIndexChanged[str].connect(self.zoneChanged)
        self.zone_lon = QtGui.QLabel(('Time zone calculated from MERRA data'))
        self.grid.addWidget(self.zoneCombo, 2, 1)
        self.grid.addWidget(self.zone_lon, 2, 2, 1, 3)
        self.grid.addWidget(QtGui.QLabel('Solar Format:'), 3, 0)
        self.fmatcombo = QtGui.QComboBox(self)
        self.fmats = ['csv', 'smw']
        for i in range(len(self.fmats)):
            self.fmatcombo.addItem(self.fmats[i])
        self.fmatcombo.setCurrentIndex(1)
        self.grid.addWidget(self.fmatcombo, 3, 1)
        self.grid.addWidget(QtGui.QLabel('Coordinates:'), 4, 0)
        self.coords = QtGui.QPlainTextEdit()
        self.grid.addWidget(self.coords, 4, 1, 1, 4)
        self.grid.addWidget(QtGui.QLabel('Copy folder down:'), 5, 0)
        self.checkbox = QtGui.QCheckBox()
        self.checkbox.setCheckState(QtCore.Qt.Checked)
        self.grid.addWidget(self.checkbox, 5, 1)
        self.grid.addWidget(QtGui.QLabel('If checked will copy solar folder changes down to others'), 5, 2, 1, 3)
        cur_dir = os.getcwd()
        self.dir_labels = ['Solar Source', 'Wind Source', 'Target']
        self.dirs = [None, None, None, None]
        for i in range(3):
            self.grid.addWidget(QtGui.QLabel(self.dir_labels[i] + ' Folder:'), 6 + i, 0)
            self.dirs[i] = ClickableQLabel()
            self.dirs[i].setText(cur_dir)
            self.dirs[i].setFrameStyle(6)
            self.connect(self.dirs[i], QtCore.SIGNAL('clicked()'), self.dirChanged)
            self.grid.addWidget(self.dirs[i], 6 + i, 1, 1, 4)
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, 9, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        dosolar = QtGui.QPushButton('Produce Solar Files', self)
        wdth = dosolar.fontMetrics().boundingRect(dosolar.text()).width() + 9
        self.grid.addWidget(dosolar, 9, 1)
        dosolar.clicked.connect(self.dosolarClicked)
        dowind = QtGui.QPushButton('Produce Wind Files', self)
        dowind.setMaximumWidth(wdth)
        self.grid.addWidget(dowind, 9, 2)
        dowind.clicked.connect(self.dowindClicked)
        help = QtGui.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        self.grid.addWidget(help, 9, 3)
        help.clicked.connect(self.helpClicked)
        QtGui.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        info = QtGui.QPushButton('File Info', self)
        info.setMaximumWidth(wdth)
        self.grid.addWidget(info, 9, 4)
        info.clicked.connect(self.infoClicked)
      #   self.grid.setColumnStretch(4, 2)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('makeweatherfiles - Make weather files from MERRA data')
        self.center()
        self.resize(int(self.sizeHint().width()* 1.07), int(self.sizeHint().height() * 1.07))
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
        centerPoint = QtGui.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def zoneChanged(self, val):
        if self.zoneCombo.currentIndex() == 0:
            lon = '(Time zone calculated from MERRA data)'
        else:
            lw = int(self.zoneCombo.currentIndex() - 13) * 15 - 7.5
            le = int(self.zoneCombo.currentIndex() - 13) * 15 + 7.5
            if lw < -180:
                lw = lw + 360
            if le > 180:
                le = le - 360
            lon = '(Approx. Longitude: %s to %s)' % ('{:0.1f}'.format(lw), '{:0.1f}'.format(le))
        self.zone_lon.setText(lon)

    def zoomChanged(self, val):
        self.zoomScale.setText('(' + scale[int(val)] + ')')
        self.zoomScale.adjustSize()

    def dirChanged(self):
        for i in range(3):
            if self.dirs[i].hasFocus():
                break
        curdir = self.dirs[i].text()
        newdir = str(QtGui.QFileDialog.getExistingDirectory(self, 'Choose ' + self.dir_labels[i] + ' Folder',
                 curdir, QtGui.QFileDialog.ShowDirsOnly))
        if newdir != '':
            self.dirs[i].setText(newdir)
            if self.checkbox.isChecked():
                if i == 0:
                    self.dirs[1].setText(newdir)
                    self.dirs[2].setText(newdir)

    def helpClicked(self):
        dialog = AnObject(QtGui.QDialog(), self.help,
                 title='makeweatherfiles - Help', section='makeweather')
        dialog.exec_()

    def quitClicked(self):
        self.close()

    def dosolarClicked(self):
        coords = str(self.coords.toPlainText())
        if coords != '':
            if coords.find(' ') >= 0:
                if coords.find(',') < 0:
                    coords = coords.replace(' ', ',')
                else:
                    coords = coords.replace(' ', '')
        if self.wrapbox.isChecked():
            wrap = 'y'
        else:
            wrap = ''
        if self.zoneCombo.currentIndex() == 0:
            zone = 'auto'
        else:
            zone = str(self.zoneCombo.currentIndex() - 13)
        solar = makeWeather(str(self.yearSpin.value()), zone, str(self.dirs[0].text()), \
                            str(self.dirs[1].text()), str(self.dirs[2].text()), str(self.fmatcombo.currentText()), \
                            wrap, coords)
        dialr = RptDialog(str(self.yearSpin.value()), zone, str(self.dirs[0].text()), \
                          str(self.dirs[1].text()), str(self.dirs[2].text()), str(self.fmatcombo.currentText()), \
                          wrap, coords, solar.returnCode(), solar.getLog())
        dialr.exec_()
        del solar
        del dialr

    def dowindClicked(self):
        coords = str(self.coords.toPlainText())
        if coords != '':
            if coords.find(' ') >= 0:
                if coords.find(',') < 0:
                    coords = coords.replace(' ', ',')
                else:
                    coords = coords.replace(' ', '')
        if self.wrapbox.isChecked():
            wrap = 'y'
        else:
            wrap = ''
        if self.zoneCombo.currentIndex() == 0:
            zone = 'auto'
        else:
            zone = str(self.zoneCombo.currentIndex() - 13)
        wind = makeWeather(str(self.yearSpin.value()), zone, str(self.dirs[0].text()), \
                            str(self.dirs[1].text()), str(self.dirs[2].text()), \
                            'wind', wrap, coords)
        dialr = RptDialog(str(self.yearSpin.value()), zone, str(self.dirs[0].text()), \
                          str(self.dirs[1].text()), str(self.dirs[2].text()), 'srw', \
                          wrap, coords, wind.returnCode(), wind.getLog())
        dialr.exec_()
        del wind
        del dialr

    def infoClicked(self):
        coords = str(self.coords.toPlainText())
        if coords != '':
            if coords.find(' ') >= 0:
                if coords.find(',') < 0:
                    coords = coords.replace(' ', ',')
                else:
                    coords = coords.replace(' ', '')
        if self.wrapbox.isChecked():
            wrap = 'y'
        else:
            wrap = ''
        if self.zoneCombo.currentIndex() == 0:
            zone = 'auto'
        else:
            zone = str(self.zoneCombo.currentIndex() - 13)
        wind = makeWeather(str(self.yearSpin.value()), zone, str(self.dirs[0].text()), \
                            str(self.dirs[1].text()), str(self.dirs[2].text()), \
                            'both', wrap, coords, info=True)
        dialr = RptDialog(str(self.yearSpin.value()), zone, str(self.dirs[0].text()), \
                          str(self.dirs[1].text()), str(self.dirs[2].text()), 'srw', \
                          wrap, coords, wind.returnCode(), wind.getLog())
        dialr.exec_()
        del wind
        del dialr


class RptDialog(QtGui.QDialog):
    def __init__(self, year, zone, solar_dir, wind_dir, tgt_dir, fmat, wrap, coords, return_code, output):
        super(RptDialog, self).__init__()
        self.parms = [str(year), str(zone), wrap, fmat, tgt_dir, solar_dir, wind_dir]
        if wrap == 'y':
            wrapy = 'Yes'
        else:
            wrapy = 'No'
        max_line = 0
        self.lines = 'Parameters:\n    Year: %s\n    Wrap year: %s\n    Time Zone: %s\n    Output Format: %s\n' \
                     % (year, wrapy, zone, fmat)
        if fmat != 'srw':
            self.lines += '    Solar Files: %s\n' % solar_dir
        self.lines += '    Wind Files: %s\n' % wind_dir
        self.lines += '    Target Folder: %s\n' % tgt_dir
        if coords != '':
            self.lines += '    Coordinates: %s\n' % coords
        self.lines += 'Return Code:\n    %s\n' % return_code
        self.lines += 'Output:\n'
        self.lines += output
        lenem = self.lines.split('\n')
        line_cnt = len(lenem)
        for i in range(line_cnt):
            max_line = max(max_line, len(lenem[i]))
        del lenem
        QtGui.QDialog.__init__(self)
        self.saveButton = QtGui.QPushButton(self.tr('&Save'))
        self.cancelButton = QtGui.QPushButton(self.tr('Cancel'))
        buttonLayout = QtGui.QHBoxLayout()
        buttonLayout.addStretch(1)
        buttonLayout.addWidget(self.saveButton)
        buttonLayout.addWidget(self.cancelButton)
        self.connect(self.saveButton, QtCore.SIGNAL('clicked()'), self,
                     QtCore.SLOT('accept()'))
        self.connect(self.cancelButton, QtCore.SIGNAL('clicked()'),
                     self, QtCore.SLOT('reject()'))
        self.widget = QtGui.QTextEdit()
        self.widget.setFont(QtGui.QFont('Courier New', 11))
        fnt = self.widget.fontMetrics()
        ln = (max_line + 5) * fnt.maxWidth()
        ln2 = (line_cnt + 2) * fnt.height()
        screen = QtGui.QDesktopWidget().availableGeometry()
        if ln > screen.width() * .67:
            ln = int(screen.width() * .67)
        if ln2 > screen.height() * .67:
            ln2 = int(screen.height() * .67)
        self.widget.setSizePolicy(QtGui.QSizePolicy(QtGui.QSizePolicy.Expanding,
            QtGui.QSizePolicy.Expanding))
        self.widget.resize(ln, ln2)
        self.widget.setPlainText(self.lines)
        layout = QtGui.QVBoxLayout()
        layout.addWidget(self.widget)
        layout.addLayout(buttonLayout)
        self.setLayout(layout)
        self.setWindowTitle('makeweatherfiles - Output')
        size = self.geometry()
        self.setGeometry(1, 1, ln + 10, ln2 + 35)
        size = self.geometry()
        self.move((screen.width() - size.width()) / 2,
            (screen.height() - size.height()) / 2)
        self.widget.show()

    def accept(self):
        i = sys.argv[0].rfind('/')  # fudge to see if program has a directory to use as an alternative
        j = sys.argv[0].rfind('.')
        if i > 0:
            save_filename = sys.argv[0][i + 1:j]
        else:
            save_filename = sys.argv[0][:j]
        last_bit = ''
        for k in range(len(self.parms)):
            if self.parms[k] == '':
                continue
            i = self.parms[k].rfind('/')
            if i > 0:
                if self.parms[k][i + 1:] != last_bit:
                    save_filename += '_' + self.parms[k][i + 1:]
                    last_bit = self.parms[k][i + 1:]
            else:
                if self.parms[k] != last_bit:
                    save_filename += '_' + self.parms[k]
                    last_bit = self.parms[k]
        save_filename += '.txt'
        fileName = QtGui.QFileDialog.getSaveFileName(self,
                                         self.tr("QFileDialog.getSaveFileName()"),
                                         save_filename,
                                         self.tr("All Files (*);;Text Files (*.txt)"))
        if not fileName.isEmpty():
            s = open(fileName, 'w')
            s.write(self.lines)
            s.close()
        self.close()


if "__main__" == __name__:
    app = QtGui.QApplication(sys.argv)
    if len(sys.argv) > 1:  # arguments
        src_lat_lon = ''
        src_year = 2014
        wrap = ''
        src_zone = 0
        fmat = 'srw'
        src_dir_s = ''
        src_dir_w = ''
        tgt_dir = ''
        for i in range(1, len(sys.argv)):
            if sys.argv[i][:5] == 'year=':
                src_year = int(sys.argv[i][5:])
            elif sys.argv[i][:5] == 'wrap=':
                wrap = int(sys.argv[i][5:])
            elif sys.argv[i][:7] == 'latlon=' or sys.argv[i][:7] == 'coords=':  # lat and lon
                src_lat_lon = sys.argv[i][7:]
            elif sys.argv[i][:5] == 'zone=':
                src_zone = int(sys.argv[i][5:])
            elif sys.argv[i][:9] == 'timezone=':
                src_zone = int(sys.argv[i][9:])
            elif sys.argv[i][:5] == 'fmat=':
                fmat = sys.argv[i][5:]
            elif sys.argv[i][:7] == 'format=':
                fmat = sys.argv[i][7:]
            elif sys.argv[i][:6] == 'solar=':
                src_dir_s = sys.argv[i][6:]
            elif sys.argv[i][:7] == 'source=' or sys.argv[i][:7] == 'srcdir=':
                src_dir_w = sys.argv[i][7:]
            elif sys.argv[i][:5] == 'wind=':
                src_dir_w = sys.argv[i][5:]
            elif sys.argv[i][:7] == 'target=' or sys.argv[i][:7] == 'tgtdir=':
                tgt_dir = sys.argv[i][7:]
        files = makeWeather(src_year, src_zone, src_dir_s, src_dir_w, tgt_dir, fmat, wrap, src_lat_lon)
        dialr = RptDialog(files.returnCode(), files.getLog())
        dialr.exec_()
    else:
        ex = getParms()
        app.exec_()
        app.deleteLater()
        sys.exit()
