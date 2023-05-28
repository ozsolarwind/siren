#!/usr/bin/python3
#
#  Copyright (C) 2015-2023 Sustainable Energy Now Inc., Angus King
#
#  makeweatherfiles.py - This file is part of SIREN.
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
from datetime import datetime, timedelta
import gzip
from math import *
import os
import sys
import time
from netCDF4 import Dataset
import configparser   # decode .ini file
from PyQt5 import QtCore, QtGui, QtWidgets
if sys.platform == 'win32' or sys.platform == 'cygwin':
    from win32api import GetFileVersionInfo, LOWORD, HIWORD

from credits import fileVersion
from editini import SaveIni
from getmodels import getModelFile
from senutils import ClickableQLabel, getUser, extrapolateWind
from sammodels import getDNI, getDHI


class ShowHelp(QtWidgets.QDialog):

    def __init__(self, dialog, anobject, title=None, section=None):
        super(ShowHelp, self).__init__()
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
            quit = QtWidgets.QPushButton('Close', self)
            width = quit.fontMetrics().boundingRect('Close').width() + 10
            quit.setMaximumWidth(width)
        else:
            quit = QtWidgets.QPushButton('Quit', self)
        grid.addWidget(quit, i + 1, 0)
        quit.clicked.connect(self.quitClicked)
        self.setLayout(grid)
        screen = QtWidgets.QDesktopWidget().availableGeometry()
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
        grid = QtWidgets.QGridLayout()
        self.web = QtWidgets.QTextEdit()
        self.web.acceptRichText = True
        if os.path.exists(self.anobject):
            htf = open(self.anobject, 'r')
            html = htf.read()
            htf.close()
            if self.section is not None:
                b = html.find('<body')
                s = html.find('id="' + self.section)
                s = html[:s].rfind('<h')
                e = html[s + 1:].find(html[s:s+3]) + s
                html = html[:b] + '<body>' + html[s:e] + '</body></html>'
            self.web.setHtml(html)
        else:
            html = self.anobject
            if self.anobject[:5] == '<html':
                self.anobject = self.anobject.replace('[VERSION]', fileVersion())
                self.web.setHtml(self.anobject)
            else:
                self.web.setPlainText(self.anobject)
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
            screen = QtWidgets.QDesktopWidget().availableGeometry()
            if widths[0] > screen.width() * .67:
                heights = int(heights / .67)
                widths[0] = int(screen.width() * .67)
        self.web.setReadOnly(True)
        i = 1
        grid.addWidget(self.web, 0, 0)
        self.set_stuff(grid, widths, heights, i)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)

    def quitClicked(self):
        self.close()


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
                    lm.append(round(sqrt(um[hr][lat][lon] * um[hr][lat][lon] +
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
                        theta = atan(um[hr][lat][lon] / vm[hr][lat][lon])
                        theta = degrees(theta)
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

    def getGHI(self, pmi, watts=True):
        try:
            pm = pmi.data
        except:
            pm = pmi[:]
        ps = []
        for hr in range(len(pm)):   # 24 or 8760 hours
            hp = []
            for lat in range(len(pm[hr])):   # n latitudes
                lp = []
                for lon in range(len(pm[hr][lat])):   # m longitudes
                    if watts:
                        lp.append(pm[hr][lat][lon])
                    else:
                        lp.append(pm[hr][lat][lon] / 3600) # joules so / seconds in an hour
                hp.append(lp)
            ps.append(hp)
        return ps

    def getAlbedo(self, alb1, alb2i=None):
        try:
            alb = alb1.data
        except:
            alb = alb1[:]
        if alb2i is not None:
            try:
                alb2 = alb2i.data
            except:
                alb2 = alb2i[:]
            do_2 = True
        else:
            do_2 = false
        tmp = []
        for hr in range(len(alb)):   # 24 or 8760 hours
            at = []
            for lat in range(len(alb[hr])):   # n latitudes
                ap = []
                for lon in range(len(alb[hr][lat])):   # m longitudes
                    ap.append(alb[hr][lat][lon])
                    if do_2:
                        ap[-1] = (ap[-1] + alb2[hr][lat][lon]) / 2
                at.append(ap)
            tmp.append(at)
        return tmp

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
            cdf_file = Dataset(unzip_file, 'r')
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
        self.tims = cdf_file.variables['time'][:]
        self.lat_lon_ndx += [len(self.lati)] * len(self.tims)
        lats = cdf_file.variables[self.vars['latitude']][:]
        self.lati.append([])
        for lat in lats:
            self.lati[-1].append(lat)
        lons = cdf_file.variables[self.vars['longitude']][:]
        self.longi.append([])
        for lon in lons:
            self.longi[-1].append(lon)
        for lat in self.lati[-1]:
            try:
                self.lats.index(lat)
            except:
                self.lats.append(lat)
        for lon in self.longi[-1]:
            try:
                self.lons.index(lon)
            except:
                self.lons.append(lon)
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

    def get_era5_data(self, inp_file, frst_hour, last_hour):
        unzip_file = self.unZip(inp_file)
        if self.return_code != 0:
            return
        try:
            cdf_file = Dataset(unzip_file, 'r')
        except:
            self.decodeError(inp_file)
            return

     #   Variable Description                            Units
     #   -------- -------------------------------------- --------
     #   sp       Surface pressure                       Pa
     #   ssrd     Surface solar radiation downwards      W m-2
     #   ssr      Surface net solar radiation            W m-2
     #   t2m      2m temperature                         K
     #   u10      10m u-component of wind                m/s
     #   u100     10m u-component of wind                m/s
     #   v10      10m v-component of wind                m/s
     #   v100     100m u-component of wind               m/s
     #   alnip    near_ir_albedo_for_direct_radiation
     #   aluvp    uv_visible_albedo_for_direct_radiation
        expver = False
        if 'expver' in cdf_file.variables.keys():
            self.logMsg('ERA5 and ERA5T data in {}'.format(inp_file[inp_file.rfind('/') + 1:]))
            expver = True
        self.tims = cdf_file.variables['time'][:]
        t1 = -1
        t2 = len(self.tims)
        for hr in range(len(self.tims)):
            if t1 < 0 and self.tims[hr] >= frst_hour:
                t1 = hr
            if self.tims[hr] >= last_hour:
                t2 = hr
                break
        self.lat_lon_ndx += [len(self.lati)] * (t2 - t1 + 1)
        lats = cdf_file.variables[self.vars['latitude']][:]
        self.lati.append([])
        for lat in lats:
            self.lati[-1].append(lat)
        lons = cdf_file.variables[self.vars['longitude']][:]
        self.longi.append([])
        for lon in lons:
            self.longi[-1].append(lon)
        for lat in self.lati[-1]:
            try:
                self.lats.index(lat)
            except:
                self.lats.append(lat)
        for lon in self.longi[-1]:
            try:
                self.lons.index(lon)
            except:
                self.lons.append(lon)
        if self.show_progress:
            self.caller.daybar.setValue(0)
            self.caller.daybar.setMaximum(7)
            QtCore.QCoreApplication.processEvents()
        if expver:
            # need to find valid value in the two expver dimensions
            tmp_var = []
            for t in range(len(self.tims)):
                tmp_var.append([])
                for la in range(len(self.lats)):
                    tmp_var[-1].append([])
                    for lo in range(len(self.lons)):
                        tmp_var[-1][-1].append([])
            tmi = cdf_file.variables[self.vars['t2m']][t1 : t2]
            for t in range(len(tmp_var)):
                for la in range(len(tmp_var[t])):
                    for lo in range(len(tmp_var[t][la])):
                        if isinstance(tmi[t][0][la][lo], float):
                            tmp_var[t][la][lo] = float(tmi[t][0][la][lo])
                        else:
                            tmp_var[t][la][lo] = float(tmi[t][1][la][lo])
            self.t_2m += self.getTemp(tmp_var)
            if self.show_progress:
                self.caller.daybar.setValue(1)
                QtCore.QCoreApplication.processEvents()
            tmp_var = []
            for t in range(len(self.tims)):
                tmp_var.append([])
                for la in range(len(self.lats)):
                    tmp_var[-1].append([])
                    for lo in range(len(self.lons)):
                        tmp_var[-1][-1].append([])
            tmi = cdf_file.variables[self.vars['v10']][t1 : t2]
            for t in range(len(tmp_var)):
                for la in range(len(tmp_var[t])):
                    for lo in range(len(tmp_var[t][la])):
                        if isinstance(tmi[t][0][la][lo], float):
                            tmp_var[t][la][lo] = float(tmi[t][0][la][lo])
                        else:
                            tmp_var[t][la][lo] = float(tmi[t][1][la][lo])
            tmp_var2 = []
            for t in range(len(self.tims)):
                tmp_var2.append([])
                for la in range(len(self.lats)):
                    tmp_var2[-1].append([])
                    for lo in range(len(self.lons)):
                        tmp_var2[-1][-1].append([])
            tmi = cdf_file.variables[self.vars['u10']][t1 : t2]
            for t in range(len(tmp_var2)):
                for la in range(len(tmp_var2[t])):
                    for lo in range(len(tmp_var2[t][la])):
                        if isinstance(tmi[t][0][la][lo], float):
                            tmp_var2[t][la][lo] = float(tmi[t][0][la][lo])
                        else:
                            tmp_var2[t][la][lo] = float(tmi[t][1][la][lo])
            self.s10m += self.getSpeed(tmp_var, tmp_var2)
            if self.show_progress:
                self.caller.daybar.setValue(2)
                QtCore.QCoreApplication.processEvents()
            self.d10m += self.getDirn(tmp_var, tmp_var2)
            if self.show_progress:
                self.caller.daybar.setValue(3)
                QtCore.QCoreApplication.processEvents()
            tmp_var = []
            for t in range(len(self.tims)):
                tmp_var.append([])
                for la in range(len(self.lats)):
                    tmp_var[-1].append([])
                    for lo in range(len(self.lons)):
                        tmp_var[-1][-1].append([])
            tmi = cdf_file.variables[self.vars['sp']][t1 : t2]
            for t in range(len(tmp_var)):
                for la in range(len(tmp_var[t])):
                    for lo in range(len(tmp_var[t][la])):
                        if isinstance(tmi[t][0][la][lo], float):
                            tmp_var[t][la][lo] = float(tmi[t][0][la][lo])
                        else:
                            tmp_var[t][la][lo] = float(tmi[t][1][la][lo])
            self.p_s += self.getPress(tmp_var)
            if self.show_progress:
                self.caller.daybar.setValue(4)
                QtCore.QCoreApplication.processEvents()
            try:
                if self.vars[self.swg] not in cdf_file.variables.keys():
                    self.swg = 'swgnt'
                tmp_var = []
                for t in range(len(self.tims)):
                    tmp_var.append([])
                    for la in range(len(self.lats)):
                        tmp_var[-1].append([])
                        for lo in range(len(self.lons)):
                            tmp_var[-1][-1].append([])
                tmi = cdf_file.variables[self.vars[self.swg]][t1 : t2]
                for t in range(len(tmp_var)):
                    for la in range(len(tmp_var[t])):
                        for lo in range(len(tmp_var[t][la])):
                            if isinstance(tmi[t][0][la][lo], float):
                                tmp_var[t][la][lo] = float(tmi[t][0][la][lo])
                            else:
                                tmp_var[t][la][lo] = float(tmi[t][1][la][lo])
                self.ghi += self.getGHI(tmp_var, watts=False)
                if self.show_progress:
                    self.caller.daybar.setValue(5)
                    QtCore.QCoreApplication.processEvents()
            except:
                pass
            if self.make_wind:
                tmp_var = []
                for t in range(len(self.tims)):
                    tmp_var.append([])
                    for la in range(len(self.lats)):
                        tmp_var[-1].append([])
                        for lo in range(len(self.lons)):
                            tmp_var[-1][-1].append([])
                tmi = cdf_file.variables[self.vars['v100']][t1 : t2]
                for t in range(len(tmp_var)):
                    for la in range(len(tmp_var[t])):
                        for lo in range(len(tmp_var[t][la])):
                            if isinstance(tmi[t][0][la][lo], float):
                                tmp_var[t][la][lo] = float(tmi[t][0][la][lo])
                            else:
                                tmp_var[t][la][lo] = float(tmi[t][1][la][lo])
                tmp_var2 = []
                for t in range(len(self.tims)):
                    tmp_var2.append([])
                    for la in range(len(self.lats)):
                        tmp_var2[-1].append([])
                        for lo in range(len(self.lons)):
                            tmp_var2[-1][-1].append([])
                tmi = cdf_file.variables[self.vars['u100']][t1 : t2]
                for t in range(len(tmp_var2)):
                    for la in range(len(tmp_var2[t])):
                        for lo in range(len(tmp_var2[t][la])):
                            if isinstance(tmi[t][0][la][lo], float):
                                tmp_var2[t][la][lo] = float(tmi[t][0][la][lo])
                            else:
                                tmp_var2[t][la][lo] = float(tmi[t][1][la][lo])
                self.s100m += self.getSpeed(tmp_var, tmp_var2)
                if self.show_progress:
                    self.caller.daybar.setValue(6)
                    QtCore.QCoreApplication.processEvents()
                self.d100m += self.getDirn(tmp_var, tmp_var2)
                if self.show_progress:
                    self.caller.daybar.setValue(7)
                    QtCore.QCoreApplication.processEvents()
        else:
            self.t_2m += self.getTemp(cdf_file.variables[self.vars['t2m']][t1 : t2])
            if self.show_progress:
                self.caller.daybar.setValue(1)
                QtCore.QCoreApplication.processEvents()
            self.s10m += self.getSpeed(cdf_file.variables[self.vars['v10']][t1 : t2] , cdf_file.variables[self.vars['u10']][t1 : t2])
            if self.show_progress:
                self.caller.daybar.setValue(2)
                QtCore.QCoreApplication.processEvents()
            self.d10m += self.getDirn(cdf_file.variables[self.vars['v10']][t1 : t2], cdf_file.variables[self.vars['u10']][t1 : t2])
            if self.show_progress:
                self.caller.daybar.setValue(3)
                QtCore.QCoreApplication.processEvents()
            self.p_s += self.getPress(cdf_file.variables[self.vars['sp']][t1 : t2])
            if self.show_progress:
                self.caller.daybar.setValue(4)
                QtCore.QCoreApplication.processEvents()
            try:
                if self.vars[self.swg] not in cdf_file.variables.keys():
                    self.swg = 'swgnt'
                self.ghi += self.getGHI(cdf_file.variables[self.vars[self.swg]][t1 : t2], watts=False)
                if self.show_progress:
                    self.caller.daybar.setValue(5)
                    QtCore.QCoreApplication.processEvents()
            except:
                pass
            if self.make_wind:
                self.s100m += self.getSpeed(cdf_file.variables[self.vars['v100']][t1 : t2], cdf_file.variables[self.vars['u100']][t1 : t2])
                if self.show_progress:
                    self.caller.daybar.setValue(6)
                    QtCore.QCoreApplication.processEvents()
                self.d100m += self.getDirn(cdf_file.variables[self.vars['v100']][t1 : t2], cdf_file.variables[self.vars['u100']][t1 : t2])
                if self.show_progress:
                    self.caller.daybar.setValue(7)
                    QtCore.QCoreApplication.processEvents()
      #      else:
       #         if self.vars['alb'] in cdf_file.variables.keys():
        #            if self.vars['alb2'] in cdf_file.variables.keys():
         #               self.alb += self.getAlbedo(cdf_file.variables[self.vars['alb']][t1 : t2], cdf_file.variables[self.vars['alb2']][t1 : t2])
          #          else:
           #             self.alb += self.getAlbedo(cdf_file.variables[self.vars['alb']][t1 : t2])
            #    elif self.vars['alb2'] in cdf_file.variables.keys():
             #       self.alb += self.getAlbedo(cdf_file.variables[self.vars['alb2']][t1 : t2])
        cdf_file.close()

    def get_rad_data(self, inp_file):
        unzip_file = self.unZip(inp_file)
        if self.return_code != 0:
            return
        try:
            cdf_file = Dataset(unzip_file, 'r')
        except:
            self.decodeError(inp_file)
            return
     #   Variable Description                          Units
     #   -------- ------------------------------------ --------
     #   swgdn    Surface incoming shortwave flux flux W m-2
     #   swgnt    Surface net downward shortwave flux  W m-2
     #   albedo   surface albedo
        self.tims = cdf_file.variables['time'][:]
        lats = cdf_file.variables[self.vars['latitude']][:]
        self.latsi.append([])
        for lat in lats:
            self.latsi[-1].append(lat)
        lons = cdf_file.variables[self.vars['longitude']][:]
        self.longsi.append([])
        for lon in lons:
            self.longsi[-1].append(lon)
        if self.vars[self.swg] in cdf_file.variables:
            self.ghi += self.getGHI(cdf_file.variables[self.vars[self.swg]])
        else:
            self.swg = 'swgnt'
            self.ghi += self.getGHI(cdf_file.variables[self.vars['swgnt']])
        if self.vars['alb'] in cdf_file.variables:
            self.alb += self.getAlbedo(cdf_file.variables[self.vars['alb']])
        cdf_file.close()

    def checkZone(self):
        self.return_code = 0
        if self.longrange[0] is not None:
            if int(round(self.longrange[0] / 15)) != self.src_zone:
                self.log += '%s west longitude (%s) in different time zone: %s\n' % (self.dataset,
                            '{:0.4f}'.format(self.longrange[0]), int(round(self.longrange[0] / 15)))
                self.return_code = 1
        if self.longrange[1] is not None:
            if int(round(self.longrange[1] / 15)) != self.src_zone:
                self.log += '%s east longitude (%s) in different time zone: %s\n' % (self.dataset,
                            '{:0.4f}'.format(self.longrange[1]), int(round(self.longrange[1] / 15)))
                self.return_code = 1
        return

    def close(self):
        return

    def getLog(self):
        if self.gaplog != '':
            self.log += self.gaplog
        return self.log

    def logMsg(self, msg, progress=False):
        now = datetime.now()
        self.log += now.strftime('%Y-%m-%d %H:%M:%S') + '. ' + msg + '\n'
        if self.show_progress:
            self.caller.progresslabel.setText(msg)
            if progress:
                self.caller.progress(progress)
            QtCore.QCoreApplication.processEvents()

    def returnCode(self):
        return str(self.return_code)

    def findFile(self, inp_strt, wind=True, quiet=False):
        if self.era5:
            for p in range(len(self.src_s_pfx)):
                inp_file = self.src_dir_s + self.src_s_pfx[p] + inp_strt + self.src_s_sfx[p]
                if os.path.exists(inp_file):
                    break
                else:
                    if self.yearly[0]:
                        inp_file = self.src_dir_s[:-5] + str(int(inp_strt[:4])) + '/' + self.src_s_pfx[p] + inp_strt + self.src_s_sfx[p]
                        if os.path.exists(inp_file):
                            break
            else:
                if not quiet:
                    self.log += 'No ERA5 file found for ' + inp_strt + '\n'
                return None
        elif wind:
            for p in range(len(self.src_w_pfx)):
                inp_file = self.src_dir_w + self.src_w_pfx[p] + inp_strt + self.src_w_sfx[p]
                if os.path.exists(inp_file):
                    break
                else:
                    if self.yearly[1]:
                        inp_file = self.src_dir_w[:-5] + str(int(inp_strt[:4])) + '/' + self.src_w_pfx[p] + inp_strt + self.src_w_sfx[p]
                        if os.path.exists(inp_file):
                            break
                    if inp_file.find('MERRA300') >= 0:
                        inp_file = inp_file.replace('MERRA300', 'MERRA301')
                        if os.path.exists(inp_file):
                            break
            else:
                if not quiet:
                    self.log += 'No Wind file found for ' + inp_strt + '\n'
                return None
        else:
            for p in range(len(self.src_s_pfx)):
                inp_file = self.src_dir_s + self.src_s_pfx[p] + inp_strt + self.src_s_sfx[p]
                if os.path.exists(inp_file):
                    break
                else:
                    if self.yearly[0]:
                        inp_file = self.src_dir_s[:-5] + str(int(inp_strt[:4])) + '/' + self.src_s_pfx[p] + inp_strt + self.src_s_sfx[p]
                        if os.path.exists(inp_file):
                            break
                    if inp_file.find('MERRA300') >= 0:
                        inp_file = inp_file.replace('MERRA300', 'MERRA301')
                        if os.path.exists(inp_file):
                            break
            else:
                if not quiet:
                    self.log += 'No Solar file found for ' + inp_strt + '\n'
                return None
        return inp_file

    def getInfo(self, inp_file):
        if inp_file is None:
            self.log += '\nInput file missing!\n    '
            return
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
        cdf_file = Dataset(unzip_file, 'r')
        i = unzip_file.rfind('/')
        self.log += '\nFile:\n    '
        self.log += unzip_file[i + 1:] + '\n'
        self.log += ' Format:\n    '
        try:
            self.log += cdf_file.Format + '\n'
        except:
            try:
                self.log += cdf_file.file_format + '\n'
            except:
                self.log += '?\n'
        self.log += ' Dimensions:\n    '
        vals = ''
        keys = list(cdf_file.dimensions.keys())
        values = list(cdf_file.dimensions.values())
        if type(cdf_file.dimensions) is dict:
            for i in range(len(keys)):
                try:
                    vals += keys[i] + ': ' + str(values[i].size) + ', '
                except:
                    vals += keys[i] + ': ' + str(values[i]) + ', '
        else:
            for i in range(len(keys)):
                bits = str(values[i]).strip().split()
                vals += keys[i] + ': ' + bits[-1] + ', '
        self.log += vals[:-2] + '\n'
        times = [cdf_file.variables[self.vars['time']][0], cdf_file.variables[self.vars['time']][-1]]
        if times[0] > 0:
            self.log += ' Times:\n    '
            strt_time = datetime(1900, 1, 1, 0, 0)
            frst_hour = strt_time + timedelta(hours=int(times[0]))
            last_hour = strt_time + timedelta(hours=int(times[-1]))
            self.log += frst_hour.strftime('%Y-%m-%d %H:%M') + ' to ' + last_hour.strftime('%Y-%m-%d %H:%M') + '\n'
        self.log += ' Variables:\n    '
        vals = ''
        for key in iter(sorted(cdf_file.variables.keys())):
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
        zones = {}
        longi = longitude[:]
        for val in longi:
            vals += '{:0.4f}'.format(val) + ', '
            zone = int(round(val / 15))
            try:
                zones[zone] = zones[zone] + 1
            except:
                zones[zone] = 1
        self.log += vals[:-2] + '\n'
        if len(zones) > 1:
            self.log += ' Timezones:\n    '
            for key, value in zones.items():
                self.log += str(key) + ' (' + str(value) + ')  '
            self.log += '\n'
        cdf_file.close()
        return

    def process_era5(self):
        dys = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if self.show_progress:
            self.caller.daybar.setValue(0)
            self.caller.progresslabel.setText('Reading input data')
            QtCore.QCoreApplication.processEvents()
        date_1900 = datetime(1900, 1, 1, 0)
        year_strt = datetime(self.src_year, 1, 1)
        year_end = datetime(self.src_year + 1, 1, 1)
        frst_hour = year_strt - date_1900
        frst_hour = frst_hour.days * 24 - self.src_zone + 1 # weather files are 1 hour different
        last_hour = year_end - date_1900
        last_hour = last_hour.days * 24 - self.src_zone + 1
         # get wind and solar data
        if self.src_zone > 0: # go back to last year if needed
            inp_strt = '{:04d}'.format(self.src_year - 1) + '12'
            inp_file = self.findFile(inp_strt, quiet=True)
            if inp_file is None:
                inp_strt = '{:04d}'.format(self.src_year - 1)
                inp_file = self.findFile(inp_strt, quiet=True)
                if inp_file is None:
                    inp_strt = '{:04d}'.format(self.src_year - 1) + '1231'
                    inp_file = self.findFile(inp_strt, quiet=True)
            if inp_file is None:
                return
            self.get_era5_data(inp_file, frst_hour, last_hour)
            if self.return_code != 0:
                return
        inp_strt = '{:04d}'.format(self.src_year)
        inp_file = self.findFile(inp_strt, quiet=True)
        if inp_file is None: # monthly files
            if self.wrap:
                yrs = 1
            else:
                yrs = 2
            year = self.src_year
            for mt in range(12):
                self.logMsg('Processing month %s' % str(mt + 1), mt / 12.)
                inp_strt = '{:04d}{:02d}'.format(year, mt + 1)
                inp_file = self.findFile(inp_strt, quiet=True)
                if inp_file is None:
                    if yrs == 2:
                        return
                    yrs == 2
                    year -= 1
                    self.logMsg('Wrapping to prior year - %.4d-%.2d-%.2d' % (year, mt + 1, 1))
                    last_hour = datetime(year + 1, 1, 1)
                    last_hour = last_hour - date_1900
                    last_hour = int(last_hour.days) * 24 - self.src_zone + 1
                    inp_strt = '{:04d}{:02d}'.format(year, mt + 1)
                    inp_file = self.findFile(inp_strt, quiet=True)
                    if inp_file is None:
                        # try last full year
                        inp_strt = '{:04d}'.format(year)
                        inp_file = self.findFile(inp_strt, quiet=True)
                        if inp_file is None:
                            return
                        self.logMsg('Processing months %s to 12' % str(mt + 1))
                        frst_hour = datetime(year, mt + 1, 1)
                        frst_hour = frst_hour - date_1900
                        frst_hour = int(frst_hour.days) * 24
                        if mt < 2 and year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                            last_hour = datetime(year + 1, 2, 28)
                            last_hour += timedelta(hours=1)
                            last_hour = last_hour - date_1900
                            last_hour = int(last_hour.days) * 24
                            self.get_era5_data(inp_file, frst_hour, last_hour)
                            if self.return_code != 0:
                                return
                            frst_hour = datetime(year, 3, 1)
                            frst_hour = frst_hour - date_1900
                            frst_hour = int(frst_hour.days) * 24
                        self.get_era5_data(inp_file, frst_hour, last_hour)
                        if self.return_code != 0:
                            return
                        break
                fst_hour = datetime(year, mt + 1, 1)
                fst_hour = fst_hour - date_1900
                fst_hour = int(fst_hour.days) * 24
                if fst_hour < frst_hour:
                    fst_hour = frst_hour
                lst_hour = fst_hour + dys[mt] * 24
                if lst_hour > last_hour:
                    lst_hour = last_hour
                self.get_era5_data(inp_file, fst_hour, lst_hour)
                if self.return_code != 0:
                    return
        else: # year file
            if (last_hour - frst_hour) > 8760: # leap year?
                last_prt1 = frst_hour + 59 * 24
                self.get_era5_data(self.findFile(inp_strt), frst_hour, last_prt1)
                if self.return_code != 0:
                    return
                self.get_era5_data(self.findFile(inp_strt), last_prt1 + 24, last_hour)
                if self.return_code != 0:
                    return
            else:
                self.get_era5_data(self.findFile(inp_strt), frst_hour, last_hour)
                if self.return_code != 0:
                    return
        if self.src_zone < 0: # go forward to next year if needed
            inp_strt = '{:04d}'.format(self.src_year + 1) + '01'
            inp_file = self.findFile(inp_strt, quiet=True)
            if inp_file is None:
                inp_strt = '{:04d}'.format(self.src_year + 1)
                inp_file = self.findFile(inp_strt, quiet=True)
                if inp_file is None:
                    inp_strt = '{:04d}'.format(self.src_year + 1) + '0101'
                    inp_file = self.findFile(inp_strt, quiet=True)
            if inp_file is None:
                return
            self.get_era5_data(inp_file, frst_hour, last_hour)
            if self.return_code != 0:
                return
        self.longrange = [self.lons[0], self.lons[-1]]
        self.checkZone()
        if self.make_wind:
            if self.show_progress:
                self.caller.daybar.setValue(0)
                self.caller.progresslabel.setText('Creating wind weather files')
                QtCore.QCoreApplication.processEvents()
            target_dir = self.tgt_dir
            self.logMsg('Target directory is %s' % target_dir)
            if not os.path.exists(target_dir):
                self.log += 'mkdir %s\n' % target_dir
                os.makedirs(target_dir)
            if self.src_lat is not None:   # specific location(s)
                for i in range(len(self.src_lat)):
                    if self.show_progress:
                        self.caller.daybar.setValue(i)
                        QtCore.QCoreApplication.processEvents()
                    out_file = self.tgt_dir + 'wind_weather_' + str(self.src_lat[i]) + '_' + \
                               str(self.src_lon[i]) + '_' + str(self.src_year) + '.' + self.fmat
                    tf = open(out_file, 'w')
                    with_gaps = 0
                    missing = False
                    hdr = 'id,<city>,<state>,<country>,%s,%s,%s,0,1,8760\n' % (str(self.src_year),
                          round(self.src_lat[i], 4), round(self.src_lon[i], 4))
                    tf.write(hdr)
                    tf.write('Wind data derived from ERA5 reanalysis-era5-single-levels' + '\n')
                    tf.write('Temperature,Pressure,Direction,Speed,Direction,Speed\n')
                    tf.write('C,atm,degrees,m/s,degrees,m/s\n')
                    tf.write('2,0,10,10,100,100\n')
                    for hr in range(len(self.s100m)):
                        for lat2 in range(len(self.lati[self.lat_lon_ndx[hr]])):
                            if self.src_lat[i] <= self.lati[self.lat_lon_ndx[hr]][lat2]:
                                break
                        for lon2 in range(len(self.longi[self.lat_lon_ndx[hr]])):
                            if self.src_lon[i] <= self.longi[self.lat_lon_ndx[hr]][lon2]:
                                break
                        if self.longrange[0] is None:
                            self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        else:
                            if self.longi[self.lat_lon_ndx[hr]][lon2] < self.longrange[0]:
                                self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        if self.longrange[1] is None:
                            self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        else:
                            if self.longi[self.lat_lon_ndx[hr]][lon2] > self.longrange[1]:
                                self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        lat1 = lat2 - 1
                        lat_rat = (self.lati[self.lat_lon_ndx[hr]][lat2] - self.src_lat[i]) / \
                                  (self.lati[self.lat_lon_ndx[hr]][lat2] - self.lati[self.lat_lon_ndx[hr]][lat1])
                        lon1 = lon2 - 1
                        lon_rat = (self.longi[self.lat_lon_ndx[hr]][lon2] - self.src_lon[i]) / \
                                  (self.longi[self.lat_lon_ndx[hr]][lon2] - self.longi[self.lat_lon_ndx[hr]][lon1])
                        tf.write(str(self.valu(self.t_2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                                 str(self.valu(self.p_s[hr], lat1, lon1, lat_rat, lon_rat, rnd=6)) + ',' +
                                 str(self.valu(self.d10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                 str(self.valu(self.s10m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                                 str(self.valu(self.d100m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                 str(self.valu(self.s100m[hr], lat1, lon1, lat_rat, lon_rat)) + '\n')
                    tf.close()
                    self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
                    if self.hub_height > 0:
                        ok = extrapolateWind(out_file, self.hub_height, law=self.law, replace=True)
                        if ok:
                            self.log += '%s updated\n' % out_file[out_file.rfind('/') + 1:]
            else: # all locations
                if self.show_progress:
                    self.caller.daybar.setMaximum(len(self.lats) * len(self.lons))
                    QtCore.QCoreApplication.processEvents()
                for la in range(len(self.lats)):
                    for lo in range(len(self.lons)):
                        if self.show_progress:
                            self.caller.daybar.setValue(len(self.lats) * la + lo)
                            QtCore.QCoreApplication.processEvents()
                        with_gaps = 0
                        missing = False
                        out_file = self.tgt_dir + 'wind_weather_' + '{:0.4f}'.format(self.lats[la]) + \
                                   '_' + '{:0.4f}'.format(self.lons[lo]) + '_' + str(self.src_year) + '.srw'
                        tf = open(out_file, 'w')
                        hdr = 'id,<city>,<state>,<country>,%s,%s,%s,0,1,8760\n' % (str(self.src_year),
                              round(self.lats[la], 4), round(self.lons[lo], 4))
                        tf.write(hdr)
                        tf.write('Wind data derived from ERA5 reanalysis-era5-single-levels' + '\n')
                        tf.write('Temperature,Pressure,Direction,Speed,Direction,Speed\n')
                        tf.write('C,atm,degrees,m/s,degrees,m/s\n')
                        tf.write('2,0,10,10,100,100\n')
                        for hr in range(len(self.s100m)):
                            try:
                                lat = self.lati[self.lat_lon_ndx[hr]].index(self.lats[la])
                                lon = self.longi[self.lat_lon_ndx[hr]].index(self.lons[lo])
                            except:
                                if self.gaps:
                                    tf.write(',,,,,,,,\n')
                                    with_gaps += 1
                                    continue
                                else:
                                    missing = True
                                    break
                            try:
                                tf.write(str(self.t_2m[hr][lat][lon]) + ',' + str(self.p_s[hr][lat][lon]) + ',' +
                                         str(self.d10m[hr][lat][lon]) + ',' + str(self.s10m[hr][lat][lon]) + ',' +
                                         str(self.d100m[hr][lat][lon]) + ',' + str(self.s100m[hr][lat][lon]) + '\n')
                            except:
                                if self.gaps:
                                    tf.write(',,,,,,,,\n')
                                    with_gaps += 1
                                    continue
                                else:
                                    missing = True
                                    break
                        tf.close()
                        if with_gaps > 0 and with_gaps < 504:
                            self.log += '%s created with gaps (%s days)\n' % ( \
                                        out_file[out_file.rfind('/') + 1:], str(int(with_gaps / 24)))
                        elif missing or with_gaps > 0:
                            os.remove(out_file)
                            self.gaplog += '%s not created due to data gaps\n' % out_file[out_file.rfind('/') + 1:]
                            continue
                        else:
                            self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
                        if self.hub_height > 0:
                            ok = extrapolateWind(out_file, self.hub_height, law=self.law, replace=True)
                            if ok:
                                self.log += '%s updated\n' % out_file[out_file.rfind('/') + 1:]
            if self.show_progress:
                self.caller.daybar.setValue(self.caller.daybar.maximum())
                self.caller.progresslabel.setText('All done')
                QtCore.QCoreApplication.processEvents()
            return # that's it for wind
        #for solar we already have the data - same file (format)
        if self.show_progress:
            self.caller.daybar.setValue(0)
            self.caller.progresslabel.setText('Creating solar weather files')
            QtCore.QCoreApplication.processEvents()
        target_dir = self.tgt_dir
        self.logMsg('Target directory is %s' % target_dir)
        if not os.path.exists(target_dir):
            self.log += 'mkdir %s\n' % target_dir
            os.makedirs(target_dir)
        if self.src_lat is not None:  # specific location(s)
            if self.show_progress:
                self.caller.daybar.setMaximum(len(self.src_lat) - 1)
                QtCore.QCoreApplication.processEvents()
            for i in range(len(self.src_lat)):
                if self.show_progress:
                    self.caller.daybar.setValue(i)
                    QtCore.QCoreApplication.processEvents()
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
                    hour = 0
                    for hr in range(len(self.s10m)):
                        for lat2 in range(len(self.lati[self.lat_lon_ndx[hr]])):
                            if self.src_lat[i] <= self.lati[self.lat_lon_ndx[hr]][lat2]:
                                break
                        for lon2 in range(len(self.longi[self.lat_lon_ndx[hr]])):
                            if self.src_lon[i] <= self.longi[self.lat_lon_ndx[hr]][lon2]:
                                break
                        if self.longrange[0] is None:
                            self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        else:
                            if self.longi[self.lat_lon_ndx[hr]][lon2] < self.longrange[0]:
                                self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        if self.longrange[1] is None:
                            self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        else:
                            if self.longi[self.lat_lon_ndx[hr]][lon2] > self.longrange[1]:
                                self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        lat1 = lat2 - 1
                        lat_rat = (self.lati[self.lat_lon_ndx[hr]][lat2] - self.src_lat[i]) / \
                                  (self.lati[self.lat_lon_ndx[hr]][lat2] - self.lati[self.lat_lon_ndx[hr]][lat1])
                        lon1 = lon2 - 1
                        lon_rat = (self.longi[self.lat_lon_ndx[hr]][lon2] - self.src_lon[i]) / \
                                  (self.longi[self.lat_lon_ndx[hr]][lon2] - self.longi[self.lat_lon_ndx[hr]][lon1])
                        ghi = self.valu(self.ghi[hr], lat1, lon1, lat_rat, lon_rat)
                        dni = getDNI(ghi, hour=hr + 1, lat=self.src_lat[i], lon=self.src_lon[i],
                              press=self.valu(self.p_s[hr], lat1,
                              lon1, lat_rat, lon_rat, rnd=0), zone=self.src_zone)
                        dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.src_lat[i])
                        tf.write(str(self.src_year) + ',' + str(mth + 1) + ',' +
                        str(day) + ',' + str(hour) + ',' +
                        str(int(ghi)) + ',' + str(int(dni)) + ',' + str(int(dhi)) + ',' +
                        str(self.valu(self.t_2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                        str(self.valu(self.p_s[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                        str(self.valu(self.s10m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                        str(self.valu(self.d10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + '\n')
                        hour += 1
                        if hour > 23:
                            hour = 0
                            day += 1
                            if day > dys[mth]:
                                mth += 1
                                day = 1
                                hour = 0
                else:
                    hdr = 'id,<city>,<state>,%s,%s,%s,0,3600.0,%s,0:30:00\n' % (str(self.src_zone),
                          round(self.src_lat[i], 4),
                          round(self.src_lon[i], 4), str(self.src_year))
                    tf.write(hdr)
                    for hr in range(len(self.s10m)):
                        for lat2 in range(len(self.lati[self.lat_lon_ndx[hr]])):
                            if self.src_lat[i] <= self.lati[self.lat_lon_ndx[hr]][lat2]:
                                break
                        for lon2 in range(len(self.longi[self.lat_lon_ndx[hr]])):
                            if self.src_lon[i] <= self.longi[self.lat_lon_ndx[hr]][lon2]:
                                break
                        if self.longrange[0] is None:
                            self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        else:
                            if self.longi[self.lat_lon_ndx[hr]][lon2] < self.longrange[0]:
                                self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        if self.longrange[1] is None:
                            self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        else:
                            if self.longi[self.lat_lon_ndx[hr]][lon2] > self.longrange[1]:
                                self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        lat1 = lat2 - 1
                        lat_rat = (self.lati[self.lat_lon_ndx[hr]][lat2] - self.src_lat[i]) / \
                                  (self.lati[self.lat_lon_ndx[hr]][lat2] - self.lati[self.lat_lon_ndx[hr]][lat1])
                        lon1 = lon2 - 1
                        lon_rat = (self.longi[self.lat_lon_ndx[hr]][lon2] - self.src_lon[i]) / \
                                  (self.longi[self.lat_lon_ndx[hr]][lon2] - self.longi[self.lat_lon_ndx[hr]][lon1])
                        ghi = self.valu(self.ghi[hr], lat1, lon1, lat_rat, lon_rat)
                        dni = getDNI(ghi, hour=hr + 1, lat=self.src_lat[i], lon=self.src_lon[i],
                              press=self.valu(self.p_s[hr], lat1, lon1, lat_rat,
                              lon_rat, rnd=0), zone=self.src_zone)
                        dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.src_lat[i])
                        if len(self.alb) > 0:
                            alb = str(self.valu(self.alb[hr], lat1, lon1, lat_rat, lon_rat))
                        else:
                            alb = '-999'
                        tf.write(str(self.valu(self.t_2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) +
                        ',-999,-999,-999,' +
                        str(self.valu(self.s10m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                        str(self.valu(self.d10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                        str(self.valu(self.p_s[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                        str(int(ghi)) + ',' + str(int(dni)) + ',' + str(int(dhi)) + ',' + alb + ',-999,\n')
                tf.close()
                self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
        else: # all locations
            if self.show_progress:
                self.caller.daybar.setMaximum(len(self.lats) * len(self.lons))
                QtCore.QCoreApplication.processEvents()
            for la in range(len(self.lats)):
                for lo in range(len(self.lons)):
                    if self.show_progress:
                        self.caller.daybar.setValue(len(self.lats) * la + lo)
                        QtCore.QCoreApplication.processEvents()
                    with_gaps = 0
                    missing = False
                    out_file = self.tgt_dir + 'solar_weather_' + '{:0.4f}'.format(self.lats[la]) + \
                            '_' + '{:0.4f}'.format(self.lons[lo]) + '_' + str(self.src_year) + '.' + self.fmat
                    tf = open(out_file, 'w')
                    if self.fmat == 'csv':
                        hdr = 'Location,City,Region,Country,Latitude,Longitude,Time Zone,Elevation,Source\n' + \
                              'id,<city>,<state>,<country>,%s,%s,%s,0,IWEC\n' % (round(self.lats[la], 4),
                              round(self.lons[lo], 4), str(self.src_zone))
                        tf.write(hdr)
                        tf.write('Year,Month,Day,Hour,GHI,DNI,DHI,Tdry,Pres,Wspd,Wdir' + '\n')
                        mth = 0
                        day = 1
                        hour = 0
                        for hr in range(len(self.s10m)):
                            try:
                                lat = self.lati[self.lat_lon_ndx[hr]].index(self.lats[la])
                                lon = self.longi[self.lat_lon_ndx[hr]].index(self.lons[lo])
                            except:
                                if self.gaps:
                                    tf.write(',,,,,,,,\n')
                                    with_gaps += 1
                                    continue
                                else:
                                    missing = True
                                    break
                            ghi = self.ghi[hr][lat][lon]
                            dni = getDNI(ghi, hour=hr + 1, lat=self.lati[self.lat_lon_ndx[hr]][lat],
                                  lon=self.longi[self.lat_lon_ndx[hr]][lon],
                                  press=self.p_s[hr][lat][lon], zone=self.src_zone)
                            dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.lati[self.lat_lon_ndx[hr]][lat])
                            tf.write(str(self.src_year) + ',' + '{:02d}'.format(mth + 1) + ',' +
                                '{:02d}'.format(day) + ',' + '{:02d}'.format(hour) + ',' +
                                '{:0.1f}'.format(ghi) + ',' + '{:0.1f}'.format(dni) + ',' +
                                '{:0.1f}'.format(dhi) + ',' +
                                str(self.t_2m[hr][lat][lon]) + ',' +
                                str(self.p_s[hr][lat][lon]) + ',' +
                                str(self.s10m[hr][lat][lon]) + ',' +
                                str(self.d10m[hr][lat][lon]) + '\n')
                            hour += 1
                            if hour > 23:
                                hour = 0
                                day += 1
                                if day > dys[mth]:
                                    mth += 1
                                    day = 1
                                    hour = 0
                    else:
                        hdr = 'id,<city>,<state>,%s,%s,%s,0,3600.0,%s,0:30:00\n' % (str(self.src_zone),
                              round(self.lats[la], 4),
                              round(self.lons[lo], 4), str(self.src_year))
                        tf.write(hdr)
                        for hr in range(len(self.s10m)):
                            try:
                                lat = self.lati[self.lat_lon_ndx[hr]].index(self.lats[la])
                                lon = self.longi[self.lat_lon_ndx[hr]].index(self.lons[lo])
                                la2 = self.lati[self.lat_lon_ndx[hr]].index(self.lats[la])
                                lo2 = self.longi[self.lat_lon_ndx[hr]].index(self.lons[lo])
                            except:
                                if self.gaps:
                                    tf.write(',,,,,,,,\n')
                                    with_gaps += 1
                                    continue
                                else:
                                    missing = True
                                    break
                            ghi = self.ghi[hr][lat][lon]
                            dni = getDNI(ghi, hour=hr + 1, lat=self.lati[self.lat_lon_ndx[hr]][lat],
                                      lon=self.longi[self.lat_lon_ndx[hr]][lon],
                                      press=self.p_s[hr][lat][lon], zone=self.src_zone)
                            dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.lati[self.lat_lon_ndx[hr]][lat])
                            if len(self.alb) > 0:
                                try: # bug here but let's fix later
                                    alb = str(self.valu(self.alb[hr], lat1, lon1, lat_rat, lon_rat))
                                except:
                                    alb = '-999'
                            else:
                                alb = '-999'
                            tf.write(str(self.t_2m[hr][lat][lon]) +
                                ',-999,-999,-999,' +
                                str(self.s10m[hr][lat][lon]) + ',' +
                                str(self.d10m[hr][lat][lon]) + ',' +
                                str(self.p_s[hr][lat][lon]) + ',' +
                                str(int(ghi)) + ',' + str(int(dni)) + ',' + str(int(dhi)) + ',' + alb + ',-999,\n')
                    tf.close()
                    if with_gaps > 0 and with_gaps < 504:
                        self.log += '%s created with gaps (%s days)\n' % ( \
                                    out_file[out_file.rfind('/') + 1:], str(int(with_gaps / 24)))
                    elif missing or with_gaps > 0:
                  #      os.remove(out_file)
                        self.gaplog += '%s not created due to data gaps\n' % out_file[out_file.rfind('/') + 1:]
                        continue
                    else:
                        self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
        if self.show_progress:
            self.caller.daybar.setValue(self.caller.daybar.maximum())
            self.caller.progresslabel.setText('All done')
            QtCore.QCoreApplication.processEvents()
        return

    def __init__(self, caller, src_year, src_zone, src_dir_s, src_dir_w, tgt_dir, fmat, swg='swgdn',
                 wrap=None, gaps=None, src_lat_lon=None, info=False, hub_height=0, law='l'):
      #  self.last_time = datetime.now()
        if caller is None:
            self.show_progress = False
        else:
            self.show_progress = True
            self.caller = caller
        self.log = ''
        self.gaplog = ''
        self.return_code = 0
        self.src_year = int(src_year)
        self.the_year = self.src_year  # start with their year
        self.src_zone = src_zone
        self.src_dir_s = src_dir_s
        self.src_dir_w = src_dir_w
        self.yearly = [False, False]
        if self.src_dir_s[-5:] == '/' + str(src_year):
            self.yearly[0] = True
        if self.src_dir_w[-5:] == '/' + str(src_year):
            self.yearly[1] = True
        self.tgt_dir = tgt_dir
        self.fmat = fmat
        self.swg = swg
        self.hub_height = hub_height
        self.law = law
        self.wrap = False
        if wrap is None or wrap == '':
            pass
        elif wrap[0].lower() == 'y' or wrap[0].lower() == 't' or wrap[:2].lower() == 'on':
            self.wrap = True
        self.gaps = False
        if gaps is None or gaps == '':
            pass
        elif gaps[0].lower() == 'y' or gaps[0].lower() == 't' or gaps[:2].lower() == 'on':
            self.gaps = True
        self.src_lat_lon = src_lat_lon
        self.src_s_pfx = []
        self.src_s_sfx = []
        self.dataset = 'MERRA-2'
        self.era5 = False
        merra300 = False
        self.vars = {'latitude': 'lat', 'longitude': 'lon', 'ps': 'PS', 'swgdn': 'SWGDN', 'swgnt': 'SWGNT',
                     'time': 'time', 't2m': 'T2M', 't10m': 'T10M', 't50m': 'T50M', 'u2m': 'U2M',
                     'u10m': 'U10M', 'u50m': 'U50M', 'v2m': 'V2M', 'v10m': 'V10M', 'v50m': 'V50M',
                     'alb': 'albedo'}
        if self.src_dir_s != '':
            self.src_dir_s += '/'
            fils = os.listdir(self.src_dir_s)
            for fil in fils:
                if fil.find('MERRA') >= 0:
                    j = fil.find('.tavg1_2d_rad_Nx.')
                    if j > 0:
                        if fil[:j + 17] not in self.src_s_pfx:
                            self.src_s_pfx.append(fil[:j + 17])
                            if self.src_s_pfx[-1].find('MERRA3') >= 0:
                                merra300 = True
                                self.src_s_pfx[-1] = self.src_s_pfx[-1].replace('MERRA301', 'MERRA300')
                            self.src_s_sfx.append(fil[j + 17 + 8:])
                     #   break
            del fils
        if merra300:
            for key in list(self.vars.keys()):
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
            for key in list(self.vars.keys()):
                self.vars[key] = self.vars[key].lower()
            self.vars['latitude'] = 'latitude'
            self.vars['longitude'] = 'longitude'
        if len(self.src_s_pfx) == 0 and len(self.src_w_pfx) == 0:
            self.era5 = True
            self.dataset = 'ERA5'
            self.vars = {'latitude': 'latitude', 'longitude': 'longitude', 'sp': 'sp', 'swgdn': 'ssrd',
                         'swgnt': 'ssr', 'time': 'time', 't2m': 't2m', 'u10': 'u10', 'u100': 'u100',
                         'v10': 'v10', 'v100': 'v100', 'alb': 'alnip', 'alb2': 'aluvp'}
            ini_file = getModelFile('getfiles.ini')
            config = configparser.RawConfigParser()
            config.read(ini_file)
            era_file = config.get('getera5', 'filename').split('$year$')
            self.src_s_pfx = [era_file[0]]
            self.src_s_sfx = [era_file[1]]
            self.src_w_pfx = self.src_s_pfx[:]
            self.src_w_sfx = self.src_s_sfx[:]
        if self.tgt_dir != '':
            self.tgt_dir += '/'
        if info:
            if self.era5:
                inp_strt = '{:04d}'.format(self.src_year)
                self.log += '\nERA5 file for: ' + inp_strt + '\n'
                inp_file = self.findFile(inp_strt, quiet=True)
                if inp_file is None:
                    inp_file = self.findFile(inp_strt + '01', quiet=True)
                # get variables from ERA5 file
                self.getInfo(inp_file)
            else:
                inp_strt = '{:04d}'.format(self.src_year) + '0101'
                self.log += '\nSolar file for: ' + inp_strt + '\n'
                # get variables from "solar" file
                self.getInfo(self.findFile(inp_strt, False))
                self.log += '\nWind file for: ' + inp_strt + '\n'
                # get variables from "wind" file
                self.getInfo(self.findFile(inp_strt, True))
            return
        if str(self.src_zone).lower() in ['auto', 'best']:
           # self.auto_zone = True
            if self.era5:
                inp_strt = '{:04d}'.format(self.src_year)
                inp_file = self.findFile(inp_strt, True, quiet=True)
                if inp_file is None:
                    inp_strt += '01'
                    inp_file = self.findFile(inp_strt, True, quiet=True)
                    if inp_file is None:
                        inp_strt += '01'
                        inp_file = self.findFile(inp_strt, True, quiet=True)
            else:
                inp_strt = '{:04d}'.format(self.src_year) + '0101'
                inp_file = self.findFile(inp_strt, True, quiet=True)
             # get longitude from "wind" file
            unzip_file = self.unZip(self.findFile(inp_strt, True))
            if self.return_code != 0:
                return
            cdf_file = Dataset(unzip_file, 'r')
            longitude = cdf_file.variables[self.vars['longitude']][:]
            self.src_zone = int(round(longitude[0] / 15))
            if str(self.src_zone).lower() == 'best':
                if self.src_zone != int(round(longitude[-1] / 15)):
                    zones = {}
                    longi = longitude[:]
                    for val in longi:
                        zone = int(round(val / 15))
                        try:
                            zones[zone] = zones[zone] + 1
                        except:
                            zones[zone] = 1
                    self.src_zone = max(zones, key=lambda k: zones[k])
            self.log += 'Time zone: %s based on %s (west) longitude (%s to %s)\n' % (str(self.src_zone),
                        self.dataset, '{:0.4f}'.format(longitude[0]), '{:0.4f}'.format(longitude[-1]))
            cdf_file.close()
        else:
            self.src_zone = int(self.src_zone)
          #  self.auto_zone = False
        if self.src_lat_lon != '':
            self.src_lat = []
            self.src_lon = []
            latlon = self.src_lat_lon.replace('_',',')
            latlon = latlon.replace(' ','').split(',')
            try:
                for j in range(0, len(latlon), 2):
                    self.src_lat.append(float(latlon[j]))
                    self.src_lon.append(float(latlon[j + 1]))
            except:
                self.log += 'Error with Coordinates field'
                self.return_code = 2
                return
            self.logMsg('Processing %s, %s' % (self.src_lat, self.src_lon))
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
        self.lats = []
        self.lati = []
        self.latsi = []
        self.lons = []
        self.longi = []
        self.longsi = []
        self.lat_lon_ndx = []
        self.longrange = [None, None]
        self.tims = []
        self.s10m = []
        self.d10m = []
        self.t_10m = []
        self.p_s = []
        self.ghi = []
        self.s2m = []
        self.s50m = []
        self.s100m = []
        self.d2m = []
        self.d50m = []
        self.d100m = []
        self.t_2m = []
        self.alb = []
        dys = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if self.era5:
            self.process_era5()
            return
        # MERRA_2 data
        if self.src_zone > 0:
            inp_strt = '{:04d}'.format(self.src_year - 1) + '1231'
        else:
            inp_strt = '{:04d}'.format(self.src_year) + '0101'
        # get variables from "wind" files
        self.get_data(self.findFile(inp_strt, True))   # get wind data
        if self.return_code != 0:
            return
        if self.src_zone != 0:
            # need last self.src_zone hours of either last day of previous year
            # or first day of this year
            self.p_s = self.p_s[-self.src_zone:]
            self.lat_lon_ndx = self.lat_lon_ndx[-self.src_zone:]
            self.lati = self.lati[-self.src_zone:]
            self.longi = self.longi[-self.src_zone:]
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
            if len(self.alb) > 0:
                self.alb = self.alb[-self.src_zone:]
        if self.wrap:
            yrs = 2
        else:
            yrs = 1
        for mt in range(len(dys)):
            if self.show_progress:
                if self.make_wind:
                    progress = mt / 12.
                else:
                    progress = mt * 2 / 36.
                self.caller.daybar.setMaximum(dys[mt])
            else:
                progress = False
            self.logMsg('Processing month %s wind' % str(mt + 1), progress)
            for dy in range(1, dys[mt] + 1):
                if self.src_zone <= 0:
                    if mt == 0 and dy == 1:
                        continue
                if self.show_progress:
                    self.caller.daybar.setValue(dy)
                    QtCore.QCoreApplication.processEvents()
                for yr in range(yrs):
                    inp_strt = '{:04d}'.format(self.the_year) + '{:02d}'.format(mt + 1) + \
                               '{:02d}'.format(dy)
                    inp_file = self.findFile(inp_strt, True)
                    if inp_file is None:
                        if self.wrap and self.the_year == self.src_year: # check if need to go back a year
                            self.the_year = self.src_year - 1
                            self.logMsg('Wrapping to prior year - %.4d-%.2d-%.2d' % (self.the_year, (mt + 1), dy))
                            yrs = 1
                self.get_data(inp_file)   # get wind data
                if self.return_code != 0:
                    return
        if self.src_zone < 0:
            inp_strt = '{:04d}'.format(self.the_year + 1) + '0101'
            inp_file = self.findFile(inp_strt, True)
            if inp_file is None:
                self.logMsg('Wrapping to prior year - %.4d-%.2d-%.2d' % (self.the_year, 1, 1))
                inp_strt = '{:04d}'.format(self.the_year) + '0101'
                inp_file = self.findFile(inp_strt, True)
            self.get_data(inp_file)
            if self.return_code != 0:
                return
        if len(self.p_s) > 8760: # will be the case if src_zone != 0
            del self.p_s[len(self.p_s) - (len(self.p_s) - 8760):]
            del self.lat_lon_ndx[len(self.lat_lon_ndx) - (len(self.lat_lon_ndx) - 8760):]
            del self.lati[len(self.lati) - (len(self.lati) - 8760):]
            del self.longi[len(self.longi) - (len(self.longi) - 8760):]
            if len(self.s10m) > 0:
                del self.s10m[len(self.s10m) - (len(self.s10m) - 8760):]
                del self.d10m[len(self.d10m) - (len(self.d10m) - 8760):]
                try:
                    del self.t_10m[len(self.t_10m) - (len(self.t_10m) - 8760):]
                except:
                    pass
            if self.make_wind:
                try:
                    del self.s2m[len(self.s2m) - (len(self.s2m) - 8760):]
                    del self.s50m[len(self.s50m) - (len(self.s50m) - 8760):]
                    del self.d2m[len(self.d2m) - (len(self.d2m) - 8760):]
                    del self.d50m[len(self.d50m) - (len(self.d50m) - 8760):]
                except:
                    pass
                del self.t_2m[len(self.t_2m) - (len(self.t_2m) - 8760):]
        self.longrange = [self.lons[0], self.lons[-1]]
        self.checkZone()
        if self.make_wind:
            if self.show_progress:
                self.caller.daybar.setValue(0)
                self.caller.progresslabel.setText('Creating wind weather files')
                QtCore.QCoreApplication.processEvents()
            target_dir = self.tgt_dir
            self.logMsg('Target directory is %s' % target_dir)
            if not os.path.exists(target_dir):
                self.log += 'mkdir %s\n' % target_dir
                os.makedirs(target_dir)
            if self.src_lat is not None:   # specific location(s)
                for i in range(len(self.src_lat)):
                    if self.show_progress:
                        self.caller.daybar.setValue(i)
                        QtCore.QCoreApplication.processEvents()
                    out_file = self.tgt_dir + 'wind_weather_' + str(self.src_lat[i]) + '_' + \
                               str(self.src_lon[i]) + '_' + str(self.src_year) + '.' + self.fmat
                    tf = open(out_file, 'w')
                    hdr = 'id,<city>,<state>,<country>,%s,%s,%s,0,1,8760\n' % (str(self.src_year),
                          round(self.src_lat[i], 4), round(self.src_lon[i], 4))
                    tf.write(hdr)
                    tf.write('Wind data derived from MERRA-2 tavg1_2d_slv_Nx' + '\n')
                    if len(self.s10m) > 0:
                        tf.write('Temperature,Pressure,Direction,Speed,Temperature,Direction,Speed,' +
                                 'Direction,Speed' + '\n')
                        tf.write('C,atm,degrees,m/s,C,degrees,m/s,degrees,m/s' + '\n')
                        tf.write('2,0,2,2,10,10,10,50,50' + '\n')
                        for hr in range(len(self.s50m)):
                            for lat2 in range(len(self.lati[self.lat_lon_ndx[hr]])):
                                if self.src_lat[i] <= self.lati[self.lat_lon_ndx[hr]][lat2]:
                                    break
                            for lon2 in range(len(self.longi[self.lat_lon_ndx[hr]])):
                                if self.src_lon[i] <= self.longi[self.lat_lon_ndx[hr]][lon2]:
                                    break
                            if self.longrange[0] is None:
                                self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                            else:
                                if self.longi[self.lat_lon_ndx[hr]][lon2] < self.longrange[0]:
                                    self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                            if self.longrange[1] is None:
                                self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                            else:
                                if self.longi[self.lat_lon_ndx[hr]][lon2] > self.longrange[1]:
                                    self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                            lat1 = lat2 - 1
                            lat_rat = (self.lati[self.lat_lon_ndx[hr]][lat2] - self.src_lat[i]) / \
                                      (self.lati[self.lat_lon_ndx[hr]][lat2] - self.lati[self.lat_lon_ndx[hr]][lat1])
                            lon1 = lon2 - 1
                            lon_rat = (self.longi[self.lat_lon_ndx[hr]][lon2] - self.src_lon[i]) / \
                                      (self.longi[self.lat_lon_ndx[hr]][lon2] - self.longi[self.lat_lon_ndx[hr]][lon1])
                            tf.write(str(self.valu(self.t_2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                                str(self.valu(self.p_s[hr], lat1, lon1, lat_rat, lon_rat, rnd=6)) + ',' +
                                str(self.valu(self.d2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                str(self.valu(self.s2m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                                str(self.valu(self.t_10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                                str(self.valu(self.s10m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                                str(self.valu(self.d50m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                str(self.valu(self.s50m[hr], lat1, lon1, lat_rat, lon_rat)) + '\n')
                    else:
                        tf.write('Temperature,Pressure,Direction,Speed,Direction,Speed' + '\n')
                        tf.write('C,atm,degrees,m/s,degrees,m/s' + '\n')
                        tf.write('2,0,2,2,50,50' + '\n')
                        for hr in range(len(self.s50m)):
                            for lat2 in range(len(self.lati[self.lat_lon_ndx[hr]])):
                                if self.src_lat[i] <= self.lati[self.lat_lon_ndx[hr]][lat2]:
                                    break
                            for lon2 in range(len(self.longi[self.lat_lon_ndx[hr]])):
                                if self.src_lon[i] <= self.longi[self.lat_lon_ndx[hr]][lon2]:
                                    break
                            if self.longrange[0] is None:
                                self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                            else:
                                if self.longi[self.lat_lon_ndx[hr]][lon2] < self.longrange[0]:
                                    self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                            if self.longrange[1] is None:
                                self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                            else:
                                if self.longi[self.lat_lon_ndx[hr]][lon2] > self.longrange[1]:
                                    self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                            lat1 = lat2 - 1
                            lat_rat = (self.lati[self.lat_lon_ndx[hr]][lat2] - self.src_lat[i]) / (self.lati[self.lat_lon_ndx[hr]][lat2] - self.lati[self.lat_lon_ndx[hr]][lat1])
                            lon1 = lon2 - 1
                            lon_rat = (self.longi[self.lat_lon_ndx[hr]][lon2] - self.src_lon[i]) / (self.longi[self.lat_lon_ndx[hr]][lon2] - self.longi[self.lat_lon_ndx[hr]][lon1])
                            tf.write(str(self.valu(self.t_2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                                str(self.valu(self.p_s[hr], lat1, lon1, lat_rat, lon_rat, rnd=6)) + ',' +
                                str(self.valu(self.d2m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                str(self.valu(self.s2m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                                str(self.valu(self.d50m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                                str(self.valu(self.s50m[hr], lat1, lon1, lat_rat, lon_rat)) + '\n')
                    tf.close()
                    self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
                    if self.hub_height > 0:
                        ok = extrapolateWind(out_file, self.hub_height, law=self.law, replace=True)
                        if ok:
                            self.log += '%s updated\n' % out_file[out_file.rfind('/') + 1:]
            else: # all locations
                if self.show_progress:
                    self.caller.daybar.setMaximum(len(self.lats) * len(self.lons))
                    QtCore.QCoreApplication.processEvents()
                for la in range(len(self.lats)):
                    for lo in range(len(self.lons)):
                        if self.show_progress:
                            self.caller.daybar.setValue(len(self.lats) * la + lo)
                            QtCore.QCoreApplication.processEvents()
                        with_gaps = 0
                        missing = False
                        out_file = self.tgt_dir + 'wind_weather_' + '{:0.4f}'.format(self.lats[la]) + \
                                   '_' + '{:0.4f}'.format(self.lons[lo]) + '_' + str(self.src_year) + '.srw'
                        tf = open(out_file, 'w')
                        hdr = 'id,<city>,<state>,<country>,%s,%s,%s,0,1,8760\n' % (str(self.src_year),
                              round(self.lats[la], 4), round(self.lons[lo], 4))
                        tf.write(hdr)
                        if self.era5:
                            tf.write('Wind data derived from ERA5 reanalysis-era5-single-levels' + '\n')
                        else:
                            tf.write('Wind data derived from MERRA-2 tavg1_2d_slv_Nx' + '\n')
                        if len(self.s10m) > 0:
                            tf.write('Temperature,Pressure,Direction,Speed,Temperature,Direction,' +
                                     'Speed,Direction,Speed' + '\n')
                            tf.write('C,atm,degrees,m/s,C,degrees,m/s,degrees,m/s' + '\n')
                            tf.write('2,0,2,2,10,10,10,50,50' + '\n')
                            for hr in range(len(self.s50m)):
                                try:
                                    lat = self.lati[self.lat_lon_ndx[hr]].index(self.lats[la])
                                    lon = self.longi[self.lat_lon_ndx[hr]].index(self.lons[lo])
                                except:
                                    if self.gaps:
                                        tf.write(',,,,,,,,\n')
                                        with_gaps += 1
                                        continue
                                    else:
                                        missing = True
                                        break
                                try:
                                    tf.write(str(self.t_2m[hr][lat][lon]) + ',' + str(self.p_s[hr][lat][lon]) + ',' +
                                        str(self.d2m[hr][lat][lon]) + ',' + str(self.s2m[hr][lat][lon]) + ',' +
                                        str(self.t_10m[hr][lat][lon]) + ',' + str(self.d10m[hr][lat][lon]) + ',' +
                                        str(self.s10m[hr][lat][lon]) + ',' +
                                        str(self.d50m[hr][lat][lon]) + ',' + str(self.s50m[hr][lat][lon]) + '\n')
                                except:
                                    if self.gaps:
                                        tf.write(',,,,,,,,\n')
                                        with_gaps += 1
                                        continue
                                    else:
                                        missing = True
                                        break
                        else:
                            tf.write('Temperature,Pressure,Direction,Speed,Direction,Speed' + '\n')
                            tf.write('C,atm,degrees,m/s,degrees,m/s' + '\n')
                            tf.write('2,0,2,2,50,50' + '\n')
                            for hr in range(len(self.s50m)):
                                try:
                                    lat = self.lati[self.lat_lon_ndx[hr]].index(self.lats[la])
                                    lon = self.longi[self.lat_lon_ndx[hr]].index(self.lons[lo])
                                except:
                                    if self.gaps:
                                        continue
                                    else:
                                        missing = True
                                        break
                                try:
                                    tf.write(str(self.t_2m[hr][lat][lon]) + ',' + str(self.p_s[hr][lat][lon]) + ',' +
                                        str(self.d2m[hr][lat][lon]) + ',' + str(self.s2m[hr][lat][lon]) + ',' +
                                        str(self.d50m[hr][lat][lon]) + ',' + str(self.s50m[hr][lat][lon]) + '\n')
                                except:
                                    if self.gaps:
                                        tf.write(',,,,,,,,\n')
                                        with_gaps += 1
                                        continue
                                    else:
                                        missing = True
                                        break
                        tf.close()
                        if with_gaps > 0 and with_gaps < 504:
                            self.log += '%s created with gaps (%s days)\n' % ( \
                                        out_file[out_file.rfind('/') + 1:], str(int(with_gaps / 24)))
                        elif missing or with_gaps > 0:
                            os.remove(out_file)
                            self.gaplog += '%s not created due to data gaps\n' % out_file[out_file.rfind('/') + 1:]
                            continue
                        else:
                            self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
                        if self.hub_height > 0:
                            ok = extrapolateWind(out_file, self.hub_height, law=self.law, replace=True)
                            if ok:
                                self.log += '%s updated\n' % out_file[out_file.rfind('/') + 1:]
            return  # that's it for wind
        # get variable from solar files
        if self.src_zone > 0:
            inp_strt = '{:04d}'.format(self.src_year - 1) + '1231'
        elif self.src_zone <= 0:
            inp_strt = '{:04d}'.format(self.src_year) + '0101'
        self.get_rad_data(self.findFile(inp_strt, False))  # get solar data
        if self.return_code != 0:
            return
        if self.src_zone != 0:
            # need last self.src_zone hours of either last day of previous year
            # or first day of this year
            self.latsi = self.latsi[-self.src_zone:]
            self.longsi = self.longsi[-self.src_zone:]
            self.ghi = self.ghi[-self.src_zone:]
        self.the_year = self.src_year  # start with their year
        if self.wrap:
            yrs = 2
        else:
            yrs = 1
        for mt in range(len(dys)):
            if self.show_progress:
                progress = (mt + 24) / 36.
                self.caller.daybar.setMaximum(dys[mt])
            else:
                progress= False
            self.logMsg('Processing month %s solar' % str(mt + 1), progress)
            for dy in range(1, dys[mt] + 1):
                if self.src_zone <= 0:
                    if mt == 0 and dy == 1:
                        continue
                if self.show_progress:
                    self.caller.daybar.setValue(dy)
                    QtCore.QCoreApplication.processEvents()
                found = False
                for yr in range(yrs):
                    inp_strt = '{:04d}'.format(self.the_year) + '{:02d}'.format(mt + 1) + \
                               '{:02d}'.format(dy)
                    inp_file = self.findFile(inp_strt, False)
                    if inp_file is None:
                        if self.wrap and self.the_year == self.src_year:
                            self.the_year = self.src_year - 1
                            self.logMsg('Wrapping to prior year - %.4d-%.2d-%.2d' % (self.the_year, (mt + 1), dy))
                            yrs = 1
                    if found:
                        break
                self.get_rad_data(inp_file)  # get solar data
                if self.return_code != 0:
                    return
        if self.src_zone < 0:
            inp_strt = '{:04d}'.format(self.the_year + 1) + '0101'
            inp_file = self.findFile(inp_strt, False)
            if inp_file is None:
                self.logMsg('Wrapping to prior year - %.4d-%.2d-%.2d' % (self.the_year, 1, 1))
                inp_strt = '{:04d}'.format(self.the_year) + '0101'
                inp_file = self.findFile(inp_strt, False)
            self.get_rad_data(inp_file)
            if self.return_code != 0:
                return
        if len(self.ghi) > 8760: # will be the case if src_zone != 0
            del self.ghi[len(self.ghi) - (len(self.ghi) - 8760):]
            del self.latsi[len(self.latsi) - (len(self.latsi) - 8760):]
            del self.longsi[len(self.longsi) - (len(self.longsi) - 8760):]
        if self.show_progress:
            self.caller.daybar.setValue(0)
            self.caller.progresslabel.setText('Creating solar weather files')
            QtCore.QCoreApplication.processEvents()
        target_dir = self.tgt_dir
        self.logMsg('Target directory is %s' % target_dir)
        if not os.path.exists(target_dir):
            self.log += 'mkdir %s\n' % target_dir
            os.makedirs(target_dir)
        if self.src_lat is not None:  # specific location(s)
            if self.show_progress:
                self.caller.daybar.setMaximum(len(self.src_lat) - 1)
                QtCore.QCoreApplication.processEvents()
            for i in range(len(self.src_lat)):
                if self.show_progress:
                    self.caller.daybar.setValue(i)
                    QtCore.QCoreApplication.processEvents()
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
                    hour = 0
                    for hr in range(len(self.s10m)):
                        for lat2 in range(len(self.lati[self.lat_lon_ndx[hr]])):
                            if self.src_lat[i] <= self.lati[self.lat_lon_ndx[hr]][lat2]:
                                break
                        for lon2 in range(len(self.longi[self.lat_lon_ndx[hr]])):
                            if self.src_lon[i] <= self.longi[self.lat_lon_ndx[hr]][lon2]:
                                break
                        if self.longrange[0] is None:
                            self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        else:
                            if self.longi[self.lat_lon_ndx[hr]][lon2] < self.longrange[0]:
                                self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        if self.longrange[1] is None:
                            self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        else:
                            if self.longi[self.lat_lon_ndx[hr]][lon2] > self.longrange[1]:
                                self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        if len(self.alb) > 0:
                            try: # bug here but let's fix later
                                    alb = self.valu(self.alb[hr], lat1, lon1, lat_rat, lon_rat)
                            except:
                                    alb = '-999'
                        else:
                            alb = -999
                        lat1 = lat2 - 1
                        lat_rat = (self.lati[self.lat_lon_ndx[hr]][lat2] - self.src_lat[i]) / (self.lati[self.lat_lon_ndx[hr]][lat2] - self.lati[self.lat_lon_ndx[hr]][lat1])
                        lon1 = lon2 - 1
                        lon_rat = (self.longi[self.lat_lon_ndx[hr]][lon2] - self.src_lon[i]) / (self.longi[self.lat_lon_ndx[hr]][lon2] - self.longi[self.lat_lon_ndx[hr]][lon1])
                        ghi = self.valu(self.ghi[hr], lat1, lon1, lat_rat, lon_rat)
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
                        if hour > 23:
                            hour = 0
                            day += 1
                            if day > dys[mth]:
                                mth += 1
                                day = 1
                                hour = 0
                else:
                    hdr = 'id,<city>,<state>,%s,%s,%s,0,3600.0,%s,0:30:00\n' % (str(self.src_zone),
                          round(self.src_lat[i], 4),
                          round(self.src_lon[i], 4), str(self.src_year))
                    tf.write(hdr)
                    for hr in range(len(self.s10m)):
                        for lat2 in range(len(self.lati[self.lat_lon_ndx[hr]])):
                            if self.src_lat[i] <= self.lati[self.lat_lon_ndx[hr]][lat2]:
                                break
                        for lon2 in range(len(self.longi[self.lat_lon_ndx[hr]])):
                            if self.src_lon[i] <= self.longi[self.lat_lon_ndx[hr]][lon2]:
                                break
                        if self.longrange[0] is None:
                            self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        else:
                            if self.longi[self.lat_lon_ndx[hr]][lon2] < self.longrange[0]:
                                self.longrange[0] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        if self.longrange[1] is None:
                            self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        else:
                            if self.longi[self.lat_lon_ndx[hr]][lon2] > self.longrange[1]:
                                self.longrange[1] = self.longi[self.lat_lon_ndx[hr]][lon2]
                        lat1 = lat2 - 1
                        lat_rat = (self.lati[self.lat_lon_ndx[hr]][lat2] - self.src_lat[i]) / (self.lati[self.lat_lon_ndx[hr]][lat2] - self.lati[self.lat_lon_ndx[hr]][lat1])
                        lon1 = lon2 - 1
                        lon_rat = (self.longi[self.lat_lon_ndx[hr]][lon2] - self.src_lon[i]) / (self.longi[self.lat_lon_ndx[hr]][lon2] - self.longi[self.lat_lon_ndx[hr]][lon1])
                        ghi = self.valu(self.ghi[hr], lat1, lon1, lat_rat, lon_rat)
                        dni = getDNI(ghi, hour=hr + 1, lat=self.src_lat[i], lon=self.src_lon[i],
                              press=self.valu(self.p_s[hr], lat1, lon1, lat_rat,
                              lon_rat, rnd=0), zone=self.src_zone)
                        dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.src_lat[i])
                        if len(self.alb) > 0:
                            alb = self.valu(self.alb[hr], lat1, lon1, lat_rat, lon_rat)
                        else:
                            alb = '-999'
                        tf.write(str(self.valu(self.t_10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) +
                        ',-999,-999,-999,' +
                        str(self.valu(self.s10m[hr], lat1, lon1, lat_rat, lon_rat)) + ',' +
                        str(self.valu(self.d10m[hr], lat1, lon1, lat_rat, lon_rat, rnd=0)) + ',' +
                        str(self.valu(self.p_s[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + ',' +
                        str(int(ghi)) + ',' + str(int(dni)) + ',' + str(int(dhi)) + ',' + alb + ',-999,\n')
                tf.close()
                self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
        else: # all locations
            if self.show_progress:
                self.caller.daybar.setMaximum(len(self.lats) * len(self.lons))
                QtCore.QCoreApplication.processEvents()
            for la in range(len(self.lats)):
                for lo in range(len(self.lons)):
                    if self.show_progress:
                        self.caller.daybar.setValue(len(self.lats) * la + lo)
                        QtCore.QCoreApplication.processEvents()
                    with_gaps = 0
                    missing = False
                    out_file = self.tgt_dir + 'solar_weather_' + '{:0.4f}'.format(self.lats[la]) + \
                            '_' + '{:0.4f}'.format(self.lons[lo]) + '_' + str(self.src_year) + '.' + self.fmat
                    tf = open(out_file, 'w')
                    if self.fmat == 'csv':
                        hdr = 'Location,City,Region,Country,Latitude,Longitude,Time Zone,Elevation,Source\n' + \
                              'id,<city>,<state>,<country>,%s,%s,%s,0,IWEC\n' % (round(self.lats[la], 4),
                              round(self.lons[lo], 4), str(self.src_zone))
                        tf.write(hdr)
                        tf.write('Year,Month,Day,Hour,GHI,DNI,DHI,Tdry,Pres,Wspd,Wdir' + '\n')
                        mth = 0
                        day = 1
                        hour = 0
                        for hr in range(len(self.s10m)):
                            try:
                                lat = self.latsi[self.lat_lon_ndx[hr]].index(self.lats[la])
                                lon = self.longsi[self.lat_lon_ndx[hr]].index(self.lons[lo])
                                la2 = self.lati[self.lat_lon_ndx[hr]].index(self.lats[la])
                                lo2 = self.longi[self.lat_lon_ndx[hr]].index(self.lons[lo])
                            except:
                                if self.gaps:
                                    tf.write(',,,,,,,,\n')
                                    with_gaps += 1
                                    continue
                                else:
                                    missing = True
                                    break
                            ghi = self.ghi[hr][lat][lon]
                            dni = getDNI(ghi, hour=hr + 1, lat=self.latsi[self.lat_lon_ndx[hr]][lat],
                                  lon=self.longsi[self.lat_lon_ndx[hr]][lon],
                                  press=self.p_s[hr][lat][lon], zone=self.src_zone)
                            dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.latsi[self.lat_lon_ndx[hr]][lat])
                            tf.write(str(self.src_year) + ',' + '{:02d}'.format(mth + 1) + ',' +
                                '{:02d}'.format(day) + ',' + '{:02d}'.format(hour) + ',' +
                                '{:0.1f}'.format(ghi) + ',' + '{:0.1f}'.format(dni) + ',' +
                                '{:0.1f}'.format(dhi) + ',' +
                                str(self.t_10m[hr][la2][lo2]) + ',' +
                                str(self.p_s[hr][la2][lo2]) + ',' +
                                str(self.s10m[hr][la2][lo2]) + ',' +
                                str(self.d10m[hr][la2][lo2]) + '\n')
                            hour += 1
                            if hour > 23:
                                hour = 0
                                day += 1
                                if day > dys[mth]:
                                    mth += 1
                                    day = 1
                                    hour = 0
                    else:
                        hdr = 'id,<city>,<state>,%s,%s,%s,0,3600.0,%s,0:30:00\n' % (str(self.src_zone),
                              round(self.lats[la], 4),
                              round(self.lons[lo], 4), str(self.src_year))
                        tf.write(hdr)
                        for hr in range(len(self.s10m)):
                            try:
                                lat = self.latsi[self.lat_lon_ndx[hr]].index(self.lats[la])
                                lon = self.longsi[self.lat_lon_ndx[hr]].index(self.lons[lo])
                                la2 = self.lati[self.lat_lon_ndx[hr]].index(self.lats[la])
                                lo2 = self.longi[self.lat_lon_ndx[hr]].index(self.lons[lo])
                            except:
                                if self.gaps:
                                    tf.write(',,,,,,,,\n')
                                    with_gaps += 1
                                    continue
                                else:
                                    missing = True
                                    break
                            if len(self.alb) > 0:
                                try:
                                    alb = self.alb[hr][lat][lon]
                                except:
                                    alb = '-999'
                            else:
                                alb = '-999'
                            ghi = self.ghi[hr][lat][lon]
                            dni = getDNI(ghi, hour=hr + 1, lat=self.latsi[self.lat_lon_ndx[hr]][lat],
                                  lon=self.longsi[self.lat_lon_ndx[hr]][lon],
                                  press=self.p_s[hr][lat][lon], zone=self.src_zone)
                            dhi = getDHI(ghi, dni, hour=hr + 1, lat=self.latsi[self.lat_lon_ndx[hr]][lat])
                            tf.write(str(self.t_10m[hr][la2][lo2]) +
                                ',-999,-999,-999,' +
                                str(self.s10m[hr][la2][lo2]) + ',' +
                                str(self.d10m[hr][la2][lo2]) + ',' +
                                str(self.p_s[hr][la2][lo2]) + ',' +
                                str(int(ghi)) + ',' + str(int(dni)) + ',' + str(int(dhi)) + ',' + alb + ',-999,\n')
                    tf.close()
                    if with_gaps > 0 and with_gaps < 504:
                        self.log += '%s created with gaps (%s days)\n' % ( \
                                    out_file[out_file.rfind('/') + 1:], str(int(with_gaps / 24)))
                    elif missing or with_gaps > 0:
                        os.remove(out_file)
                        self.gaplog += '%s not created due to data gaps\n' % out_file[out_file.rfind('/') + 1:]
                        continue
                    else:
                        self.log += '%s created\n' % out_file[out_file.rfind('/') + 1:]
        if self.show_progress:
            self.caller.daybar.setValue(self.caller.daybar.maximum())
            self.caller.progresslabel.setText('All done')
            QtCore.QCoreApplication.processEvents()

class getParms(QtWidgets.QWidget):

    def __init__(self, help='makeweatherfiles.html', ini_file='getfiles.ini'):
        super(getParms, self).__init__()
        self.help = help
        self.ini_file = getModelFile(ini_file)
        self.config = configparser.RawConfigParser()
        self.config_file = self.ini_file
        self.config.read(self.config_file)
        try:
            self.help = self.config.get('Files', 'help')
        except:
            pass
        self.ignore = False
        try:
            self.law = self.config.get('makeweatherfiles', 'extrapolate')
        except:
            self.law = 'l'
        rw = 0
        self.grid = QtWidgets.QGridLayout()
        self.grid.addWidget(QtWidgets.QLabel('Year:'), 0, 0)
        self.yearSpin = QtWidgets.QSpinBox()
        now = datetime.now()
        self.yearSpin.setRange(1979, now.year)
        self.yearSpin.setValue(now.year - 1)
        self.last_year = str(self.yearSpin.value())
        self.grid.addWidget(self.yearSpin, rw, 1)
        self.yearSpin.valueChanged.connect(self.yearChanged)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Wrap to prior year:'), rw, 0)
        self.wrapbox = QtWidgets.QCheckBox()
        self.wrapbox.setCheckState(QtCore.Qt.Checked)
        self.grid.addWidget(self.wrapbox, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('If checked will wrap back to prior year'), rw, 2, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Ignore data gaps:'), rw, 0)
        self.gapsbox = QtWidgets.QCheckBox()
        self.gapsbox.setCheckState(QtCore.Qt.Unchecked)
        self.grid.addWidget(self.gapsbox, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('If checked will ignore data gaps (but not missing days)'), rw, 2, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Time Zone:'), rw, 0)
        self.zoneCombo = QtWidgets.QComboBox()
        self.zoneCombo.addItem('Auto')
        self.zoneCombo.addItem('Best')
        for i in range(-12, 13):
            self.zoneCombo.addItem(str(i))
        self.zoneCombo.currentIndexChanged[str].connect(self.zoneChanged)
        self.zone_lon = QtWidgets.QLabel(('Time zone calculated from MERRA-2 / ERA5 data'))
        self.grid.addWidget(self.zoneCombo, rw, 1)
        self.grid.addWidget(self.zone_lon, rw, 2, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Solar Format:'), rw, 0)
        self.fmatcombo = QtWidgets.QComboBox(self)
        self.fmats = ['csv', 'smw']
        for i in range(len(self.fmats)):
            self.fmatcombo.addItem(self.fmats[i])
        self.fmatcombo.setCurrentIndex(1)
        self.grid.addWidget(self.fmatcombo, rw, 1)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Solar Variable:'), rw, 0)
        self.swgcombo = QtWidgets.QComboBox(self)
        self.swgs = ['swgdn', 'swgnt']
        for i in range(len(self.swgs)):
            self.swgcombo.addItem(self.swgs[i])
        self.swgcombo.setCurrentIndex(0) # default is swgdn/ssrd
        self.grid.addWidget(self.swgcombo, rw, 1)
        # ssrd: https://gis.stackexchange.com/questions/389174/calculating-global-horizontal-irradiance-ghi-from-surface-solar-radiation-down
        # swgdn: https://earthscience.stackexchange.com/questions/19062/how-to-get-ghi-data-from-merra-2
        self.grid.addWidget(QtWidgets.QLabel("ERA5 variables are 'swgdn=ssrd' and 'swgnt=ssr'"), rw, 2, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Hub height:'), rw, 0)
        self.hub_height = QtWidgets.QSpinBox()
        self.hub_height.setRange(0, 200)
        self.hub_height.setSingleStep(10)
        try:
            self.hub_height.setValue(int(self.config.get('makeweatherfiles', 'hub_height')))
        except:
            pass
        self.grid.addWidget(self.hub_height, rw, 1)
        self.hub_height.valueChanged.connect(self.hub_heightChanged)
        self.grid.addWidget(QtWidgets.QLabel('To extrapolate wind data above 50 / 100 metres; using:'), rw, 2, 1, 2)
        self.lawcombo = QtWidgets.QComboBox(self)
        self.laws = ['logarithmic', 'hellman']
        for i in range(len(self.laws)):
            self.lawcombo.addItem(self.laws[i])
            if self.law == self.laws[i][0]  :
                self.lawcombo.setCurrentIndex(i)
        self.grid.addWidget(self.lawcombo, rw, 4)
        self.lawcombo.currentIndexChanged.connect(self.lawChanged)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Coordinates:'), rw, 0)
        self.coords = QtWidgets.QPlainTextEdit()
        self.grid.addWidget(self.coords, rw, 1, 1, 4)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Copy folder down:'), rw, 0)
        self.checkbox = QtWidgets.QCheckBox()
        self.checkbox.setCheckState(QtCore.Qt.Checked)
        self.grid.addWidget(self.checkbox, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('If checked will copy solar folder changes down to others'), rw, 2, 1, 3)
        rw += 1
        cur_dir = os.getcwd()
        self.dir_labels = ['Solar', 'Wind', 'Target']
        xtr_labels = ['/ERA5', '', '']
        self.dirs = [None, None, None, None]
        for i in range(3):
            self.grid.addWidget(QtWidgets.QLabel(self.dir_labels[i] + xtr_labels[i] + ' Folder:'), rw, 0)
            self.dirs[i] = ClickableQLabel()
            try:
                self.dirs[i].setText(self.config.get('makeweatherfiles', self.dir_labels[i].lower() + '_files').replace('$USER$', getUser()))
            except:
                try:
                    self.dirs[i].setText(self.config.get('Files', self.dir_labels[i].lower() + '_files').replace('$USER$', getUser()))
                except:
                    self.dirs[i].setText(cur_dir)
            self.dirs[i].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
            self.dirs[i].clicked.connect(self.dirChanged)
            self.grid.addWidget(self.dirs[i], rw, 1, 1, 4)
            rw += 1
        self.daybar = QtWidgets.QProgressBar()
        self.daybar.setMinimum(0)
        self.daybar.setMaximum(31)
        self.daybar.setValue(0)
        #06A9D6 #FFC709
        self.daybar.setStyleSheet('QProgressBar {border: 1px solid grey; border-radius: 2px; text-align: center;}' \
                                       + 'QProgressBar::chunk { background-color: #FFC709;}')
        self.daybar.setHidden(True)
        self.grid.addWidget(self.daybar, rw, 0)
        self.progressbar = QtWidgets.QProgressBar()
        self.progressbar.setMinimum(0)
        self.progressbar.setMaximum(100)
        self.progressbar.setValue(0)
        #06A9D6 #FFC709
        self.progressbar.setStyleSheet('QProgressBar {border: 1px solid grey; border-radius: 2px; text-align: center;}' \
                                       + 'QProgressBar::chunk { background-color: #06A9D6;}')
        self.grid.addWidget(self.progressbar, rw, 1, 1, 4)
        self.progressbar.setHidden(True)
        self.progresslabel = QtWidgets.QLabel('')
        self.grid.addWidget(self.progresslabel, rw, 1, 1, 2)
        self.progresslabel.setHidden(True)
        rw += 1
        quit = QtWidgets.QPushButton('Done', self)
        self.grid.addWidget(quit, rw, 0)
        quit.clicked.connect(self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        dosolar = QtWidgets.QPushButton('Produce Solar Files', self)
        wdth = dosolar.fontMetrics().boundingRect(dosolar.text()).width() + 9
        quit.setMaximumWidth(wdth)
        dosolar.setMaximumWidth(wdth)
        self.grid.addWidget(dosolar, rw, 1)
        dosolar.clicked.connect(self.dosolarClicked)
        dowind = QtWidgets.QPushButton('Produce Wind Files', self)
        dowind.setMaximumWidth(wdth)
        self.grid.addWidget(dowind, rw, 2)
        dowind.clicked.connect(self.dowindClicked)
        updwind = QtWidgets.QPushButton('Update Wind Files', self)
        updwind.setMaximumWidth(wdth)
        self.grid.addWidget(updwind, rw, 3)
        updwind.clicked.connect(self.updwindClicked)
        info = QtWidgets.QPushButton('File Info', self)
        info.setMaximumWidth(wdth)
        self.grid.addWidget(info, rw, 4)
        info.clicked.connect(self.infoClicked)
        help = QtWidgets.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        self.grid.addWidget(help, rw, 5)
        help.clicked.connect(self.helpClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        self.yearSpin.setMaximumWidth(wdth)
        self.zoneCombo.setMaximumWidth(wdth)
        self.fmatcombo.setMaximumWidth(wdth)
        self.swgcombo.setMaximumWidth(wdth)
        self.lawcombo.setMaximumWidth(wdth)
        self.hub_height.setMaximumWidth(wdth)
      #   self.grid.setColumnStretch(4, 2)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('makeweatherfiles (' + fileVersion() + ') - Make weather files from MERRA-2 or ERA5 data')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.center()
        self.resize(int(self.sizeHint().width()* 1.27), int(self.sizeHint().height() * 1.07))
        self.updated = False
        self.hubupdated = False
        self.lawupdated = False
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        centerPoint = QtWidgets.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def hub_heightChanged(self):
        self.hubupdated = True
        if self.hub_height.value() > 0 and self.hub_height.value() < 80:
            if self.hub_height.value() == 70:
                self.hub_height.setValue(0)
            else:
                self.hub_height.setValue(80)

    def lawChanged(self):
        self.lawupdated = True
        self.law = self.lawcombo.currentText()[0].lower()

    def yearChanged(self):
        yr = str(self.yearSpin.value())
        for i in range(3):
            if self.dirs[i].text()[-len(self.last_year):] == self.last_year:
                self.dirs[i].setText(self.dirs[i].text()[:-len(self.last_year)] + yr)
                self.updated = True
        self.last_year = yr

    def zoneChanged(self, val):
        if self.zoneCombo.currentIndex() <= 1:
            lon = '(Time zone calculated from MERRA-2 or ERA5 data)'
        else:
            lw = int(self.zoneCombo.currentIndex() - 14) * 15 - 7.5
            le = int(self.zoneCombo.currentIndex() - 14) * 15 + 7.5
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
        newdir = QtWidgets.QFileDialog.getExistingDirectory(self, 'Choose ' + self.dir_labels[i] + ' Folder',
                 curdir, QtWidgets.QFileDialog.ShowDirsOnly)
        if newdir != '':
            self.dirs[i].setText(newdir)
            if self.checkbox.isChecked():
                if i == 0:
                    self.dirs[1].setText(newdir)
                    self.dirs[2].setText(newdir)
            self.updated = True

    def helpClicked(self):
        dialog = ShowHelp(QtWidgets.QDialog(), self.help,
                 title='makeweatherfiles (' + fileVersion() + ') - Help', section='makeweather')
        dialog.exec_()

    def quitClicked(self):
        if self.updated or self.hubupdated or self.lawupdated:
            updates = {}
            lines = []
            if self.updated:
                for i in range(3):
                    lines.append(self.dir_labels[i].lower() + '_files=' + \
                                 self.dirs[i].text().replace(getUser(), '$USER$'))
            if self.hubupdated:
                lines.append('hub_height=')
                if self.hub_height.value() > 0:
                    lines[-1] += str(self.hub_height.value())
            if self.lawupdated:
                lines.append('extrapolate=')
                if self.law != 'l':
                    lines[-1] += self.law
            updates['makeweatherfiles'] = lines
            SaveIni(updates, ini_file=self.config_file)
        self.close()

    def dosolarClicked(self):
        self.progressbar.setHidden(False)
        self.progresslabel.setHidden(False)
        self.daybar.setHidden(False)
        coords = self.coords.toPlainText()
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
        if self.gapsbox.isChecked():
            gaps = 'y'
        else:
            gaps = ''
        if self.zoneCombo.currentIndex() == 0:
            zone = 'auto'
        elif self.zoneCombo.currentIndex() == 1:
            zone = 'best'
        else:
            zone = str(self.zoneCombo.currentIndex() - 14)
        solar = makeWeather(self, str(self.yearSpin.value()), zone, self.dirs[0].text(), \
                            self.dirs[1].text(), self.dirs[2].text(), self.fmatcombo.currentText(), \
                            self.swgcombo.currentText(), wrap, gaps, coords)
        dialr = RptDialog(str(self.yearSpin.value()), zone, self.dirs[0].text(), \
                          self.dirs[1].text(), self.dirs[2].text(), self.fmatcombo.currentText(), \
                          self.swgcombo.currentText(), wrap, gaps, coords, solar.returnCode(), solar.getLog())
        dialr.exec_()
        del solar
        del dialr
        self.progressbar.setValue(0)
        self.progressbar.setHidden(True)
        self.progresslabel.setHidden(True)
        self.daybar.setMaximum(31)
        self.daybar.setValue(0)
        self.daybar.setHidden(True)

    def dowindClicked(self):
        self.progressbar.setHidden(False)
        self.progresslabel.setHidden(False)
        self.daybar.setHidden(False)
        coords = self.coords.toPlainText()
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
        if self.gapsbox.isChecked():
            gaps = 'y'
        else:
            gaps = ''
        if self.zoneCombo.currentIndex() == 0:
            zone = 'auto'
        elif self.zoneCombo.currentIndex() == 1:
            zone = 'best'
        else:
            zone = str(self.zoneCombo.currentIndex() - 14)
        wind = makeWeather(self, str(self.yearSpin.value()), zone, self.dirs[0].text(), \
                           self.dirs[1].text(), self.dirs[2].text(), \
                            'wind', '', wrap, gaps, coords, hub_height=self.hub_height.value(), law=self.law)
        dialr = RptDialog(str(self.yearSpin.value()), zone, self.dirs[0].text(), \
                          self.dirs[1].text(), self.dirs[2].text(), 'srw', \
                          '', wrap, gaps, coords, wind.returnCode(), wind.getLog())
        dialr.exec_()
        del wind
        del dialr
        self.progressbar.setValue(0)
        self.progressbar.setHidden(True)
        self.progresslabel.setHidden(False)
        self.daybar.setMaximum(31)
        self.daybar.setValue(0)
        self.daybar.setHidden(True)

    def updwindClicked(self):
        if self.hub_height.value() < 80:
            return
        self.progresslabel.setHidden(False)
        if not os.path.exists(self.dirs[2].text()):
            self.progresslabel.setText("Target folder doesn't exist")
            return
        self.progresslabel.setText('Updating wind weather files')
        fils = os.listdir(self.dirs[2].text())
        for f in range(len(fils) -1, -1, -1):
            if fils[f][-4:] != '.srw':
                del fils[f]
        ctr = 0
        for f in range(len(fils)):
            ok = extrapolateWind(self.dirs[2].text() + '/' + fils[f], self.hub_height.value(), law=self.law, replace=True)
            if ok:
                self.progresslabel.setText('Updated ' + fils[f])
                ctr += 1
            QtCore.QCoreApplication.processEvents()
        if ctr < len(fils):
            msg = '{:d} (of {:d})'.format(ctr, len(fils))
        else:
            msg = str(ctr)
        self.progresslabel.setText('Updated ' + msg + ' wind weather files')

    def infoClicked(self):
        coords = self.coords.toPlainText()
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
        if self.gapsbox.isChecked():
            gaps = 'y'
        else:
            gaps = ''
        if self.zoneCombo.currentIndex() == 0:
            zone = 'auto'
        elif self.zoneCombo.currentIndex() == 1:
            zone = 'best'
        else:
            zone = str(self.zoneCombo.currentIndex() - 14)
        wind = makeWeather(self, str(self.yearSpin.value()), zone, self.dirs[0].text(), \
                           self.dirs[1].text(), self.dirs[2].text(), \
                           'both', self.swgcombo.currentText(), wrap, gaps, coords, info=True)
        dialr = RptDialog(str(self.yearSpin.value()), zone, self.dirs[0].text(), \
                          self.dirs[1].text(), self.dirs[2].text(), 'srw', \
                          self.swgcombo.currentText(), wrap, gaps, coords, wind.returnCode(), wind.getLog())
        dialr.exec_()
        del wind
        del dialr

    def progress(self, pct):
        self.progressbar.setValue(int(pct * 100.))  #  @QtCore.pyqtSlot()

class RptDialog(QtWidgets.QDialog):
    def __init__(self, year, zone, solar_dir, wind_dir, tgt_dir, fmat, swg, wrap, gaps, coords, return_code, output):
        super(RptDialog, self).__init__()
        self.parms = [str(year), str(zone), swg, wrap, gaps, fmat, tgt_dir, solar_dir, wind_dir]
        self.tgt_dir = tgt_dir
        if wrap == 'y':
            wrapy = 'Yes'
        else:
            wrapy = 'No'
        if gaps == 'y':
            gapsy = 'Yes'
        else:
            gapsy = 'No'
        max_line = 0
        if swg == 'swgdn':
            eraswg = 'ssrd'
        else:
            eraswg = 'ssr'
        self.lines = 'Parameters:\n    Year: %s\n    Wrap year: %s\n    Ignore gaps: %s\n    Time Zone: %s\n    Output Format: %s\n' \
                     % (year, wrapy, gapsy, zone, fmat)
        if fmat != 'srw':
            self.lines += '    Solar Files: %s\n' % solar_dir
        self.lines += '    Solar Variable: %s (ERA5 variable is %s)\n' % (swg, eraswg)
        self.lines += '    Wind/ERA5 Files: %s\n' % wind_dir
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
        QtWidgets.QDialog.__init__(self)
        self.saveButton = QtWidgets.QPushButton(self.tr('&Save'))
        self.cancelButton = QtWidgets.QPushButton(self.tr('Cancel'))
        buttonLayout = QtWidgets.QHBoxLayout()
        buttonLayout.addStretch(1)
        buttonLayout.addWidget(self.saveButton)
        buttonLayout.addWidget(self.cancelButton)
        self.saveButton.clicked.connect(self.accept)
        self.cancelButton.clicked.connect(self.reject)
        self.widget = QtWidgets.QTextEdit()
        self.widget.setFont(QtGui.QFont('Courier New', 11))
        fnt = self.widget.fontMetrics()
        ln = (max_line + 5) * fnt.maxWidth()
        ln2 = (line_cnt + 2) * fnt.height()
        screen = QtWidgets.QDesktopWidget().availableGeometry()
        if ln > screen.width() * .67:
            ln = int(screen.width() * .67)
        if ln2 > screen.height() * .67:
            ln2 = int(screen.height() * .67)
        self.widget.setSizePolicy(QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding,
            QtWidgets.QSizePolicy.Expanding))
        self.widget.resize(ln, ln2)
        self.widget.setPlainText(self.lines)
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.widget)
        layout.addLayout(buttonLayout)
        self.setLayout(layout)
        self.setWindowTitle('makeweatherfiles - Output')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        size = self.geometry()
        self.setGeometry(1, 1, ln + 10, ln2 + 35)
        size = self.geometry()
        self.move(int((screen.width() - size.width()) / 2),
                  int((screen.height() - size.height()) / 2))
        self.widget.show()

    def accept(self):
        i = sys.argv[0].rfind('/')  # fudge to see if program has a directory to use as an alternative
        j = sys.argv[0].rfind('.')
        save_filename = self.tgt_dir + '/'
      #  if i > 0:
       #     save_filename += sys.argv[0][i + 1:j]
        #else:
         #   save_filename += sys.argv[0][:j]
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
        fileName = QtWidgets.QFileDialog.getSaveFileName(self,
                                         self.tr('Save makeweatherfiles Report'),
                                         save_filename,
                                         self.tr('All Files (*);;Text Files (*.txt)'))[0]
        if fileName != '':
            s = open(fileName, 'w')
            s.write(self.lines)
            s.close()
        self.close()


if "__main__" == __name__:
    app = QtWidgets.QApplication(sys.argv)
    if len(sys.argv) > 1:  # arguments
        src_lat_lon = ''
        src_year = 2014
        swg = 'swgdn'
        wrap = ''
        gaps = ''
        src_zone = 'auto'
        fmat = 'srw'
        src_dir_s = ''
        src_dir_w = ''
        tgt_dir = ''
        hub_height = 0
        law = 'l'
        for i in range(1, len(sys.argv)):
            if sys.argv[i][:5] == 'year=':
                src_year = int(sys.argv[i][5:])
            elif sys.argv[i][:5] == 'wrap=':
                wrap = sys.argv[i][5:]
            elif sys.argv[i][:5] == 'gaps=':
                gaps = sys.argv[i][5:]
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
            elif sys.argv[i][:11] == 'hub_height=':
                hub_height = int(sys.argv[i][11:])
            elif sys.argv[i][:12] == 'extrapolate=':
                law = sys.argv[i][12:]
            elif sys.argv[i][:4] == 'law=':
                law = sys.argv[i][4:]
            elif sys.argv[i][:4] == 'swg=':
                swg = sys.argv[i][4:]
            elif sys.argv[i][:6] == 'solar=':
                src_dir_s = sys.argv[i][6:]
            elif sys.argv[i][:7] == 'erasrc=':
                src_dir_s = sys.argv[i][6:]
            elif sys.argv[i][:7] == 'source=' or sys.argv[i][:7] == 'srcdir=':
                src_dir_w = sys.argv[i][7:]
            elif sys.argv[i][:5] == 'wind=':
                src_dir_w = sys.argv[i][5:]
            elif sys.argv[i][:7] == 'target=' or sys.argv[i][:7] == 'tgtdir=':
                tgt_dir = sys.argv[i][7:]
        if fmat == 'solar':
            fmat = 'smw'
        if fmat in ['csv', 'smw'] and src_dir_s == '':
            src_dir_s = src_dir_w
        makeWeather(None, src_year, src_zone, src_dir_s, src_dir_w, tgt_dir, fmat, swg, wrap, gaps, src_lat_lon, hub_height=hub_height, law=law)
        dialr = RptDialog(str(src_year), src_zone, src_dir_s, src_dir_w, tgt_dir, fmat, swg, wrap, gaps, src_lat_lon, files.returnCode(), files.getLog())
        dialr.exec_()
    else:
        ex = getParms()
        app.exec_()
        app.deleteLater()
        sys.exit()
