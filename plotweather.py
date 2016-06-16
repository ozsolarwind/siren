#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
#
#  plotweather.py - This file is part of SIREN.
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

from math import asin, ceil, cos, radians, sin, sqrt
import pylab as plt
from matplotlib.font_manager import FontProperties
import matplotlib.lines as mlines
import csv
import os
import sys
import xlrd

import ConfigParser  # decode .ini file
from PyQt4 import QtGui, QtCore
from PyQt4.QtCore import Qt

from senuser import getUser
from sammodels import getZenith

the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]


class whatPlots(QtGui.QDialog):
    def __init__(self, plots, plot_order, hdrs, spacers, base_year, comment):
        self.plots = plots
        self.plot_order = plot_order
        self.hdrs = hdrs
        self.spacers = spacers
        self.base_year = int(base_year)
        self.comment = comment
        super(whatPlots, self).__init__()
        self.initUI()

    def initUI(self):
        self.grid = QtGui.QGridLayout()
        self.checkbox = []
        i = 0
        bold = QtGui.QFont()
        bold.setBold(True)
        for plot in range(len(self.plot_order)):
            if self.plot_order[plot] in self.spacers:
                label = QtGui.QLabel(self.spacers[self.plot_order[plot]])
                label.setFont(bold)
                self.grid.addWidget(label, i, 0)
                i += 1
            self.checkbox.append(QtGui.QCheckBox(self.hdrs[self.plot_order[plot]], self))
            if self.plots[self.plot_order[plot]]:
                self.checkbox[plot].setCheckState(Qt.Checked)
            self.grid.addWidget(self.checkbox[-1], i, 0)
            i += 1
        self.grid.connect(self.checkbox[0], QtCore.SIGNAL('stateChanged(int)'), self.check_all)
        show = QtGui.QPushButton('Proceed', self)
        show.clicked.connect(self.showClicked)
        self.grid.addWidget(show, i, 0)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        commnt = QtGui.QLabel('Nearest weather files:\n' + self.comment)
        self.layout.addWidget(commnt)
        self.setWindowTitle('SIREN - Weather dialog for ' + str(self.base_year))
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        self.show_them = False
        self.show()

    def check_all(self):
        if self.checkbox[0].isChecked():
            for i in range(len(self.checkbox)):
                self.plots[self.plot_order[i]] = True
                self.checkbox[i].setCheckState(Qt.Checked)
        else:
            for i in range(len(self.checkbox)):
                self.plots[self.plot_order[i]] = True
                self.checkbox[i].setCheckState(Qt.Unchecked)

    def closeEvent(self, event):
        if not self.show_them:
            self.plots = None
        event.accept()

    def quitClicked(self):
        self.plots = None
        self.close()

    def showClicked(self):
        for plot in range(len(self.checkbox)):
            if self.checkbox[plot].checkState() == Qt.Checked:
                self.plots[self.plot_order[plot]] = True
            else:
                self.plots[self.plot_order[plot]] = False
        self.show_them = True
        self.close()

    def getValues(self):
        return self.plots


class PlotWeather():
    def haversine(self, lat1, lon1, lat2, lon2):
        """
        Calculate the great circle distance between two points
        on the earth (specified in decimal degrees)
        """
   #     convert decimal degrees to radians
        lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])

   #     haversine formula
        dlon = lon2 - lon1
        dlat = lat2 - lat1
        a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
        c = 2 * asin(sqrt(a))

   #     6367 km is the radius of the Earth
        km = 6367 * c
        return km

    def find_closest(self, latitude, longitude, wind=False, rain=False):
        dist = 99999
        closest = ''
        close_lat = 0
        close_lon = 0
        if wind:
            filetype = '.srw'
            technology = 'wind_index'
            index_file = self.wind_index
            folder = self.wind_files
        elif rain:
            filetype = '.csv'
            technology = 'rain_index'
            index_file = ''
            folder = self.rain_files
        else:
            filetype = '.smw'
            technology = 'solar_index'
            index_file = self.solar_index
            folder = self.solar_files
        if folder != '':
            if index_file == '':
                fils = os.listdir(folder)
                for fil in fils:
                    if fil[-4:] == filetype or fil[-4:] == '.csv':
                        bit = fil.split('_')
                        if bit[-1][:4] == str(self.base_year):
                            dist1 = self.haversine(float(bit[-3]), float(bit[-2]), latitude, longitude)
                            if dist1 < dist:
                                closest = fil
                                dist = dist1
                                close_lat = bit[-3]
                                close_lon = bit[-2]
            else:
                fils = []
                ndx_file = ''
                if os.path.exists(index_file):
                    ndx_file = index_file
                elif os.path.exists(folder + os.sep + index_file):
                    ndx_file = folder + os.sep + index_file
                if ndx_file != '':
                    if ndx_file[-4:] == '.xls' or ndx_file[-5:] == '.xlsx':
                        var = {}
                        xl_file = xlrd.open_workbook(ndx_file)
                        worksheet = xl_file.sheet_by_index(0)
                        num_rows = worksheet.nrows - 1
                        num_cols = worksheet.ncols - 1
#                       get column names
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
                    else:
                        dft_variables = csv.DictReader(open(ndx_file))
                        for var in dft_variables:
                            lat = float(var['Latitude'])
                            lon = float(var['Longitude'])
                            fil = var['Filename']
                            fils.append([lat, lon, fil])
                    for fil in fils:
                        dist1 = self.haversine(fil[0], fil[1], latitude, longitude)
                        if dist1 < dist:
                            closest = fil[2]
                            dist = dist1
        if __name__ == '__main__':
            print closest
        return closest, dist, close_lat, close_lon

    def showGraphs(self, ly, x, locn):
        def dayPlot(self, period, data, locn, per_labels=None, x_labels=None):
            plt.figure(period)
            plt.suptitle(self.hdrs[period] + locn, fontsize=16)
            maxy = 0
            maxw = 0
            if len(data[0]) > 4:
                p1 = 3
                p2 = 4
                xl = 8
                yl = [0, 4, 8]
                yl2 = [3, 7, 11]
            else:
                p1 = p2 = 2
                if len(data[0]) == 4:
                    xl = 2
                    yl = [0, 2]
                    yl2 = [1, 3]
                else:
                    xl = 0
                    yl = yl2 = [0, 1]
            wi = -1
            ra = -1
            te = -1
            i = -1
            for key, value in iter(sorted(self.ly.iteritems())):
                i += 1
                if key == 'wind':
                    wi = i
                elif key == 'temp':
                    te = i
                elif key == 'rain':
                    ra = i
            for p in range(len(data[0])):
                for i in range(len(ly)):
                    maxy = max(maxy, max(data[i][p]))
                    if i == wi or i == te or i == ra:
                        maxw = max(maxw, max(data[i][p]))
            for p in range(len(data[0])):
                px = plt.subplot(p1, p2, p + 1)
                i = -1
                if self.two_axes:
                    px2 = px.twinx()
                for key, value in iter(sorted(self.ly.iteritems())):
                    i += 1
                    lw = 2.0
                    if self.two_axes and key in self.ylabel2[0]:
                        px2.plot(x24, data[i][p], linewidth=lw, label=key, color=self.colours[key])
                    else:
                        px.plot(x24, data[i][p], linewidth=lw, label=key, color=self.colours[key])
                    plt.title(per_labels[p])
                plt.xticks(range(0, 25, 4))
                px.set_xticklabels(x_labels)
                px.set_ylim([0, maxy])
                if self.two_axes:
                    px2.set_ylim(0, maxw)
                    if p in yl2:
                        px2.set_ylabel(self.ylabel2[2])
                if p >= xl:
                    px.set_xlabel('Hour of the Day')
                if p in yl:
                    px.set_ylabel(self.ylabel[1])
            if self.two_axes:
                lines = []
                for key in self.ly:
                    lines.append(mlines.Line2D([], [], color=self.colours[key], label=key))
                px.legend(handles=lines, loc='best')
            else:
                px.legend(loc='best')
            if self.maximise:
                mng = plt.get_current_fig_manager()
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if plt.get_backend() == 'TkAgg':
                        mng.window.state('zoomed')
                    elif plt.get_backend() == 'Qt4Agg':
                        mng.window.showMaximized()
                else:
                    mng.resize(*mng.window.maxsize())
            if self.plots['block']:
                plt.show(block=True)
            else:
                plt.draw()

        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        seasons = [[], [], [], []]
        periods = [[], []]
        self.maximise = False
        try:
            items = config.items('Power')
            for item, values in items:
                if item[:6] == 'season':
                    if item == 'season':
                        continue
                    i = int(item[6:]) - 1
                    seasons[i] = values.split(',')
                    for j in range(1, len(seasons[i])):
                        seasons[i][j] = int(seasons[i][j]) - 1
                if item[:6] == 'period':
                    if item == 'period':
                        continue
                    i = int(item[6:]) - 1
                    periods[i] = values.split(',')
                    for j in range(1, len(periods[i])):
                        periods[i][j] = int(periods[i][j]) - 1
                if item == 'maximise':
                    if values.lower() in ['true', 'on', 'yes']:
                        self.maximise = True
        except:
            seasons[0] = ['Summer', 11, 0, 1]
            seasons[1] = ['Autumn', 2, 3, 4]
            seasons[2] = ['Winter', 5, 6, 7]
            seasons[3] = ['Spring', 8, 9, 10]
            periods[0] = ['Winter', 5, 6, 7, 8, 9, 10]
            periods[1] = ['Summer', 11, 12, 1, 2, 3, 4]
        mth_labels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        ssn_labels = []
        for i in range(len(seasons)):
            ssn_labels.append('%s (%s-%s)' % (seasons[i][0], mth_labels[seasons[i][1]],
                               mth_labels[seasons[i][-1]]))
      #   ssn_labels = ['Summer (Dec, Jan-Feb)', 'Autumn (Mar-May)', 'Winter (Jun-Aug)', 'Spring(Sep-Nov)']
        smp_labels = []
        for i in range(len(periods)):
            smp_labels.append('%s (%s-%s)' % (periods[i][0], mth_labels[periods[i][1]],
                               mth_labels[periods[i][-1]]))
     #    smp_labels = ['Winter (May-Oct)', 'Summer (Nov-Apr)']
        labels = ['0:00', '4:00', '8:00', '12:00', '16:00', '20:00', '24:00']
        mth_xlabels = ['0:', '4:', '8:', '12:', '16:', '20:', '24:']
        m = 0
        d = 1
        day_labels = []
        while m < len(the_days):
            day_labels.append('%s %s' % (str(d), mth_labels[m]))
            d += 7
            if d > the_days[m]:
                d = d - the_days[m]
                m += 1
        lbl_font = FontProperties()
        lbl_font.set_size('small')
        x24 = []
        l24 = []
        m24 = []
        q24 = []
        s24 = []
        t12 = []
        for i in range(24):
            x24.append(i + 1)
        for i in range(len(self.ly)):
            if self.plots['total']:
                l24.append([])
                for j in range(24):
                    l24[i].append(0.)
            if self.plots['month']:
                m24.append([])
                for m in range(12):
                    m24[i].append([])
                    for j in range(24):
                        m24[i][m].append(0.)
            if self.plots['season']:
                q24.append([])
                for q in range(4):
                    q24[i].append([])
                    for j in range(24):
                        q24[i][q].append(0.)
            if self.plots['period']:
                s24.append([])
                for s in range(2):
                    s24[i].append([])
                    for j in range(24):
                        s24[i][s].append(0.)
            if self.plots['monthly']:
                t12.append([])
                for m in range(14):
                    t12[i].append(0.)
        the_qtrs = [the_days[0] + the_days[1] + the_days[11],
                    the_days[2] + the_days[3] + the_days[4],
                    the_days[5] + the_days[6] + the_days[7],
                    the_days[8] + the_days[9] + the_days[10]]
        the_qtrs = []
        for i in range(len(seasons)):
            d = 0
            for j in range(1, len(seasons[i])):
                d += the_days[seasons[i][j]]
            the_qtrs.append(d)
        the_ssns = [the_days[4] + the_days[5] + the_days[6] + the_days[7] + the_days[8] + the_days[9],
                    the_days[10] + the_days[11] + the_days[0] + the_days[1] + the_days[2] + the_days[3]]
        the_ssns = []
        for i in range(len(periods)):
            d = 0
            for j in range(1, len(periods[i])):
                d += the_days[periods[i][j]]
            the_ssns.append(d)
        the_hours = [0]
        i = 0
        for m in range(len(the_days)):
            i = i + the_days[m] * 24
            the_hours.append(i)
        d = -1
        for i in range(0, len(x), 24):
            m = 11
            d += 1
            while i < the_hours[m] and m > 0:
                m -= 1
            for k in range(24):
                j = -1
                for key, value in iter(sorted(self.ly.iteritems())):
                    j += 1
                    if self.plots['total']:
                        l24[j][k] += value[i + k]
                    if self.plots['month']:
                        m24[j][m][k] = m24[j][m][k] + value[i + k]
                    if self.plots['season']:
                        for q in range(len(seasons)):
                            if m in seasons[q]:
                                break
                        q24[j][q][k] = q24[j][q][k] + value[i + k]
                    if self.plots['period']:
                        for s in range(len(periods)):
                            if m in periods[s]:
                                break
                        s24[j][s][k] = s24[j][s][k] + value[i + k]
                    if self.plots['monthly']:
                        t12[j][m + 1] = t12[j][m + 1] + value[i + k]
        for i in range(len(ly)):
            for k in range(24):
                if self.plots['total']:
                    l24[i][k] = l24[i][k] / 365
                if self.plots['month']:
                    for m in range(12):
                        m24[i][m][k] = m24[i][m][k] / the_days[m]
                if self.plots['season']:
                    for q in range(4):
                        q24[i][q][k] = q24[i][q][k] / the_qtrs[q]
                if self.plots['period']:
                    for s in range(2):
                        s24[i][s][k] = s24[i][s][k] / the_ssns[s]
        if self.plots['hour']:
            fig = plt.figure('hour')
            plt.grid(True)
            hx = fig.add_subplot(111)
            plt.title(self.hdrs['hour'] + ' - ' + locn)
            maxy = 0
            if self.two_axes:
                hx2 = hx.twinx()
            for key, value in iter(sorted(self.ly.iteritems())):
                lw = 2.0
                if self.two_axes and key in self.ylabel2[0]:
                    hx2.plot(x, value, linewidth=lw, label=key, color=self.colours[key])
                else:
                    hx.plot(x, value, linewidth=lw, label=key, color=self.colours[key])
                    maxy = max(maxy, max(value))
            hx.set_ylim([0, maxy])
            plt.xlim([0, len(x)])
            plt.xticks(range(12, len(x), 168))
            hx.set_xticklabels(day_labels, rotation='vertical')
            hx.set_xlabel('Month of the year')
            hx.set_ylabel(self.ylabel[0])
            hx.legend(loc='best')
            if self.two_axes:
                ylim = hx2.get_ylim()
                hx2.set_ylim(0, ylim[1])
                lines = []
                for key in self.ly:
                    lines.append(mlines.Line2D([], [], color=self.colours[key], label=key))
                hx.legend(handles=lines, loc='best')
                hx2.set_ylabel(self.ylabel2[1])
            if self.maximise:
                mng = plt.get_current_fig_manager()
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if plt.get_backend() == 'TkAgg':
                        mng.window.state('zoomed')
                    elif plt.get_backend() == 'Qt4Agg':
                        mng.window.showMaximized()
                else:
                    mng.resize(*mng.window.maxsize())
            if self.plots['block']:
                plt.show(block=True)
            else:
                plt.draw()
        if self.plots['total']:
            figt = plt.figure('total')
            plt.grid(True)
            tx = figt.add_subplot(111)
            plt.title(self.hdrs['total'] + ' - ' + locn)
            maxy = 0
            i = -1
            if self.two_axes:
                tx2 = tx.twinx()
            for key, value in iter(sorted(self.ly.iteritems())):
                i += 1
                lw = 2.0
                if self.two_axes and key in self.ylabel2[0]:
                    tx2.plot(x24, l24[i], linewidth=lw, label=key, color=self.colours[key])
                else:
                    tx.plot(x24, l24[i], linewidth=lw, label=key, color=self.colours[key])
                    maxy = max(maxy, max(l24[i]))
            tx.set_ylim([0, maxy])
            plt.xlim([1, 25])
            plt.xticks(range(0, 25, 4))
            tx.set_xticklabels(labels)
            tx.set_xlabel('Hour of the Day')
            tx.set_ylabel(self.ylabel[0])
            tx.legend(loc='best')
            if self.two_axes:
                ylim = tx2.get_ylim()
                tx2.set_ylim(0, ylim[1])
                lines = []
                for key in self.ly:
                    lines.append(mlines.Line2D([], [], color=self.colours[key], label=key))
                tx.legend(handles=lines, loc='best')
                tx2.set_ylabel(self.ylabel2[1])
            if self.maximise:
                mng = plt.get_current_fig_manager()
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if plt.get_backend() == 'TkAgg':
                        mng.window.state('zoomed')
                    elif plt.get_backend() == 'Qt4Agg':
                        mng.window.showMaximized()
                else:
                    mng.resize(*mng.window.maxsize())
            if self.plots['block']:
                plt.show(block=True)
            else:
                plt.draw()
        if self.plots['month']:
            dayPlot(self, 'month', m24, locn, mth_labels, mth_xlabels)
        if self.plots['season']:
            dayPlot(self, 'season', q24, locn, ssn_labels, labels)
        if self.plots['period']:
            dayPlot(self, 'period', s24, locn, smp_labels, labels)
        if 'pdf' in self.plots.keys():
            if self.plots['pdf'] and self.plots['wind']:
                j = int(ceil(max(self.ly['wind'])))
                figp = plt.figure('pdf')
                plt.grid(True)
                px = figp.add_subplot(111)
                plt.title(self.hdrs['pdf'] + ' - ' + locn)
                px.hist(self.ly['wind'], j)
                px.set_ylabel('Number of occurences')
                px.set_xlabel('Wind Speed (m/s)')
            if self.plots['block']:
                plt.show(block=True)
            else:
                plt.draw()
        if self.plots['monthly']:
            figt = plt.figure('monthly')
            plt.grid(True)
            tx = figt.add_subplot(111)
            plt.title(self.hdrs['monthly'] + ' - ' + locn)
            maxy = 0
            i = -1
            if self.two_axes:
                tx2 = tx.twinx()
            for key, value in iter(sorted(self.ly.iteritems())):
                i += 1
                lw = 2.0
                if self.two_axes and key in self.ylabel2[0]:
                    tx2.step(x24[:14], t12[i], linewidth=lw, label=key, color=self.colours[key])
                else:
                    tx.step(x24[:14], t12[i], linewidth=lw, label=key, color=self.colours[key])
                    maxy = max(maxy, max(t12[i]) + 1)
            tx.set_ylim([0, maxy])
            tick_spot = []
            for i in range(12):
                tick_spot.append(i + 1.5)
            tx.set_xticks(tick_spot)
            tx.set_xticklabels(mth_labels)
            tx.set_xlabel('Month')
            tx.set_ylabel(self.ylabel[0])
            tx.legend(loc='best')
            if self.two_axes:
                ylim = tx2.get_ylim()
                tx2.set_ylim(0, ylim[1])
                lines = []
                for key in self.ly:
                    lines.append(mlines.Line2D([], [], color=self.colours[key], label=key))
                tx.legend(handles=lines, loc='best')
                tx2.set_ylabel(self.ylabel2[1])
            if self.maximise:
                mng = plt.get_current_fig_manager()
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if plt.get_backend() == 'TkAgg':
                        mng.window.state('zoomed')
                    elif plt.get_backend() == 'Qt4Agg':
                        mng.window.showMaximized()
                else:
                    mng.resize(*mng.window.maxsize())
            if self.plots['block']:
                plt.show(block=True)
            else:
                plt.draw()
        if not self.plots['block']:
            plt.show()

    def __init__(self, latitude, longitude, year=None, adjust_wind=None):
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        if year is None:
            try:
                self.base_year = config.get('Base', 'year')
                self.base_year = str(self.base_year)
            except:
                self.base_year = '2012'
        else:
            self.base_year = year
        parents = []
        try:
            parents = config.items('Parents')
        except:
            pass
        try:
            self.rain_files = config.get('Files', 'rain_files')
            for key, value in parents:
                self.rain_files = self.rain_files.replace(key, value)
            self.rain_files = self.rain_files.replace('$USER$', getUser())
            self.rain_files = self.rain_files.replace('$YEAR$', str(self.base_year))
        except:
            self.rain_files = ''
        try:
            self.solar_files = config.get('Files', 'solar_files')
            for key, value in parents:
                self.solar_files = self.solar_files.replace(key, value)
            self.solar_files = self.solar_files.replace('$USER$', getUser())
            self.solar_files = self.solar_files.replace('$YEAR$', str(self.base_year))
        except:
            self.solar_files = ''
        try:
            self.solar_index = config.get('Files', 'solar_index')
            for key, value in parents:
                self.solar_index = self.solar_index.replace(key, value)
            self.solar_index = self.solar_index.replace('$USER$', getUser())
            self.solar_index = self.solar_index.replace('$YEAR$', str(self.base_year))
        except:
            self.solar_index = ''
        try:
            self.wind_files = config.get('Files', 'wind_files')
            for key, value in parents:
                self.wind_files = self.wind_files.replace(key, value)
            self.wind_files = self.wind_files.replace('$USER$', getUser())
            self.wind_files = self.wind_files.replace('$YEAR$', str(self.base_year))
        except:
            self.wind_files = ''
        try:
            self.wind_index = config.get('Files', 'wind_index')
            for key, value in parents:
                self.wind_index = self.wind_index.replace(key, value)
            self.wind_index = self.wind_index.replace('$USER$', getUser())

            self.wind_index = self.wind_index.replace('$YEAR$', str(self.base_year))
        except:
            self.wind_index = ''
        rain = False
        try:
            variable = config.get('View', 'resource_rainfall')
            if variable.lower() in ['true', 'yes', 'on']:
                rain = True
        except:
            pass
        self.windy = adjust_wind
       # find closest solar file
        self.solar_file, dist, lat, lon = self.find_closest(latitude, longitude)
        if os.path.exists(self.solar_files + os.sep + self.solar_file):
            comment = 'Solar: %s\n            at %s, %s (%s Km away)' % (self.solar_file, lat, lon, '{:0,.0f}'.format(dist))
        self.wind_file, dist, lat, lon = self.find_closest(latitude, longitude, wind=True)
        if os.path.exists(self.wind_files + os.sep + self.wind_file):
            if comment != '':
                comment += '\n'
            comment += 'Wind: %s\n            at %s, %s (%s Km away)' % (self.wind_file, lat, lon, '{:0,.0f}'.format(dist))
        plot_order = ['show_menu', 'dhi', 'dni', 'ghi', 'temp', 'wind']
        if rain:
            plot_order.append('rain')
        plot_order2 = ['hour', 'total', 'month', 'season', 'period', 'block']  # , 'pdf']
        for plt in plot_order2:
            plot_order.append(plt)
        if rain:
            plot_order.append('monthly')
        self.hdrs = {'show_menu': 'Check / Uncheck all',
                'dhi': 'Solar - DHI (Diffuse)',
                'dni': 'Solar - DNI (Beam)',
                'ghi': 'Solar - GHI (Global)',
                'temp': 'Solar - Temperature',
                'wind': 'Wind - Speed',
                'hour': 'Hourly weather',
                'total': 'Daily average',
                'month': 'Daily average by month',
                'season': 'Daily average by season',
                'period': 'Daily average by Winter-Summer',
                'block': 'Show plots one at a time',
                'pdf': 'Probability Density Function'}
        if rain:
             self.hdrs['rain'] = 'Rainfall - mm'
             self.hdrs['monthly'] = 'Monthly Totals'
        spacers = {'dhi': 'Weather values',
                   'hour': 'Choose plots (all use a full year of data)'}
        self.plots = {}
        for i in range(len(plot_order)):
            self.plots[plot_order[i]] = False
        self.plots['total'] = True
        what_plots = whatPlots(self.plots, plot_order, self.hdrs, spacers, self.base_year, comment)
        what_plots.exec_()
        self.plots = what_plots.getValues()
        if self.plots is None:
            return
        if 'rain' not in self.plots.keys():
            self.plots['monthly'] = False
            self.plots['rain'] = False
        self.x = []
        self.ly = {}
        self.text = ''
        rain_col = -1
        if self.plots['dhi'] or self.plots['dni'] or self.plots['ghi'] or self.plots['temp'] or self.plots['rain']:
            if os.path.exists(self.solar_files + os.sep + self.solar_file):
                tf = open(self.solar_files + os.sep + self.solar_file, 'r')
                lines = tf.readlines()
                tf.close()
                fst_row = len(lines) - 8760
                if self.plots['dhi']:
                    self.ly['dhi'] = []
                if self.plots['dni']:
                    self.ly['dni'] = []
                if self.plots['ghi']:
                    self.ly['ghi'] = []
                if self.plots['temp']:
                    self.ly['temp'] = []
                if self.plots['wind']:  # on the off chance there's no wind file we'll use what we can from solar
                    self.ly['wind'] = []
                    wind_col = -1
                if self.plots['rain']:
                    self.ly['rain'] = []
                if self.solar_file[-4:] == '.smw':
                    dhi_col = 9
                    dni_col = 8
                    ghi_col = 7
                    temp_col = 0
                elif self.solar_file[-10:] == '(TMY2).csv' or self.solar_file[-10:] == '(TMY3).csv' \
                  or self.solar_file[-10:] == '(INTL).csv' or self.solar_file[-4:] == '.csv':
                    ghi_col = -1
                    if fst_row < 3:
                        bits = lines[0].split(',')
                        src_lat = float(bits[4])
                        src_lon = float(bits[5])
                        src_zne = float(bits[6])
                    else:
                        cols = lines[fst_row - 3].split(',')
                        bits = lines[fst_row - 2].split(',')
                        for i in range(len(cols)):
                            if cols[i].lower() in ['latitude', 'lat']:
                                src_lat = float(bits[i])
                            elif cols[i].lower() in ['longitude', 'lon', 'long', 'lng']:
                                src_lon = float(bits[i])
                            elif cols[i].lower() in ['tz', 'timezone', 'time zone']:
                                src_zne = float(bits[i])
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
                            temp_col = i
                        elif cols[i].lower() in ['wspd', 'wind speed']:
                            wind_col = i
                        elif cols[i].lower() in ['rain', 'rainfall', 'rainfall (mm)']:
                            rain_col = i
                for i in range(fst_row, len(lines)):
                    bits = lines[i].split(',')
                    if self.plots['dhi']:
                        self.ly['dhi'].append(float(bits[dhi_col]))
                    if self.plots['dni']:
                        self.ly['dni'].append(float(bits[dni_col]))
                    if self.plots['ghi']:
                        if ghi_col < 0:
                            zenith = getZenith(i - fst_row + 1, src_lat, src_lon, src_zne)
                            ghi_val = int(self.ly['dni'][-1] * cos(radians(zenith)) + self.ly['dhi'][-1])
                            self.ly['ghi'].append(ghi_val)
                        else:
                            self.ly['ghi'].append(float(bits[ghi_col]))
                    if self.plots['temp']:
                        self.ly['temp'].append(float(bits[temp_col]))
                    if self.plots['wind'] and wind_col >= 0:
                        self.ly['wind'].append(float(bits[wind_col]))
                    if self.plots['rain'] and rain_col >= 0:
                        self.ly['rain'].append(float(bits[rain_col]))
            else:
                return
        if self.plots['wind']:
            if self.wind_file != '':
                if os.path.exists(self.wind_files + os.sep + self.wind_file):
                    tf = open(self.wind_files + os.sep + self.wind_file, 'r')
                    lines = tf.readlines()
                    tf.close()
                    fst_row = len(lines) - 8760
                    self.ly['wind'] = []  # we'll override and wind from the solar file
                    if self.windy is None:
                        pass
                    else:
                        self.ly['wind2'] = []
                    if self.wind_file[-4:] == '.srw':
                        units = lines[3].strip().split(',')
                        heights = lines[4].strip().split(',')
                        col = -1
                        for j in range(len(units)):
                            if units[j] == 'm/s':
                               if heights[j] == '50':
                                   col = j
                                   break
                        for i in range(fst_row, len(lines)):
                            bits = lines[i].split(',')
                            self.ly['wind'].append(float(bits[col]))
                            if self.windy is None:
                                pass
                            else:
                                self.ly['wind2'].append(float(bits[col]) * (self.windy[1] / self.windy[0]) ** 0.143)
                else:
                    return
        if self.plots['rain'] and rain_col < 0:
            if self.rain_files != '':
                self.rain_file, dist, lat, lon = self.find_closest(latitude, longitude, wind=True)
                if os.path.exists(self.rain_files + os.sep + self.rain_file):
                    if comment != '':
                        comment += '\n'
                    comment += 'Rain: %s\n            at %s, %s (%s Km away)' % (self.rain_file, lat, lon, '{:0,.0f}'.format(dist))
                    tf = open(self.rain_files + os.sep + self.rain_file, 'r')
                    lines = tf.readlines()
                    tf.close()
                    fst_row = len(lines) - 8760
                    self.ly['rain'] = []  # we'll override and wind from the solar file
                    cols = lines[fst_row - 1].strip().split(',')
                    for i in range(len(cols)):
                        if cols[i].lower() in ['rain', 'rainfall', 'rainfall (mm)']:
                            rain_col = i
                            break
                    for i in range(fst_row, len(lines)):
                        bits = lines[i].split(',')
                        self.ly['rain'].append(float(bits[rain_col]))
        len_x = 8760
        for i in range(len_x):
            self.x.append(i)
        self.colours = {'dhi': 'r', 'dni': 'y', 'ghi': 'orange', 'rain': 'c', 'temp': 'g', 'wind': 'b', 'wind2': 'black'}
        self.two_axes = False
        self.ylabel = ['Irradiance (W/m2)', 'Irrad (W/m2)']
        if (self.plots['dhi'] or self.plots['dni'] or self.plots['ghi']):
            if self.plots['temp']:
                self.two_axes = True
                self.ylabel2 = [['temp'], 'Temperature. (oC)', 'Temp. (oC)']
            if self.plots['wind']:
                if self.two_axes:
                    self.ylabel2 = [['temp', 'wind'], 'Wind (m/s) & Temp. (oC)', 'Wind & Temp.']
                else:
                    self.two_axes = True
                    self.ylabel2 = [['wind'], 'Wind Speed (m/s)', 'Wind (m/s)']
        elif self.plots['temp']:
            self.ylabel = ['Temperature. (oC)', 'Temp. (oC)']
            if self.plots['wind']:
                self.two_axes = True
                self.ylabel2 = [['wind'], 'Wind Speed (m/s)', 'Wind (m/s)']
        elif self.plots['wind']:
            self.ylabel = ['Wind Speed (m/s)', 'Wind (m/s)']
        elif self.plots['rain']:
            self.ylabel = ['Rainfall (mm)', 'Rain (mm)']
        self.showGraphs(self.ly, self.x, ' for location %s, %s - %s' % (latitude, longitude, self.base_year))
