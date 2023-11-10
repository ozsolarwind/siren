#!/usr/bin/python3
#
#  Copyright (C) 2015-2023 Sustainable Energy Now Inc., Angus King
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
import matplotlib
if matplotlib.__version__ > '3.5.1':
    matplotlib.use('Qt5Agg')
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import matplotlib.lines as mlines
import os
import sys

import configparser  # decode .ini file
from PyQt5 import QtCore, QtGui, QtWidgets

import displaytable
from getmodels import getModelFile
from zoompan import ZoomPanX
from senutils import getParents, getUser, WorkBook
from sammodels import getZenith

the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]


class whatPlots(QtWidgets.QDialog):
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
        self.grid = QtWidgets.QGridLayout()
        self.checkbox = []
        i = 0
        bold = QtGui.QFont()
        bold.setBold(True)
        for plot in range(len(self.plot_order)):
            if self.plot_order[plot] in self.spacers:
                label = QtWidgets.QLabel(self.spacers[self.plot_order[plot]])
                label.setFont(bold)
                self.grid.addWidget(label, i, 0)
                i += 1
            self.checkbox.append(QtWidgets.QCheckBox(self.hdrs[self.plot_order[plot]], self))
            if self.plots[self.plot_order[plot]]:
                self.checkbox[plot].setCheckState(QtCore.Qt.Checked)
            self.grid.addWidget(self.checkbox[-1], i, 0)
            i += 1
        self.checkbox[0].stateChanged[int].connect(self.check_all)
        show = QtWidgets.QPushButton('Proceed', self)
        show.clicked.connect(self.showClicked)
        self.grid.addWidget(show, i, 0)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        commnt = QtWidgets.QLabel('Nearest weather files:\n' + self.comment)
        self.layout.addWidget(commnt)
        self.setWindowTitle('SIREN - Weather dialog for ' + str(self.base_year))
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        self.show_them = False
        self.show()

    def check_all(self):
        if self.checkbox[0].isChecked():
            for i in range(len(self.checkbox)):
                self.plots[self.plot_order[i]] = True
                self.checkbox[i].setCheckState(QtCore.Qt.Checked)
        else:
            for i in range(len(self.checkbox)):
                self.plots[self.plot_order[i]] = True
                self.checkbox[i].setCheckState(QtCore.Qt.Unchecked)

    def closeEvent(self, event):
        if not self.show_them:
            self.plots = None
        event.accept()

    def quitClicked(self):
        self.plots = None
        self.close()

    def showClicked(self):
        for plot in range(len(self.checkbox)):
            if self.checkbox[plot].checkState() == QtCore.Qt.Checked:
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
        lon1, lat1, lon2, lat2 = list(map(radians, [lon1, lat1, lon2, lat2]))

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
            filetype = ['.srw']
            technology = 'wind_index'
            index_file = self.wind_index
            folder = self.wind_files
        elif rain:
            filetype = ['.csv']
            technology = 'rain_index'
            index_file = ''
            folder = self.rain_files
        else:
            filetype = ['.csv', '.smw']
            technology = 'solar_index'
            index_file = self.solar_index
            folder = self.solar_files
        if folder != '':
            if index_file == '':
                fils = os.listdir(folder)
                for fil in fils:
                    if fil[-4:] in filetype:
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
                elif os.path.exists(folder + '/' + index_file):
                    ndx_file = folder + '/' + index_file
                if ndx_file != '':
                    w_file = WorkBook()
                    w_file.open_workbook(ndx_file)
                    worksheet = w_file.sheet_by_index(0)
                    num_rows = worksheet.nrows - 1
                    num_cols = worksheet.ncols - 1
                    var = {}
#                   get column names
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
        return closest, dist, close_lat, close_lon

    def showGraphs(self, ly, x, locn):
        def dayPlot(self, period, data, locn, per_labels=None, x_labels=None):
            plt.figure(period)
            plt.suptitle(self.hdrs[period] + locn, fontsize=16)
            maxy = 0
            maxw = 0
            if len(data[0]) > 9:
                p_y = 3
                p_x = 4
                xl = 8
                yl = [0, 4, 8]
                yl2 = [3, 7, 11]
            elif len(data[0]) > 6:
                p_y = 3
                p_x = 3
                xl = 6
                yl = [0, 3]
                yl2 = [2, 5]
            elif len(data[0]) > 4:
                p_y = 2
                p_x = 3
                xl = 3
                yl = [0, 3]
                yl2 = [2, 5]
            elif len(data[0]) > 2:
                p_y = 2
                p_x = 2
                xl = 2
                yl = [0, 2]
                yl2 = [1, 3]
            else:
                p_y = 1
                p_x = 2
                xl = 0
                yl = [0, 1]
                yl2 = [0, 1]
            wi = -1
            ra = -1
            te = -1
            i = -1
            for key, value in iter(sorted(self.ly.items())):
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
                px = plt.subplot(p_y, p_x, p + 1)
                i = -1
                if self.two_axes:
                    px2 = px.twinx()
                for key, value in iter(sorted(self.ly.items())):
                    i += 1
                    lw = 2.0
                    if self.two_axes and key in self.ylabel2[0]:
                        px2.plot(x24, data[i][p], linewidth=lw, label=self.labels[key], color=self.colours[key])
                    else:
                        px.plot(x24, data[i][p], linewidth=lw, label=self.labels[key], color=self.colours[key])
                    plt.title(per_labels[p])
                plt.xlim([1, 24])
                plt.xticks(list(range(4, 25, 4)))
       #         px.set_xticklabels(labels])
  #              plt.xticks(range(0, 25, 4))
           #     plt.grid(axis='x')
                px.set_xticklabels(x_labels[1:])
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
                    lines.append(mlines.Line2D([], [], color=self.colours[key], label=self.labels[key]))
                px.legend(handles=lines, loc='best')
            else:
                px.legend(loc='best')
            if self.maximise:
                mng = plt.get_current_fig_manager()
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if plt.get_backend() == 'TkAgg':
                        mng.window.state('zoomed')
                    elif plt.get_backend() == 'Qt5Agg':
                        mng.window.showMaximized()
                else:
                    mng.resize(*mng.window.maxsize())
            if self.plots['block']:
                plt.show(block=True)
            else:
                plt.draw()
        papersizes = {'a0': '33.1,46.8', 'a1': '23.4,33.1', 'a2': '16.5,23.4',
                    'a3': '11.7,16.5', 'a4': '8.3,11.7', 'a5': '5.8,8.3',
                    'a6': '4.1,5.8', 'a7': '2.9,4.1', 'a8': '2,2.9',
                    'a9': '1.5,2', 'a10': '1,1.5', 'b0': '39.4,55.7',
                    'b1': '27.8,39.4', 'b2': '19.7,27.8', 'b3': '13.9,19.7',
                    'b4': '9.8,13.9', 'b5': '6.9,9.8', 'b6': '4.9,6.9',
                    'b7': '3.5,4.9', 'b8': '2.4,3.5', 'b9': '1.7,2.4',
                    'b10': '1.2,1.7', 'foolscap': '8.0,13.0', 'ledger': '8.5,14.0',
                    'legal': '8.5,14.09', 'letter': '8.5,11.0'}
        landscape = False
        papersize = ''
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = getModelFile('SIREN.ini')
        config.read(config_file)
        self.maximise = False
        seasons = []
        periods = []
        parents = []
        try:
            parents = getParents(config.items('Parents'))
        except:
            pass
        try:
            self.scenarios = config.get('Files', 'scenarios')
            for key, value in parents:
                self.scenarios = self.scenarios.replace(key, value)
            self.scenarios = self.scenarios.replace('$USER$', getUser())
            self.scenarios = self.scenarios.replace('$YEAR$', str(self.base_year))
            i = self.scenarios.rfind('/')
            self.scenarios = self.scenarios[:i + 1]
        except:
            self.scenarios = ''
        try:
            items = config.items('Power')
            for item, values in items:
                if item[:6] == 'season':
                    if item == 'season':
                        continue
                    i = int(item[6:]) - 1
                    if i >= len(seasons):
                        seasons.append([])
                    seasons[i] = values.split(',')
                    for j in range(1, len(seasons[i])):
                        seasons[i][j] = int(seasons[i][j]) - 1
                elif item[:6] == 'period':
                    if item == 'period':
                        continue
                    i = int(item[6:]) - 1
                    if i >= len(periods):
                        periods.append([])
                    periods[i] = values.split(',')
                    for j in range(1, len(periods[i])):
                        periods[i][j] = int(periods[i][j]) - 1
                elif item == 'maximise':
                    if values.lower() in ['true', 'on', 'yes']:
                        self.maximise = True
                elif item == 'save_format':
                    plt.rcParams['savefig.format'] = values
                elif item == 'figsize':
                    try:
                        papersize = papersizes[values]
                    except:
                        papersize = values
                elif item == 'orientation':
                    if values.lower()[0] == 'l':
                        landscape = True
        except:
            pass
        if papersize != '':
            if landscape:
                bit = papersize.split(',')
                plt.rcParams['figure.figsize'] = bit[1] + ',' + bit[0]
            else:
                plt.rcParams['figure.figsize'] = papersize
        if len(seasons) == 0:
            seasons = [['Summer', 11, 0, 1], ['Autumn', 2, 3, 4], ['Winter', 5, 6, 7], ['Spring', 8, 9, 10]]
        if len(periods) == 0:
            periods = [['Winter', 4, 5, 6, 7, 8, 9], ['Summer', 10, 11, 0, 1, 2, 3]]
        for i in range(len(periods)):
            for j in range(len(seasons)):
                if periods[i][0] == seasons[j][0]:
                    periods[i][0] += '2'
                    break
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
                for q in range(len(seasons)):
                    q24[i].append([])
                    for j in range(24):
                        q24[i][q].append(0.)
            if self.plots['period']:
                s24.append([])
                for s in range(len(periods)):
                    s24[i].append([])
                    for j in range(24):
                        s24[i][s].append(0.)
            if self.plots['monthly'] or self.plots['mthavg']:
                t12.append([])
                for m in range(14):
                    t12[i].append(0.)
        the_qtrs = []
        for i in range(len(seasons)):
            d = 0
            for j in range(1, len(seasons[i])):
                d += the_days[seasons[i][j]]
            the_qtrs.append(d)
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
                for key, value in iter(sorted(self.ly.items())):
                    if len(value) == 0:
                        continue
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
                    if self.plots['monthly'] or self.plots['mthavg']:
                        t12[j][m + 1] = t12[j][m + 1] + value[i + k]
        for i in range(len(ly)):
            for k in range(24):
                if self.plots['total']:
                    l24[i][k] = l24[i][k] / 365
                if self.plots['month']:
                    for m in range(12):
                        m24[i][m][k] = m24[i][m][k] / the_days[m]
                if self.plots['season']:
                    for q in range(len(seasons)):
                        q24[i][q][k] = q24[i][q][k] / the_qtrs[q]
                if self.plots['period']:
                    for s in range(len(periods)):
                        s24[i][s][k] = s24[i][s][k] / the_ssns[s]
        if self.plots['hour']:
            if self.plots['save_plot']:
                vals = ['hour']
                data = []
                data.append(x)
            fig = plt.figure('hour')
            plt.grid(True)
            hx = plt.subplot(111)
            plt.title(self.hdrs['hour'] + ' - ' + locn)
            maxy = 0
            if self.two_axes:
                hx2 = hx.twinx()
            for key, value in iter(sorted(self.ly.items())):
                lw = 2.0
                if self.two_axes and key in self.ylabel2[0]:
                    hx2.plot(x, value, linewidth=lw, label=self.labels[key], color=self.colours[key])
                else:
                    hx.plot(x, value, linewidth=lw, label=self.labels[key], color=self.colours[key])
                    maxy = max(maxy, max(value))
                if self.plots['save_plot']:
                    vals.append(key)
                    data.append(value)
            if self.plots['save_plot']:
                titl = 'hour'
                dialog = displaytable.Table(list(map(list, list(zip(*data)))), title=titl, fields=vals, save_folder=self.scenarios)
                dialog.exec_()
                del dialog, data, vals
            hx.set_ylim([0, maxy])
            plt.xlim([0, len(x)])
            plt.xticks(list(range(12, len(x), 168)))
            hx.set_xticklabels(day_labels, rotation='vertical')
            hx.set_xlabel('Month of the year')
            hx.set_ylabel(self.ylabel[0])
            hx.legend(loc='best')
            if self.two_axes:
                ylim = hx2.get_ylim()
                hx2.set_ylim(0, ylim[1])
                lines = []
                for key in self.ly:
                    lines.append(mlines.Line2D([], [], color=self.colours[key], label=self.labels[key]))
                hx.legend(handles=lines, loc='best')
                hx2.set_ylabel(self.ylabel2[1])
            if self.maximise:
                mng = plt.get_current_fig_manager()
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if plt.get_backend() == 'TkAgg':
                        mng.window.state('zoomed')
                    elif plt.get_backend() == 'Qt5Agg':
                        mng.window.showMaximized()
                else:
                    mng.resize(*mng.window.maxsize())
            zp = ZoomPanX()
            if self.two_axes:
                f = zp.zoom_pan(hx2, base_scale=1.2) # enable scrollable zoom
            else:
                f = zp.zoom_pan(hx, base_scale=1.2) # enable scrollable zoom
            if self.plots['block']:
                plt.show(block=True)
            else:
                plt.draw()
            del zp
        if self.plots['total']:
            if self.plots['save_plot']:
                vals = ['hour of the day']
                decpts = [0]
                data = []
                data.append(x24)
            figt = plt.figure('total')
            plt.grid(True)
            tx = plt.subplot(111)
            plt.title(self.hdrs['total'] + ' - ' + locn)
            maxy = 0
            i = -1
            lw = 2.0
            if self.two_axes:
                tx2 = tx.twinx()
            for key, value in iter(sorted(self.ly.items())):
                i += 1
                if self.two_axes and key in self.ylabel2[0]:
                    tx2.plot(x24, l24[i], linewidth=lw, label=self.labels[key], color=self.colours[key])
                else:
                    tx.plot(x24, l24[i], linewidth=lw, label=self.labels[key], color=self.colours[key])
                    maxy = max(maxy, max(l24[i]))
                if self.plots['save_plot']:
                    vals.append(key)
                    data.append(l24[i])
                    if key == 'wind':
                        decpts.append(2)
                    else:
                        decpts.append(1)
            if self.plots['save_plot']:
                titl = 'total'
                dialog = displaytable.Table(list(map(list, list(zip(*data)))), title=titl, fields=vals, save_folder=self.scenarios, decpts=decpts)
                dialog.exec_()
                del dialog, data, vals
            tx.set_ylim([0, maxy])
            plt.xlim([1, 24])
            plt.xticks(list(range(4, 25, 4)))
            tx.set_xticklabels(labels[1:])
            tx.set_xlabel('Hour of the Day')
            tx.set_ylabel(self.ylabel[0])
            tx.legend(loc='best')
            if self.two_axes:
                ylim = tx2.get_ylim()
                tx2.set_ylim(0, ylim[1])
                lines = []
                for key in self.ly:
                    lines.append(mlines.Line2D([], [], color=self.colours[key], label=self.labels[key]))
                tx.legend(handles=lines, loc='best')
                tx2.set_ylabel(self.ylabel2[1])
            if self.maximise:
                mng = plt.get_current_fig_manager()
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if plt.get_backend() == 'TkAgg':
                        mng.window.state('zoomed')
                    elif plt.get_backend() == 'Qt5Agg':
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
        if 'pdf' in list(self.plots.keys()):
            if self.plots['pdf'] and self.plots['wind']:
                j = int(ceil(max(self.ly['wind'])))
                figp = plt.figure('pdf')
                plt.grid(True)
                px = plt.subplot(111)
                plt.title(self.hdrs['pdf'] + ' - ' + locn)
                px.hist(self.ly['wind'], j)
                px.set_ylabel('Number of occurences')
                px.set_xlabel('Wind Speed (m/s)')
            if self.plots['block']:
                plt.show(block=True)
            else:
                plt.draw()
        if self.plots['monthly']:
            if self.plots['save_plot']:
                vals = ['monthly']
                data = []
                data.append(x24[:12])
                decpts = [0]
            figt = plt.figure('monthly')
            plt.grid(True)
            tx = plt.subplot(111)
            plt.title(self.hdrs['monthly'] + ' - ' + locn)
            maxy = 0
            i = -1
            if self.two_axes:
                tx2 = tx.twinx()
            for key, value in iter(sorted(self.ly.items())):
                i += 1
                lw = 2.0
                if self.two_axes and key in self.ylabel2[0]:
                    tx2.step(x24[:14], t12[i], linewidth=lw, label=self.labels[key], color=self.colours[key])
                else:
                    tx.step(x24[:14], t12[i], linewidth=lw, label=self.labels[key], color=self.colours[key])
                    maxy = max(maxy, max(t12[i]) + 1)
                if self.plots['save_plot']:
                    vals.append(key)
                    data.append(t12[i][1:-1])
                    if key == 'wind':
                        decpts.append(2)
                    else:
                        decpts.append(1)
            if self.plots['save_plot']:
                titl = 'monthly'
                dialog = displaytable.Table(list(map(list, list(zip(*data)))), title=titl, fields=vals, save_folder=self.scenarios, decpts=decpts)
                dialog.exec_()
                del dialog, data, vals
            tx.set_ylim([0, maxy])
            tick_spot = []
            for i in range(12):
                tick_spot.append(i + 1.5)
            tx.set_xticks(tick_spot)
            tx.set_xticklabels(mth_labels)
            plt.xlim([1, 13.0])
            tx.set_xlabel('Month')
            tx.set_ylabel(self.ylabel[0])
            tx.legend(loc='best')
            if self.two_axes:
                ylim = tx2.get_ylim()
                tx2.set_ylim(0, ylim[1])
                lines = []
                for key in self.ly:
                    lines.append(mlines.Line2D([], [], color=self.colours[key], label=self.labels[key]))
                tx.legend(handles=lines, loc='best')
                tx2.set_ylabel(self.ylabel2[1])
            if self.maximise:
                mng = plt.get_current_fig_manager()
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if plt.get_backend() == 'TkAgg':
                        mng.window.state('zoomed')
                    elif plt.get_backend() == 'Qt5Agg':
                        mng.window.showMaximized()
                else:
                    mng.resize(*mng.window.maxsize())
            if self.plots['block']:
                plt.show(block=True)
            else:
                plt.draw()
        if self.plots['mthavg']:
            if self.plots['save_plot']:
                vals = ['monthly_average']
                data = []
                data.append(x24[:12])
                decpts = [0]
            figt = plt.figure('monthly_average')
            plt.grid(True)
            tx = plt.subplot(111)
            plt.title(self.hdrs['mthavg'] + ' - ' + locn)
            maxy = 0
            if self.two_axes:
                tx2 = tx.twinx()
            m12 = []
            for i in range(len(t12)):
                m12.append([])
                for m in range(1, 13):
                    m12[-1].append(t12[i][m] / the_days[m - 1] / 24.)
            i = -1
            lw = 2.0
            for key, value in iter(sorted(self.ly.items())):
                i += 1
                if self.two_axes and key in self.ylabel2[0]:
                    tx2.plot(x24[:12], m12[i], linewidth=lw, label=self.labels[key], color=self.colours[key])
                else:
                    tx.plot(x24[:12], m12[i], linewidth=lw, label=self.labels[key], color=self.colours[key])
                    maxy = max(maxy, max(m12[i]) + 1)
                if self.plots['save_plot']:
                    vals.append(key)
                    data.append(m12[i])
                    if key == 'wind':
                        decpts.append(2)
                    else:
                        decpts.append(1)
            if self.plots['save_plot']:
                titl = 'mthavg'
                dialog = displaytable.Table(list(map(list, list(zip(*data)))), title=titl, fields=vals, save_folder=self.scenarios, decpts=decpts)
                dialog.exec_()
                del dialog, data, vals
            tx.set_ylim([0, maxy])
            plt.xlim([1, 12])
            plt.xticks(list(range(1, 13, 1)))
            tx.set_xticklabels(mth_labels)
            tx.set_xlabel('Month')
            tx.set_ylabel(self.ylabel[0])
            tx.legend(loc='best')
            if self.two_axes:
                ylim = tx2.get_ylim()
                tx2.set_ylim(0, ylim[1])
                lines = []
                for key in self.ly:
                    lines.append(mlines.Line2D([], [], color=self.colours[key], label=self.labels[key]))
                tx.legend(handles=lines, loc='best')
                tx2.set_ylabel(self.ylabel2[1])
            if self.maximise:
                mng = plt.get_current_fig_manager()
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if plt.get_backend() == 'TkAgg':
                        mng.window.state('zoomed')
                    elif plt.get_backend() == 'Qt5Agg':
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
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = getModelFile('SIREN.ini')
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
            parents = getParents(config.items('Parents'))
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
        if os.path.exists(self.solar_files + '/' + self.solar_file):
            comment = 'Solar: %s\n            at %s, %s (%s Km away)' % (self.solar_file, lat, lon, '{:0,.0f}'.format(dist))
        self.wind_file, dist, lat, lon = self.find_closest(latitude, longitude, wind=True)
        if os.path.exists(self.wind_files + '/' + self.wind_file):
            if comment != '':
                comment += '\n'
            comment += 'Wind: %s\n            at %s, %s (%s Km away)' % (self.wind_file, lat, lon, '{:0,.0f}'.format(dist))
        plot_order = ['show_menu', 'dhi', 'dni', 'ghi', 'temp', 'wind']
        if rain:
            plot_order.append('rain')
        plot_order2 = ['hour', 'total', 'month', 'season', 'period', 'mthavg', 'block'] #, 'pdf']
        for plt in plot_order2:
            plot_order.append(plt)
        if rain:
            plot_order.append('monthly')
        plot_order.append('save_plot')
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
                'period': 'Daily average by period',
                'mthavg': 'Monthly average',
                'block': 'Show charts one at a time',
                'pdf': 'Probability Density Function',
                'save_plot': 'Save chart data to a file'}
        self.labels = {'dhi': 'DHI', 'dni': 'DNI', 'ghi': 'GHI', 'rain': 'Rain', 'temp': 'Temp.', 'wind': 'Wind', 'wind2': 'Wind 2'}
        if rain:
             self.hdrs['rain'] = 'Rainfall - mm'
             self.hdrs['monthly'] = 'Monthly totals'
        spacers = {'dhi': 'Weather values',
                   'hour': 'Choose charts (all use a full year of data)',
                   'save_plot': 'Save chart data'}
        weathers = ['dhi', 'dni', 'ghi', 'rain', 'temp', 'wind']
        self.plots = {}
        for i in range(len(plot_order)):
            self.plots[plot_order[i]] = False
        self.plots['total'] = True
        what_plots = whatPlots(self.plots, plot_order, self.hdrs, spacers, self.base_year, comment)
        what_plots.exec_()
        self.plots = what_plots.getValues()
        if self.plots is None:
            return
        if 'rain' not in list(self.plots.keys()):
            self.plots['monthly'] = False
            self.plots['rain'] = False
        self.x = []
        self.ly = {}
        self.text = ''
        rain_col = -1
        if self.plots['dhi'] or self.plots['dni'] or self.plots['ghi'] or self.plots['temp'] or self.plots['rain']:
            if os.path.exists(self.solar_files + '/' + self.solar_file):
                tf = open(self.solar_files + '/' + self.solar_file, 'r')
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
                for key in weathers:
                    try:
                        if len(self.ly[key]) == 0:
                            self.ly.pop(key)
                    except:
                        pass
                if len(self.ly) == 0:
                    return
            else:
                return
        if self.plots['wind']:
            if self.wind_file != '':
                if os.path.exists(self.wind_files + '/' + self.wind_file):
                    tf = open(self.wind_files + '/' + self.wind_file, 'r')
                    lines = tf.readlines()
                    tf.close()
                    fst_row = len(lines) - 8760
                    self.ly['wind'] = []  # we'll override any wind from the solar file
                    if self.windy is None:
                        pass
                    else:
                        self.ly['wind2'] = []
                    if self.wind_file[-4:] == '.srw':
                        units = lines[3].strip().split(',')
                        heights = lines[4].strip().split(',')
                        cols = []
                        col2 = -1
                        col = -1
                        for j in range(len(units)):
                            if units[j] == 'm/s':
                               try:
                                   h = int(heights[j])
                                   cols.append([h, j])
                               except:
                                   pass
                        if len(cols) > 1:
                            cols.sort(key=lambda x: x[1], reverse=True)
                            if cols[0][0] > 50:
                                col = cols[1][1]
                                self.labels['wind'] = 'Wind {:d}m'.format(cols[1][0])
                                col2 = cols[0][1]
                                self.labels['wind2'] = 'Wind {:d}m'.format(cols[0][0])
                                self.ly['wind2'] = []
                            else:
                                col = cols[0][1]
                        else:
                            col = col[0][1]
                        for i in range(fst_row, len(lines)):
                            bits = lines[i].split(',')
                            self.ly['wind'].append(float(bits[col]))
                            if col2 > 0:
                                self.ly['wind2'].append(float(bits[col2]))
                            elif self.windy is None:
                                pass
                            else:
                                self.ly['wind2'].append(float(bits[col]) * (self.windy[1] / self.windy[0]) ** 0.143)
                else:
                    return
        if self.plots['rain'] and rain_col < 0:
            if self.rain_files != '':
                self.rain_file, dist, lat, lon = self.find_closest(latitude, longitude, wind=True)
                if os.path.exists(self.rain_files + '/' + self.rain_file):
                    if comment != '':
                        comment += '\n'
                    comment += 'Rain: %s\n            at %s, %s (%s Km away)' % (self.rain_file, lat, lon, '{:0,.0f}'.format(dist))
                    tf = open(self.rain_files + '/' + self.rain_file, 'r')
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
                    self.ylabel2 = [['temp', 'wind', 'wind2'], 'Wind (m/s) & Temp. (oC)', 'Wind & Temp.']
                else:
                    self.two_axes = True
                    self.ylabel2 = [['wind', 'wind2'], 'Wind Speed (m/s)', 'Wind (m/s)']
        elif self.plots['temp']:
            self.ylabel = ['Temperature. (oC)', 'Temp. (oC)']
            if self.plots['wind']:
                self.two_axes = True
                self.ylabel2 = [['wind', 'wind2'], 'Wind Speed (m/s)', 'Wind (m/s)']
        elif self.plots['wind']:
            self.ylabel = ['Wind Speed (m/s)', 'Wind (m/s)']
        elif self.plots['rain']:
            self.ylabel = ['Rainfall (mm)', 'Rain (mm)']
        self.showGraphs(self.ly, self.x, ' for location %s, %s - %s' % (latitude, longitude, self.base_year))
