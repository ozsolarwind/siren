#!/usr/bin/python
#
#  Copyright (C) 2015-2018 Sustainable Energy Now Inc., Angus King
#
#  powermodel.py - This file is part of SIREN.
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

from math import asin, ceil, cos, fabs, floor, log10, pow, radians, sin, sqrt
import pylab as plt
from matplotlib.font_manager import FontProperties
import numpy.linalg as linalg
from numpy import *
import csv
import openpyxl as oxl
import os
import sys
import ssc
import xlrd
import xlwt

import ConfigParser  # decode .ini file
from PyQt4 import Qt, QtGui, QtCore

from senuser import getUser
import displayobject
import displaytable
from editini import SaveIni
from grid import Grid
from sirenicons import Icons
# import Station
from turbine import Turbine
from visualise import Visualise

the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]


def split_array(array):
    arry = []
    varbl = array.replace('(', '[')
    varbl2 = varbl.replace(')', ']')
    varbl2 = varbl2.replace('[', '')
    varbl2 = varbl2.replace(']', '')
    if ',' in varbl2:
        bits = varbl2.split(',')
    else:
        bits = varbl2.split(';')
    if '.' in varbl:
        for bit in bits:
            arry.append(float(bit))
    else:
        for bit in bits:
            arry.append(int(bit))
    return arry


def split_matrix(matrix):
    mtrx = []
    varbl = matrix.replace('(', '[')
    varbl2 = varbl.replace(')', ']')
    varbl2 = varbl2.replace('[[', '')
    varbl2 = varbl2.replace(']]', '')
    if '],[' in varbl2:
        arrys = varbl2.split('],[')
    else:
        arrys = varbl2.split('][')
    for arr1 in arrys:
        arr2 = arr1.replace('[', '')
        arry = arr2.replace(']', '')
        mtrx.append([])
        if ',' in arry:
            bits = arry.split(',')
        else:
            bits = arry.split(';')
        if '.' in varbl:
            for bit in bits:
                mtrx[-1].append(float(bit))
        else:
            for bit in bits:
                mtrx[-1].append(int(bit))
    return mtrx


def the_date(year, h):
    mm = 0
    dy, hr = divmod(h, 24)
    dy += 1
    while dy > the_days[mm]:
        dy -= the_days[mm]
        mm += 1
    return '%s-%s-%s %s:00' % (year, str(mm + 1).zfill(2), str(dy).zfill(2), str(hr).zfill(2))


class PowerSummary:
    def __init__(self, name, technology, generation, capacity, transmitted=None):
        self.name = name
        self.technology = technology
        self.generation = int(round(generation))
        self.capacity = capacity
        try:
            self.cf = round(generation / (capacity * 8760), 2)
        except:
            pass
        if transmitted is not None:
            self.transmitted = int(round(transmitted))
        else:
            self.transmitted = None


class ColumnData:
    def __init__(self, hour, period, value, values=None):
        self.hour = hour
        self.period = period
        if isinstance(value, list):
            for i in range(len(value)):
                if values is not None:
                    setattr(self, values[i], round(value[i], 2))
                else:
                    setattr(self, 'value' + str(i + 1), round(value[i], 2))
        else:
            if values is not None:
                setattr(self, values, round(value, 2))
            else:
                setattr(self, 'value', round(value, 2))


class DailyData:
    def __init__(self, day, date, value, values=None):
        self.day = day
        self.date = date
        if isinstance(value, list):
            for i in range(len(value)):
                if values is not None:
                    setattr(self, values[i], round(value[i], 2))
                else:
                    setattr(self, 'value' + str(i + 1), round(value[i], 2))
        else:
            if values is not None:
                setattr(self, values, round(value, 2))
            else:
                setattr(self, 'value', round(value, 2))


class whatPlots(QtGui.QDialog):
    def __init__(self, plots, plot_order, hdrs, spacers, load_growth, base_year, load_year,
                 iterations, storage, discharge, recharge, initials=None, initial=False, helpfile=None):
        self.plots = plots
        self.plot_order = plot_order
        self.hdrs = hdrs
        self.spacers = spacers
        self.load_growth = load_growth * 100
        self.base_year = int(base_year)
        self.load_year = int(load_year)
        self.iterations = iterations
        self.storage = storage
        self.discharge = discharge
        self.recharge = recharge
        self.initial = initial
        self.helpfile = helpfile
        if self.initial:
            self.initials = None
        else:
            self.initials = initials
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
                if self.plot_order[plot] == 'save_plot':  # fudge to add in growth stuff
                    self.percentLabel = QtGui.QLabel('        Growth. Set annual '
                                                     + 'Load growth & target year')
                    self.percentSpin = QtGui.QDoubleSpinBox()
                    self.percentSpin.setDecimals(2)
                    self.percentSpin.setSingleStep(.1)
                    self.percentSpin.setRange(-100, 100)
                    self.percentSpin.setValue(self.load_growth)
                    self.counterSpin = QtGui.QSpinBox()
                    self.counterSpin.setRange(min(self.base_year, 2015), 2100)
                    self.totalOutput = QtGui.QLabel('')
                    self.grid.addWidget(self.percentLabel, i, 0)
                    self.grid.addWidget(self.percentSpin, i, 1)
                    self.grid.addWidget(self.counterSpin, i, 2)
                    self.grid.addWidget(self.totalOutput, i, 3)
                    self.percentSpin.valueChanged[str].connect(self.growthChanged)
                    self.counterSpin.valueChanged[str].connect(self.growthChanged)
                    self.counterSpin.setValue(int(self.load_year))
                    i += 1
                    label = QtGui.QLabel('Storage')
                    label.setFont(bold)
                    self.grid.addWidget(label, i, 0)
                    i += 1
                    label = QtGui.QLabel('        Storage capacity (GWh) & initial value')
                    self.storageSpin = QtGui.QDoubleSpinBox()
                    self.storageSpin.setDecimals(3)
                    self.storageSpin.setRange(0, 500)
                    self.storageSpin.setValue(self.storage[0])
                    self.storageSpin.setSingleStep(5)
                    self.storpctSpin = QtGui.QDoubleSpinBox()
                    self.storpctSpin.setDecimals(3)
                    self.storpctSpin.setRange(0, 500)
                    self.storpctSpin.setValue(self.storage[1])
                    self.storpctSpin.setSingleStep(5)
                    self.grid.addWidget(label, i, 0)
                    self.grid.addWidget(self.storageSpin, i, 1)
                    self.grid.addWidget(self.storpctSpin, i, 2)
                    i += 1
                    label = QtGui.QLabel('        Discharge cap (MW) & loss (%)')
                    self.dischargeSpin = QtGui.QDoubleSpinBox()
                    self.dischargeSpin.setDecimals(2)
                    self.dischargeSpin.setRange(0, 50000)  # max is 10% of capacity
                    self.dischargeSpin.setValue(self.discharge[0])
                    self.dischargeSpin.setSingleStep(5)
                    self.dischargepctSpin = QtGui.QSpinBox()
                    self.dischargepctSpin.setRange(0, 50)
                    self.dischargepctSpin.setValue(int(100 - self.discharge[1] * 100))
                    self.grid.addWidget(label, i, 0)
                    self.grid.addWidget(self.dischargeSpin, i, 1)
                    self.grid.addWidget(self.dischargepctSpin, i, 2)
                    i += 1
                    label = QtGui.QLabel('        Recharge cap (MW) & loss (%)')
                    self.rechargeSpin = QtGui.QDoubleSpinBox()
                    self.rechargeSpin.setDecimals(2)
                    self.rechargeSpin.setRange(0, 50000)  # max is 10% of capacity
                    self.rechargeSpin.setValue(self.recharge[0])
                    self.rechargeSpin.setSingleStep(5)
                    self.rechargepctSpin = QtGui.QSpinBox()
                    self.rechargepctSpin.setRange(0, 50)
                    self.rechargepctSpin.setValue(int(100 - self.recharge[1] * 100))
                    self.grid.addWidget(label, i, 0)
                    self.grid.addWidget(self.rechargeSpin, i, 1)
                    self.grid.addWidget(self.rechargepctSpin, i, 2)
                    i += 1
                elif self.plot_order[plot] == 'summary':  # fudge to add in iterations stuff
                    self.iterLabel = QtGui.QLabel('        Shortfall. Choose analysis iterations')
                    self.iterSpin = QtGui.QSpinBox()
                    self.iterSpin.setRange(0, 3)
                    self.iterSpin.setValue(self.iterations)
                    self.grid.addWidget(self.iterLabel, i, 0)
                    self.grid.addWidget(self.iterSpin, i, 1)
                    i += 1
                label = QtGui.QLabel(self.spacers[self.plot_order[plot]])
                label.setFont(bold)
                self.grid.addWidget(label, i, 0)
                i += 1
            self.checkbox.append(QtGui.QCheckBox(self.hdrs[self.plot_order[plot]], self))
            if self.plots[self.plot_order[plot]]:
                self.checkbox[plot].setCheckState(QtCore.Qt.Checked)
            if not self.initial:
                if self.plot_order[plot] in self.initials:
                    self.checkbox[plot].setEnabled(False)
            self.grid.addWidget(self.checkbox[-1], i, 0)
            if self.plot_order[plot] == 'save_balance':
                self.grid.connect(self.checkbox[-1], QtCore.SIGNAL('stateChanged(int)'), self.check_balance)
            i += 1
        self.grid.connect(self.checkbox[0], QtCore.SIGNAL('stateChanged(int)'), self.check_all)
        show = QtGui.QPushButton('Proceed', self)
        show.clicked.connect(self.showClicked)
        self.grid.addWidget(show, i, 0)
        if self.initial:
            save = QtGui.QPushButton('Save Options', self)
            save.clicked.connect(self.saveClicked)
            self.grid.addWidget(save, i, 1)
        else:
            if self.plots['financials']:
                quitxt = 'Do Financials'
            else:
                quitxt = 'All Done'
            quit = QtGui.QPushButton(quitxt, self)
            quit.clicked.connect(self.quitClicked)
            self.grid.addWidget(quit, i, 1)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        #
        menubar = QtGui.QMenuBar()
        help = QtGui.QAction(QtGui.QIcon('help.png'), 'Help', self)
        help.setShortcut('F1')
        help.setStatusTip('Help')
        help.triggered.connect(self.showHelp)
        helpMenu = menubar.addMenu('&Help')
        helpMenu.addAction(help)
        self.layout.setMenuBar(menubar)
        #
        screen = QtGui.QDesktopWidget().availableGeometry()
        h = int(screen.height() * .9)
        self.resize(600, h)
        self.setWindowTitle('SIREN - Power dialog for ' + str(self.base_year))
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        self.show_them = False
        self.show()

    def showHelp(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.helpfile, title='Powermodel Help', section='popts')
        dialog.exec_()

    def growthChanged(self, val):
        summ = pow(1 + self.percentSpin.value() / 100, (self.counterSpin.value() - self.base_year))
        summ = '{:0.1f}%'.format((summ - 1) * 100)
        self.totalOutput.setText(summ)
        self.totalOutput.adjustSize()

    def check_all(self):
        if self.checkbox[0].isChecked():
            for i in range(len(self.checkbox)):
                if self.plot_order[i] == 'show_menu':
                    continue
                if not self.initial:
                    if self.plot_order[i] in self.initials:
                        continue
                self.checkbox[i].setCheckState(QtCore.Qt.Checked)
        else:
            for i in range(len(self.checkbox)):
                if self.plot_order[i] == 'show_menu':
                    continue
                if not self.initial:
                    if self.plot_order[i] in self.initials:
                        continue
                self.checkbox[i].setCheckState(QtCore.Qt.Unchecked)

    def check_balance(self, event):
        if event:
            for i in range(len(self.checkbox)):
                if self.plot_order[i] == 'show_load' or self.plot_order[i] == 'grid_losses':
                    self.checkbox[i].setCheckState(QtCore.Qt.Checked)

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
        try:
            self.iterations = self.iterSpin.value()
        except:
            self.iterations = 0
        self.show_them = True
        self.close()

    def saveClicked(self):
        for plot in range(len(self.checkbox)):
            if self.checkbox[plot].checkState() == QtCore.Qt.Checked:
                self.plots[self.plot_order[plot]] = True
            else:
                self.plots[self.plot_order[plot]] = False
        updates = {}
        store_lines = []
        store_lines.append('storage=%s,%s' % (str(self.storageSpin.value()),
                           str(self.storpctSpin.value())))
        store_lines.append('discharge_max=%s' % str(self.dischargeSpin.value()))
        store_lines.append('discharge_eff=%s' % str(self.dischargepctSpin.value() / 100.))
        store_lines.append('recharge_max=%s' % str(self.rechargeSpin.value()))
        store_lines.append('recharge_eff=%s' % str(self.rechargepctSpin.value() / 100.))
        updates['Storage'] = store_lines
        power_lines = []
        for key in self.plots:
            if key == 'show_menu':
                continue
            if self.plots[key]:
                power_lines.append(key + '=True')
            else:
                power_lines.append(key + '=False')
        power_lines.append('load_growth=%s%%' % str(self.percentSpin.value()))
        power_lines.append('shortfall_iterations=%s' % str(self.iterSpin.value()))
        updates['Power'] = power_lines
        SaveIni(updates)

    def getValues(self):
        load_multiplier = pow(1 + self.percentSpin.value() / 100,
                              (self.counterSpin.value() - self.base_year)) - 1
        return self.plots, self.percentSpin.value() / 100, str(self.counterSpin.value()), \
               load_multiplier, \
               self.iterations, [self.storageSpin.value(), self.storpctSpin.value()], \
               [self.dischargeSpin.value(), (1. - self.dischargepctSpin.value() / 100.)], \
               [self.rechargeSpin.value(), (1. - self.rechargepctSpin.value() / 100.)]


class whatStations(QtGui.QDialog):
    def __init__(self, stations, gross_load=False, actual=False, helpfile=None):
        self.stations = stations
        self.gross_load = gross_load
        self.actual = actual
        super(whatStations, self).__init__()
        self.initUI()

    def initUI(self):
        self.chosen = []
        self.grid = QtGui.QGridLayout()
        self.checkbox = []
        self.checkbox.append(QtGui.QCheckBox('Check / Uncheck all', self))
        self.grid.addWidget(self.checkbox[-1], 0, 0)
        i = 0
        c = 0
        icons = Icons()
        for stn in sorted(self.stations, key=lambda station: station.name):
            if stn.technology[:6] == 'Fossil' and not self.actual:
                continue
            if stn.technology == 'Rooftop PV' and stn.scenario == 'Existing' and not self.gross_load:
                continue
            self.checkbox.append(QtGui.QCheckBox(stn.name, self))
            icon = icons.getIcon(stn.technology)
            if icon != '':
                self.checkbox[-1].setIcon(QtGui.QIcon(icon))
            i += 1
            self.grid.addWidget(self.checkbox[-1], i, c)
            if i > 25:
                i = 0
                c += 1
        self.grid.connect(self.checkbox[0], QtCore.SIGNAL('stateChanged(int)'), self.check_all)
        show = QtGui.QPushButton('Choose', self)
        self.grid.addWidget(show, i + 1, c)
        show.clicked.connect(self.showClicked)
        self.setLayout(self.grid)
        self.setWindowTitle('SIREN - Power Stations dialog')
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        self.show_them = False
        self.show()

    def check_all(self):
        if self.checkbox[0].isChecked():
            for i in range(len(self.checkbox)):
                self.checkbox[i].setCheckState(QtCore.Qt.Checked)
        else:
            for i in range(len(self.checkbox)):
                self.checkbox[i].setCheckState(QtCore.Qt.Unchecked)

    def closeEvent(self, event):
        if not self.show_them:
            self.chosen = None
        event.accept()

    def quitClicked(self):
        self.close()

    def showClicked(self):
        for stn in range(1, len(self.checkbox)):
            if self.checkbox[stn].checkState() == QtCore.Qt.Checked:
                self.chosen.append(str(self.checkbox[stn].text()))
        self.show_them = True
        self.close()

    def getValues(self):
        return self.chosen


class whatFinancials(QtGui.QDialog):
    def __init__(self, helpfile=None):
        super(whatFinancials, self).__init__()
        self.proceed = False
        self.helpfile = helpfile
        self.financials = [['analysis_period', 'Analyis period', 0, 50, 30],
                      ['federal_tax_rate', 'Federal tax rate', 0, 30., 30.],
                      ['real_discount_rate', 'Real discount rate', 0, 20., 0],
                      ['inflation_rate', 'Inflation rate', 0, 20., 0],
                      ['insurance_rate', 'Insurance rate', 0, 15., 0],
                      ['loan_term', 'Loan term', 0, 60, 0],
                      ['loan_rate', 'Loan rate', 0, 30., 0],
                      ['debt_fraction', 'Debt percentage', 0, 100, 0],
                      ['depr_fed_type', 'Federal depreciation type 2=straight line', 0, 2, 2],
                      ['depr_fed_sl_years', 'Federal depreciation straight-line Years', 0, 60, 20],
                      ['market', 'Commercial PPA (on), Utility IPP (off)', 0, 1, 0],
                      ['bid_price_esc', 'Bid Price escalation', 0, 100, 0],
                      ['salvage_percentage', 'Salvage value percentage', 0, 100, 0],
                      ['min_dscr_target', 'Minimum required DSCR', 0, 5., 1.4],
                      ['min_irr_target', 'Minimum required IRR', 0, 30., 15],
                      ['ppa_escalation', 'PPA escalation', 0, 100., 0.6],
                      ['min_dscr_required', 'Minimum DSCR required', 0, 1, 1],
                      ['positive_cashflow_required', 'Positive cash flow required', 0, 1, 1],
                      ['optimize_lcoe_wrt_debt_fraction', 'Optimize LCOE with respect to debt' +
                       ' percent', 0, 1, 0],
                      ['optimize_lcoe_wrt_ppa_escalation', 'Optimize LCOE with respect to PPA' +
                       ' escalation', 0, 1, 0],
                      ['grid_losses', 'Reduce power by Grid losses', False, True, False],
                      ['grid_costs', 'Include Grid costs in LCOE', False, True, False],
                      ['grid_path_costs', 'Include full grid path in LCOE', False, True, False]]
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        beans = []
        try:
            beans = config.items('Financial')
            for key, value in beans:
                for i in range(len(self.financials)):
                    if key == self.financials[i][0]:
                        if value[-1] == '%':
                            self.financials[i][4] = float(value[:-1])
                        elif '.' in value:
                            self.financials[i][4] = float(value)
                        elif isinstance(self.financials[i][4], bool):
                            if value == 'True':
                                self.financials[i][4] = True
                            else:
                                self.financials[i][4] = False
                        else:
                            self.financials[i][4] = int(value)
                        break
        except:
            pass
        self.grid = QtGui.QGridLayout()
        self.labels = []
        self.spin = []
        i = 0
        for item in self.financials:
            if isinstance(item[2], bool) or isinstance(item[3], bool):
                continue
            if isinstance(item[2], int) or isinstance(item[3], int):
                if item[2] == 0 and item[3] == 1:
                    continue
            self.labels.append(QtGui.QLabel(item[1]))
            if isinstance(item[2], float) or isinstance(item[3], float):
                self.spin.append(QtGui.QDoubleSpinBox())
                self.spin[-1].setDecimals(1)
                self.spin[-1].setSingleStep(.1)
            else:
                self.spin.append(QtGui.QSpinBox())
            self.spin[-1].setRange(item[2], item[3])
            self.spin[-1].setValue(item[4])
            self.grid.addWidget(self.labels[-1], i, 0)
            self.grid.addWidget(self.spin[-1], i, 1)
            i += 1
        self.checkbox = []
        for item in self.financials:
            if isinstance(item[2], bool) and isinstance(item[3], bool):
                self.checkbox.append(QtGui.QCheckBox(item[1], self))
                self.grid.addWidget(self.checkbox[-1], i, 0)
                if item[4]:
                    self.checkbox[-1].setCheckState(QtCore.Qt.Checked)
                i += 1
            elif isinstance(item[2], int) and isinstance(item[3], int):
                if item[2] == 0 and item[3] == 1:
                    self.checkbox.append(QtGui.QCheckBox(item[1], self))
                    self.grid.addWidget(self.checkbox[-1], i, 0)
                    if item[4] == 1:
                        self.checkbox[-1].setCheckState(QtCore.Qt.Checked)
                    i += 1
        show = QtGui.QPushButton('Proceed', self)
        show.clicked.connect(self.showClicked)
        save = QtGui.QPushButton('Save Options', self)
        save.clicked.connect(self.saveClicked)
        quit = QtGui.QPushButton('All Done', self)
        quit.clicked.connect(self.quitClicked)
        self.grid.addWidget(show, i + 1, 0)
        self.grid.addWidget(save, i + 1, 1)
        self.grid.addWidget(quit, i + 1, 2)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        #
        menubar = QtGui.QMenuBar()
        help = QtGui.QAction(QtGui.QIcon('help.png'), 'Help', self)
        help.setShortcut('F1')
        help.setStatusTip('Help')
        help.triggered.connect(self.showHelp)
        helpMenu = menubar.addMenu('&Help')
        helpMenu.addAction(help)
        self.layout.setMenuBar(menubar)
        #
        screen = QtGui.QDesktopWidget().availableGeometry()
        h = int(screen.height() * .9)
        self.resize(600, h)
        self.setWindowTitle('SIREN - Financial Parameters dialog')
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        self.show()

    def showHelp(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.helpfile, title='Powermodel Help', section='fopts')
        dialog.exec_()

    def closeEvent(self, event):
        event.accept()

    def quitClicked(self):
        self.close()

    def showClicked(self):
        self.proceed = True
        self.close()

    def saveClicked(self):
        self.proceed = True
        values = self.getValues()
        finance_lines = []
        for key in values:
            if isinstance(values[key], bool):
                if values[key]:
                    finance_lines.append(key + '=True')
                else:
                    finance_lines.append(key + '=False')
            else:
                finance_lines.append(key + '=' + str(values[key]))
        beans = {'Financial': finance_lines}
        SaveIni(beans)
        self.proceed = False

    def getValues(self):
        if not self.proceed:
            return None
        values = {}
        for i in range(len(self.spin)):
            for item in self.financials:
                if item[1] == self.labels[i].text():
                    values[item[0]] = self.spin[i].value()
        for i in range(len(self.checkbox)):
            for item in self.financials:
                if item[1] == self.checkbox[i].text():
                    if isinstance(item[2], bool) and isinstance(item[3], bool):
                        if self.checkbox[i].checkState() == QtCore.Qt.Checked:
                            values[item[0]] = True
                        else:
                            values[item[0]] = False
                    else:
                        if self.checkbox[i].checkState() == QtCore.Qt.Checked:
                            values[item[0]] = 1
                        else:
                            values[item[0]] = 0
                    break
        return values


class Adjustments(QtGui.QDialog):
    def __init__(self, keys, load_key=None, load=None, data=None, base_year=None):
        super(Adjustments, self).__init__()
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        self.seasons = []
        self.periods = []
        self.daily = True
        self.adjust_cap = 25
        self.opt_load = False
        try:
            items = config.items('Power')
            for item, values in items:
                if item[:6] == 'season':
                    if item == 'season':
                        continue
                    i = int(item[6:]) - 1
                    if i >= len(self.seasons):
                        self.seasons.append([])
                    self.seasons[i] = values.split(',')
                    for j in range(1, len(self.seasons[i])):
                        self.seasons[i][j] = int(self.seasons[i][j]) - 1
                elif item[:6] == 'period':
                    if item == 'period':
                        continue
                    i = int(item[6:]) - 1
                    if i >= len(self.periods):
                        self.periods.append([])
                    self.periods[i] = values.split(',')
                    for j in range(1, len(self.periods[i])):
                        self.periods[i][j] = int(self.periods[i][j]) - 1
                elif item == 'optimise':
                    if values[0].lower() == 'h': # hourly
                        self.daily = False
                elif item == 'optimise_load':
                    if values.lower() in ['true', 'yes', 'on']:
                        self.opt_load = True
                elif item == 'adjust_cap':
                    self.adjust_cap = float(values)
        except:
            pass
        if len(self.seasons) == 0:
            self.seasons = [['Summer', 11, 0, 1], ['Autumn', 2, 3, 4], ['Winter', 5, 6, 7], ['Spring', 8, 9, 10]]
        if len(self.periods) == 0:
            self.periods = [['Winter', 4, 5, 6, 7, 8, 9], ['Summer', 10, 11, 0, 1, 2, 3]]
        for i in range(len(self.periods)):
            for j in range(len(self.seasons)):
                if self.periods[i][0] == self.seasons[j][0]:
                    self.periods[i][0] += '2'
                    break
        self.adjusts = {}
        self.checkbox = {}
        self.results = None
        self.grid = QtGui.QGridLayout()
        ctr = 0
        self.skeys = []
        self.lkey = load_key
        if len(keys) > 10:
            octr = 0
            ctr += 2
        else:
            octr = -1
        if load is not None:
            self.grid.addWidget(QtGui.QLabel('Check / Uncheck all zeroes'), ctr, 0)
            self.zerobox = QtGui.QCheckBox('', self)
            self.zerobox.setCheckState(QtCore.Qt.Unchecked)
            self.zerobox.stateChanged.connect(self.zeroCheck)
            self.grid.addWidget(self.zerobox, ctr, 2)
        ctr += 1
        for key in sorted(keys):
            if key == 'Generation':
                continue
            if key[:4] == 'Load':
                self.lkey = key
                continue
            self.adjusts[key] = QtGui.QDoubleSpinBox()
            self.adjusts[key].setRange(0, max(self.adjust_cap, self.adjusts[key].value()))
            if type(keys) is dict:
                self.adjusts[key].setValue(keys[key])
            else:
                self.adjusts[key].setValue(1.)
            self.adjusts[key].setDecimals(2)
            self.adjusts[key].setSingleStep(.1)
            self.skeys.append(key)
            self.grid.addWidget(QtGui.QLabel(key), ctr, 0)
            self.grid.addWidget(self.adjusts[key], ctr, 1)
            if load is not None:
                self.checkbox[key] = QtGui.QCheckBox('', self)
                self.checkbox[key].setCheckState(QtCore.Qt.Checked)
                self.grid.addWidget(self.checkbox[key], ctr, 2)
            ctr += 1
        if octr >= 0:
            ctr = 0
        if load is not None:
            for key in data.keys():
                if key[:4] == 'Load':
                    self.lkey = key
                    break
            self.load = load
            self.data = data
         # add pulldown list to optimise for year, month, none
            self.grid.addWidget(QtGui.QLabel('Optimise for:'), ctr, 0)
            self.periodCombo = QtGui.QComboBox()
            self.periodCombo.addItem('None')
            self.periodCombo.addItem(base_year)
            for i in range(1, 13):
                self.periodCombo.addItem(base_year + '-' + '{0:02d}'.format(i))
            self.l_m = self.periodCombo.count()
            for i in range(len(self.seasons)):
                self.periodCombo.addItem(base_year + '-' + self.seasons[i][0])
            self.l_s = self.periodCombo.count()
            for i in range(len(self.periods)):
                self.periodCombo.addItem(base_year + '-' + self.periods[i][0])
            self.l_p = self.periodCombo.count()
            self.grid.addWidget(self.periodCombo, ctr, 1)
            self.optmsg = QtGui.QLabel('')
            self.grid.addWidget(self.optmsg, ctr, 2)
            ctr += 1
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, ctr, 0)
        quit.clicked.connect(self.quitClicked)
        show = QtGui.QPushButton('Proceed', self)
        self.grid.addWidget(show, ctr, 1)
        show.clicked.connect(self.showClicked)
        if load is not None:
            optimise = QtGui.QPushButton('Optimise', self)
            self.grid.addWidget(optimise, ctr, 2)
            optimise.clicked.connect(self.optimiseClicked)
  #       self.setLayout(self.grid)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Gen Adj. multiplier')
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        self.show()

    def closeEvent(self, event):
        event.accept()

    def quitClicked(self):
        self.close()

    def zeroCheck(self):
        for key in self.skeys:
            if self.adjusts[key].value() == 0:
                self.checkbox[key].setCheckState(self.zerobox.checkState())

    def optimiseClicked(self):
        self.optmsg.setText('')
        try:
            if self.periodCombo.currentIndex() == 0:
                self.zerobox.setCheckState(QtCore.Qt.Unchecked)
                for key in self.skeys:
                    self.adjusts[key].setValue(1.)
                    self.checkbox[key].setCheckState(QtCore.Qt.Checked)
                return
            elif self.lkey is None:
                return
        except:
            return
        strt = [0]
        stop = [365]
        l = self.periodCombo.currentIndex()
        if l == 1:
            pass
        elif l < self.l_m:
            for i in range(l - 2):
                strt[0] += the_days[i]
            stop[0] = strt[0] + the_days[l - 2]
        elif l < self.l_s:
            i = l - self.l_m
            for j in range(len(self.seasons[i]) - 2):
                strt.append(0)
                stop.append(365)
            for j in range(1, len(self.seasons[i])):
                for k in range(self.seasons[i][j]):
                    strt[j - 1] += the_days[k]
                stop[j - 1] = strt[j - 1] + the_days[self.seasons[i][j]]
        else:
            i = l - self.l_s
            for j in range(len(self.periods[i]) - 2):
                strt.append(0)
                stop.append(365)
            for j in range(1, len(self.periods[i])):
                for k in range(self.periods[i][j]):
                    strt[j - 1] += the_days[k]
                stop[j - 1] = strt[j - 1] + the_days[self.periods[i][j]]
        load = []
        gen = []
        if self.daily:
            for i in range(24):
                load.append(0.)
            for i in range(24):
                gen.append([])
        okeys = []
        for key in self.skeys:
            if self.checkbox[key].isChecked():
                okeys.append(key)
                if self.daily:
                    for i in range(len(gen)):
                        gen[i].append(0.)
                else:
                    gen.append([])
            elif not self.opt_load:
                self.adjusts[key].setValue(0.)
        for m in range(len(strt)):
            strt[m] = strt[m] * 24
            stop[m] = stop[m] * 24
            if self.daily:
                h = 0
                for i in range(strt[m], stop[m]):
                    load[h] += self.load[i]
                    if h == 23:
                        h = 0
                    else:
                        h += 1
            else:
                for i in range(strt[m], stop[m]):
                    load.append(self.load[i])
            for key in self.skeys:
                if self.checkbox[key].isChecked():
                    n = okeys.index(key)
                    if self.daily:
                        h = 0
                        for i in range(strt[m], stop[m]):
                            gen[h][n] += self.data[key][i]
                            if h == 23:
                                h = 0
                            else:
                                h += 1
                    else:
                        for i in range(strt[m], stop[m]):
                            gen[n].append(self.data[key][i])
                elif self.opt_load:
                    if self.daily:
                        h = 0
                        for i in range(strt[m], stop[m]):
                            load[h] -= self.data[key][i] * self.adjusts[key].value()
                            if h == 23:
                                h = 0
                            else:
                                h += 1
                    else:
                        for i in range(strt[m], stop[m]):
                            load[i - strt[m]] -= self.data[key][i] * self.adjusts[key].value()
        if len(gen) == 0:
            self.optmsg.setText('None chosen')
            return
        B = array(load)
        if self.daily:
            pass
        else:
            gen2 = []
            for i in range(len(gen[0])):
                gen2.append([])
                for j in range(len(gen)):
                    gen2[-1].append(0.)
            for i in range(len(gen[0])):
                for j in range(len(gen)):
                    gen2[i][j] = gen[j][i]
            gen = gen2
        A = array(gen)
        res = linalg.lstsq(A, B)  # least squares solution of the generation vs load
        x = res[0]
        do_corr = True
        for i in range(len(okeys)):
            if x[i] < 0:
               self.optmsg.setText('-ve results')
               do_corr = False
            if x[i] > self.adjusts[okeys[i]].maximum():
                self.adjusts[okeys[i]].setMaximum(round(x[i], 2))
            self.adjusts[okeys[i]].setValue(round(x[i], 2))
        if do_corr:
            cgen = []
            for i in range(len(gen)):
                cgen.append(0.)
            for i in range(len(gen)):
                for j in range(len(gen[i])):
                    cgen[i] += gen[i][j] * x[j]
            corr = corrcoef(load, cgen)
            self.optmsg.setText('Corr: %.2f' % corr[0][1])
        self.zeroCheck()

    def showClicked(self):
        self.results = {}
        for key in self.adjusts.keys():
            self.results[key] = round(self.adjusts[key].value(), 2)
        self.close()

    def getValues(self):
        return self.results


class SuperPower():
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
            if index_file[-4:] == '.xls' or index_file[-5:] == '.xlsx':
                do_excel = True
            else:
                do_excel = False
            if self.default_files[technology] is None:
                dft_file = index_file
                if os.path.exists(dft_file):
                    pass
                else:
                    dft_file = folder + '/' + index_file
                if os.path.exists(dft_file):
                    if do_excel:
                        self.default_files[technology] = xlrd.open_workbook(dft_file)
                    else:
                        self.default_files[technology] = open(dft_file)
                else:
                    return closest
            else:
                if do_excel:
                    pass
                else:
                   self.default_files[technology].seek(0)
            if do_excel:
                var = {}
                worksheet = self.default_files[technology].sheet_by_index(0)
                num_rows = worksheet.nrows - 1
                num_cols = worksheet.ncols - 1
#               get column names
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
                dft_variables = csv.DictReader(self.default_files[technology])
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
        return closest

    def do_defaults(self, station):
        if 'PV' in station.technology:
            technology = 'PV'
        elif 'Wind' in station.technology:
            technology = 'Wind'
        else:
            technology = station.technology
        if self.defaults[technology][-4:] == '.xls' or \
          self.defaults[technology][-5:] == '.xlsx':
            do_excel = True
        else:
            do_excel = False
        if self.default_files[technology] is None:
            dft_file = self.variable_files + '/' + self.defaults[technology]
            if os.path.exists(dft_file):
                if do_excel:
                    self.default_files[technology] = xlrd.open_workbook(dft_file)
                else:
                    self.default_files[technology] = open(dft_file)
            else:
                return
        else:
            if do_excel:
                pass
            else:
                self.default_files[technology].seek(0)
        if do_excel:
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
                if (worksheet.cell_value(curr_row, var['TYPE']) == 'SSC_INPUT' or \
                  worksheet.cell_value(curr_row, var['TYPE']) == 'SSC_INOUT') and \
                  worksheet.cell_value(curr_row, var['DEFAULT']) != '' and \
                  str(worksheet.cell_value(curr_row, var['DEFAULT'])).lower() != 'input':
                    if worksheet.cell_value(curr_row, var['DATA']) == 'SSC_STRING':
                        self.data.set_string(worksheet.cell_value(curr_row, var['NAME']),
                          worksheet.cell_value(curr_row, var['DEFAULT']))
                    elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_ARRAY':
                        arry = split_array(worksheet.cell_value(curr_row, var['DEFAULT']))
                        self.data.set_array(worksheet.cell_value(curr_row, var['NAME']), arry)
                    elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_NUMBER':
                        if isinstance(worksheet.cell_value(curr_row, var['DEFAULT']), float):
                            self.data.set_number(worksheet.cell_value(curr_row, var['NAME']),
                              float(worksheet.cell_value(curr_row, var['DEFAULT'])))
                        else:
                            self.data.set_number(worksheet.cell_value(curr_row, var['NAME']),
                              worksheet.cell_value(curr_row, int(var['DEFAULT'])))
                    elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_MATRIX':
                        mtrx = split_matrix(worksheet.cell_value(curr_row, var['DEFAULT']))
                        self.data.set_matrix(worksheet.cell_value(curr_row, var['NAME']), mtrx)
        else:
            dft_variables = csv.DictReader(self.default_files[technology])
            for var in dft_variables:
                if (var['TYPE'] == 'SSC_INPUT' or var['TYPE'] == 'SSC_INOUT') and \
                  var['DEFAULT'] != '' and var['DEFAULT'].lower() != 'input':
                    if var['DATA'] == 'SSC_STRING':
                        self.data.set_string(var['NAME'], var['DEFAULT'])
                    elif var['DATA'] == 'SSC_ARRAY':
                        arry = split_array(var['DEFAULT'])
                        self.data.set_array(var['NAME'], arry)
                    elif var['DATA'] == 'SSC_NUMBER':
                        if var['DEFAULT'].find('.') >= 0:
                            self.data.set_number(var['NAME'], float(var['DEFAULT']))
                        else:
                            self.data.set_number(var['NAME'], int(var['DEFAULT']))
                    elif var['DATA'] == 'SSC_MATRIX':
                        mtrx = split_matrix(var['DEFAULT'])
                        self.data.set_matrix(var['NAME'], mtrx)

    def __init__(self, stations, plots, show_progress=None, parent=None, year=None, selected=None, status=None):
        self.stations = stations
        self.plots = plots
        self.show_progress = show_progress
        self.power_summary = []
        self.selected = selected
        self.status = status
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
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
            aparents = config.items('Parents')
            for key, value in aparents:
                for key2, value2 in aparents:
                    if key2 == key:
                        continue
                    value = value.replace(key2, value2)
                parents.append((key, value))
            del aparents
        except:
            pass
        try:
            self.biomass_multiplier = float(config.get('Biomass', 'multiplier'))
        except:
            self.biomass_multiplier = 8.25
        try:
            resource = config.get('Geothermal', 'resource')
            if resource.lower()[0:1] == 'hy':
                self.geores = 0
            else:
                self.geores = 1
        except:
            self.geores = 0
        self.dc_ac_ratio = [1.1] * 5
        try:
            self.dc_ac_ratio = [float(config.get('PV', 'dc_ac_ratio'))] * 5
        except:
            pass
        try:
            self.dc_ac_ratio[0] = float(config.get('Fixed PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.dc_ac_ratio[1] = float(config.get('Rooftop PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.dc_ac_ratio[2] = float(config.get('Single Axis PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.dc_ac_ratio[3] = float(config.get('Backtrack PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.dc_ac_ratio[4] = float(config.get('Tracking PV', 'dc_ac_ratio'))
        except:
            pass
        try:
            self.dc_ac_ratio[4] = float(config.get('Dual Axis PV', 'dc_ac_ratio'))
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
        self.turbine_spacing = [8, 8] # onshore and offshore winds
        self.row_spacing = [8, 8]
        self.offset_spacing = [4, 4]
        self.wind_farm_losses_percent = [2, 2]
        try:
            self.turbine_spacing[0] = int(config.get('Wind', 'turbine_spacing'))
        except:
            try:
                self.turbine_spacing[0] = int(config.get('Onshore Wind', 'turbine_spacing'))
            except:
                pass
        try:
            self.row_spacing[0] = int(config.get('Wind', 'row_spacing'))
        except:
            try:
                self.row_spacing[0] = int(config.get('Onshore Wind', 'row_spacing'))
            except:
                pass
        try:
            self.offset_spacing[0] = int(config.get('Wind', 'offset_spacing'))
        except:
            try:
                self.offset_spacing[0] = int(config.get('Onshore Wind', 'offset_spacing'))
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
            self.turbine_spacing[1] = int(config.get('Offshore Wind', 'turbine_spacing'))
        except:
            pass
        try:
            self.row_spacing[1] = int(config.get('Offshore Wind', 'row_spacing'))
        except:
            pass
        try:
            self.offset_spacing[1] = int(config.get('Offshore Wind', 'offset_spacing'))
        except:
            pass
        try:
            self.wind_farm_losses_percent[1] = int(config.get('Offshore Wind', 'wind_farm_losses_percent').strip('%'))
        except:
            pass
        self.gross_net = 0.87
        try:
            self.gross_net = float(config.get('Solar Thermal', 'gross_net'))
        except:
            pass
        self.tshours = 0
        try:
            self.tshours = float(config.get('Solar Thermal', 'tshours'))
        except:
            pass
        self.volume = 12.9858
        try:
            self.volume = float(config.get('Solar Thermal', 'volume'))
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
                    tec = tec.replace('_', ' ').title()
                    tec = tec.replace('Pv', 'PV')
                    self.defaults[tec] = default
                    self.default_files[tec] = None
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
                self.subs_loss = float(subs_loss[:-1]) / 100.
            else:
                self.subs_loss = float(subs_loss) / 10.
            line_loss = config.get('Grid', 'line_loss')
            if line_loss[-1] == '%':
                self.line_loss = float(line_loss[:-1]) / 100000.
            else:
                self.line_loss = float(line_loss) / 1000.
        except:
            self.subs_loss = 0.
            self.line_loss = 0.
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
        elif self.plots['save_tech'] or self.plots['save_balance']:
            self.stn_outs = []
            self.stn_tech = []
        elif self.plots['visualise']:
            self.stn_outs = []
            self.stn_pows = []
        len_x = 8760
        for i in range(len_x):
            self.x.append(i)
        if self.plots['grid_losses']:
            self.ly['Generation'] = []
            for i in range(len_x):
                self.ly['Generation'].append(0.)
     #   if self.plots['by_station']:
     #       self.plots['show_pct'] = False
        if not self.show_progress:
            self.getPowerLoop()
            self.getPowerDone()

    def getPowerLoop(self):
        self.all_done = False
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
            stn = self.stations[st]
            if stn.technology[:6] == 'Fossil' and not self.plots['actual']:
                try:
                    value = self.progressbar.value() + 1
                    self.progressbar.setValue(value)
                except:
                    pass
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
                if stn.technology == 'Rooftop PV' and stn.scenario == 'Existing':
                    key = 'Existing Rooftop PV'
                else:
                    key = stn.technology
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
            elif self.plots['save_tech'] or self.plots['save_balance']:
                self.stn_outs.append(stn.name)
                self.stn_tech.append(stn.technology)
            elif self.plots['visualise']:
                self.stn_outs.append(stn.name)
                self.stn_pows.append([])
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
                            enrgy = power[i] * (1 - self.line_loss * stn.grid_path_len - self.subs_loss)
                        else:
                            enrgy = power[i] * (1 - self.subs_loss)
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
            self.power_summary.append(pt)
            if self.show_progress:
                value = self.progressbar.value() + 1
                if value > self.progressbar.maximum():
                    self.all_done = True
                    break
                self.progressbar.setValue(value)
                QtGui.qApp.processEvents()
                if not self._active:
                    break
        else:
            self.all_done = True
        if self.show_progress:
            self.close()

    def getPowerDone(self):
        return

    def getStationPower(self, station):
        farmpwr = []
        if self.plots['actual'] and self.actual_power != '':
            if self.actual_power[-4:] == '.xls' or self.actual_power[-5:] == '.xlsx':
                do_excel = True
            else:
                do_excel = False
            if self.default_files['actual'] is None:
                if os.path.exists(self.scenarios + self.actual_power):
                    if do_excel:
                        self.default_files['actual'] = xlrd.open_workbook(self.scenarios +
                                                       self.actual_power)
                    else:
                        self.default_files['actual'] = open(self.scenarios + self.actual_power)
            if do_excel:
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
                if col < worksheet.ncols:
                    for i in range(1, worksheet.nrows):
                        farmpwr.append(worksheet.cell_value(i, col) * 1000.)
                return farmpwr
            else:
                self.default_files['actual'].seek(0)
                pwr_data = csv.DictReader(self.default_files[technology])
                try:
                    h = pwr_data.fieldnames.index('Hourly energy | (kWh)')
                except:
                    try:
                        h = pwr_data.fieldnames.index('Hourly Energy | (kW)')
                    except:
                        try:
                            h = pwr_data.fieldnames.index('Power generated by system | (kW)')
                        except:
                            h = pwr_data.fieldnames.index(station.name)
                for data in pwr_data:
                    farmpwr.append(float(data[pwr_data.fieldnames[h]]))
                return farmpwr
        elif station.power_file is not None and station.power_file != '':
            if os.path.exists(self.scenarios + station.power_file):
                if station.power_file[-4:] == '.csv':
                    csv_file = open(self.scenarios + station.power_file, 'r')
                    pwr_data = csv.DictReader(csv_file)
                    try:
                        h = pwr_data.fieldnames.index('Hourly energy | (kWh)')
                    except:
                        try:
                            h = pwr_data.fieldnames.index('Hourly Energy | (kW)')
                        except:
                            try:
                                h = pwr_data.fieldnames.index('Power generated by system | (kW)')
                            except:
                                h = pwr_data.fieldnames.index(station.name)
                    for data in pwr_data:
                        farmpwr.append(float(data[pwr_data.fieldnames[h]]))
                    return farmpwr
                else:
                    xl_file = xlrd.open_workbook(self.scenarios + station.power_file)
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
            self.status.emit(QtCore.SIGNAL('log'), 'Processing ' + station.name)
        self.data = None
        self.data = ssc.Data()
        if station.technology[-4:] == 'Wind':
            if station.technology[:3] == 'Off': # offshore?
                wtyp = 1
            else:
                wtyp = 0
            closest = self.find_closest(station.lat, station.lon, wind=True)
            self.data.set_string('wind_resource_filename', self.wind_files + '/' + closest)
            turbine = Turbine(station.turbine)
            no_turbines = int(station.no_turbines)
            if station.scenario == 'Existing' and (no_turbines * turbine.capacity) != (station.capacity * 1000):
                loss = round(1. - (station.capacity * 1000) / (no_turbines * turbine.capacity), 2)
                loss = loss * 100
                if loss < 0.:
                    loss = self.wind_farm_losses_percent[wtyp]
                self.data.set_number('system_capacity', station.capacity * 1000000)
                self.data.set_number('wind_farm_losses_percent', loss)
            else:
                self.data.set_number('system_capacity', no_turbines * turbine.capacity * 1000)
                self.data.set_number('wind_farm_losses_percent', self.wind_farm_losses_percent[wtyp])
            pc_wind = turbine.speeds
            pc_power = turbine.powers
            self.data.set_array('wind_turbine_powercurve_windspeeds', pc_wind)
            self.data.set_array('wind_turbine_powercurve_powerout', pc_power)
            t_rows = int(ceil(sqrt(no_turbines)))
            ctr = no_turbines
            wt_x = []
            wt_y = []
            for r in range(t_rows):
                for c in range(t_rows):
                    wt_x.append(r * self.row_spacing[wtyp] * turbine.rotor)
                    wt_y.append(c * self.turbine_spacing[wtyp] * turbine.rotor +
                                (r % 2) * self.offset_spacing[wtyp] * turbine.rotor)
                    ctr -= 1
                    if ctr < 1:
                        break
                if ctr < 1:
                    break
            self.data.set_array('wind_farm_xCoordinates', wt_x)
            self.data.set_array('wind_farm_yCoordinates', wt_y)
            self.data.set_number('wind_turbine_rotor_diameter', turbine.rotor)
            self.data.set_number('wind_turbine_cutin', turbine.cutin)
            self.do_defaults(station)
            module = ssc.Module('windpower')
            if (module.exec_(self.data)):
                farmpwr = self.data.get_array('gen')
                del module
                return farmpwr
            else:
                if self.status:
                   self.status.emit(QtCore.SIGNAL('log'), 'Errors encountered processing ' + station.name)
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    if self.status:
                       self.status.emit(QtCore.SIGNAL('log'), 'windpower error [' + str(idx) + ']: ' + msg)
                    else:
                        print 'windpower error [', idx, ' ]: ', msg
                    idx += 1
                    msg = module.log(idx)
                del module
                return None
        elif station.technology == 'Solar Thermal':
            closest = self.find_closest(station.lat, station.lon)
            base_capacity = 100
            self.data.set_string('solar_resource_file', self.solar_files + '/' + closest)
            self.data.set_number('system_capacity', base_capacity * 1000)
            self.data.set_number('P_ref', base_capacity / self.gross_net)
            if station.storage_hours is None:
                tshours = self.tshours
            else:
                tshours = station.storage_hours
            self.data.set_number('tshours', tshours)
            sched = [[1.] * 24] * 12
            self.data.set_matrix('weekday_schedule', sched[:])
            self.data.set_matrix('weekend_schedule', sched[:])
            if ssc.API().version() >= 159:
                self.data.set_matrix('dispatch_sched_weekday', sched[:])
                self.data.set_matrix('dispatch_sched_weekend', sched[:])
            else:
                self.data.set_number('Design_power', base_capacity / self.gross_net)
                self.data.set_number('W_pb_design', base_capacity / self.gross_net)
                vol_tank = base_capacity * tshours * self.volume
                self.data.set_number('vol_tank', vol_tank)
                f_tc_cold = self.data.get_number('f_tc_cold')
                V_tank_hot_ini = vol_tank * (1. - f_tc_cold)
                self.data.set_number('V_tank_hot_ini', V_tank_hot_ini)
            self.do_defaults(station)
            module = ssc.Module('tcsmolten_salt')
            if (module.exec_(self.data)):
                farmpwr = self.data.get_array('gen')
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    idx += 1
                    msg = module.log(idx)
                del module
                if station.capacity != base_capacity:
                    for i in range(len(farmpwr)):
                        farmpwr[i] = farmpwr[i] * station.capacity / float(base_capacity)
                return farmpwr
            else:
                if self.status:
                   self.status.emit(QtCore.SIGNAL('log'), 'Errors encountered processing ' + station.name)
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    if self.status:
                       self.status.emit(QtCore.SIGNAL('log'), 'tcsmolten_salt error [' + str(idx) + ']: ' + msg)
                    else:
                        print 'tcsmolten_salt error [', idx, ' ]: ', msg
                    idx += 1
                    msg = module.log(idx)
                del module
                return None
        elif 'PV' in station.technology:
            closest = self.find_closest(station.lat, station.lon)
            self.data.set_string('solar_resource_file', self.solar_files + '/' + closest)
            dc_ac_ratio = self.dc_ac_ratio[0]
            if station.technology[:5] == 'Fixed':
                dc_ac_ratio = self.dc_ac_ratio[0]
                self.data.set_number('array_type', 0)
            elif station.technology[:7] == 'Rooftop':
                dc_ac_ratio = self.dc_ac_ratio[1]
                self.data.set_number('array_type', 1)
            elif station.technology[:11] == 'Single Axis':
                dc_ac_ratio = self.dc_ac_ratio[2]
                self.data.set_number('array_type', 2)
            elif station.technology[:9] == 'Backtrack':
                dc_ac_ratio = self.dc_ac_ratio[3]
                self.data.set_number('array_type', 3)
            elif station.technology[:8] == 'Tracking' or station.technology[:9] == 'Dual Axis':
                dc_ac_ratio = self.dc_ac_ratio[4]
                self.data.set_number('array_type', 4)
            self.data.set_number('system_capacity', station.capacity * 1000 * dc_ac_ratio)
            self.data.set_number('dc_ac_ratio', dc_ac_ratio)
            try:
                self.data.set_number('tilt', fabs(station.tilt))
            except:
                self.data.set_number('tilt', fabs(station.lat))
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
            self.data.set_number('azimuth', azi)
            self.data.set_number('losses', self.pv_losses)
            self.do_defaults(station)
            module = ssc.Module('pvwattsv5')
            if (module.exec_(self.data)):
                farmpwr = self.data.get_array('gen')
                del module
                return farmpwr
            else:
                if self.status:
                   self.status.emit(QtCore.SIGNAL('log'), 'Errors encountered processing ' + station.name)
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    if self.status:
                       self.status.emit(QtCore.SIGNAL('log'), 'pvwattsv5 error [' + str(idx) + ']: ' + msg)
                    else:
                        print 'pvwattsv5 error [', idx, ' ]: ', msg
                    idx += 1
                    msg = module.log(idx)
                del module
                return None
        elif station.technology == 'Biomass':
            closest = self.find_closest(station.lat, station.lon)
            self.data.set_string('file_name', self.solar_files + '/' + closest)
            self.data.set_number('system_capacity', station.capacity * 1000)
            self.data.set_number('biopwr.plant.nameplate', station.capacity * 1000)
            feedstock = station.capacity * 1000 * self.biomass_multiplier
            self.data.set_number('biopwr.feedstock.total', feedstock)
            self.data.set_number('biopwr.feedstock.total_biomass', feedstock)
            carbon_pct = self.data.get_number('biopwr.feedstock.total_biomass_c')
            self.data.set_number('biopwr.feedstock.total_c', feedstock * carbon_pct / 100.)
            self.do_defaults(station)
            module = ssc.Module('biomass')
            if (module.exec_(self.data)):
                energy = self.data.get_number('annual_energy')
                farmpwr = self.data.get_array('gen')
                del module
                return farmpwr
            else:
                if self.status:
                   self.status.emit(QtCore.SIGNAL('log'), 'Errors encountered processing ' + station.name)
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    if self.status:
                       self.status.emit(QtCore.SIGNAL('log'), 'biomass error [' + str(idx) + ']: ' + msg)
                    else:
                        print 'biomass error [', idx, ' ]: ', msg
                    idx += 1
                    msg = module.log(idx)
                del module
                return None
        elif station.technology == 'Geothermal':
            closest = self.find_closest(station.lat, station.lon)
            self.data.set_string('file_name', self.solar_files + '/' + closest)
            self.data.set_number('nameplate', station.capacity * 1000)
            self.data.set_number('resource_potential', station.capacity * 10)
            self.data.set_number('resource_type', self.geores)
            self.data.set_string('hybrid_dispatch_schedule', '1' * 24 * 12)
            self.do_defaults(station)
            module = ssc.Module('geothermal')
            if (module.exec_(self.data)):
                energy = self.data.get_number('annual_energy')
                pwr = self.data.get_array('monthly_energy')
                for i in range(12):
                    for j in range(the_days[i] * 24):
                        farmpwr.append(pwr[i] / (the_days[i] * 24))
                del module
                return farmpwr
            else:
                if self.status:
                   self.status.emit(QtCore.SIGNAL('log'), 'Errors encountered processing ' + station.name)
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    if self.status:
                       self.status.emit(QtCore.SIGNAL('log'), 'geothermal error [' + str(idx) + ']: ' + msg)
                    else:
                        print 'geothermal error [', idx, ' ]: ', msg
                    idx += 1
                    msg = module.log(idx)
                del module
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
            config = ConfigParser.RawConfigParser()
            if len(sys.argv) > 1:
                config_file = sys.argv[1]
            else:
                config_file = 'SIREN.ini'
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
                    formulb = propty['formula'].lower().split(' ')
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
                            for key in propty.keys():
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

    def getVisual(self):
        return self.model.getVisual()

class FinancialSummary:
    def __init__(self, name, technology, capacity, generation, cf, capital_cost, lcoe_real,
                 lcoe_nominal, npv, grid_cost):
        self.name = name
        self.technology = technology
        self.capacity = capacity
        self.generation = int(round(generation))
        try:
            self.cf = round(generation / (capacity * 8760), 2)
        except:
            self.cf = 0.
        self.capital_cost = int(round(capital_cost))
        self.lcoe_real = round(lcoe_real, 2)
        self.lcoe_nominal = round(lcoe_nominal, 2)
        self.npv = int(round(npv))
        self.grid_cost = int(round(grid_cost))


class FinancialModel():
    def get_variables(self, xl_file, overrides=None):
        data = None
        data = ssc.Data()
        var = {}
        workfile = xlrd.open_workbook(xl_file)
        worksheet = workfile.sheet_by_index(0)
        num_rows = worksheet.nrows - 1
        num_cols = worksheet.ncols - 1
   # get column names
        curr_col = -1
        while curr_col < num_cols:
            curr_col += 1
            var[worksheet.cell_value(0, curr_col)] = curr_col
        curr_row = 0
        output_variables = []
        while curr_row < num_rows:
            curr_row += 1
            if worksheet.cell_value(curr_row, var['TYPE']) == 'SSC_INPUT' and \
              worksheet.cell_value(curr_row, var['DEFAULT']) != '' and \
              str(worksheet.cell_value(curr_row, var['DEFAULT'])).lower() != 'input':
                if worksheet.cell_value(curr_row, var['DATA']) == 'SSC_STRING':
                    data.set_string(worksheet.cell_value(curr_row, var['NAME']),
                    worksheet.cell_value(curr_row, var['DEFAULT']))
                elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_ARRAY':
                    arry = split_array(worksheet.cell_value(curr_row, var['DEFAULT']))
                    data.set_array(worksheet.cell_value(curr_row, var['NAME']), arry)
                elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_NUMBER':
                    if overrides is not None and worksheet.cell_value(curr_row, var['NAME']) \
                      in overrides:
                        data.set_number(worksheet.cell_value(curr_row, var['NAME']),
                          overrides[worksheet.cell_value(curr_row, var['NAME'])])
                    else:
                        if isinstance(worksheet.cell_value(curr_row, var['DEFAULT']), float):
                            data.set_number(worksheet.cell_value(curr_row, var['NAME']),
                              float(worksheet.cell_value(curr_row, var['DEFAULT'])))
                        else:
                            data.set_number(worksheet.cell_value(curr_row, var['NAME']),
                              worksheet.cell_value(curr_row, int(var['DEFAULT'])))
                elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_MATRIX':
                    mtrx = split_matrix(worksheet.cell_value(curr_row, var['DEFAULT']))
                    data.set_matrix(worksheet.cell_value(curr_row, var['NAME']), mtrx)
            elif worksheet.cell_value(curr_row, var['TYPE']) == 'SSC_OUTPUT':
                output_variables.append([worksheet.cell_value(curr_row, var['NAME']),
                                        worksheet.cell_value(curr_row, var['DATA'])])
        return data, output_variables

    def __init__(self, name, technology, capacity, power, grid, path, year=None, status=None):
        def set_grid_variables():
            self.dispatchable = None
            self.line_loss = 0.
            self.subs_cost = 0.
            self.subs_loss = 0.
            try:
                itm = config.get('Grid', 'dispatchable')
                itm = itm.replace('_', ' ').title()
                self.dispatchable = itm.replace('Pv', 'PV')
                line_loss = config.get('Grid', 'line_loss')
                if line_loss[-1] == '%':
                    self.line_loss = float(line_loss[:-1]) / 100000.
                else:
                    self.line_loss = float(line_loss) / 1000.
                line_loss = config.get('Grid', 'substation_loss')
                if line_loss[-1] == '%':
                    self.subs_loss = float(line_loss[:-1]) / 100.
                else:
                    self.subs_loss = float(line_loss)
            except:
                pass

        def stn_costs():
            if technology[stn] not in costs:
                try:
                    cap_cost = config.get(technology[stn], 'capital_cost')
                    if cap_cost[-1] == 'K':
                        cap_cost = float(cap_cost[:-1]) * pow(10, 3)
                    elif cap_cost[-1] == 'M':
                        cap_cost = float(cap_cost[:-1]) * pow(10, 6)
                except:
                    cap_cost = 0.
                try:
                    o_m_cost = config.get(technology[stn], 'o_m_cost')
                    if o_m_cost[-1] == 'K':
                        o_m_cost = float(o_m_cost[:-1]) * pow(10, 3)
                    elif o_m_cost[-1] == 'M':
                        o_m_cost = float(o_m_cost[:-1]) * pow(10, 6)
                except:
                    o_m_cost = 0.
                o_m_cost = o_m_cost * pow(10, -3)
                try:
                    fuel_cost = config.get(technology[stn], 'fuel_cost')
                    if fuel_cost[-1] == 'K':
                        fuel_cost = float(fuel_cost[:-1]) * pow(10, 3)
                    elif fuel_cost[-1] == 'M':
                        fuel_cost = float(fuel_cost[:-1]) * pow(10, 6)
                    else:
                        fuel_cost = float(fuel_cost)
                except:
                    fuel_cost = 0.
                costs[technology[stn]] = [cap_cost, o_m_cost, fuel_cost]
            capital_cost = capacity[stn] * costs[technology[stn]][0]
            if do_grid_cost or do_grid_path_cost:
                if technology[stn] in self.dispatchable:
                    cost, line_table = self.grid.Line_Cost(capacity[stn], capacity[stn])
                else:
                    cost, line_table = self.grid.Line_Cost(capacity[stn], 0.)
                if do_grid_path_cost:
                    grid_cost = cost * path[stn]
                else:
                    grid_cost = cost * grid[stn]
                try:
                    grid_cost += self.grid.Substation_Cost(line_table)
                except:
                    pass
            else:
                grid_cost = 0
            return capital_cost, grid_cost

        def do_ippppa():
            capital_cost, grid_cost = stn_costs()
            ippppa_data.set_number('system_capacity', capacity[stn] * 1000)
            ippppa_data.set_array('gen', net_hourly)
            ippppa_data.set_number('construction_financing_cost', capital_cost + grid_cost)
            ippppa_data.set_number('total_installed_cost', capital_cost + grid_cost)
            ippppa_data.set_array('om_capacity', [costs[technology[stn]][1]])
            if technology[stn] == 'Biomass':
                ippppa_data.set_number('om_opt_fuel_1_usage', self.biomass_multiplier
                                       * capacity[stn] * 1000)
                ippppa_data.set_array('om_opt_fuel_1_cost', [costs[technology[stn]][2]])
                ippppa_data.set_number('om_opt_fuel_1_cost_escal',
                                       ippppa_data.get_number('inflation_rate'))
            module = ssc.Module('ippppa')
            if (module.exec_(ippppa_data)):
             # return the relevant outputs desired
                energy = ippppa_data.get_array('gen')
                generation = 0.
                for i in range(len(energy)):
                    generation += energy[i]
                generation = generation * pow(10, -3)
                lcoe_real = ippppa_data.get_number('lcoe_real')
                lcoe_nom = ippppa_data.get_number('lcoe_nom')
                npv = ippppa_data.get_number('npv')
                self.stations.append(FinancialSummary(name[stn], technology[stn], capacity[stn],
                  generation, 0, round(capital_cost), lcoe_real, lcoe_nom, npv, round(grid_cost)))
            else:
                if self.status:
                   self.status.emit(QtCore.SIGNAL('log'), 'Errors encountered processing ' + name[stn])
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    if self.status:
                       self.status.emit(QtCore.SIGNAL('log'), 'ippppa error [' + str(idx) + ']: ' + msg)
                    else:
                        print 'ippppa error [', idx, ' ]: ', msg
                    idx += 1
                    msg = module.log(idx)
            del module

        self.stations = []
        self.status = status
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
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
            aparents = config.items('Parents')
            for key, value in aparents:
                for key2, value2 in aparents:
                    if key2 == key:
                        continue
                    value = value.replace(key2, value2)
                parents.append((key, value))
            del aparents
        except:
            pass
        try:
            self.helpfile = config.get('Files', 'help')
            for key, value in parents:
                self.helpfile = self.helpfile.replace(key, value)
            self.helpfile = self.helpfile.replace('$USER$', getUser())
            self.helpfile = self.helpfile.replace('$YEAR$', self.base_year)
        except:
            self.helpfile = ''
        try:
            self.biomass_multiplier = float(config.get('Biomass', 'multiplier'))
        except:
            self.biomass_multiplier = 8.25
        try:
            variable_files = config.get('Files', 'variable_files')
            for key, value in parents:
                variable_files = variable_files.replace(key, value)
            variable_files = variable_files.replace('$USER$', getUser())
            annual_file = config.get('SAM Modules', 'annualoutput_variables')
            annual_file = variable_files + '/' + annual_file
            ippppa_file = config.get('SAM Modules', 'ippppa_variables')
            ippppa_file = variable_files + '/' + ippppa_file
        except:
            annual_file = 'annualoutput_variables.xls'
            ippppa = 'ippppa_variables.xls'
        annual_data, annual_outputs = self.get_variables(annual_file)
        what_beans = whatFinancials(helpfile=self.helpfile)
        what_beans.exec_()
        ippas = what_beans.getValues()
        if ippas is None:
            self.stations = None
            return
        ssc_api = ssc.API()
# to suppress messages
        if not self.expert:
            ssc_api.set_print(0)
        ippppa_data, ippppa_outputs = self.get_variables(ippppa_file, overrides=ippas)
        costs = {}
        do_grid_loss = False
        do_grid_cost = False
        do_grid_path_cost = False
        if 'grid_losses' in ippas or 'grid_costs' in ippas or 'grid_path_costs' in ippas:
            set_grid_variables()
            try:
                if ippas['grid_losses']:
                    do_grid_loss = True
            except:
                pass
            if 'grid_costs' in ippas or 'grid_path_costs' in ippas:
                self.grid = Grid()  # open grid here to access cost table
                try:
                    if ippas['grid_costs']:
                        do_grid_cost = True
                except:
                    pass
                try:
                    if ippas['grid_path_costs']:
                        do_grid_path_cost = True
                except:
                    pass
        for stn in range(len(name)):
            if len(power[stn]) != 8760:
                capital_cost, grid_cost = stn_costs()
                self.stations.append(FinancialSummary(name[stn], technology[stn], capacity[stn],
                  0., 0, round(capital_cost), 0., 0., 0., round(grid_cost)))
                continue
            energy = []
            if do_grid_loss and grid[stn] != 0:
                if do_grid_path_cost:
                    for hr in range(len(power[stn])):
                        energy.append(power[stn][hr] * 1000 * (1 - self.line_loss * path[stn] -
                                      self.subs_loss))
                else:
                    for hr in range(len(power[stn])):
                        energy.append(power[stn][hr] * 1000 * (1 - self.line_loss * grid[stn] -
                                      self.subs_loss))
            else:
                for hr in range(len(power[stn])):
                    energy.append(power[stn][hr] * 1000)
            annual_data.set_array('system_hourly_energy', energy)
            net_hourly = None
            module = ssc.Module('annualoutput')
            if (module.exec_(annual_data)):
             # return the relevant outputs desired
                net_hourly = annual_data.get_array('hourly_energy')
                net_annual = annual_data.get_array('annual_energy')
                degradation = annual_data.get_array('annual_degradation')
                del module
                do_ippppa()
            else:
                if self.status:
                   self.status.emit(QtCore.SIGNAL('log'), 'Errors encountered processing ' + name[stn])
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    if self.status:
                       self.status.emit(QtCore.SIGNAL('log'), 'annualoutput error [' + str(idx) + ']: ' + msg)
                    else:
                        print 'annualoutput error [', idx, ' ]: ', msg
                    idx += 1
                    msg = module.log(idx)
                del module

    def getValues(self):
        return self.stations


class ProgressModel(QtGui.QDialog):
    def __init__(self, stations, plots, show_progress, year=None, selected=None, status=None):
        super(ProgressModel, self).__init__()
        self.plots = plots
        self.model = SuperPower(stations, self.plots, False, year=year, selected=selected, status=status)
        self._active = False
        self.power_summary = []
        self.model.show_progress = show_progress
        self.progressbar = QtGui.QProgressBar()
        self.progressbar.setMinimum(1)
        try:
            self.progressbar.setMaximum(len(self.model.selected))
        except:
            self.progressbar.setMaximum(len(self.model.stations))
        self.button = QtGui.QPushButton('Start')
        self.button.clicked.connect(self.handleButton)
        self.progress_stn = QtGui.QLabel('Note: Solar Thermal Stations take a while to process')
        main_layout = QtGui.QGridLayout()
        main_layout.addWidget(self.button, 0, 0)
        main_layout.addWidget(self.progressbar, 0, 1)
        main_layout.addWidget(self.progress_stn, 1, 1)
        self.setLayout(main_layout)
        self.setWindowTitle('SIREN - Power Model Progress')
        self.resize(250, 30)

    def handleButton(self):
        if not self._active:
            self._active = True
            if self.progressbar.value() == self.progressbar.maximum():
                # self.progressbar.reset()
                self.button.setText('Finished')
            else:
                self.button.setText('Stop')
                self.model.getPower()
                self.getPowerLoop()
        else:
            self.close()
            self._active = False
            self.button.setText('Start')

    def closeEvent(self, event):
        self._active = False
        event.accept()

    def getPowerLoop(self):
        self.model.all_done = False
        to_do = []
        for st in range(len(self.model.stations)):
            if self.model.plots['by_station']:
                if self.model.stations[st].name not in self.model.selected:
                    continue
            if self.model.stations[st].technology == 'Rooftop PV' \
              and self.model.stations[st].scenario == 'Existing' \
              and not self.plots['gross_load']:
                continue
            if self.model.stations[st].technology[:6] == 'Fossil' \
              and not self.model.plots['actual']:
                continue
            to_do.append(st)
        for st in to_do:
            self.progress_stn.setText('Processing ' + self.model.stations[st].name)
            stn = self.model.stations[st]
            if stn.technology[:6] == 'Fossil' and not self.model.plots['actual']:
                value = self.progressbar.value() + 1
                self.progressbar.setValue(value)
                continue
            if self.model.plots['by_station']:
                if stn.name not in self.model.selected:
                    continue
            if self.model.plots['save_data'] or self.model.plots['financials'] or self.plots['save_detail']:
                self.model.stn_outs.append(stn.name)
                self.model.stn_tech.append(stn.technology)
                self.model.stn_size.append(stn.capacity)
                self.model.stn_pows.append([])
                if stn.grid_len is not None:
                    self.model.stn_grid.append(stn.grid_len)
                else:
                    self.model.stn_grid.append(0.)
                if stn.grid_path_len is not None:
                    self.model.stn_path.append(stn.grid_path_len)
                else:
                    self.model.stn_path.append(0.)
            elif self.model.plots['save_tech'] or self.plots['save_balance']:
                self.model.stn_outs.append(stn.name)
                self.model.stn_tech.append(stn.technology)
            elif self.plots['visualise']:
                self.model.stn_outs.append(stn.name)
                self.model.stn_pows.append([])
            if stn.technology == 'Rooftop PV' and stn.scenario == 'Existing' \
              and not self.model.plots['gross_load']:
                continue
            if self.model.plots['by_station']:
                if stn.name not in self.model.selected:
                    continue
                key = stn.name
            else:
                if stn.technology == 'Rooftop PV' and stn.scenario == 'Existing':
                    key = 'Existing Rooftop PV'
                else:
                    key = stn.technology
            power = self.model.getStationPower(stn)
            total_power = 0.
            total_energy = 0.
            if power is None:
                pass
            else:
                if key in self.model.ly:
                    pass
                else:
                    self.model.ly[key] = []
                    for i in range(len(self.model.x)):
                        self.model.ly[key].append(0.)
                for i in range(len(power)):
                    if self.model.plots['grid_losses']:
                        if stn.grid_path_len is not None:
                            enrgy = power[i] * (1 - self.model.line_loss * stn.grid_path_len -
                                    self.model.subs_loss)
                        else:
                            enrgy = power[i] * (1 - self.model.subs_loss)
                        self.model.ly[key][i] += enrgy / 1000.
                        total_energy += enrgy / 1000.
                        self.model.ly['Generation'][i] += power[i] / 1000.
                    else:
                        self.model.ly[key][i] += power[i] / 1000.
                    total_power += power[i] / 1000.
                    if self.model.plots['save_data'] or self.model.plots['financials'] or \
                      self.plots['save_detail'] or self.plots['visualise']:
                        self.model.stn_pows[-1].append(power[i] / 1000.)
            if total_energy > 0:
                pt = PowerSummary(stn.name, stn.technology, total_power, stn.capacity,
                                  total_energy)
            else:
                pt = PowerSummary(stn.name, stn.technology, total_power, stn.capacity)
            self.power_summary.append(pt)
            value = self.progressbar.value() + 1
            if value > self.progressbar.maximum():
                self.model.all_done = True
                break
            self.progressbar.setValue(value)
            QtGui.qApp.processEvents()
            if not self._active:
                break
        else:
            self.model.all_done = True
        self.close()
        self.model.getPowerDone()

    def getValues(self):
        return self.power_summary

    def getPct(self):
        return self.model.getPct()

    def getLy(self):
        return self.model.getLy()

    def getStnOuts(self):
        return self.model.getStnOuts()

    def getStnTech(self):
        return self.model.getStnTech()

    def getStnPows(self):
        return self.model.getStnPows()

class PowerModel():
    powerExit = QtCore.pyqtSignal(str)

    def showGraphs(self, ydata, x):
        def shrinkKey(key):
            remove = ['Biomass', 'Community', 'Farm', 'Fixed', 'Geothermal', 'Hydro', 'Pumped',
                      'PV', 'Rooftop', 'Solar', 'Station', 'Thermal', 'Tracking', 'Wave', 'Wind']
            oukey = key
            for i in range(len(remove)):
                oukey = oukey.replace(remove[i], '')
            oukey = ' '.join(oukey.split())
            if oukey == '' or oukey == 'Existing':
                return key
            else:
                return oukey

        def stepPlot(self, period, data, x_labels=None):
            k1 = data.keys()[0]
            if self.plots['cumulative']:
                pc = 1
            else:
                pc = 0
            if self.plots['gross_load']:
                pc += 1
            if self.plots['shortfall']:
                pc += 1
            fig = plt.figure(self.hdrs['by_' + period].title() + self.suffix)
            plt.grid(True)
            bbdx = fig.add_subplot(111)
            plt.title(self.hdrs['by_' + period].title() + self.suffix)
            maxy = 0
            miny = 0
            xs = []
            for i in range(len(data[k1]) + 1):
                xs.append(i)
            if self.plots['save_plot']:
                sp_data = []
                sp_data.append(xs[1:])
                if period == 'day':
                    sp_vals = [period, 'Date']
                    sp_data.append([])
                    mm = 0
                    dy = 1
                    for d in range(1, len(xs)):
                        sp_data[-1].append('%s-%s-%s' % (self.load_year, str(mm + 1).zfill(2), str(dy).zfill(2)))
                        dy += 1
                        if dy > the_days[mm]:
                            mm += 1
                            dy = 1
                else:
                    sp_vals = ['No.', period]
                    sp_data.append(x_labels)
            if self.plots['cumulative']:
                cumulative = [0.]
                for i in range(len(data[k1])):
                    cumulative.append(0.)
            load = []
            i = -1
            storage = None
            if self.plots['show_pct']:
                load_sum = 0.
                gen_sum = 0.
            for key, value in iter(sorted(ydata.iteritems())):
                if key == 'Generation':
                    continue
                dval = [0.]
                if self.plots['show_pct']:
                    for d in range(len(data[key])):
                        if key[:4] == 'Load':
                            for k in range(len(data[key][d])):
                                load_sum += data[key][d][k]
                        elif key == 'Storage':
                            pass
                        else:
                            for k in range(len(data[key][d])):
                                gen_sum += data[key][d][k]
                                if self.plots['gross_load'] and key == 'Existing Rooftop PV':
                                    load_sum += data[key][d][k]
                for d in range(len(data[key])):
                   dval.append(0.)
                   for k in range(len(data[key][0])):
                       dval[-1] += data[key][d][k] / 1000
                maxy = max(maxy, max(dval))
                lw = self.other_width
                if self.plots['cumulative'] and key[:4] != 'Load' and key != 'Storage':
                    lw = 1.0
                    for j in range(len(xs)):
                        cumulative[j] += dval[j]
                bbdx.step(xs, dval, linewidth=lw, label=shrinkKey(key), color=self.colours[key],
                          linestyle=self.linestyle[key])
                if self.plots['save_plot']:
                    sp_vals.append(shrinkKey(key))
                    sp_data.append(dval[1:])
                if (self.plots['shortfall'] or self.plots['show_load']) and key[:4] == 'Load':
                    load = dval[:]
                if self.plots['shortfall'] and key == 'Storage':
                    storage = dval[:]
            if self.plots['show_pct']:
                self.gen_pct = ' (%s%% of load)' % '{:0,.1f}'.format(gen_sum * 100. / load_sum)
                plt.title(self.hdrs['by_' + period].title() + self.suffix + self.gen_pct)
            if self.plots['cumulative']:
                bbdx.step(xs, cumulative, linewidth=self.other_width, label='Tot. Generation', color=self.colours['cumulative'])
                maxy = max(maxy, max(cumulative))
                if self.plots['save_plot']:
                    sp_vals.append('Tot. Generation')
                    sp_data.append(cumulative[1:])
            try:
                rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                maxy = ceil(maxy / rndup) * rndup
            except:
                pass
            if (self.plots['shortfall'] and self.do_load):
                load2 = []
                if storage is None:
                    for i in range(len(cumulative)):
                        load2.append(cumulative[i] - load[i])
                        if load2[-1] < miny:
                            miny = load2[-1]
                else:
                    for i in range(len(cumulative)):
                        load2.append(cumulative[i] + storage[i] - load[i])
                        if load2[-1] < miny:
                            miny = load2[-1]
                bbdx.step(xs, load2, linewidth=self.other_width, label='Shortfall', color=self.colours['shortfall'])
                plt.axhline(0, color='black')
                if rndup != 0 and miny < 0:
                    miny = -ceil(-miny / rndup) * rndup
                if self.plots['save_plot']:
                    sp_vals.append('Shortfall')
                    sp_data.append(load2[1:])
            else:
                miny = 0
            if self.plots['save_plot']:
                titl = 'by_' + period
                dialog = displaytable.Table(map(list, zip(*sp_data)), title=titl, fields=sp_vals, save_folder=self.scenarios)
                dialog.exec_()
                del dialog, sp_data, sp_vals
            plt.ylim([miny, maxy])
            plt.xlim([0, len(data[k1])])
            if (len(ydata) + pc) > 9:
                 # Shrink current axis by 5%
                box = bbdx.get_position()
                bbdx.set_position([box.x0, box.y0, box.width * 0.95, box.height])
                 # Put a legend to the right of the current axis
                bbdx.legend(loc='center left', bbox_to_anchor=(1, 0.5), prop=lbl_font)
            else:
                bbdx.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(ydata) + pc),
                            prop=lbl_font)
            rotn = 'horizontal'
            if len(data[k1]) > 12:
                stp = 7
                rotn = 'vertical'
            else:
                stp = 1
            plt.xticks(range(0, len(data[k1]), stp))
            tick_spot = []
            for i in range(0, len(data[k1]), stp):
                tick_spot.append(i + .5)
            bbdx.set_xticks(tick_spot)
            bbdx.set_xticklabels(x_labels, rotation=rotn)
            bbdx.set_xlabel(period.title() + ' of the year')
            bbdx.set_ylabel('Energy (MWh)')
            if self.plots['maximise']:
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

        def dayPlot(self, period, data, per_labels=None, x_labels=None):
            k1 = data.keys()[0]
            if self.plots['cumulative']:
                pc = 1
            else:
                pc = 0
            if self.plots['gross_load']:
                pc += 1
            if self.plots['shortfall']:
                pc += 1
            hdr = self.hdrs[period].replace('Power - ', '')
            plt.figure(hdr + self.suffix)
            plt.suptitle(self.hdrs[period] + self.suffix, fontsize=16)
            maxy = 0
            miny = 0
            if len(data[k1]) > 9:
                p1 = 3
                p2 = 4
                xl = 8
                yl = [0, 4, 8]
            elif len(data[k1]) > 6:
                p1 = 3
                p2 = 3
                xl = 6
                yl = [0, 3, 6]
            elif len(data[k1]) > 4:
                p1 = 2
                p2 = 3
                xl = 3
                yl = [0, 3]
            elif len(data[k1]) > 2:
                p1 = 2
                p2 = 2
                xl = 2
                yl = [0, 2]
            else:
                p1 = 1
                p2 = 2
                xl = 0
                yl = [0, 1]
            for key in data.keys():
                for p in range(len(data[key])):
                    maxy = max(maxy, max(data[key][p]))
            if self.plots['show_pct']:
                load_sum = []
                gen_sum = []
                for p in range(len(data[k1])):
                    load_sum.append(0.)
                    gen_sum.append(0.)
            for p in range(len(data[k1])):
                if self.plots['cumulative']:
                    cumulative = []
                    for i in range(len(x24)):
                        cumulative.append(0.)
                if self.plots['gross_load']:
                    gross_load = []
                    for i in range(len(x24)):
                        gross_load.append(0.)
                px = plt.subplot(p1, p2, p + 1)
                l_k = ''
                for key, value in iter(sorted(ydata.iteritems())):
                    if key == 'Generation':
                        continue
                    if key[:4] == 'Load':
                        l_k = key
                    if self.plots['show_pct']:
                        for d in range(len(data[key][p])):
                            if key[:4] == 'Load':
                                load_sum[p] += data[key][p][d]
                            elif key == 'Storage':
                                pass
                            else:
                                gen_sum[p] += data[key][p][d]
                                if self.plots['gross_load'] and key == 'Existing Rooftop PV':
                                    load_sum[p] += data[key][p][d]
                    lw = self.other_width
                    if self.plots['cumulative'] and key[:4] != 'Load' and key != 'Storage':
                        lw = 1.0
                        for j in range(len(x24)):
                            cumulative[j] += data[key][p][j]
                    if self.plots['gross_load'] and (key[:4] == 'Load' or
                      key == 'Existing Rooftop PV'):
                        for j in range(len(x24)):
                            gross_load[j] += data[key][p][j]
                    px.plot(x24, data[key][p], linewidth=lw, label=shrinkKey(key),
                            color=self.colours[key], linestyle=self.linestyle[key])
                    plt.title(per_labels[p])
                if self.plots['cumulative']:
                    px.plot(x24, cumulative, linewidth=self.other_width, label='Tot. Generation',
                            color=self.colours['cumulative'])
                    maxy = max(maxy, max(cumulative))
                if self.plots['gross_load'] and 'Existing Rooftop PV' in ydata.keys():
                    px.plot(x24, gross_load, linewidth=1.0, label='Gross Load', color=self.colours['gross_load'])
                    maxy = max(maxy, max(gross_load))
                if self.plots['shortfall'] and self.do_load:
                    load2 = []
                    for i in range(len(x24)):
                        load2.append(cumulative[i] - data[l_k][p][i])
                    miny = min(miny, min(load2))
                    px.plot(x24, load2, linewidth=self.other_width, label='Shortfall', color=self.colours['shortfall'])
                    plt.axhline(0, color='black')
                plt.xticks(range(4, 25, 4))
                px.set_xticklabels(x_labels[1:])
                plt.xlim([1, 24])
                if p >= xl:
                    px.set_xlabel('Hour of the Day')
                if p in yl:
                    px.set_ylabel('Power (MW)')
            try:
                rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                maxy = ceil(maxy / rndup) * rndup
            except:
                pass
            if self.plots['shortfall']:
                if rndup != 0 and miny < 0:
                    miny = -ceil(-miny / rndup) * rndup
            for p in range(len(data[k1])):
                px = plt.subplot(p1, p2, p + 1)
                plt.ylim([miny, maxy])
                plt.xlim([1, 24])
                if self.plots['show_pct']:
                    pct = ' (%s%%)' % '{:0,.1f}'.format(gen_sum[p] * 100. / load_sum[p])
                    titl = px.get_title()
                    px.set_title(titl + pct)
                    #  px.annotate(pct, xy=(1.0, 3.0))
            px = plt.subplot(p1, p2, len(data[k1]))
         #    px.legend(bbox_to_anchor=[1., -0.15], loc='best', ncol=min((len(ly) + pc), 9),
         # prop=lbl_font)
            if (len(ydata) + pc) > 9:
                if len(data[k1]) > 9:
                    do_in = [1, 5, 9, 2, 6, 10, 3, 7, 11, 4, 8, 12]
                elif len(data[k1]) > 6:
                    do_in = [1, 4, 7, 2, 5, 8, 3, 6, 9]
                elif len(data[k1]) > 4:
                    do_in = [1, 4, 2, 5, 3, 6]
                elif len(data[k1]) > 2:
                    do_in = [1, 3, 2, 4]
                else:
                    do_in = [1, 2]
                do_in = do_in[:len(data[k1])]
                for i in range(len(do_in)):
                    px = plt.subplot(p1, p2, do_in[i])
                 # Shrink current axis by 5%
                    box = px.get_position()
                    px.set_position([box.x0, box.y0, box.width * 0.95, box.height])
                 # Put a legend to the right of the current axis
                px.legend(loc='center left', bbox_to_anchor=(1, 0.5), prop=lbl_font)
            else:
                px.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(ydata) + pc),
                          prop=lbl_font)
            if self.plots['show_pct']:
                for p in range(1, len(gen_sum)):
                    load_sum[0] += load_sum[p]
                    gen_sum[0] += gen_sum[p]
                self.gen_pct = ' (%s%%) of load)' % '{:0,.1f}'.format(gen_sum[0] * 100. / load_sum[0])
                titl = px.get_title()
                plt.suptitle(self.hdrs[period] + self.suffix + self.gen_pct, fontsize=16)
            if self.plots['maximise']:
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

        def saveBalance(self, shortstuff):
            data_file = 'Powerbalance_data_%s.xls' % (
                    str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'yyyy-MM-dd_hhmm')))
            data_file = QtGui.QFileDialog.getSaveFileName(None, 'Save Powerbalance data file',
                        self.scenarios + data_file, 'Excel Files (*.xls*);;CSV Files (*.csv)')
            if data_file != '':
                if data_file[-4:] == '.csv' or data_file[-4:] == '.xls' \
                  or data_file[-5:] == '.xlsx':
                    pass
                else:
                    data_file += '.xls'
                if os.path.exists(data_file):
                    if os.path.exists(data_file + '~'):
                        os.remove(data_file + '~')
                    os.rename(data_file, data_file + '~')
                stns = {}
                techs = {}
                for i in range(len(self.power_summary)):
                    stns[self.power_summary[i].name] = i
                    techs[self.power_summary[i].technology] = [0., 0., 0.]
                if data_file[-4:] == '.csv':
                    tf = open(data_file, 'w')
                    line = 'Generation Summary Table'
                    tf.write(line + '\n')
                    line = 'Name,Technology,Capacity (MW),CF,Generation'
                    if getattr(self.power_summary[0], 'transmitted') != None:
                        line += ',Transmitted'
                    tf.write(line + '\n')
                    for key, value in iter(sorted(stns.iteritems())):
                        if self.power_summary[value].generation > 0:
                            cf = '{:0.2f}'.format(self.power_summary[value].generation /
                                 (self.power_summary[value].capacity * 8760))
                        else:
                            cf = ''
                        if self.power_summary[value].transmitted is not None:
                            ts = '{:0.2f}'.format(self.power_summary[value].transmitted)
                            techs[self.power_summary[value].technology][2] += self.power_summary[value].transmitted
                        else:
                            ts = ''
                        line = '"%s",%s,%s,%s,%s,%s' % (self.power_summary[value].name,
                                               self.power_summary[value].technology,
                                               '{:0.2f}'.format(self.power_summary[value].capacity),
                                               cf,
                                               '{:0.0f}'.format(self.power_summary[value].generation),
                                               ts)
                        techs[self.power_summary[value].technology][0] += self.power_summary[value].capacity
                        techs[self.power_summary[value].technology][1] += self.power_summary[value].generation
                        tf.write(line + '\n')
                    total = [0., 0., 0.]
                    for key, value in iter(sorted(techs.iteritems())):
                        total[0] += value[0]
                        total[1] += value[1]
                        if value[2] > 0:
                            v2 = ',{:0.0f}'.format(value[2])
                            total[2] += value[2]
                        else:
                            v2 = ''
                        line = ',%s,%s,,%s%s' % (key,
                                               '{:0.2f}'.format(value[0]),
                                               '{:0.0f}'.format(value[1]),
                                               v2)
                        tf.write(line + '\n')
                    if total[2] > 0:
                        v2 = ',{:0.0f}'.format(total[2])
                        total[2] += value[2]
                    else:
                        v2 = ''
                    line = ',Total,%s,,%s%s' % ('{:0.2f}'.format(total[0]),
                                               '{:0.0f}'.format(total[1]),
                                               v2)
                    tf.write(line + '\n')
                    line = '\nHourly Shortfall Table'
                    tf.write(line + '\n')
                    line = 'Hour,Period,Shortfall'
                    tf.write(line + '\n')
                    for i in range(len(shortstuff)):
                        line = '%s,%s,%s' % (str(shortstuff[i].hour), shortstuff[i].period,
                                             '{:0.2f}'.format(shortstuff[i].shortfall))
                        tf.write(line + '\n')
                    tf.close()
                else:
                    wb = xlwt.Workbook()
                    fnt = xlwt.Font()
                    fnt.bold = True
                    styleb = xlwt.XFStyle()
                    styleb.font = fnt
                    style2d = xlwt.XFStyle()
                    style2d.num_format_str = '#,##0.00'
                    style0d = xlwt.XFStyle()
                    style0d.num_format_str = '#,##0'
                    pattern = xlwt.Pattern()
                    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
                    pattern.pattern_fore_colour = xlwt.Style.colour_map['ice_blue']
                    style2db = xlwt.XFStyle()
                    style2db.num_format_str = '#,##0.00'
                    style2db.pattern = pattern
                    style0db = xlwt.XFStyle()
                    style0db.num_format_str = '#,##0'
                    style0db.pattern = pattern
                    ws = wb.add_sheet('Powerbalance')
                    xl_lens = {}
                    row = 0
                    col = 0
                    ws.write(row, col, 'Hourly Shortfall Table', styleb)
                    row += 1
                    shrt_cols = ['Hour', 'Period', 'Shortfall']
                    for i in range(len(shrt_cols)):
                        ws.write(row, col + i, shrt_cols[i], styleb)
                        xl_lens[col + i] = 0
                    row += 1
                    for i in range(len(shortstuff)):
                        ws.write(row, col, shortstuff[i].hour)
                        ws.write(row, col + 1, shortstuff[i].period)
                        xl_lens[col + 1] = max(xl_lens[col + 1], len(shortstuff[i].period))
                        ws.write(row, col + 2, shortstuff[i].shortfall, style2db)
                        row += 1
                    row = 0
                    col = len(shrt_cols) + 1
                    ws.write(row, col, 'Generation Summary Table', styleb)
                    sum_cols = ['Name', 'Technology', 'Capacity (MW)', 'CF', 'Generated\n(to be\ncosted)']
                    if getattr(self.power_summary[0], 'transmitted') != None:
                        sum_cols.append('Transmitted\n(reduces\nShortfall)')
                    for i in range(len(sum_cols)):
                        ws.write(1, col + i, sum_cols[i], styleb)
                        j = sum_cols[i].find('\n') - 1
                        if j < 0:
                            j = len(sum_cols[i])
                        xl_lens[col + i] = j
                    for key, value in iter(stns.iteritems()):
                        techs[self.power_summary[value].technology][0] += self.power_summary[value].capacity
                        techs[self.power_summary[value].technology][1] += self.power_summary[value].generation
                        if self.power_summary[value].transmitted is not None:
                            techs[self.power_summary[value].technology][2] += self.power_summary[value].transmitted
                    total = [0., 0., 0.]
                    row = 2
                    ws.write(row, col, 'Totals', styleb)
                    row += 1
                    for key, value in iter(sorted(techs.iteritems())):
                        ws.write(row, col + 1, key)
                        ws.write(row, col + 2, value[0], style2db)
                        total[0] += value[0]
                        ws.write(row, col + 4, value[1], style0db)
                        total[1] += value[1]
                        if value[2] > 0:
                            ws.write(row, col + 5, value[2], style0d)
                            total[2] += value[2]
                        row += 1
                    ws.write(row, col + 1, 'Total', styleb)
                    ws.write(row, col + 2, total[0], style2db)
                    ws.write(row, col + 4, total[1], style0db)
                    if total[2] > 0:
                        ws.write(row, col + 5, total[2], style0d)
                    row += 1
                    ws.write(row, col, 'Stations', styleb)
                    row += 1
                    for key, value in iter(sorted(stns.iteritems())):
                        ws.write(row, col, self.power_summary[value].name)
                        xl_lens[col] = max(xl_lens[col], len(self.power_summary[value].name))
                        ws.write(row, col + 1, self.power_summary[value].technology)
                        xl_lens[col + 1] = max(xl_lens[col + 1], len(self.power_summary[value].technology))
                        ws.write(row, col + 2, self.power_summary[value].capacity, style2d)
                        if self.power_summary[value].generation > 0:
                            ws.write(row, col + 3, self.power_summary[value].generation /
                                     (self.power_summary[value].capacity * 8760), style2d)
                        ws.write(row, col + 4, self.power_summary[value].generation, style0d)
                        if self.power_summary[value].transmitted is not None:
                            ws.write(row, col + 5, self.power_summary[value].transmitted, style0d)
                        row += 1
                    for key in xl_lens:
                        if xl_lens[key] * 275 > ws.col(key).width:
                            ws.col(key).width = xl_lens[key] * 275
                    ws.row(1).height_mismatch = True
                    ws.row(1).height = 256 * 3
                    ws.set_panes_frozen(True)
                    ws.set_horz_split_pos(2)
                    ws.set_remove_splits(True)
                    wb.save(data_file)

        def saveBalance2(self, shortstuff):
            def cell_format(cell, new_cell):
                if cell.has_style:
                    new_cell.number_format = cell.number_format

         #   for i in range(len(shortstuff)):
          #             ws.write(row, col, shortstuff[i].hour)
            ts = oxl.load_workbook(self.pb_template)
            ws = ts.active
            type_tags = ['name', 'tech', 'cap', 'cf', 'gen', 'tmit', 'hrly']
            tech_tags = ['load', 'wind', 'offw', 'roof', 'fixed', 'single', 'dual', 'biomass', 'geotherm', 'other1', 'cst']
            tech_names = ['Load', 'Wind', 'Offshore Wind', 'Rooftop PV', 'Fixed PV', 'Single Axis PV', 'Tracking PV', 'Biomass',
                       'Geothermal', 'Other1', 'CST']
            tech_names2 = [''] * len(tech_names)
            tech_names2[tech_names.index('Wind')] = 'Onshore Wind'
            tech_names2[tech_names.index('Tracking PV')] = 'Dual Axis PV'
            tech_names2[tech_names.index('CST')] = 'Solar Thermal'
            st_row = []
            st_col = []
            tech_row = []
            tech_col = []
            for i in range(len(type_tags)):
                st_row.append(0)
                st_col.append(0)
            for j in range(len(tech_tags)):
                tech_row.append([])
                tech_col.append([])
                for i in range(len(type_tags)):
                    tech_row[-1].append(0)
                    tech_col[-1].append(0)
            per_row = [0, 0]
            per_col= [0, 0]
            for row in range(1, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    try:
                        if ws.cell(row=row, column=col).value[0] != '<':
                            continue
                        if ws.cell(row=row, column=col).value == '<title>':
                            titl = ''
                            for stn in self.stations:
                                if stn.scenario not in titl:
                                    titl += stn.scenario + '; '
                            try:
                                titl = titl[:-2]
                                titl = titl.replace('.xls', '')
                                ws.cell(row=row, column=col).value = titl
                            except:
                                ws.cell(row=row, column=col).value = None
                        elif ws.cell(row=row, column=col).value == '<date>':
                            dte = str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(),
                                    'yyyy-MM-dd hh:mm'))
                            ws.cell(row=row, column=col).value = dte
                        elif ws.cell(row=row, column=col).value == '<period>':
                            per_row[1] = row
                            per_col[1] = col
                            ws.cell(row=row, column=col).value = None
                        elif ws.cell(row=row, column=col).value == '<hour>':
                            per_row[0] = row
                            per_col[0] = col
                            ws.cell(row=row, column=col).value = None
                        elif ws.cell(row=row, column=col).value == '<year>':
                            ws.cell(row=row, column=col).value = str(self.base_year)
                        elif ws.cell(row=row, column=col).value == '<growth>':
                            if self.load_multiplier != 0:
                                ws.cell(row=row, column=col).value = self.load_multiplier
                            else:
                                ws.cell(row=row, column=col).value = None
                        elif ws.cell(row=row, column=col).value[:5] == '<stn_':
                            bit = str(ws.cell(row=row, column=col).value)[:-1].split('_')
                            ty = type_tags.index(bit[-1])
                            st_row[ty] = row
                            st_col[ty] = col
                            ws.cell(row=row, column=col).value = None
                        elif ws.cell(row=row, column=col).value.find('_') > 0:
                            bit = str(ws.cell(row=row, column=col).value)[1:-1].split('_')
                            te = tech_tags.index(bit[0])
                            ty = type_tags.index(bit[-1])
                            tech_row[te][ty] = row
                            tech_col[te][ty] = col
                            ws.cell(row=row, column=col).value = None
                    except:
                        pass
            data_file = 'Powerbalance_data_%s.xlsx' % (
                    str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'yyyy-MM-dd_hhmm')))
            data_file = str(QtGui.QFileDialog.getSaveFileName(None, 'Save Powerbalance data file',
                        self.scenarios + data_file, 'Excel Files (*.xlsx)'))
            if data_file == '':
                return
            if data_file[-5:] != '.xlsx':
                data_file += '.xlsx'
            if os.path.exists(data_file):
                if os.path.exists(data_file + '~'):
                    os.remove(data_file + '~')
                os.rename(data_file, data_file + '~')
            stns = {}
            techs = {}
            for i in range(len(self.power_summary)):
                stns[self.power_summary[i].name] = i
                techs[self.power_summary[i].technology] = [0., 0., 0.]
            st = 0
            for key, value in iter(sorted(stns.iteritems())):
                ws.cell(row=st_row[0] + st, column=st_col[0]).value = self.power_summary[value].name
                ws.cell(row=st_row[1] + st, column=st_col[1]).value = self.power_summary[value].technology
                ws.cell(row=st_row[2] + st, column=st_col[2]).value = self.power_summary[value].capacity
                if self.power_summary[value].generation > 0:
                    ws.cell(row=st_row[3] + st, column=st_col[3]).value = self.power_summary[value].generation / \
                             (self.power_summary[value].capacity * 8760)
                ws.cell(row=st_row[4] + st, column=st_col[4]).value = self.power_summary[value].generation
                if self.power_summary[value].transmitted is not None:
                    ws.cell(row=st_row[5] + st, column=st_col[5]).value = self.power_summary[value].transmitted
                st += 1
            for key, value in iter(stns.iteritems()):
                techs[self.power_summary[value].technology][0] += self.power_summary[value].capacity
                techs[self.power_summary[value].technology][1] += self.power_summary[value].generation
                if self.power_summary[value].transmitted is not None:
                    techs[self.power_summary[value].technology][2] += self.power_summary[value].transmitted
            for key, value in iter(techs.iteritems()):
                try:
                    te = tech_names.index(key)
                except:
                    try:
                        te = tech_names2.index(key)
                    except:
                        continue
                ws.cell(row=tech_row[te][2], column=tech_col[te][2]).value = value[0]
                ws.cell(row=tech_row[te][4], column=tech_col[te][4]).value = value[1]
                if self.plots['grid_losses']:
                    ws.cell(row=tech_row[te][5], column=tech_col[te][5]).value = value[2]
                if value[1] > 0:
                    ws.cell(row=tech_row[te][3], column=tech_col[te][3]).value = \
                      value[1] / (value[0] * 8760)
            if per_row[0] > 0:
                for i in range(8760):
                    ws.cell(row=per_row[0] + i, column=per_col[0]).value = shortstuff[i].hour
                    cell_format(ws.cell(row=per_row[0], column=per_col[0]), ws.cell(row=per_row[0] + i, column=per_col[0]))
            if per_row[1] > 0:
                for i in range(8760):
                    ws.cell(row=per_row[1] + i, column=per_col[1]).value = shortstuff[i].period
                    cell_format(ws.cell(row=per_row[1], column=per_col[1]), ws.cell(row=per_row[1] + i, column=per_col[1]))
            if tech_row[0][6] > 0:
                for i in range(8760):
                    ws.cell(row=tech_row[0][6] + i, column=tech_col[0][6]).value = shortstuff[i].load
                    cell_format(ws.cell(row=tech_row[0][6], column=tech_col[0][6]), ws.cell(row=tech_row[0][6] + i, column=tech_col[0][6]))
            ly_keys = []
            for t in range(len(tech_names)):
                ly_keys.append([])
            if self.plots['by_station']:
                for t in range(len(self.stn_tech)):
                    try:
                        i = tech_names.index(self.stn_tech[t])
                    except:
                        try:
                            i = tech_names2.index(self.stn_tech[t])
                        except:
                            continue
                    ly_keys[i].append(self.stn_outs[t])
            else:
                for t in range(len(tech_names)):
                    if tech_names[t] in techs.keys():
                        ly_keys[t].append(tech_names[t])
                    if tech_names2[t] != '':
                        if tech_names2[t] in techs.keys():
                            ly_keys[t].append(tech_names2[t])
            for te in range(len(tech_row)):
                if tech_row[te][6] == 0 or len(ly_keys[te]) == 0:
                    continue
                hrly = [0.] * 8760
                doit = False
                for key in ly_keys[te]:
                    try:
                        values = self.ly[key]
                        for h in range(len(hrly)):
                            hrly[h] += values[h]
                            if hrly[h] != 0:
                                doit = True
                    except:
                        pass
                if doit or not doit:
                    for h in range(len(hrly)):
                        ws.cell(row=tech_row[te][6] + h, column=tech_col[te][6]).value = hrly[h]
                        cell_format(ws.cell(row=tech_row[te][6], column=tech_col[te][6]), \
                                    ws.cell(row=tech_row[te][6] + h, column=tech_col[te][6]))
            ts.save(data_file)

        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        try:
            mapc = config.get('Map', 'map_choice')
        except:
            mapc = ''
        self.colours = {'cumulative': '#006400', 'gross_load': '#a9a9a9', 'load': '#000000',
                        'shortfall': '#8b0000', 'wind': '#6688bb'}
        try:
            colors = config.items('Colors')
            for item, colour in colors:
                if item in self.technologies or self.colours.has_key(item):
                    itm = item.replace('_', ' ').title()
                    itm = itm.replace('Pv', 'PV')
                    self.colours[itm] = colour
        except:
            pass
        if mapc != '':
            try:
                colors = config.items('Colors' + mapc)
                for item, colour in colors:
                    if item in self.technologies or self.colours.has_key(item):
                        itm = item.replace('_', ' ').title()
                        itm = itm.replace('Pv', 'PV')
                        self.colours[itm] = colour
            except:
                pass
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
        self.other_width = 2.
        seasons = [[], [], [], []]
        periods = [[], []]
        try:
            items = config.items('Power')
        except:
            seasons[0] = ['Summer', 11, 0, 1]
            seasons[1] = ['Autumn', 2, 3, 4]
            seasons[2] = ['Winter', 5, 6, 7]
            seasons[3] = ['Spring', 8, 9, 10]
            periods[0] = ['Winter', 4, 5, 6, 7, 8, 9]
            periods[1] = ['Summer', 10, 11, 0, 1, 2, 3]
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
            elif item == 'other_width':
                try:
                    self.other_width = float(values)
                except:
                    pass
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
        if papersize != '':
            if landscape:
                bit = papersize.split(',')
                plt.rcParams['figure.figsize'] = bit[1] + ',' + bit[0]
            else:
                plt.rcParams['figure.figsize'] = papersize
        try:
            self.pb_template = config.get('Power', 'pb_template')
        except:
            try:
                self.pb_template = config.get('Files', 'pb_template')
            except:
                self.pb_template = False
        if self.pb_template:
            try:
                parents = []
                aparents = config.items('Parents')
                for key, value in aparents:
                    for key2, value2 in aparents:
                        if key2 == key:
                            continue
                        value = value.replace(key2, value2)
                    parents.append((key, value))
                del aparents
                for key, value in parents:
                    self.pb_template = self.pb_template.replace(key, value)
                self.pb_template = self.pb_template.replace('$USER$', getUser())
                if not os.path.exists(self.pb_template):
                    self.pb_template = False
            except:
                pass
        mth_labels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct',
                      'Nov', 'Dec']
        ssn_labels = []
        for i in range(len(seasons)):
            if len(seasons[i]) == 2:
                ssn_labels.append('%s (%s)' % (seasons[i][0], mth_labels[seasons[i][1]]))
            else:
                ssn_labels.append('%s (%s-%s)' % (seasons[i][0], mth_labels[seasons[i][1]],
                                   mth_labels[seasons[i][-1]]))
        smp_labels = []
        for i in range(len(periods)):
            if len(periods[i]) == 2:
                smp_labels.append('%s (%s)' % (periods[i][0], mth_labels[periods[i][1]]))
            else:
                smp_labels.append('%s (%s-%s)' % (periods[i][0], mth_labels[periods[i][1]],
                                   mth_labels[periods[i][-1]]))
        labels = ['0:00', '4:00', '8:00', '12:00', '16:00', '20:00', '24:00']
        mth_xlabels = ['0:', '4:', '8:', '12:', '16:', '20:', '24:']
        pct_labels = ['0%', '20%', '40%', '60%', '80%', '100%']
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
        l24 = {}
        m24 = {}
        q24 = {}
        s24 = {}
        d365 = {}
        for i in range(24):
            x24.append(i + 1)
        for key in ydata.keys():
            if self.plots['total']:
                l24[key] = []
                for j in range(24):
                    l24[key].append(0.)
            if self.plots['month'] or self.plots['by_month']:
                m24[key] = []
                for m in range(12):
                    m24[key].append([])
                    for j in range(24):
                        m24[key][m].append(0.)
            if self.plots['season'] or self.plots['by_season']:
                q24[key] = []
                for q in range(len(seasons)):
                    q24[key].append([])
                    for j in range(24):
                        q24[key][q].append(0.)
            if self.plots['period'] or self.plots['by_period']:
                s24[key] = []
                for s in range(len(periods)):
                    s24[key].append([])
                    for j in range(24):
                        s24[key][s].append(0.)
            if self.plots['by_day']:
                d365[key] = []
                for j in range(365):
                    d365[key].append([0.])
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
                for key, value in iter(sorted(ydata.iteritems())):
                    if key == 'Generation':
                        continue
                    if self.plots['total']:
                        l24[key][k] += value[i + k]
                    if self.plots['by_day']:
                        d365[key][d][0] += value[i + k]
                    if self.plots['month'] or self.plots['by_month']:
                        m24[key][m][k] = m24[key][m][k] + value[i + k]
                    if self.plots['season'] or self.plots['by_season']:
                        for q in range(len(seasons)):
                            if m in seasons[q]:
                                break
                        q24[key][q][k] = q24[key][q][k] + value[i + k]
                    if self.plots['period'] or self.plots['by_period']:
                        for s in range(len(periods)):
                            if m in periods[s]:
                                break
                        s24[key][s][k] = s24[key][s][k] + value[i + k]
        if self.plots['cumulative']:
            pc = 1
        else:
            pc = 0
        if self.plots['gross_load']:
            pc += 1
        if self.plots['shortfall']:
            pc += 1
        colours = ['r', 'g', 'b', 'c', 'm', 'y', 'orange', 'darkcyan', 'darkmagenta',
                   'darkolivegreen', 'darkorange', 'darkturquoise', 'darkviolet', 'violet']
        colour_index = 0
        linestyles = ['-', '--', '-.', ':']
        line_index = 0
        self.linestyle = {}
        for key in self.colours:
            self.linestyle[key] = '-'
        for key in ydata:
            if key not in self.colours:
                if key[:4] == 'Load':
                    try:
                        self.colours[key] = self.colours['load']
                    except:
                        self.colours[key] = 'black'
                    self.linestyle[key] = '-'
                else:
                    self.colours[key] = colours[colour_index]
                    self.linestyle[key] = linestyles[line_index]
                    colour_index += 1
                    if colour_index >= len(colours):
                        colour_index = 0
                        line_index += 1
                        if line_index >= len(linestyles):
                            line_index = 0
        if self.plots['by_day']:
            stepPlot(self, 'day', d365, day_labels)
        if self.plots['by_month']:
            stepPlot(self, 'month', m24, mth_labels)
        if self.plots['by_season']:
            stepPlot(self, 'season', q24, ssn_labels)
        if self.plots['by_period']:
            stepPlot(self, 'period', s24, smp_labels)
        for key in ydata.keys():
            for k in range(24):
                if self.plots['total']:
                    l24[key][k] = l24[key][k] / 365
                if self.plots['month']:
                    for m in range(12):
                        m24[key][m][k] = m24[key][m][k] / the_days[m]
                if self.plots['season']:
                    for q in range(len(seasons)):
                        q24[key][q][k] = q24[key][q][k] / the_qtrs[q]
                if self.plots['period']:
                    for s in range(len(periods)):
                        s24[key][s][k] = s24[key][s][k] / the_ssns[s]
        if self.plots['hour']:
            if self.plots['save_plot']:
                sp_vals = ['hour']
                sp_data = []
                sp_data.append(x[1:])
                sp_data[-1].append(len(x))
                sp_vals.append('period')
                sp_data.append([])
                for i in range(len(x)):
                    sp_data[-1].append(the_date(self.load_year, i))
            hdr = self.hdrs['hour'].replace('Power - ', '')
            fig = plt.figure(hdr + self.suffix)
            plt.grid(True)
            hx = fig.add_subplot(111)
            plt.title(self.hdrs['hour'] + self.suffix)
            maxy = 0
            storage = None
            if self.plots['cumulative']:
                cumulative = []
                for i in range(len(x)):
                    cumulative.append(0.)
            if self.plots['gross_load']:
                gross_load = []
                for i in range(len(x)):
                    gross_load.append(0.)
            if self.plots['show_pct']:
                load_sum = 0.
                gen_sum = 0.
            for key, value in iter(sorted(ydata.iteritems())):
                if key == 'Generation':
                    continue
                if self.plots['show_pct']:
                    for i in range(len(x)):
                        if key[:4] == 'Load':
                            load_sum += value[i]
                        elif key == 'Storage':
                            pass
                        else:
                            gen_sum += value[i]
                            if self.plots['gross_load'] and key == 'Existing Rooftop PV':
                                load_sum += value[i]
                maxy = max(maxy, max(value))
                lw = self.other_width
                if self.plots['cumulative'] and key[:4] != 'Load' and key != 'Storage':
                    lw = 1.0
                    for i in range(len(x)):
                        cumulative[i] += value[i]
                if self.plots['gross_load'] and (key[:4] == 'Load' or
                      key == 'Existing Rooftop PV'):
                    for i in range(len(x)):
                        gross_load[i] += value[i]
                if self.plots['shortfall'] and key[:4] == 'Load':
                    load = value
                if self.plots['shortfall'] and key == 'Storage':
                    storage = value
                hx.plot(x, value, linewidth=lw, label=shrinkKey(key), color=self.colours[key],
                        linestyle=self.linestyle[key])
                if self.plots['save_plot']:
                    sp_vals.append(shrinkKey(key))
                    sp_data.append(value)
            if self.plots['cumulative']:
                hx.plot(x, cumulative, linewidth=self.other_width, label='Tot. Generation', color=self.colours['cumulative'])
                maxy = max(maxy, max(cumulative))
                if self.plots['save_plot']:
                    sp_vals.append('Tot. Generation')
                    sp_data.append(cumulative)
            if self.plots['gross_load'] and 'Existing Rooftop PV' in ydata.keys():
                hx.plot(x, gross_load, linewidth=1.0, label='Gross Load', color=self.colours['gross_load'])
                maxy = max(maxy, max(gross_load))
                if self.plots['save_plot']:
                    sp_vals.append('Gross Load')
                    sp_data.append(gross_load)
            try:
                rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                maxy = ceil(maxy / rndup) * rndup
            except:
                rndup = 0
            if self.plots['shortfall'] and self.do_load:
                load2 = []
                if storage is None:
                    for i in range(len(cumulative)):
                        load2.append(cumulative[i] - load[i])
                else:
                    for i in range(len(cumulative)):
                        load2.append(cumulative[i] + storage[i] - load[i])
                hx.plot(x, load2, linewidth=self.other_width, label='Shortfall', color=self.colours['shortfall'])
                plt.axhline(0, color='black')
                miny = min(load2)
                if rndup != 0 and miny < 0:
                    miny = -ceil(-miny / rndup) * rndup
                if self.plots['save_plot']:
                    sp_vals.append('Shortfall')
                    sp_data.append(load2)
            else:
                miny = 0
            if self.plots['save_plot']:
                titl = 'hour'
                dialog = displaytable.Table(map(list, zip(*sp_data)), title=titl, fields=sp_vals, save_folder=self.scenarios)
                dialog.exec_()
                del dialog, sp_data, sp_vals
            plt.ylim([miny, maxy])
            plt.xlim([0, len(x)])
            if (len(ydata) + pc) > 9:
                 # Shrink current axis by 5%
                box = hx.get_position()
                hx.set_position([box.x0, box.y0, box.width * 0.95, box.height])
                 # Put a legend to the right of the current axis
                hx.legend(loc='center left', bbox_to_anchor=(1, 0.5), prop=lbl_font)
            else:
                hx.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(ydata) + pc),
                          prop=lbl_font)
            plt.xticks(range(12, len(x), 168))
            hx.set_xticklabels(day_labels, rotation='vertical')
            hx.set_xlabel('Month of the year')
            hx.set_ylabel('Power (MW)')
            if self.plots['show_pct']:
                self.gen_pct = ' (%s%% of load)' % '{:0,.1f}'.format(gen_sum * 100. / load_sum)
                plt.title(self.hdrs['hour'] + self.suffix + self.gen_pct)
            if self.plots['maximise']:
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
        if self.plots['augment'] and self.do_load:
            hdr = self.hdrs['augment'].replace('Power - ', '')
            fig = plt.figure(hdr + self.suffix)
            plt.grid(True)
            hx = fig.add_subplot(111)
            plt.title(self.hdrs['augment'] + self.suffix)
            maxy = 0
            miny = 0
            storage = None
            cumulative = []
            for i in range(len(x)):
                cumulative.append(0.)
            if self.plots['gross_load']:
                gross_load = []
                for i in range(len(x)):
                    gross_load.append(0.)
            if self.plots['show_pct']:
                load_sum = 0.
                gen_sum = 0.
            for key, value in iter(sorted(ydata.iteritems())):
                if key == 'Generation' or key == 'Excess': # might need to keep excess
                    continue
                if self.plots['show_pct']:
                    for i in range(len(x)):
                        if key[:4] == 'Load':
                            load_sum += value[i]
                        elif key == 'Storage':
                            pass
                        else:
                            gen_sum += value[i]
                            if self.plots['gross_load'] and key == 'Existing Rooftop PV':
                                load_sum += value[i]
                maxy = max(maxy, max(value))
                lw = self.other_width
                if key[:4] != 'Load' and key != 'Storage':
                    for i in range(len(x)):
                        cumulative[i] += value[i]
                if self.plots['gross_load'] and (key[:4] == 'Load' or
                      key == 'Existing Rooftop PV'):
                    for i in range(len(x)):
                        gross_load[i] += value[i]
                if key[:4] == 'Load':
                    load = value
                if key == 'Storage':
                    storage = value
            maxy = max(maxy, max(cumulative))
            try:
                rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                maxy = ceil(maxy / rndup) * rndup
            except:
                rndup = 0
            regen = cumulative[:]
            for r in range(len(regen)):
                if regen[r] > load[r]:
                    regen[r] = load[r]
            hx.fill_between(x, 0, regen, color=self.colours['cumulative']) #'#004949')
            if storage is not None:
                for r in range(len(storage)):
                    storage[r] += cumulative[r]
                for r in range(len(storage)):
                    if storage[r] > load[r]:
                        storage[r] = load[r]
                hx.fill_between(x, regen, storage, color=self.colours['wind']) #'#006DDB')
                hx.fill_between(x, storage, load, color=self.colours['shortfall']) #'#920000')
            else:
                hx.fill_between(x, load, regen, color=self.colours['shortfall']) #'#920000')
            hx.plot(x, cumulative, linewidth=self.other_width, label='RE', linestyle='--', color=self.colours['gross_load'])
            if self.plots['save_plot']:
                sp_vals = ['hour']
                sp_data = []
                sp_data.append(x[1:])
                sp_data[-1].append(len(x))
                sp_vals.append('period')
                sp_data.append([])
                for i in range(len(x)):
                    sp_data[-1].append(the_date(self.load_year, i))
                sp_vals.append('load')
                sp_data.append(load)
                l = len(sp_data) - 1
                sp_vals.append('renewable')
                sp_data.append(regen)
                r = len(sp_data) - 1
                if storage is not None:
                    sp_vals.append('storage')
                    sp_data.append(storage)
                    s = len(sp_data) - 1
                else:
                    s = 0
                sp_vals.append('re gen.')
                sp_data.append(cumulative)
                e = len(sp_data) - 1
                titl = 'augmented'
                dialog = displaytable.Table(map(list, zip(*sp_data)), title=titl, fields=sp_vals, save_folder=self.scenarios)
                dialog.exec_()
                del dialog
                if s > 0:
                    for i in range(len(sp_data[s])):
                        sp_data[s][i] = sp_data[s][i] - sp_data[r][i]
                sp_data.append([])
                for i in range(len(sp_data[r])):
                    sp_data[-1].append(sp_data[e][i] - sp_data[r][i])
                sp_vals.append('excess')
                sp_vals[e] = 'augment'
                if s > 0:
                    for i in range(len(sp_data[e])):
                        sp_data[e][i] = sp_data[l][i] - sp_data[r][i] - sp_data[s][i]
                else:
                    for i in range(len(sp_data[e])):
                        sp_data[e][i] = sp_data[l][i] - sp_data[r][i]
                titl = 'augmented2'
                dialog = displaytable.Table(map(list, zip(*sp_data)), title=titl, fields=sp_vals, save_folder=self.scenarios)
                dialog.exec_()
                del dialog, sp_vals, sp_data
            plt.ylim([miny, maxy])
            plt.xlim([0, len(x)])
            plt.xticks(range(12, len(x), 168))
            hx.set_xticklabels(day_labels, rotation='vertical')
            hx.set_xlabel('Month of the year')
            hx.set_ylabel('Power (MW)')
            if self.plots['show_pct']:
                self.gen_pct = ' (%s%% of load)' % '{:0,.1f}'.format(gen_sum * 100. / load_sum)
                plt.title(self.hdrs['hour'] + self.suffix + self.gen_pct)
            if self.plots['maximise']:
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
    #        shortstuff = []
    #        vals = ['load', 'renewable', 'storage', 'cumulative']
    #        for i in range(0, len(load)):
    #            if storage is None:
    #                shortstuff.append(ColumnData(i + 1, the_date(self.load_year, i), [load[i],
    #                                  regen[i], 0., cumulative[i]], values=vals))
    #            else:
    #                shortstuff.append(ColumnData(i + 1, the_date(self.load_year, i), [load[i],
    #                                  regen[i], storage[i], cumulative[i]], values=vals))
    #        vals.insert(0, 'period')
    #        vals.insert(0, 'hour')
    #        dialog = displaytable.Table(shortstuff, title='Augmented',
    #                                        save_folder=self.scenarios, fields=vals)
    #        dialog.exec_()
    #        del dialog
        if self.plots['duration']:
            hdr = self.hdrs['duration'].replace('Power - ', '')
            fig = plt.figure(hdr + self.suffix)
            plt.grid(True)
            dx = fig.add_subplot(111)
            plt.title(self.hdrs['duration'] + self.suffix)
            maxy = 0
            if self.plots['cumulative']:
                cumulative = []
                for i in range(len(x)):
                    cumulative.append(0.)
            if self.plots['show_pct']:
                load_sum = 0.
                gen_sum = 0.
            for key, value in iter(sorted(ydata.iteritems())):
                if key == 'Generation':
                    continue
                if self.plots['show_pct']:
                    for i in range(len(x)):
                        if key[:4] == 'Load':
                            load_sum += value[i]
                        elif key == 'Storage':
                            pass
                        else:
                            gen_sum += value[i]
                            if self.plots['gross_load'] and key == 'Existing Rooftop PV':
                                load_sum += value[i]
                sortydata = sorted(value, reverse=True)
                lw = self.other_width
                if self.plots['cumulative'] and key[:4] != 'Load' and key != 'Storage':
                    lw = 1.0
                    for i in range(len(x)):
                        cumulative[i] += value[i]
                maxy = max(maxy, max(sortydata))
                dx.plot(x, sortydata, linewidth=lw, label=shrinkKey(key), color=self.colours[key],
                        linestyle=self.linestyle[key])
            if self.plots['cumulative']:
                sortydata = sorted(cumulative, reverse=True)
                dx.plot(x, sortydata, linewidth=self.other_width, label='Tot. Generation', color=self.colours['cumulative'])
            try:
                rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                maxy = ceil(maxy / rndup) * rndup
            except:
                pass
            plt.ylim([0, maxy])
            plt.xlim([0, len(x)])
            if (len(ydata) + pc) > 9:
                 # Shrink current axis by 10%
                box = dx.get_position()
                dx.set_position([box.x0, box.y0, box.width * 0.95, box.height])

                 # Put a legend to the right of the current axis
                dx.legend(loc='center left', bbox_to_anchor=(1, 0.5), prop=lbl_font)
            else:
                dx.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(ydata) + pc),
                          prop=lbl_font)
            tics = int(len(x) / (len(pct_labels) - 1))
            plt.xticks(range(0, len(x), tics))
            dx.set_xticklabels(pct_labels)
            dx.set_xlabel('Percentage of Year')
            dx.set_ylabel('Power (MW)')
            if self.plots['show_pct']:
                self.gen_pct = ' (%s%% of load)' % '{:0,.1f}'.format(gen_sum * 100. / load_sum)
                plt.title(self.hdrs['duration'] + self.suffix + self.gen_pct)
            if self.plots['maximise']:
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
            if self.do_load:
                hdr = self.hdrs['duration'].replace('Power - ', '')
           #      fig = plt.figure(hdr + self.suffix)
                plt.figure(hdr + ' 2')
                plt.grid(True)
                plt.title(self.hdrs['duration'] + ' with renewable contribution')
                lx = plt.subplot(111)
                maxy = 0
                miny = 0
                load = []  # use for this and next graph
                rgen = []  # use for this and next graph
                rgendiff = []
                for i in range(len(self.x)):
                    rgen.append(0.)
                    rgendiff.append(0.)
                if self.plots['show_pct']:
                    load_sum = 0.
                    gen_sum = 0.
                for key, value in ydata.iteritems():
                    if key == 'Generation':
                        continue
                    if self.plots['show_pct']:
                        for i in range(len(value)):
                            if key[:4] == 'Load':
                                load_sum += value[i]
                            elif key == 'Storage':
                                pass
                            else:
                                gen_sum += value[i]
                                if self.plots['gross_load'] and key == 'Existing Rooftop PV':
                                    load_sum += value[i]
                    if key[:4] == 'Load':
                        load = value
                    else:
                        for i in range(len(value)):
                            rgen[i] += value[i]
                for i in range(len(load)):
                    rgendiff[i] = load[i] - rgen[i]
                sortly1 = sorted(load, reverse=True)
                maxy = max(maxy, max(load))
                maxy = max(maxy, max(rgendiff))
                miny = min(miny, min(rgendiff))
                try:
                    rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                    maxy = ceil(maxy / rndup) * rndup
                    miny = -ceil(-miny / rndup) * rndup
                except:
                    pass
                if self.load_multiplier != 0:
                    load_key = 'Load ' + self.load_year
                else:
                    load_key = 'Load'
                lx.plot(x, sortly1, linewidth=self.other_width, label=load_key)
                sortly2 = sorted(rgendiff, reverse=True)
                lx.plot(x, sortly2, linewidth=self.other_width, label='Tot. Generation')
                lx.fill_between(x, sortly1, sortly2, facecolor=self.colours['cumulative'])
                plt.ylim([miny, maxy])
                plt.xlim([0, len(x)])
                lx.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=2, prop=lbl_font)
                tics = int(len(x) / (len(pct_labels) - 1))
                plt.xticks(range(0, len(x), tics))
                lx.set_xticklabels(pct_labels)
                lx.set_xlabel('Percentage of Year')
                lx.set_ylabel('Power (MW)')
                lx.axhline(0, color='black')
                if self.plots['show_pct']:
                    self.gen_pct = ' (%s%% of load)' % '{:0,.1f}'.format(gen_sum * 100. / load_sum)
                    plt.title(self.hdrs['duration'] + ' with renewable contribution' +
                              self.gen_pct)
                if self.plots['maximise']:
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
            plt.show(block=True)
        if (self.plots['shortfall_detail'] or self.plots['save_balance']) and self.do_load:
            load = []
            rgen = []
            shortfall = [[], [], [], []]
            generation = []
            for i in range(len(self.x)):
                rgen.append(0.)
                shortfall[0].append(0.)
            for key, value in ydata.iteritems():
                if key == 'Generation':
                    generation = value
                elif key[:4] == 'Load':
                    load = value
                else:
                    for i in range(len(value)):
                        rgen[i] += value[i]
            shortfall[0][0] = rgen[0] - load[0]
            for i in range(1, len(load)):
                shortfall[0][i] = shortfall[0][i - 1] + rgen[i] - load[i]
            d_short = [[], [0], [0, 0]]
            for i in range(0, len(load), 24):
                d_short[0].append(0.)
                for j in range(i, i + 24):
                    d_short[0][-1] += rgen[i] - load[i]
            if self.iterations > 0:
                for i in range(1, len(d_short[0])):
                    d_short[1].append((d_short[0][i - 1] + d_short[0][i]) / 2)
                for i in range(2, len(d_short[0])):
                    d_short[2].append((d_short[0][i - 2] + d_short[0][i - 1] + d_short[0][i]) / 3)
                d_short[1][0] = d_short[1][1]
                d_short[2][0] = d_short[2][1] = d_short[2][2]
                shortstuff = []
                vals = ['shortfall', 'iteration 1', 'iteration 2']
                for i in range(len(d_short[0])):
                    shortstuff.append(DailyData(i + 1, the_date(self.load_year, i * 24)[:10],
                                      [d_short[0][i], d_short[1][i], d_short[2][i]], values=vals))
                vals.insert(0, 'date')
                vals.insert(0, 'day')
                dialog = displaytable.Table(shortstuff, title='Daily Shortfall',
                         save_folder=self.scenarios, fields=vals)
                dialog.exec_()
                del dialog
                del shortstuff
                xs = []
                for i in range(len(d_short[0])):
                    xs.append(i + 1)
                plt.figure('daily shortfall')
                plt.grid(True)
                plt.title('Daily Shortfall')
                sdfx = plt.subplot(111)
                for i in range(self.iterations):
                    sdfx.step(xs, d_short[i], linewidth=self.other_width, label=str(i + 1) + ' day average',
                              color=colours[i])
                plt.xticks(range(0, len(xs), 7))
                tick_spot = []
                for i in range(0, len(xs), 7):
                    tick_spot.append(i + .5)
                sdfx.set_xticks(tick_spot)
                sdfx.set_xticklabels(day_labels, rotation='vertical')
                sdfx.set_xlabel('Day of the year')
                sdfx.set_ylabel('Power (MW)')
                plt.xlim([0, len(xs)])
                sdfx.legend(loc='best')
                for i in range(len(d_short)):
                    lin = min(d_short[i])
                    sdfx.axhline(lin, linestyle='--', color=colours[i])
                    lin = max(d_short[i])
                    sdfx.axhline(lin, linestyle='--', color=colours[i])
                lin = sum(d_short[0]) / len(d_short[0])
                sdfx.axhline(lin, linestyle='--', color='black')
                if self.plots['maximise']:
                    mng = plt.get_current_fig_manager()
                    if sys.platform == 'win32' or sys.platform == 'cygwin':
                        if plt.get_backend() == 'TkAgg':
                            mng.window.state('zoomed')
                        elif plt.get_backend() == 'Qt4Agg':
                            mng.window.showMaximized()
                    else:
                        mng.resize(*mng.window.maxsize())
                plt.show(block=True)
                h_storage = [-(shortfall[0][-1] / len(shortfall[0]))]  # average shortfall
                for s in range(1, self.iterations + 1):
                    for i in range(len(self.x)):
                        shortfall[s].append(0.)
                    ctr = 0
                    still_short = [0, 0]
                    if rgen[0] - load[0] + h_storage[-1] < 0:
                        still_short[0] += rgen[0] - load[0] + h_storage[-1]
                        ctr += 1
                    else:
                        still_short[1] += rgen[0] - load[0] + h_storage[-1]
                    shortfall[s][0] = rgen[0] - load[0] + h_storage[-1]
                    for i in range(1, len(load)):
                        shortfall[s][i] = shortfall[s][i - 1] + rgen[i] - load[i] + h_storage[-1]
                        if rgen[i] - load[i] + h_storage[-1] < 0:
                            still_short[0] += rgen[i] - load[i] + h_storage[-1]
                            ctr += 1
                        else:
                            still_short[1] += rgen[i] - load[i] + h_storage[-1]
    #                 h_storage.append(h_storage[-1] - still_short[0] / len(self.x))
                    h_storage.append(-(shortfall[0][-1] + still_short[0]) / len(self.x))
                dimen = log10(fabs(shortfall[0][-1]))
                unit = 'MW'
                if dimen > 11:
                    unit = 'PW'
                    div = 9
                elif dimen > 8:
                    unit = 'TW'
                    div = 6
                elif dimen > 5:
                    unit = 'GW'
                    div = 3
                else:
                    div = 0
                if div > 0:
                    for s in range(self.iterations + 1):
                        for i in range(len(shortfall[s])):
                            shortfall[s][i] = shortfall[s][i] / pow(10, div)
                plt.figure('cumulative shortfall')
                plt.grid(True)
                plt.title('Cumulative Shortfall')
                sfx = plt.subplot(111)
                sfx.plot(x, shortfall[0], linewidth=self.other_width, label='Shortfall', color=self.colours['shortfall'])
                for s in range(1, self.iterations + 1):
                    lbl = 'iteration %s - add %s MW to generation' % \
                          (s, '{:0,.0f}'.format(h_storage[s - 1]))
                    sfx.plot(x, shortfall[s], linewidth=self.other_width, label=lbl, color=colours[s])
                plt.xticks(range(0, len(x), 168))
                tick_spot = []
                for i in range(0, len(x), 168):
                    tick_spot.append(i + .5)
                box = sfx.get_position()
                sfx.set_position([box.x0, box.y0, box.width, box.height])
                sfx.set_xticks(tick_spot)
                sfx.set_xticklabels(day_labels, rotation='vertical')
                plt.xlim([0, len(x)])
                sfx.set_xlabel('Day of the year')
                sfx.set_ylabel('Power (' + unit + ')')
                sfx.legend(loc='best', prop=lbl_font)
                if self.plots['maximise']:
                    mng = plt.get_current_fig_manager()
                    if sys.platform == 'win32' or sys.platform == 'cygwin':
                        if plt.get_backend() == 'TkAgg':
                            mng.window.state('zoomed')
                        elif plt.get_backend() == 'Qt4Agg':
                            mng.window.showMaximized()
                    else:
                        mng.resize(*mng.window.maxsize())
                plt.show(block=True)
                for i in range(0, len(load)):
                    shortfall[0][i] = rgen[i] - load[i]
                for s in range(1, self.iterations + 1):
                    for i in range(0, len(load)):
                        shortfall[s][i] = rgen[i] - load[i] + h_storage[s - 1]
                plt.figure('shortfall')
                plt.grid(True)
                plt.title('Shortfall')
                sfx = plt.subplot(111)
                sfx.plot(x, shortfall[0], linewidth=self.other_width, label='Shortfall', color=self.colours['shortfall'])
                for s in range(1, self.iterations + 1):
                    lbl = 'iteration %s - add %s MW to generation' % \
                          (s, '{:0,.0f}'.format(h_storage[s - 1]))
                    sfx.plot(x, shortfall[s], linewidth=self.other_width, label=lbl, color=colours[s])
                plt.axhline(0, color='black')
                plt.xticks(range(0, len(x), 168))
                tick_spot = []
                for i in range(0, len(x), 168):
                    tick_spot.append(i + .5)
                box = sfx.get_position()
                sfx.set_position([box.x0, box.y0, box.width, box.height])
                sfx.set_xticks(tick_spot)
                sfx.set_xticklabels(day_labels, rotation='vertical')
                sfx.set_xlabel('Day of the year')
                sfx.set_ylabel('Power (MW)')
                plt.xlim([0, len(x)])
                sfx.legend(loc='best', prop=lbl_font)
                if self.plots['maximise']:
                    mng = plt.get_current_fig_manager()
                    if sys.platform == 'win32' or sys.platform == 'cygwin':
                        if plt.get_backend() == 'TkAgg':
                            mng.window.state('zoomed')
                        elif plt.get_backend() == 'Qt4Agg':
                            mng.window.showMaximized()
                    else:
                        mng.resize(*mng.window.maxsize())
                plt.show(block=True)
            else:
                for i in range(0, len(load)):
                    shortfall[0][i] = rgen[i] - load[i]
            shortstuff = []
            if self.plots['grid_losses']:
                vals = ['load', 'generation', 'transmitted', 'shortfall']
                short2 = [shortfall[0][0]]
                for i in range(1, len(self.x)):
                  #   short2.append(shortfall[0][i] - shortfall[0][i - 1])
                    short2.append(shortfall[0][i])
                for i in range(0, len(load)):
                    shortstuff.append(ColumnData(i + 1, the_date(self.load_year, i), [load[i],
                                      generation[i], rgen[i], short2[i]], values=vals))
            else:
                vals = ['load', 'generation', 'shortfall']
                for i in range(0, len(load)):
                    shortstuff.append(ColumnData(i + 1, the_date(self.load_year, i), [load[i],
                                      rgen[i], shortfall[0][i]], values=vals))
            vals.insert(0, 'period')
            vals.insert(0, 'hour')
            if self.plots['shortfall_detail']:
                dialog = displaytable.Table(shortstuff, title='Hourly Shortfall',
                                            save_folder=self.scenarios, fields=vals)
                dialog.exec_()
                del dialog
            if self.plots['save_balance']:
                if self.pb_template:
                    saveBalance2(self, shortstuff)
                else:
                    saveBalance(self, shortstuff)
            del shortstuff
        if self.plots['total']:
            maxy = 0
            if self.plots['cumulative']:
                cumulative = []
                for i in range(len(x24)):
                    cumulative.append(0.)
            if self.plots['gross_load']:
                gross_load = []
                for i in range(len(x24)):
                    gross_load.append(0.)
            if self.plots['save_plot']:
                sp_data = []
                sp_data.append(x24)
                sp_vals = ['hour']
            hdr = self.hdrs['total'].replace('Power - ', '')
            plt.figure(hdr + self.suffix)
            plt.grid(True)
            plt.title(self.hdrs['total'] + self.suffix)
            tx = plt.subplot(111)
            storage = None
            if self.plots['show_pct']:
                load_sum = 0.
                gen_sum = 0.
            for key, value in iter(sorted(ydata.iteritems())):
                if key == 'Generation':
                    continue
                if self.plots['show_pct']:
                    for j in range(len(x24)):
                        if key[:4] == 'Load':
                            load_sum += l24[key][j]
                        elif key == 'Storage':
                            pass
                        else:
                            gen_sum += l24[key][j]
                            if self.plots['gross_load'] and key == 'Existing Rooftop PV':
                                load_sum += l24[key][j]
                maxy = max(maxy, max(l24[key]))
                lw = self.other_width
                if self.plots['cumulative'] and key[:4] != 'Load' and key != 'Storage':
                    lw = 1.0
                    for j in range(len(x24)):
                        cumulative[j] += l24[key][j]
                if self.plots['gross_load'] and (key[:4] == 'Load' or
                      key == 'Existing Rooftop PV'):
                    for j in range(len(x24)):
                        gross_load[j] += l24[key][j]
                if self.plots['shortfall'] and key[:4] == 'Load':
                    load = value
                if self.plots['shortfall'] and key == 'Storage':
                    storage = value
                tx.plot(x24, l24[key], linewidth=lw, label=shrinkKey(key), color=self.colours[key],
                        linestyle=self.linestyle[key])
                if self.plots['save_plot']:
                    sp_data.append(l24[key])
                    sp_vals.append(key)
            if self.plots['cumulative']:
                tx.plot(x24, cumulative, linewidth=self.other_width, label='Tot. Generation', color=self.colours['cumulative'])
                maxy = max(maxy, max(cumulative))
                if self.plots['save_plot']:
                    sp_data.append(cumulative)
                    sp_vals.append('Tot. Generation')
            if self.plots['gross_load'] and 'Existing Rooftop PV' in ydata.keys():
                tx.plot(x24, gross_load, linewidth=1.0, label='Gross Load', color=self.colours['gross_load'])
                maxy = max(maxy, max(gross_load))
                if self.plots['save_plot']:
                    sp_data.append(gross_load)
                    sp_vals.append('Gross Load')
            try:
                rndup = pow(10, round(log10(maxy * 1.5) - 1)) / 2
                maxy = ceil(maxy / rndup) * rndup
            except:
                pass
            if self.plots['shortfall'] and self.do_load:
                load2 = []
                if storage is None:
                    for i in range(len(cumulative)):
                        load2.append(cumulative[i] - load[i])
                else:
                    for i in range(len(cumulative)):
                        load2.append(cumulative[i] + storage[i] - load[i])
                tx.plot(x24, load2, linewidth=self.other_width, label='Shortfall', color=self.colours['shortfall'])
                plt.axhline(0, color='black')
                miny = min(load2)
                if rndup != 0 and miny < 0:
                    miny = -ceil(-miny / rndup) * rndup
                if self.plots['save_plot']:
                    sp_data.append(load2)
                    sp_vals.append('Shortfall')
            else:
                miny = 0
            if self.plots['save_plot']:
                titl = 'total'
                dialog = displaytable.Table(map(list, zip(*sp_data)), title=titl, fields=sp_vals, save_folder=self.scenarios)
                dialog.exec_()
                del dialog, sp_data, sp_vals
            plt.ylim([miny, maxy])
            plt.xlim([1, 25])
            plt.xticks(range(0, 25, 4))
          #   tx.legend(loc='lower left', numpoints = 2, prop=lbl_font)
            tx.set_xticklabels(labels)
            tx.set_xlabel('Hour of the Day')
            tx.set_ylabel('Power (MW)')
            if (len(ydata) + pc) > 9:
                 # Shrink current axis by 5%
                box = tx.get_position()
                tx.set_position([box.x0, box.y0, box.width * 0.95, box.height])
                 # Put a legend to the right of the current axis
                tx.legend(loc='center left', bbox_to_anchor=(1, 0.5), prop=lbl_font)
            else:
                tx.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(ydata) + pc),
                          prop=lbl_font)
            if self.plots['show_pct']:
                self.gen_pct = ' (%s%% of load)' % '{:0,.1f}'.format(gen_sum * 100. / load_sum)
                plt.title(self.hdrs['total'] + self.suffix + self.gen_pct)
            if self.plots['maximise']:
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
            dayPlot(self, 'month', m24, mth_labels, mth_xlabels)
        if self.plots['season']:
            dayPlot(self, 'season', q24, ssn_labels, labels)
        if self.plots['period']:
            dayPlot(self, 'period', s24, smp_labels, labels)
        if not self.plots['block']:
            plt.show(block=True)

    def save_detail(self, data_file, techs, keys=None):
        if self.suffix != '':
            i = data_file.rfind('.')
            if i > 0:
                data_file = data_file[:i] + '_' + self.suffix + data_file[i:]
            else:
                data_file = data_file + '_' + self.suffix
        if data_file[-4:] == '.csv' or data_file[-4:] == '.xls' \
          or data_file[-5:] == '.xlsx':
            pass
        else:
            data_file += '.xls'
        data_file = QtGui.QFileDialog.getSaveFileName(None, 'Save power data file',
                    self.scenarios + data_file, 'Excel Files (*.xls*);;CSV Files (*.csv)')
        if data_file == '':
            return
        if self.load_multiplier != 0:
            the_year = self.load_year
        else:
            the_year = self.base_year
        if data_file[-4:] == '.csv' or data_file[-4:] == '.xls' \
          or data_file[-5:] == '.xlsx':
            pass
        else:
            data_file += '.xls'
        if os.path.exists(data_file):
            if os.path.exists(data_file + '~'):
                os.remove(data_file + '~')
            os.rename(data_file, data_file + '~')
        if keys is None:
            keys = sorted(techs.keys())
        if data_file[-4:] == '.csv':
            tf = open(data_file, 'w')
            line = 'Hour,Period,'
            if self.load_multiplier != 0:
                the_year = self.load_year
            else:
                the_year = self.base_year
            max_outs = 0
            lines = []
            for i in range(8760):
                lines.append(str(i+1) + ',' + str(the_date(the_year, i)) + ',')
            for key in keys:
                if key[:4] == 'Load' and self.load_multiplier != 0:
                    line += 'Load ' + self.load_year + ','
                else:
                    line += key + ','
                for i in range(len(techs[key])):
                    lines[i] += str(round(techs[key][i], 3)) + ','
            tf.write(line + '\n')
            for i in range(len(lines)):
                tf.write(lines[i] + '\n')
            tf.close()
            del lines
        else:
            wb = xlwt.Workbook()
            ws = wb.add_sheet('Detail')
            ws.write(0, 0, 'Hour')
            ws.write(0, 1, 'Period')
            for i in range(len(self.x)):
                ws.write(i + 1, 0, i + 1)
                ws.write(i + 1, 1, the_date(the_year, i))
            if 16 * 275 > ws.col(1).width:
                ws.col(1).width = 16 * 275
            c = 2
            for key in keys:
                if key[:4] == 'Load' and self.load_multiplier != 0:
                    ws.write(0, c, 'Load ' + self.load_year)
                else:
                    ws.write(0, c, key)
                if len(key) * 275 > ws.col(c).width:
                    ws.col(c).width = len(key) * 275
                for r in range(len(techs[key])):
                    ws.write(r + 1, c, round(techs[key][r], 3))
                c += 1
            ws.set_panes_frozen(True)  # frozen headings instead of split panes
            ws.set_horz_split_pos(1)  # in general, freeze after last heading row
            ws.set_remove_splits(True)  # if user does unfreeze, dont leave a split there
            wb.save(data_file)
            del wb

    def __init__(self, stations, show_progress=None, year=None, status=None, visualise=None):
        self.something = visualise
        self.something.power_signal = self
        self.status = status
        self.stations = stations
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = 'SIREN.ini'
        config.read(config_file)
        self.expert = False
        try:
            expert = config.get('Base', 'expert_mode')
            if expert in ['true', 'on', 'yes']:
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
            aparents = config.items('Parents')
            for key, value in aparents:
                for key2, value2 in aparents:
                    if key2 == key:
                        continue
                    value = value.replace(key2, value2)
                parents.append((key, value))
            del aparents
        except:
            pass
        try:
            scenario_prefix = config.get('Files', 'scenario_prefix')
        except:
            scenario_prefix = ''
        try:
            self.scenarios = config.get('Files', 'scenarios')
            if scenario_prefix != '' :
                self.scenarios += '/' + scenario_prefix
            for key, value in parents:
                self.scenarios = self.scenarios.replace(key, value)
            self.scenarios = self.scenarios.replace('$USER$', getUser())
            self.scenarios = self.scenarios.replace('$YEAR$', self.base_year)
            i = self.scenarios.rfind('/')
            self.scenarios = self.scenarios[:i + 1]
        except:
            self.scenarios = ''
        try:
            self.load_file = config.get('Files', 'load')
            for key, value in parents:
                self.load_file = self.load_file.replace(key, value)
            self.load_file = self.load_file.replace('$USER$', getUser())
            self.load_file = self.load_file.replace('$YEAR$', self.base_year)
        except:
            self.load_file = ''
        self.data_file = ''
        try:
            self.data_file = config.get('Files', 'data_file')
        except:
            try:
                self.data_file = config.get('Power', 'data_file')
            except:
                pass
        for key, value in parents:
           self.data_file = self.data_file.replace(key, value)
        self.data_file = self.data_file.replace('$USER$', getUser())
        self.data_file = self.data_file.replace('$YEAR$', self.base_year)
        try:
            helpfile = config.get('Files', 'help')
            for key, value in parents:
                helpfile = helpfile.replace(key, value)
            helpfile = helpfile.replace('$USER$', getUser())
            helpfile = helpfile.replace('$YEAR$', self.base_year)
        except:
            helpfile = ''
        try:
            progress_bar = config.get('View', 'progress_bar')
        except:
            progress_bar = True
#
#       choose what power data to collect (once only)
#
        self.plot_order = ['show_menu', 'actual', 'cumulative', 'by_station', 'adjust',
                           'show_load', 'shortfall', 'grid_losses', 'gross_load', 'save_plot', 'visualise',
                           'maximise', 'block', 'show_pct', 'by_day', 'by_month', 'by_season',
                           'by_period', 'hour', 'total', 'month', 'season', 'period',
                           'duration', 'augment', 'shortfall_detail', 'summary', 'save_data', 'save_detail',
                           'save_tech', 'save_balance', 'financials']
        self.initials = ['actual', 'by_station', 'grid_losses', 'save_data', 'gross_load',
                         'summary', 'financials', 'show_menu']
        self.hdrs = {'show_menu': 'Check / Uncheck all',
                'actual': 'Generation - use actual generation figures',
                'cumulative': 'Generation - total (cumulative)',
                'by_station': 'Generation - from chosen stations',
                'adjust': 'Generation - adjust generation',
                'show_load': 'Generation - show Load',
                'shortfall': 'Generation - show shortfall from Load',
                'grid_losses': 'Generation - reduce generation by grid losses',
                'gross_load': 'Add Existing Rooftop PV to Load (Gross Load)',
                'save_plot': 'Save plot data',
                'visualise': 'Visualise generation',
                'maximise': 'Maximise Plot windows',
                'block': 'Show plots one at a time',
                'show_pct': 'Show generation as a percentage of load',
                'by_day': 'Energy by day',
                'by_month': 'Energy by month',
                'by_season': 'Energy by season',
                'by_period': 'Energy by period',
                'hour': 'Power by hour',
                'total': 'Power - diurnal profile',
                'month': 'Power - diurnal profile by month',
                'season': 'Power - diurnal profile by season',
                'period': 'Power - diurnal profile by period',
                'duration': 'Power - Load duration',
                'augment': 'Power - augmented by hour',
                'shortfall_detail': 'Power - Shortfall analysis',
                'summary': 'Show Summary/Other Tables',
                'save_data': 'Save initial Hourly Data Output',
                'save_detail': 'Save Hourly Data Output by Station',
                'save_tech': 'Save Hourly Data Output by Technology',
                'save_balance': 'Save Powerbalance Inputs',
                'financials': 'Run Financial Models'}
        self.spacers = {'actual': 'Show in Plot',
                   'save_plot': 'Choose plots (all use a full year of data)',
                   'summary': 'Choose tables'}
        self.plots = {}
        for i in range(len(self.plot_order)):
            self.plots[self.plot_order[i]] = False
        self.technologies = ''
        self.load_growth = 0.
        self.storage = [0., 0.]
        self.recharge = [0., 1.]
        self.discharge = [0., 1.]
        self.load_year = self.base_year
        plot_opts = []
        try:
            plot_opts = config.items('Power')
        except:
            pass
        for key, value in plot_opts:
            if key in self.plots:
                if value.lower() in ['true', 'yes', 'on']:
                    self.plots[key] = True
            elif key == 'load_growth':
                if value[-1] == '%':
                    self.load_growth = float(value[:-1]) / 100.
                else:
                    self.load_growth = float(value)
            elif key == 'storage':
                if ',' in value:
                    bits = value.split(',')
                    self.storage = [float(bits[0].strip()), float(bits[1].strip())]
                else:
                    self.storage = [float(value), 0.]
            elif key == 'technologies':
                self.technologies = value
            elif key == 'shortfall_iterations':
                self.iterations = int(value)
        try:
            storage = config.get('Storage', 'storage')
            if ',' in storage:
                bits = storage.split(',')
                self.storage = [float(bits[0].strip()), float(bits[1].strip())]
            else:
                self.storage = [float(storage), 0.]
        except:
            pass
        try:
            self.show_menu = self.plots['show_menu']
        except:
            self.show_menu = True
        try:
            self.discharge[0] = float(config.get('Storage', 'discharge_max'))
            self.discharge[1] = float(config.get('Storage', 'discharge_eff'))
            if self.discharge[1] < 0.5:
                self.discharge[1] = 1 - self.discharge[1]
            self.recharge[0] = float(config.get('Storage', 'recharge_max'))
            self.recharge[1] = float(config.get('Storage', 'recharge_eff'))
            if self.recharge[1] < 0.5:
                self.recharge[1] = 1 - self.recharge[1]
        except:
            pass
        if __name__ == '__main__':
            self.show_menu = True
            self.plots['save_data'] = True
        if not self.plots['save_data']:
            self.plot_order.remove('save_data')
        if len(self.stations) == 1:
            self.plot_order.remove('cumulative')
            self.plot_order.remove('by_station')
            self.plot_order.remove('gross_load')
        if self.show_menu:
            if __name__ == '__main__':
                app = QtGui.QApplication(sys.argv)
            what_plots = whatPlots(self.plots, self.plot_order, self.hdrs, self.spacers,
                                   self.load_growth, self.base_year, self.load_year,
                                   self.iterations, self.storage, self.discharge,
                                   self.recharge, initial=True, helpfile=helpfile)
            what_plots.exec_()
            self.plots, self.load_growth, self.load_year, self.load_multiplier, self.iterations, \
              self.storage, self.discharge, self.recharge = what_plots.getValues()
            if self.plots is None:
                self.something.power_signal = None
                return
        self.selected = None
        if self.plots['by_station']:
            self.selected = []
            if len(stations) == 1:
                self.selected.append(stations[0].name)
            else:
                selected = whatStations(stations, self.plots['gross_load'],
                                        self.plots['actual'])
                selected.exec_()
                self.selected = selected.getValues()
                if self.selected is None:
                    return
#
#       collect the data (once only)
#
        if show_progress is not None:
            progress_bar = show_progress
        self.show_progress = True
        if isinstance(progress_bar, bool):
           if not progress_bar:
               self.show_progress = False
        else:
           if progress_bar.lower() in ['false', 'no', 'off']:
               self.show_progress = False
           else:
               ctr = 0
               for st in stations:
                   if st.technology[:6] != 'Fossil':
                       ctr += 1
               try:
                   if int(progress_bar) > ctr:
                       self.show_progress = False
               except:
                   pass
        self.stn_outs = []
        if self.show_progress:
            power = ProgressModel(stations, self.plots, True, year=self.base_year,
                                  selected=self.selected, status=self.status)
            power.open()
            if __name__ == '__main__':
                app.exec_()
            else:
                power.exec_()
            if len(power.power_summary) == 0:
                return
            self.power_summary = power.power_summary
            self.gen_pct = power.getPct()
            self.ly, self.x = power.getLy()
            if self.ly is None:
                return
            if self.plots['save_data'] or self.plots['financials'] or self.plots['save_detail']:
                self.stn_outs, self.stn_tech, self.stn_size, self.stn_pows, self.stn_grid, \
                  self.stn_path = power.getStnOuts()
            elif self.plots['save_tech'] or self.plots['save_balance']:
                self.stn_outs, self.stn_tech = power.getStnTech()
            elif self.plots['visualise']:
                self.stn_outs, self.stn_pows = power.getStnPows()
        else:
            self.model = SuperPower(stations, self.plots, False, year=self.base_year,
                                    selected=self.selected, status=status)
            self.model.getPower()
            if len(self.model.power_summary) == 0:
                return
            self.power_summary = self.model.power_summary
            self.ly, self.x = self.model.getLy()
            if self.plots['save_data'] or self.plots['financials'] or self.plots['save_detail']:
                self.stn_outs, self.stn_tech, self.stn_size, self.stn_pows, self.stn_grid, \
                  self.stn_path = self.model.getStnOuts()
            elif self.plots['save_tech'] or self.plots['save_balance']:
                self.stn_outs, self.stn_tech = self.model.getStnTech()
            elif self.plots['visualise']:
                self.stn_outs, self.stn_pows = self.model.getStnPows()
        self.suffix = ''
        if len(self.stations) == 1:
            self.suffix = ' - ' + self.stations[0].name
        elif len(self.stn_outs) == 1:
            self.suffix = ' - ' + self.stn_outs[0]
        elif self.plots['by_station']:
            if len(self.ly) == 1:
                self.suffix = ' - ' + self.ly.keys()[0]
            else:
                self.suffix = ' - Chosen Stations'
        if self.plots['save_data']:
            if self.data_file == '':
                data_file = 'Power_Table_%s.xls' % \
                            str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(),
                                'yyyy-MM-dd_hhmm'))
            else:
                data_file = self.data_file
            stnsh = {}
            for i in range(len(self.stn_outs)):
                stnsh[self.stn_outs[i]] = self.stn_pows[i][:]
            self.save_detail(data_file, stnsh)
            del stnsh
        if self.plots['summary']:
            fields = ['name', 'technology', 'capacity', 'cf', 'generation']
            sumfields = ['capacity', 'generation']
            if getattr(self.power_summary[0], 'transmitted') != None:
                fields.append('transmitted')
                sumfields.append('transmitted')
            dialog = displaytable.Table(self.power_summary, sumfields=sumfields,
                     units='capacity=MW generation=MWh transmitted=MWh', sumby='technology',
                     fields=fields, save_folder=self.scenarios)
            dialog.exec_()
            del dialog
        if self.plots['financials']:
            do_financials = True
        else:
            do_financials = False
        if self.plots['save_data'] or self.plots['summary']:
            show_summ = True
        else:
            show_summ = False
        do_plots = True
#
#       loop around processing plots
#
        if do_plots:
            if plt.get_backend() != 'TkAgg':
                plt.switch_backend('TkAgg')
            self.gen_pct = None
            self.load_data = None
            if os.path.exists(self.load_file):
                self.can_do_load = True
            else:
                self.plots['show_load'] = False
                self.plots['show_pct'] = False
                self.can_do_load = False
                self.initials.append('augment')
                self.initials.append('show_load')
                self.initials.append('show_pct')
            if self.plots['save_detail']:
                pass
            else:
                self.initials.append('save_detail')
                if not self.plots['save_tech']:
                    self.initials.append('save_tech')
                if not self.plots['visualise']:
                    self.initials.append('visualise')
            self.load_key = ''
            self.adjustby = None
            while True:
                if self.plots['visualise'] and self.something is not None:
                    vis2 = Visualise(self.stn_outs, self.stn_pows, self.something, year=self.base_year)
                    vis2.setWindowModality(Qt.Qt.WindowModal)
                    vis2.setWindowFlags(vis2.windowFlags() |
                                 Qt.Qt.WindowSystemMenuHint |
                                 Qt.Qt.WindowMinMaxButtonsHint)
                    vis2.exec_()
                wrkly = {}
                summs = {}
                if self.load_key != '':
                    try:
                        del wrkly[self.load_key]
                    except:
                        pass
                    self.load_key = ''
                if (self.plots['show_load'] or self.plots['save_balance'] or self.plots['shortfall']) and self.can_do_load:
                    if self.load_data is None:
                        tf = open(self.load_file, 'r')
                        lines = tf.readlines()
                        tf.close()
                        self.load_data = []
                        bit = lines[0].rstrip().split(',')
                        if len(bit) > 0: # multiple columns
                            for b in range(len(bit)):
                                if bit[b][:4].lower() == 'load':
                                    if bit[b].lower().find('kwh') > 0: # kWh not MWh
                                        for i in range(1, len(lines)):
                                            bit = lines[i].rstrip().split(',')
                                            self.load_data.append(float(bit[b]) * 0.001)
                                    else:
                                        for i in range(1, len(lines)):
                                            bit = lines[i].rstrip().split(',')
                                            self.load_data.append(float(bit[b]))
                        else:
                            for i in range(1, len(lines)):
                                self.load_data.append(float(lines[i].rstrip()))
                    if self.load_multiplier != 0:
                        key = 'Load ' + self.load_year
                    else:
                        key = 'Load'  # lines[0].rstrip()
                    self.load_key = key
                    wrkly[key] = []
                    if self.load_multiplier != 0:
                        for i in range(len(self.load_data)):
                            wrkly[key].append(self.load_data[i] * (1 + self.load_multiplier))
                    else:
                        wrkly[key] = self.load_data[:]
                else:
                    self.plots['show_pct'] = False
                if self.plots['adjust']:
                    if self.load_key == '':
                        if self.adjustby is None:
                            adjust = Adjustments(self.ly.keys())
                        else:
                            adjust = Adjustments(self.adjustby)
                    else:
                        if self.adjustby is None:
                            adjust = Adjustments(self.ly.keys(), self.load_key, wrkly[self.load_key], self.ly, self.load_year)
                        else:
                            adjust = Adjustments(self.adjustby, self.load_key, wrkly[self.load_key], self.ly, self.load_year)
                    adjust.exec_()
                    self.adjustby = adjust.getValues()
                else:
                    self.adjustby = None
                for key in self.ly:
                    if self.adjustby is None:
                        wrkly[key] = self.ly[key][:]
                    else:
                        wrkly[key] = []
                        if key == 'Generation':
                            for i in range(len(self.ly[key])):
                                wrkly[key].append(self.ly[key][i])
                        else:
                            for i in range(len(self.ly[key])):
                                wrkly[key].append(self.ly[key][i] * self.adjustby[key])
                if self.plots['shortfall'] or self.plots['shortfall_detail'] or self.plots['save_balance']:
                    self.plots['show_load'] = True
                    self.plots['cumulative'] = True
                try:
                    del wrkly['Storage']
                except:
                    pass
                if self.load_data is None:
                    self.do_load = False
                else:
                    self.do_load = True
                if self.plots['show_load'] and self.storage[0] > 0:
                    storage_cap = self.storage[0] * 1000.
                    if self.storage[1] > self.storage[0]:
                        storage_carry = self.storage[0] * 1000.
                    else:
                        storage_carry = self.storage[1] * 1000.
                    total_gen = [0.]
                    storage_bal = []
                    storage_losses = []
                    wrkly['Storage'] = [0.]
                    wrkly['Excess'] = [0.]
                    for i in range(1, len(self.x)):
                        wrkly['Storage'].append(0.)
                        wrkly['Excess'].append(0.)
                        total_gen.append(0.)
                    for key, value in wrkly.iteritems():
                        if key == 'Generation':
                            continue
                        if key == 'Storage' or key == 'Excess':
                            continue
                        elif key[:4] == 'Load':
                            pass
                        else:
                            for i in range(len(value)):
                                total_gen[i] += value[i]
                    for i in range(len(self.x)):
                        gap = gape = total_gen[i] - wrkly[self.load_key][i]
                        storage_loss = 0.
                        if gap >= 0:  # excess generation
                            if self.recharge[0] > 0 and gap > self.recharge[0]:
                                gap = self.recharge[0]
                            if storage_carry >= storage_cap:
                                wrkly['Excess'][i] = gape
                            else:
                                if storage_carry + gap > storage_cap:
                                    gap = (storage_cap - storage_carry) * (1 / self.recharge[1])
                                storage_loss = gap - gap * self.recharge[1]
                                storage_carry += gap - storage_loss
                                if gape - gap > 0:
                                    wrkly['Excess'][i] = gape - gap
                                if storage_carry > storage_cap:
                                    storage_carry = storage_cap
                        else:
                            if self.discharge[0] > 0 and -gap > self.discharge[0]:
                                gap = -self.discharge[0]
                            if storage_carry > -gap / self.discharge[1]:  # extra storage
                                wrkly['Storage'][i] = -gap
                                storage_loss = gap * self.discharge[1] - gap
                                storage_carry += gap - storage_loss
                            else:  # not enough storage
                                if self.discharge[0] > 0 and storage_carry > self.discharge[0]:
                                    storage_carry = self.discharge[0]
                                wrkly['Storage'][i] = storage_carry * (1 / (2 - self.discharge[1]))
                                storage_loss = storage_carry - wrkly['Storage'][i]
                                storage_carry = 0
                        storage_bal.append(storage_carry)
                        storage_losses.append(storage_loss)
                    if show_summ:
                        shortstuff = []
                        summs['Shortfall'] = [0., '', 0]
                        for i in range(len(self.x)):
                            shortfall = total_gen[i] + wrkly['Storage'][i] - wrkly[self.load_key][i]
                            if shortfall > 0:
                                shortfall = 0
                            summs['Shortfall'][0] += shortfall
                            shortstuff.append(ColumnData(i + 1, the_date(self.load_year, i),
                                              [wrkly[self.load_key][i], total_gen[i],
                                              wrkly['Storage'][i], storage_losses[i],
                                              storage_bal[i], shortfall, wrkly['Excess'][i]],
                                              values=['load', 'generation', 'storage_used',
                                                      'storage_loss', 'storage_balance',
                                                      'shortfall', 'excess']))
                        dialog = displaytable.Table(shortstuff, title='Storage',
                                                    save_folder=self.scenarios,
                                                    fields=['hour', 'period', 'load', 'generation',
                                                            'storage_used', 'storage_loss',
                                                            'storage_balance', 'shortfall', 'excess'])
                        dialog.exec_()
                        del dialog
                        del shortstuff
                        summs['Shortfall'][0] = round(summs['Shortfall'][0], 1)
                if show_summ and self.adjustby is not None:
                    keys = []
                    for key in wrkly:
                        keys.append(key)
                        gen = 0.
                        for i in range(len(wrkly[key])):
                            gen += wrkly[key][i]
                        if key[:4] == 'Load':
                            incr = 1 + self.load_multiplier
                        else:
                            try:
                                incr = self.adjustby[key]
                            except:
                                incr = ''
                        try:
                            summs[key] = [round(gen, 1), round(incr, 2), 0]
                        except:
                            summs[key] = [round(gen, 1), '', 0]
                    keys.sort()
                    xtra = ['Generation', 'Load', 'Gen. - Load', 'Storage Capacity', 'Storage', 'Excess', 'Shortfall']
                    o = 0
                    gen = 0.
                    if self.storage[0] > 0:
                        summs['Storage Capacity'] = [self.storage[0] * 1000., '', 0]
                    for i in range(len(keys)):
                        if keys[i][:4] == 'Load':
                            xtra[1] = keys[i]
                        elif keys[i] in xtra:
                            continue
                        else:
                            o += 1
                            summs[keys[i]][2] = o
                            gen += summs[keys[i]][0]
                    if xtra[0] not in summs.keys():
                        summs[xtra[0]] = [gen, '', 0]
                    if xtra[1] in summs.keys():
                        summs[xtra[2]] = [round(gen - summs[xtra[1]][0], 1), '', 0]
                    for i in range(len(xtra)):
                        if xtra[i] in summs.keys():
                            o += 1
                            summs[xtra[i]][2] = o
                    try:
                        summs['Storage Used'] = summs.pop('Storage')
                    except:
                        pass
                    try:
                        summs['Excess Gen.'] = summs.pop('Excess')
                    except:
                        pass
                    dialog = displaytable.Table(summs, title='Generation Summary',
                                                save_folder=self.scenarios,
                                                fields=['component', 'generation', 'multiplier', 'row'],
                                                units='generation=MWh', sortby='row')
                    dialog.exec_()
                    del dialog
                if self.plots['save_detail'] or self.plots['save_tech']:
                    dos = []
                    if self.plots['save_detail']:
                        dos.append('')
                    if self.plots['save_tech']:
                        dos.append('Technology_')
                    for l in range(len(dos)):
                        if self.data_file == '':
                            if year is None:
                                data_file = 'Power_Detail_%s%s.xls' % ( dos[l] ,
                                str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(),
                                    'yyyy-MM-dd_hhmm')))
                            else:
                                data_file = 'Power_Detail_%s%s_%s.xls' % ( dos[l] , str(year),
                                str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(),
                                    'yyyy-MM-dd_hhmm')))
                        else:
                             data_file = self.data_file
                        keys = []
                        keys2 = []
                        if dos[l] != '':
                            techs = {}
                            for key, value in iter(wrkly.iteritems()):
                                try:
                                    i = self.stn_outs.index(key)
                                    if self.stn_tech[i] in techs.keys():
                                        for j in range(len(value)):
                                            techs[self.stn_tech[i]][j] += value[j]
                                    else:
                                        techs[self.stn_tech[i]] = value[:]
                                        keys.append(self.stn_tech[i])
                                except:
                                    techs[key] = value[:]
                                    keys2.append(key)
                            keys.sort()
                            keys2.extend(keys) # put Load first
                            self.save_detail(data_file, techs, keys=keys2)
                            del techs
                        else:
                            for key in wrkly.keys():
                                try:
                                    i = self.stn_outs.index(key)
                                    keys.append(self.stn_outs[i])
                                except:
                                    keys2.append(key)
                            keys.sort()
                            keys2.extend(keys) # put Load first
                            self.save_detail(data_file, wrkly, keys=keys2)
                self.showGraphs(wrkly, self.x)
                if __name__ == '__main__':
                    self.show_menu = True
                    self.plots['save_data'] = True
                if self.show_menu:
                    what_plots = whatPlots(self.plots, self.plot_order, self.hdrs, self.spacers,
                                 self.load_growth, self.base_year, self.load_year, self.iterations,
                                 self.storage, self.discharge, self.recharge, self.initials, helpfile=helpfile)
                    what_plots.exec_()
                    self.plots, self.load_growth, self.load_year, self.load_multiplier, \
                        self.iterations, self.storage, self.discharge, self.recharge = what_plots.getValues()
                    if self.plots is None:
                        break
                else:
                    break
#
#       loop around doing financials
#
         # run the financials model
        if do_financials:
            while True:
                self.financials = FinancialModel(self.stn_outs, self.stn_tech, self.stn_size,
                                  self.stn_pows, self.stn_grid, self.stn_path, year=self.base_year)
                if self.financials.stations is None:
                    break
                fin_fields = ['name', 'technology', 'capacity', 'generation', 'cf',
                              'capital_cost', 'lcoe_real', 'lcoe_nominal', 'npv']
                fin_sumfields = ['capacity', 'generation', 'capital_cost', 'npv']
                fin_units = 'capacity=MW generation=MWh capital_cost=$ lcoe_real=c/kWh' + \
                            ' lcoe_nominal=c/kWh npv=$'
                for stn in self.financials.stations:
                    if stn.grid_cost > 0:
                        i = fin_fields.index('capital_cost')
                        fin_fields.insert(i + 1, 'grid_cost')
                        fin_sumfields.append('grid_cost')
                        fin_units += ' grid_cost=$'
                        break
                dialog = displaytable.Table(self.financials.stations, fields=fin_fields,
                         sumfields=fin_sumfields, units=fin_units, sumby='technology',
                         save_folder=self.scenarios, title='Financials')
                dialog.exec_()
                del dialog
        self.something.power_signal = None

    def getValues(self):
        try:
            return self.power_summary
        except:
            return None

    def getPct(self):
        return self.gen_pct

    @QtCore.pyqtSlot()
    def exit(self):
        self.something.power_signal = None
        return #exit()
