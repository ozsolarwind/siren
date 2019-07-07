#!/usr/bin/python3
#
#  Copyright (C) 2015-2019 Sustainable Energy Now Inc., Angus King
#
#  powerclasses.py - This file is part of SIREN.
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

import numpy as np
import os
import sys
import ssc
import xlrd

import configparser  # decode .ini file
from PyQt4 import Qt, QtGui, QtCore

from senuser import getUser, techClean
import displayobject
from editini import SaveIni
from grid import Grid
from parents import getParents
from sirenicons import Icons

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
            if float(bit) == 0:
                try:
                    arry.append(int(bit[:bit.find('.')]))
                except:
                    arry.append(int(bit))
            else:
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
                if float(bit) == 0:
                    mtrx[-1].append(int(bit))
                else:
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
                    self.dischargeSpin.setDecimals(3)
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
                    self.rechargeSpin.setDecimals(3)
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
                elif self.plot_order[plot] == 'summary' and 'shortfall' in self.plot_order:  # fudge to add in iterations stuff
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
        if self.grid.geometry().height() > (screen.height() / 1.2):
            h = int(screen.height() * .9)
        else:
            h = int(self.grid.geometry().height() * 1.07)
        self.resize(600, h)
        self.setWindowTitle('SIREN - Power dialog for ' + str(self.base_year))
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
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
                if self.checkbox[i].isEnabled():
                    self.checkbox[i].setCheckState(QtCore.Qt.Checked)
        else:
            for i in range(len(self.checkbox)):
                if self.plot_order[i] == 'show_menu':
                    continue
                if not self.initial:
                    if self.plot_order[i] in self.initials:
                        continue
                if self.checkbox[i].isEnabled():
                    self.checkbox[i].setCheckState(QtCore.Qt.Unchecked)

    def check_balance(self, event):
        if event:
            for i in range(len(self.checkbox)):
                if self.plot_order[i] == 'show_load': # or self.plot_order[i] == 'grid_losses':
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
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
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
    def __init__(self, parms=None, helpfile=None):
        super(whatFinancials, self).__init__()
        self.proceed = False
        self.helpfile = helpfile
        self.financials = [['analysis_period', 'Analysis period (years)', 0, 50, 30],
                      ['federal_tax_rate', 'Federal tax rate (%)', 0, 30., 30.],
                      ['real_discount_rate', 'Real discount rate (%)', 0, 20., 0],
                      ['inflation_rate', 'Inflation rate (%)', 0, 20., 0],
                      ['insurance_rate', 'Insurance rate (%)', 0, 15., 0],
                      ['loan_term', 'Loan term (years)', 0, 60., 0],
                      ['loan_rate', 'Loan rate (%)', 0, 30., 0],
                      ['debt_fraction', 'Debt percentage (%)', 0, 100, 0],
                      ['depr_fed_type', 'Federal depreciation type 2=straight line', 0, 2, 2],
                      ['depr_fed_sl_years', 'Depreciation straight-line term (years)', 0, 60, 20],
                      ['market', 'Commercial PPA (on), Utility IPP (off)', 0, 1, 0],
                   #   ['bid_price_esc', 'Bid Price escalation (%)', 0, 100., 0],
                      ['salvage_percentage', 'Salvage value percentage (%)', 0, 100., 0],
                      ['min_dscr_target', 'Minimum required DSCR (ratio)', 0, 5., 1.4],
                      ['min_irr_target', 'Minimum required IRR (%)', 0, 30., 15.],
                   #   ['ppa_escalation', 'PPA escalation (%)', 0, 100., 0.6],
                      ['min_dscr_required', 'Minimum DSCR required?', 0, 1, 1],
                      ['positive_cashflow_required', 'Positive cash flow required?', 0, 1, 1],
                      ['optimize_lcoe_wrt_debt_fraction', 'Optimize LCOE with respect to debt' +
                       ' percent?', 0, 1, 0],
                   #   ['optimize_lcoe_wrt_ppa_escalation', 'Optimize LCOE with respect to PPA' +
                   #    ' escalation?', 0, 1, 0],
                      ['grid_losses', 'Reduce power by Grid losses?', False, True, False],
                      ['grid_costs', 'Include Grid costs in LCOE?', False, True, False],
                      ['grid_path_costs', 'Include full grid path in LCOE?', False, True, False]]
        if parms is None:
            config = configparser.RawConfigParser()
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
        else:
            self.financials = parms
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
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
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

    def getParms(self):
# This is in some sense a repeat of getvalues except that's a dictionary and this is a list
# It's inefficeint and untidy but in another time I'll look at it
        for i in range(len(self.spin)):
            for item in self.financials:
                if item[1] == self.labels[i].text():
                    item[4] = self.spin[i].value()
        for i in range(len(self.checkbox)):
            for item in self.financials:
                if item[1] == self.checkbox[i].text():
                    if isinstance(item[2], bool) and isinstance(item[3], bool):
                        if self.checkbox[i].checkState() == QtCore.Qt.Checked:
                            item[4] = 1
                        else:
                            item[4] = 0
                    else:
                        if self.checkbox[i].checkState() == QtCore.Qt.Checked:
                            item[4] = 1
                        else:
                            item[4] = 0
                    break
        return self.financials

class Adjustments(QtGui.QDialog):
    def __init__(self, keys, load_key=None, load=None, data=None, base_year=None):
        super(Adjustments, self).__init__()
        config = configparser.RawConfigParser()
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
            for key in list(data.keys()):
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
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
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
        B = np.array(load)
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
        A = np.array(gen)
        res = np.linalg.lstsq(A, B, rcond=None)  # least squares solution of the generation vs load
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
            corr = np.corrcoef(load, cgen)
            self.optmsg.setText('Corr: %.2f' % corr[0][1])
        self.zeroCheck()

    def showClicked(self):
        self.results = {}
        for key in list(self.adjusts.keys()):
            self.results[key] = round(self.adjusts[key].value(), 2)
        self.close()

    def getValues(self):
        return self.results


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
        try:
            workfile = xlrd.open_workbook(xl_file)
        except:
            return None, None
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
                    data.set_string(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'),
                    worksheet.cell_value(curr_row, var['DEFAULT']).encode('utf-8'))
                elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_ARRAY':
                    arry = split_array(worksheet.cell_value(curr_row, var['DEFAULT']))
                    data.set_array(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'), arry)
                elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_NUMBER':
                    if overrides is not None and worksheet.cell_value(curr_row, var['NAME']) \
                      in overrides:
                        if worksheet.cell_value(curr_row, var['DATA']) == 'SSC_ARRAY':
                            if type(overrides[worksheet.cell_value(curr_row, var['NAME'])]) is list:
                                data.set_array(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'),
                                  overrides[worksheet.cell_value(curr_row, var['NAME'])])
                            else:
                                data.set_array(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'),
                                  [overrides[worksheet.cell_value(curr_row, var['NAME'])]])
                        else:
                            data.set_number(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'),
                              overrides[worksheet.cell_value(curr_row, var['NAME'])])
                    else:
                        if isinstance(worksheet.cell_value(curr_row, var['DEFAULT']), float):
                            data.set_number(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'),
                              float(worksheet.cell_value(curr_row, var['DEFAULT'])))
                        else:
                            data.set_number(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'),
                              worksheet.cell_value(curr_row, int(var['DEFAULT'])))
                elif worksheet.cell_value(curr_row, var['DATA']) == 'SSC_MATRIX':
                    mtrx = split_matrix(worksheet.cell_value(curr_row, var['DEFAULT']))
                    data.set_matrix(worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'), mtrx)
            elif worksheet.cell_value(curr_row, var['TYPE']) == 'SSC_OUTPUT':
                output_variables.append([worksheet.cell_value(curr_row, var['NAME']).encode('utf-8'),
                                        worksheet.cell_value(curr_row, var['DATA'])])
        return data, output_variables

    def __init__(self, name, technology, capacity, power, grid, path, year=None, status=None, parms=None):
        def set_grid_variables():
            self.dispatchable = None
            self.grid_line_loss = 0.
            self.subs_cost = 0.
            self.grid_subs_loss = 0.
            try:
                itm = config.get('Grid', 'dispatchable')
                self.dispatchable = techClean(itm)
                line_loss = config.get('Grid', 'line_loss')
                if line_loss[-1] == '%':
                    self.grid_line_loss = float(line_loss[:-1]) / 100000.
                else:
                    self.grid_line_loss = float(line_loss) / 1000.
                line_loss = config.get('Grid', 'substation_loss')
                if line_loss[-1] == '%':
                    self.grid_subs_loss = float(line_loss[:-1]) / 100.
                else:
                    self.grid_subs_loss = float(line_loss)
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
            ippppa_data.set_number(b'system_capacity', capacity[stn] * 1000)
            ippppa_data.set_array(b'gen', net_hourly)
            ippppa_data.set_number(b'construction_financing_cost', capital_cost + grid_cost)
            ippppa_data.set_number(b'total_installed_cost', capital_cost + grid_cost)
            ippppa_data.set_array(b'om_capacity', [costs[technology[stn]][1]])
            if technology[stn] == 'Biomass':
                ippppa_data.set_number(b'om_opt_fuel_1_usage', self.biomass_multiplier
                                       * capacity[stn] * 1000)
                ippppa_data.set_array(b'om_opt_fuel_1_cost', [costs[technology[stn]][2]])
                ippppa_data.set_number(b'om_opt_fuel_1_cost_escal',
                                       ippppa_data.get_number(b'inflation_rate'))
            module = ssc.Module(b'ippppa')
            if (module.exec_(ippppa_data)):
             # return the relevant outputs desired
                energy = ippppa_data.get_array(b'gen')
                generation = 0.
                for i in range(len(energy)):
                    generation += energy[i]
                generation = generation * pow(10, -3)
                lcoe_real = ippppa_data.get_number(b'lcoe_real')
                lcoe_nom = ippppa_data.get_number(b'lcoe_nom')
                npv = ippppa_data.get_number(b'npv')
                self.stations.append(FinancialSummary(name[stn], technology[stn], capacity[stn],
                  generation, 0, round(capital_cost), lcoe_real, lcoe_nom, npv, round(grid_cost)))
            else:
                if self.status:
                   self.status.emit(QtCore.SIGNAL('log'), 'Errors encountered processing ' + name[stn])
                   QtGui.qApp.processEvents()
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    if self.status:
                       self.status.emit(QtCore.SIGNAL('log'), 'ippppa error [' + str(idx) + ']: ' + msg.decode())
                    else:
                        print('ippppa error [', idx, ' ]: ', msg.decode())
                    idx += 1
                    msg = module.log(idx)
            del module

        self.stations = []
        self.status = status
        self.parms = parms
        config = configparser.RawConfigParser()
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
            parents = getParents(config.items('Parents'))
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
        if annual_data is None:
            if self.status:
                self.status.emit(QtCore.SIGNAL('log'), 'Error accessing ' + annual_file)
            else:
                print('Error accessing ' + annual_file)
            self.stations = None
            return
        what_beans = whatFinancials(parms=self.parms, helpfile=self.helpfile)
        what_beans.exec_()
        ippas = what_beans.getValues()
        self.parms = what_beans.getParms()
        if ippas is None:
            self.stations = None
            return
        ssc_api = ssc.API()
# to suppress messages
        if not self.expert:
            ssc_api.set_print(0)
        ippppa_data, ippppa_outputs = self.get_variables(ippppa_file, overrides=ippas)
        if ippppa_data is None:
            if self.status:
                self.status.emit(QtCore.SIGNAL('log'), 'Error accessing ' + ippppa_file)
            else:
                print('Error accessing ' + ippppa_file)
            self.stations = None
            return
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
                        energy.append(power[stn][hr] * 1000 * (1 - self.grid_line_loss * path[stn] -
                                      self.grid_subs_loss))
                else:
                    for hr in range(len(power[stn])):
                        energy.append(power[stn][hr] * 1000 * (1 - self.grid_line_loss * grid[stn] -
                                      self.grid_subs_loss))
            else:
                for hr in range(len(power[stn])):
                    energy.append(power[stn][hr] * 1000)
            annual_data.set_array(b'system_hourly_energy', energy)
            net_hourly = None
            module = ssc.Module(b'annualoutput')
            if (module.exec_(annual_data)):
             # return the relevant outputs desired
                net_hourly = annual_data.get_array(b'hourly_energy')
                net_annual = annual_data.get_array(b'annual_energy')
                degradation = annual_data.get_array(b'annual_degradation')
                del module
                do_ippppa()
            else:
                if self.status:
                   self.status.emit(QtCore.SIGNAL('log'), 'Errors encountered processing ' + name[stn])
                idx = 0
                msg = module.log(idx)
                while (msg is not None):
                    if self.status:
                       self.status.emit(QtCore.SIGNAL('log'), 'annualoutput error [' + str(idx) + ']: ' + msg.decode())
                    else:
                        print('annualoutput error [', idx, ' ]: ', msg.decode())
                    idx += 1
                    msg = module.log(idx)
                del module

    def getValues(self):
        return self.stations

    def getParms(self):
        return self.parms
