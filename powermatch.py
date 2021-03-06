#!/usr/bin/python3
#
#  Copyright (C) 2018-2021 Sustainable Energy Now Inc., Angus King
#
#  powermatch.py - This file is part of SIREN.
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
# Add in to check maximum capacity used by the technologies
import os
import sys
import time
from PyQt5 import QtCore, QtGui, QtWidgets
import displayobject
import displaytable
from credits import fileVersion
from math import log10
import matplotlib
matplotlib.use('TkAgg') # so PyQT5 and Matplotlib windows don't interfere
import matplotlib.pyplot as plt
# This import registers the 3D projection, but is otherwise unused.
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.ticker import LinearLocator, FormatStrFormatter
import numpy as np
import openpyxl as oxl
import random
# from openpyxl.utils import get_column_letter
from parents import getParents
from senuser import getUser
from editini import SaveIni
from floaters import ProgressBar, FloatStatus
from getmodels import getModelFile
import xlrd
import configparser  # decode .ini file
from zoompan import ZoomPanX
try:
    from opt_debug import optimiseDebug
except:
    pass

tech_names = ['Load', 'Onshore Wind', 'Offshore Wind', 'Rooftop PV', 'Fixed PV', 'Single Axis PV',
              'Dual Axis PV', 'Biomass', 'Geothermal', 'Other1', 'CST', 'Shortfall']
# same order as self.file_labels
C = 0 # Constraints
G = 1 # Generators
O = 2 # Optimisation
D = 3 # Data
R = 4 # Results
col_letters = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ'
def ss_col(col, base=1):
    if base == 1:
        col -= 1
    c1 = col // 26
    c2 = col % 26
    return (col_letters[c1] + col_letters[c2 + 1]).strip()


class ListWidget(QtWidgets.QListWidget):
    def decode_data(self, bytearray):
        data = []
        ds = QtCore.QDataStream(bytearray)
        while not ds.atEnd():
            row = ds.readInt32()
            column = ds.readInt32()
            map_items = ds.readInt32()
            for i in range(map_items):
                key = ds.readInt32()
                value = QtCore.QVariant()
                ds >> value
                data.append(value.value())
        return data

    def __init__(self, parent=None):
        super(ListWidget, self).__init__(parent)
        self.setDragDropMode(self.DragDrop)
        self.setSelectionMode(self.ExtendedSelection)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            super(ListWidget, self).dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(QtCore.Qt.CopyAction)
            event.accept()
        else:
            super(ListWidget, self).dragMoveEvent(event)

    def dropEvent(self, event):
        if event.source() == self:
            event.setDropAction(QtCore.Qt.MoveAction)
            QtWidgets.QListWidget.dropEvent(self, event)
        else:
            ba = event.mimeData().data('application/x-qabstractitemmodeldatalist')
            data_items = self.decode_data(ba)
            event.setDropAction(QtCore.Qt.MoveAction)
            event.source().deleteItems(data_items)
            super(ListWidget, self).dropEvent(event)

    def deleteItems(self, items):
        for row in range(self.count() -1, -1, -1):
            if self.item(row).text() in items:
             #   r = self.row(item)
                self.takeItem(row)


class ClickableQLabel(QtWidgets.QLabel):
    clicked = QtCore.pyqtSignal()

    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtWidgets.QApplication.widgetAt(event.globalPos()).setFocus()
        self.clicked.emit()


class Constraint:
    def __init__(self, name, category, capacity_min, capacity_max, rampup_max, rampdown_max,
                 recharge_max, recharge_loss, discharge_max, discharge_loss, parasitic_loss,
                 min_run_time, warm_time):
        self.name = name
        self.category = category
        try:
            self.capacity_min = float(capacity_min) # minimum run_rate for generator; don't drain below for storage
        except:
            self.capacity_min = 0.
        try:
            self.capacity_max = float(capacity_max) # maximum run_rate for generator; don't drain more than this for storage
        except:
            self.capacity_max = 1.
        try:
            self.recharge_max = float(recharge_max) # can't charge more than this per hour
        except:
            self.recharge_max = 1.
        try:
            self.recharge_loss = float(recharge_loss)
        except:
            self.recharge_loss = 0.
        try:
            self.discharge_max = float(discharge_max) # can't discharge more than this per hour
        except:
            self.discharge_max = 1.
        try:
            self.discharge_loss = float(discharge_loss)
        except:
            self.discharge_loss = 0.
        try:
            self.parasitic_loss = float(parasitic_loss) # daily parasitic loss / hourly ?
        except:
            self.parasitic_loss = 0.
        try:
            self.rampup_max = float(rampup_max)
        except:
            self.rampup_max = 1.
        try:
            self.rampdown_max = float(rampdown_max)
        except:
            self.rampdown_max = 1.
        try:
            self.min_run_time = int(min_run_time)
        except:
            self.min_run_time = 0
        try:
            self.warm_time = float(warm_time)
            if self.warm_time >= 1:
                self.warm_time = self.warm_time / 60
                if self.warm_time > 1:
                    self.warm_time = 1
            elif self.warm_time > 0:
                if self.warm_time <= 1 / 24.:
                    self.warm_time = self.warm_time * 24
        except:
            self.warm_time = 0


class Facility:
    def __init__(self, name, order, constraint, capacity, lcoe, lcoe_cf, emissions, initial=None):
        self.name = name
        self.constraint = constraint
        try:
            self.order = int(order)
        except:
            self.order = 0.
        try:
            self.capacity = float(capacity)
        except:
            self.capacity = 0.
        try:
            self.lcoe = float(lcoe)
        except:
            self.lcoe = 0.
        try:
            self.lcoe_cf = float(lcoe_cf)
        except:
            self.lcoe_cf = 0.
        try:
            self.emissions = float(emissions)
        except:
            self.emissions = 0.
        try:
            self.initial = float(initial)
        except:
            self.initial = 0.


class Optimisation:
    def __init__(self, name, approach, values): #capacity=None, cap_min=None, cap_max=None, cap_step=None, caps=None):
        self.name = name
        self.approach = approach
        if approach == 'Discrete':
            caps = values.split()
            self.capacities = []
            cap_max = 0.
            for cap in caps:
                try:
                    self.capacities.append(float(cap))
                    cap_max += float(cap)
                except:
                    pass
            self.capacity_min = 0
            self.capacity_max = round(cap_max, 3)
            self.capacity_step = None
        elif approach == 'Range':
            caps = values.split()
            try:
                self.capacity_min = float(caps[0])
            except:
                self.capacity_min = 0.
            try:
                self.capacity_max = float(caps[1])
            except:
                self.capacity_max = 0.
            try:
                self.capacity_step = float(caps[2])
            except:
                self.capacity_step = 0.
            self.capacities = None
        else:
            self.capacity_min = 0.
            self.capacity_max = 0.
            self.capacity_step = 0.
            self.capacities = None
        self.capacity = 0.


class Adjustments(QtWidgets.QDialog):
    def setAdjValue(self, key, typ, capacity):
        if key != 'Load':
            mw = capacity * self.adjustm[key]
            if typ == 'Storage':
                unit = 'MWh'
            else:
                unit = 'MW'
            fmtstr = self.fmtstr
            dp = self.decpts
        else:
            if self.adjusts[key].value() <= 0:
                return 0, '0 MW', '0'
            dimen = log10(capacity * self.adjusts[key].value())
            unit = 'MWh'
            if dimen > 11:
                unit = 'PWh'
                div = 9
            elif dimen > 8:
                unit = 'TWh'
                div = 6
            elif dimen > 5:
                unit = 'GWh'
                div = 3
            else:
                div = 0
            mw = (capacity * self.adjusts[key].value()) / pow(10, div)
            fmtstr = '{:,.0f}'
            dp = None
        mwtxt = fmtstr.format(mw) + ' ' + unit
        mwstr = str(round(mw, dp))
        return mw, mwtxt, mwstr

    def __init__(self, parent, data, adjustin, adjust_cap, save_folder=None):
        super(Adjustments, self).__init__()
        self.adjustt = {}
        self.adjustm = {} # (actual) adjust multiplier
        self.adjusts = {} # multiplier widget (rounded to 4 digits)
        self.labels = {} # string with calculated capacity
        self.adjustval = {} # adjust capacity input
        self.save_folder = save_folder
        self.ignore = False
        self.results = None
        self.grid = QtWidgets.QGridLayout()
        self.data = {}
        ctr = 0
        self.decpts = None
        for key, typ, capacity in data:
            if key == 'Load' or capacity is None:
                continue
            dimen = log10(capacity)
            if dimen < 2.:
                if dimen < 1.:
                    self.decpts = 2
                elif self.decpts != 2:
                    self.decpts = 1
        if self.decpts is None:
            self.fmtstr = '{:,.0f}'
        else:
            self.fmtstr = '{:,.' + str(self.decpts) + 'f}'
        for key, typ, capacity in data:
            self.adjustt[key] = typ
            if key != 'Load' and capacity is None:
                continue
            self.adjusts[key] = QtWidgets.QDoubleSpinBox()
            self.adjusts[key].setRange(0, adjust_cap)
            self.adjusts[key].setDecimals(4)
            try:
                self.adjustm[key] = adjustin[key]
                self.adjusts[key].setValue(round(adjustin[key], 4))
            except:
                self.adjustm[key] = 1.
                self.adjusts[key].setValue(1.)
            self.data[key] = capacity
            self.adjusts[key].setSingleStep(.1)
            self.adjusts[key].setObjectName(key)
            self.grid.addWidget(QtWidgets.QLabel(key), ctr, 0)
            self.grid.addWidget(self.adjusts[key], ctr, 1)
            self.adjusts[key].valueChanged.connect(self.adjust)
            self.labels[key] = QtWidgets.QLabel('')
            self.labels[key].setObjectName(key + 'label')
            mw, mwtxt, mwstr = self.setAdjValue(key, typ, capacity)
            self.labels[key].setText(mwtxt)
            self.grid.addWidget(self.labels[key], ctr, 2)
            self.adjustval[key] = QtWidgets.QLineEdit()
            self.adjustval[key].setObjectName(key)
            self.adjustval[key].setText(mwstr)
            if self.decpts is None:
                self.adjustval[key].setValidator(QtGui.QIntValidator())
            else:
                self.adjustval[key].setValidator(QtGui.QDoubleValidator())
            self.grid.addWidget(self.adjustval[key], ctr, 3)
            self.adjustval[key].textChanged.connect(self.adjustst)
            ctr += 1
        quit = QtWidgets.QPushButton('Quit', self)
        self.grid.addWidget(quit, ctr, 0)
        quit.clicked.connect(self.quitClicked)
        show = QtWidgets.QPushButton('Proceed', self)
        self.grid.addWidget(show, ctr, 1)
        show.clicked.connect(self.showClicked)
        reset = QtWidgets.QPushButton('Reset', self)
        self.grid.addWidget(reset, ctr, 2)
        reset.clicked.connect(self.resetClicked)
        if save_folder is not None:
            save = QtWidgets.QPushButton('Save', self)
            self.grid.addWidget(save, ctr, 3)
            save.clicked.connect(self.saveClicked)
            restore = QtWidgets.QPushButton('Restore', self)
            self.grid.addWidget(restore, ctr, 4)
            restore.clicked.connect(self.restoreClicked)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Powermatch - Adjust generators')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        self.show()

    def adjust(self):
        key = self.sender().objectName()
        if not self.ignore:
            self.adjustm[key] = self.adjusts[key].value()
        mw, mwtxt, mwstr = self.setAdjValue(key, self.adjustt[key], self.data[key])
        self.labels[key].setText(mwtxt)
        if not self.ignore:
            self.adjustval[key].setText(mwstr)
        self.ignore = False

    def adjustst(self):
        if self.ignore:
            return
        key = self.sender().objectName()
        if self.decpts is None:
            value = int(self.sender().text())
        else:
            try:
                value = float(self.sender().text())
                value = round(value, self.decpts)
            except:
                value = 0
        if key != 'Load':
            adj = value / self.data[key]
         #   self.adjusts[key].setValue(adj)
        else:
            dimen = log10(self.data[key])
            if dimen > 11:
                mul = 9
            elif dimen > 8:
                mul = 6
            elif dimen > 5:
                mul = 3
            else:
                mul = 0
            adj = (value * pow(10, mul)) / self.data[key]
        self.adjustm[key] = adj
        self.ignore = True
        self.adjusts[key].setValue(round(adj, 4))
        self.ignore = False

    def closeEvent(self, event):
        event.accept()

    def quitClicked(self):
        self.close()

    def resetClicked(self):
        for key in self.adjusts.keys():
            self.adjusts[key].setValue(1)

    def restoreClicked(self):
        ini_file = QtWidgets.QFileDialog.getOpenFileName(self, 'Open Adjustments file',
                   self.save_folder, 'Preferences Files (*.ini)')[0]
        if ini_file != '':
            reshow = False
            config = configparser.RawConfigParser()
            config.read(ini_file)
            try:
                adjustin = config.get('Powermatch', 'adjustments')
            except:
                return
            self.resetClicked()
            bits = adjustin.split(',')
            for bit in bits:
                bi = bit.split('=')
                try:
                    self.adjusts[bi[0]].setValue(float(bi[1]))
                except:
                    pass

    def saveClicked(self):
        line = ''
        for key, value in self.adjustm.items():
            if value != 1:
                line += key + '=' + str(value) + ','
        if line != '':
            line = 'adjustments=' + line[:-1]
            updates = {'Powermatch': [line]}
            inifile = QtWidgets.QFileDialog.getSaveFileName(None, 'Save Adjustments to file',
                      self.save_folder, 'Preferences Files (*.ini)')[0]
            if inifile != '':
                if inifile[-4:] != '.ini':
                    inifile = inifile + '.ini'
                SaveIni(updates, ini_file=inifile)

    def showClicked(self):
        self.results = {}
        for key in list(self.adjusts.keys()):
            self.results[key] = self.adjustm[key]
        self.close()

    def getValues(self):
        return self.results

class powerMatch(QtWidgets.QWidget):
    log = QtCore.pyqtSignal()
    progress = QtCore.pyqtSignal()

    def get_filename(self, filename):
        if filename.find('/') == 0: # full directory in non-Windows
            return filename
        elif (sys.platform == 'win32' or sys.platform == 'cygwin') \
          and filename[1] == ':': # full directory for Windows
            return filename
        else: # subdirectory of scenarios
            return self.scenarios + filename

    def __init__(self, help='help.html'):
        super(powerMatch, self).__init__()
        self.help = help
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = getModelFile('SIREN.ini')
        config.read(config_file)
        parents = []
        try:
            parents = getParents(config.items('Parents'))
        except:
            pass
        try:
            base_year = config.get('Base', 'year')
        except:
            base_year = '2012'
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
            self.scenarios = self.scenarios.replace('$YEAR$', base_year)
            self.scenarios = self.scenarios[: self.scenarios.rfind('/') + 1]
            if self.scenarios[:3] == '../':
                ups = self.scenarios.split('../')
                me = os.getcwd().split(os.sep)
                me = me[: -(len(ups) - 1)]
                me.append(ups[-1])
                self.scenarios = '/'.join(me)
        except:
            self.scenarios = ''
        self.log_status = True
        try:
            rw = config.get('Windows', 'log_status')
            if rw.lower() in ['false', 'no', 'off']:
                self.log_status = False
        except:
            pass
        self.file_labels = ['Constraints', 'Generators', 'Optimisation', 'Data', 'Results']
        self.ifiles = [''] * 5
        self.isheets = self.file_labels[:]
        del self.isheets[-2:]
        self.more_details = False
        self.constraints = None
        self.generators = None
        self.optimisation = None
        self.adjustby = None
        self.adjust_cap = 25
        self.adjust_re = False
        self.change_res = True
        self.corrected_lcoe = True
        self.carbon_price = 0.
        self.optimise_choice = 'LCOE'
        self.optimise_generations = 20
        self.optimise_mutation = 0.005
        self.optimise_population = 50
        self.optimise_stop = 0
        self.optimise_debug = False
        self.remove_cost = True
        self.surplus_sign = 1 # Note: Preferences file has it called shortfall_sign
        # it's easier for the user to understand while for the program logic surplus is easier
        target_keys = ['lcoe', 'load_pct', 'surplus_pct', 're_pct', 'cost', 'co2']
        target_names = ['LCOE', 'Load%', 'Surplus%', 'RE%', 'Cost', 'CO2']
        target_fmats = ['$%.2f', '%d%%', '%d%%', '%d%%', '$%.1fpwr_chr', '%.1fpwr_chr']
        target_titles = ['LCOE ($)', '% Load met', 'Surplus %', 'RE %',
                         'Total Cost ($)', 'tCO2e']
        self.targets = {}
        for t in range(len(target_keys)):
             if target_keys[t] in ['re_pct', 'surplus_pct']:
                 self.targets[target_keys[t]] = [target_names[t], 0., -1, 0., 0, target_fmats[t],
                                                 target_titles[t]]
             else:
                 self.targets[target_keys[t]] = [target_names[t], 0., 0., -1, 0, target_fmats[t],
                                                 target_titles[t]]
        try:
             items = config.items('Powermatch')
             for key, value in items:
                 if key[-5:] == '_file':
                     ndx = self.file_labels.index(key[:-5].title())
                     self.ifiles[ndx] = value
                 elif key[-6:] == '_sheet':
                     ndx = self.file_labels.index(key[:-6].title())
                     self.isheets[ndx] = value
                 elif key == 'adjust_generators':
                     if value.lower() in ['true', 'on', 'yes']:
                         self.adjust_re = True
                 elif key == 'adjustments':
                     self.adjustby = {}
                     bits = value.split(',')
                     for bit in bits:
                         bi = bit.split('=')
                         self.adjustby[bi[0]] = float(bi[1])
                 elif key == 'carbon_price':
                     try:
                         self.carbon_price = float(value)
                     except:
                         pass
                 elif key == 'change_results':
                     if value.lower() in ['false', 'off', 'no']:
                         self.change_res = False
                 elif key == 'corrected_lcoe':
                     if value.lower() in ['false', 'no', 'off']:
                         self.corrected_lcoe = False
                 elif key == 'log_status':
                     if value.lower() in ['false', 'no', 'off']:
                         self.log_status = False
                 elif key == 'more_details':
                     if value.lower() in ['true', 'yes', 'on']:
                         self.more_details = True
                 elif key == 'optimise_debug':
                     if value.lower() in ['true', 'on', 'yes']:
                         self.optimise_debug = True
                 elif key == 'optimise_choice':
                     self.optimise_choice = value
                 elif key == 'optimise_generations':
                     try:
                         self.optimise_generations = int(value)
                     except:
                         pass
                 elif key == 'optimise_mutation':
                     try:
                         self.optimise_mutation = float(value)
                     except:
                         pass
                 elif key == 'optimise_population':
                     try:
                         self.optimise_population = int(value)
                     except:
                         pass
                 elif key == 'optimise_stop':
                     try:
                         self.optimise_stop = int(value)
                     except:
                         pass
                 elif key[:9] == 'optimise_':
                     try:
                         bits = value.split(',')
                         t = target_keys.index(key[9:])
                         # name, weight, minimum, maximum, widget index
                         self.targets[key[9:]] = [target_names[t], float(bits[0]), float(bits[1]),
                                                  float(bits[2]), 0, target_fmats[t],
                                                  target_titles[t]]
                     except:
                         pass
                 elif key == 'remove_cost':
                     if value.lower() in ['false', 'off', 'no']:
                         self.remove_cost = False
                 elif key == 'shortfall_sign':
                     if value[0] == '+' or value[0].lower() == 'p':
                         self.surplus_sign = -1
        except:
            pass
        try:
            self.adjust_cap =  float(config.get('Power', 'adjust_cap'))
        except:
            pass
        self.opt_progressbar = None
        self.floatstatus = None # status window
   #     self.tabs = QtGui.QTabWidget()    # Create tabs
   #     tab1 = QtGui.QWidget()
   #     tab2 = QtGui.QWidget()
   #     tab3 = QtGui.QWidget()
   #     tab4 = QtGui.QWidget()
   #     tab5 = QtGui.QWidget()
        self.grid = QtWidgets.QGridLayout()
        self.files = [None] * 5
        self.sheets = self.file_labels[:]
        del self.sheets[-2:]
        self.updated = False
        edit = [None, None, None]
        r = 0
        for i in range(5):
            self.grid.addWidget(QtWidgets.QLabel(self.file_labels[i] + ' File:'), r, 0)
            self.files[i] = ClickableQLabel()
            self.files[i].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
            self.files[i].setText(self.ifiles[i])
            self.files[i].clicked.connect(self.fileChanged)
            self.grid.addWidget(self.files[i], r, 1, 1, 3)
            if i < D:
                r += 1
                self.grid.addWidget(QtWidgets.QLabel(self.file_labels[i] + ' Sheet:'), r, 0)
                self.sheets[i] = QtWidgets.QComboBox()
                self.sheets[i].addItem(self.isheets[i])
                self.grid.addWidget(self.sheets[i], r, 1, 1, 2)
                edit[i] = QtWidgets.QPushButton(self.file_labels[i], self)
                self.grid.addWidget(edit[i], r, 3)
                edit[i].clicked.connect(self.editClicked)
            r += 1
        wdth = edit[1].fontMetrics().boundingRect(edit[1].text()).width() + 9
        self.grid.addWidget(QtWidgets.QLabel('Carbon Price:'), r, 0)
        self.carbon = QtWidgets.QDoubleSpinBox()
        self.carbon.setRange(0, 200)
        self.carbon.setDecimals(2)
        try:
            self.carbon.setValue(self.carbon_price)
        except:
            self.carbon.setValue(0.)
        self.grid.addWidget(self.carbon, r, 1)
        self.carbon.valueChanged.connect(self.cpchanged)
        self.grid.addWidget(QtWidgets.QLabel('($/tCO2e. Use only if LCOE excludes carbon price)'), r, 2, 1, 2)
        r += 1
        self.grid.addWidget(QtWidgets.QLabel('Adjust Generators:'), r, 0)
        self.adjust = QtWidgets.QCheckBox('(check to adjust/multiply generators capacity data)', self)
        if self.adjust_re:
            self.adjust.setCheckState(QtCore.Qt.Checked)
        self.grid.addWidget(self.adjust, r, 1, 1, 3)
        r += 1
        self.grid.addWidget(QtWidgets.QLabel('Dispatch Order:\n(move to right\nto exclude)'), r, 0)
        self.order = ListWidget(self) #QtWidgets.QListWidget()
      #  self.order.setDragDropMode(QtGui.QAbstractItemView.InternalMove)
        self.grid.addWidget(self.order, r, 1, 1, 2)
        self.ignore = ListWidget(self) # QtWidgets.QListWidget()
      #  self.ignore.setDragDropMode(QtGui.QAbstractItemView.InternalMove)
        self.grid.addWidget(self.ignore, r, 3, 1, 2)
        r += 1
        self.log = QtWidgets.QLabel('')
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.log.setPalette(msg_palette)
        self.grid.addWidget(self.log, r, 1, 1, 4)
        r += 1
        self.progressbar = QtWidgets.QProgressBar()
        self.progressbar.setMinimum(0)
        self.progressbar.setMaximum(10)
        self.progressbar.setValue(0)
        self.progressbar.setStyleSheet('QProgressBar {border: 1px solid grey; border-radius: 2px; text-align: center;}' \
                                       + 'QProgressBar::chunk { background-color: #6891c6;}')
        self.grid.addWidget(self.progressbar, r, 1, 1, 4)
        self.progressbar.setHidden(True)
        r += 1
        r += 1
        quit = QtWidgets.QPushButton('Done', self)
        self.grid.addWidget(quit, r, 0)
        quit.clicked.connect(self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        pms = QtWidgets.QPushButton('Summary', self)
        self.grid.addWidget(pms, r, 1)
        pms.clicked.connect(self.pmClicked)
        pm = QtWidgets.QPushButton('Powermatch', self)
     #   pm.setMaximumWidth(wdth)
        self.grid.addWidget(pm, r, 2)
        pm.clicked.connect(self.pmClicked)
        opt = QtWidgets.QPushButton('Optimise', self)
        self.grid.addWidget(opt, r, 3)
        opt.clicked.connect(self.pmClicked)
        help = QtWidgets.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        quit.setMaximumWidth(wdth)
        self.grid.addWidget(help, r, 5)
        help.clicked.connect(self.helpClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        try:
            ts = xlrd.open_workbook(self.get_filename(self.files[G].text()))
            ws = ts.sheet_by_name('Generators')
            self.getGenerators(ws)
            self.setOrder()
            ts.release_resources()
            del ts
        except:
            pass
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
    #    tab1.setLayout(self.layout)
    #    self.tabs.addTab(tab1,'Parms')
    #    self.tabs.addTab(tab2,'Constraints')
    #    self.tabs.addTab(tab3,'Generators')
    #    self.tabs.addTab(tab4,'Summary')
    #    self.tabs.addTab(tab5,'Details')
        self.setWindowTitle('SIREN - powermatch (' + fileVersion() + ') - Powermatch')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.center()
        self.resize(int(self.sizeHint().width()* 1.2), int(self.sizeHint().height() * 1.2))
        self.show_FloatStatus() # status window
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        centerPoint = QtWidgets.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def fileChanged(self):
        self.setStatus('')
        for i in range(5):
            if self.files[i].hasFocus():
                break
        curfile = self.scenarios + self.files[i].text()
        if i == R:
            if self.files[i].text() == '':
                curfile = self.scenarios + self.files[D].text()
                curfile = curfile.replace('data', 'results')
                curfile = curfile.replace('Data', 'Results')
            newfile = QtWidgets.QFileDialog.getSaveFileName(None, 'Save ' + self.file_labels[i] + ' file',
                      curfile, 'Excel Files (*.xlsx)')[0]
        else:
            newfile = QtWidgets.QFileDialog.getOpenFileName(self, 'Open ' + self.file_labels[i] + ' file',
                      curfile)[0]
        if newfile != '':
            if i < D:
                if i == C:
                    self.constraints = None
                elif i == G:
                    self.generators = None
                else:
                    self.optimisation = None
                ts = xlrd.open_workbook(newfile, on_demand = True)
                ndx = 0
                self.sheets[i].clear()
                j = -1
                for sht in ts.sheet_names():
                    j += 1
                    self.sheets[i].addItem(sht)
                    if sht == self.file_labels[i]:
                        ndx = j
                self.sheets[i].setCurrentIndex(ndx)
                if i == 1:
                    ws = ts.sheet_by_index(ndx)
                    self.getGenerators(ws)
                    self.setOrder()
                ts.release_resources()
                del ts
            if newfile[: len(self.scenarios)] == self.scenarios:
                self.files[i].setText(newfile[len(self.scenarios):])
            else:
                self.files[i].setText(newfile)
            if i == D and self.change_res:
                newfile = self.files[D].text()
                newfile = newfile.replace('data', 'results')
                newfile = newfile.replace('Data', 'Results')
                if newfile != self.files[D].text():
                    self.files[R].setText(newfile)
            self.updated = True

    def helpClicked(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.help,
                 title='Help for powermatch (' + fileVersion() + ')', section='powermatch')
        dialog.exec_()

    def cpchanged(self):
        self.updated = True
        self.carbon_price = self.carbon.value()

    def changes(self):
        self.updated = True

    def quitClicked(self):
        if self.updated:
            updates = {}
            lines = []
            lines.append('adjust_generators=' + str(self.adjust.isChecked()))
            if self.adjustby is not None:
                line = ''
                for key, value in self.adjustby.items():
                    if value != 1:
                        line += key + '=' + str(value) + ','
                if line != '':
                    lines.append('adjustments=' + line[:-1])
            lines.append('carbon_price=' + str(self.carbon_price))
            for i in range(len(self.file_labels)):
                lines.append(self.file_labels[i].lower() + '_file=' + self.files[i].text())
            for i in range(D):
                lines.append(self.file_labels[i].lower() + '_sheet=' + self.sheets[i].currentText())
            lines.append('optimise_choice=' + self.optimise_choice)
            lines.append('optimise_generations=' + str(self.optimise_generations))
            lines.append('optimise_mutation=' + str(self.optimise_mutation))
            lines.append('optimise_population=' + str(self.optimise_population))
            lines.append('optimise_stop=' + str(self.optimise_stop))
            for key, value in self.targets.items():
                line = 'optimise_{}={:.2f},{:.2f},{:.2f}'.format(key, value[1], value[2], value[3])
                lines.append(line)
            updates['Powermatch'] = lines
            SaveIni(updates)
        self.close()

    def closeEvent(self, event):
        if self.floatstatus is not None:
            self.floatstatus.exit()
        event.accept()

    def editClicked(self):

        def update_dictionary(it, source):
            new_keys = list(source.keys())
            # first we delete and add keys to match updated dictionary
            if it == C:
                old_keys = list(self.constraints.keys())
                for key in old_keys:
                    if key in new_keys:
                        del new_keys[new_keys.index(key)]
                    else:
                        del self.constraints[key]
                for key in new_keys:
                    self.constraints[key] = Constraint(key, '<category>', 0., 1., 1., 1., 1., 0.,
                                                       1., 0., 0., 0, 0)
                target = self.constraints
            elif it == G:
                old_keys = list(self.generators.keys())
                for key in old_keys:
                    if key in new_keys:
                        del new_keys[new_keys.index(key)]
                    else:
                        del self.generators[key]
                for key in new_keys:
                    self.generators[key] = Facility(key, 0, '<constraint>', 0., 0., 0., 0.)
                target = self.generators
            elif it == O:
                old_keys = list(self.optimisation.keys())
                for key in old_keys:
                    if key in new_keys:
                        del new_keys[new_keys.index(key)]
                    else:
                        del self.optimisation[key]
                for key in new_keys:
                    self.optimisation[key] = Optimisation(key, 'None', None)
                target = self.optimisation
            # now update the data
            for key in list(target.keys()):
                for prop in dir(target[key]):
                    if prop[:2] != '__' and prop[-2:] != '__':
                        try:
                            setattr(target[key], prop, source[key][prop])
                        except:
                            pass

        self.setStatus('')
        ts = None
        it = self.file_labels.index(self.sender().text())
        if it == C and self.constraints is not None:
            pass
        elif it == G and self.generators is not None:
            pass
        elif it == O and self.optimisation is not None:
            pass
        else:
            try:
                ts = xlrd.open_workbook(self.get_filename(self.files[it].text()))
                try:
                    sht = self.sheets[it].currentText()
                except:
                    self.setStatus(self.sheets[it].currentText() + ' not found in ' \
                                     + self.file_labels[it] + ' spreadsheet.')
                    return
                ws = ts.sheet_by_name(sht)
            except:
                ws = None
        if it == C: # self.constraints
            if self.constraints is None:
                try:
                    self.getConstraints(ws)
                except:
                    return
            sp_pts = [2] * 13
            sp_pts[4] = 3 # discharge loss
            sp_pts[6] = 3 # parasitic loss
            sp_pts[9] = 3 # recharge loss
            dialog = displaytable.Table(self.constraints, title=self.sender().text(),
                 save_folder=self.scenarios, edit=True, decpts=sp_pts, abbr=False)
            dialog.exec_()
            if dialog.getValues() is not None:
                update_dictionary(it, dialog.getValues())
        elif it == G: # generators
            if self.generators is None:
                try:
                    self.getGenerators(ws)
                except:
                    return
            sp_pts = [2] * 8
            sp_pts[7] = 0 # dispatch order column
            dialog = displaytable.Table(self.generators, title=self.sender().text(),
                 save_folder=self.scenarios, edit=True, decpts=sp_pts, abbr=False)
            dialog.exec_()
            if dialog.getValues() is not None:
                update_dictionary(it, dialog.getValues())
                self.setOrder()
        elif it == O: # self.optimisation
            if self.optimisation is None:
                try:
                    self.getoptimisation(ws)
                except:
                    return
            dialog = displaytable.Table(self.optimisation, title=self.sender().text(),
                     save_folder=self.scenarios, edit=True)
            dialog.exec_()
            if dialog.getValues() is not None:
                update_dictionary(it, dialog.getValues())
                for key in self.optimisation.keys():
                    if self.optimisation[key].approach == 'Discrete':
                        caps = self.optimisation[key].capacities.split()
                        self.optimisation[key].capacities = []
                        cap_max = 0.
                        for cap in caps:
                            try:
                                self.optimisation[key].capacities.append(float(cap))
                                cap_max += float(cap)
                            except:
                                pass
                        self.optimisation[key].capacity_min = 0
                        self.optimisation[key].capacity_max = round(cap_max, 3)
                        self.optimisation[key].capacity_step = None
        if ts is not None:
            ts.release_resources()
            del ts
        newfile = dialog.savedfile
        if newfile is not None:
            if newfile[: len(self.scenarios)] == self.scenarios:
                self.files[it].setText(newfile[len(self.scenarios):])
            else:
                self.files[it].setText(newfile)
            self.setStatus(self.file_labels[it] + ' spreadsheet changed.')

    def getConstraints(self, ws):
        if ws is None:
            self.constraints = {}
            self.constraints['<name>'] = Constraint('<name>', '<category>', 0., 1.,
                                              1., 1., 1., 0., 1., 0., 0., 0, 0)
            return
        wait_col = -1
        warm_col = -1
        min_run_time = 0
        warm_time = 0
        if ws.cell_value(1, 0) == 'Name' and ws.cell_value(1, 1) == 'Category':
            cat_col = 1
            for col in range(ws.ncols):
                if ws.cell_value(0, col)[:8] == 'Capacity':
                    cap_col = [col, col + 1]
                elif ws.cell_value(0, col)[:9] == 'Ramp Rate':
                    ramp_col = [col, col + 1]
                elif ws.cell_value(0, col)[:8] == 'Recharge':
                    rec_col = [col, col + 1]
                elif ws.cell_value(0, col)[:9] == 'Discharge':
                    dis_col = [col, col + 1]
                elif ws.cell_value(1, col)[:9] == 'Parasitic':
                    par_col = col
                elif ws.cell_value(1, col)[:9] == 'Wait Time':
                    wait_col = col
                elif ws.cell_value(1, col)[:11] == 'Warmup Time':
                    warm_col = col
            strt_row = 2
        elif ws.cell_value(0, 0) == 'Name': # saved file
            cap_col = [-1, -1]
            ramp_col = [-1, -1]
            rec_col = [-1, -1]
            dis_col = [-1, -1]
            for col in range(ws.ncols):
                if ws.cell_value(0, col)[:8] == 'Category':
                    cat_col = col
                elif ws.cell_value(0, col)[:8] == 'Capacity':
                    if ws.cell_value(0, col)[-3:] == 'Min':
                        cap_col[0] = col
                    else:
                        cap_col[1] = col
                elif ws.cell_value(0, col)[:6] == 'Rampup':
                    ramp_col[0] = col
                elif ws.cell_value(0, col)[:8] == 'Rampdown':
                    ramp_col[1] = col
                elif ws.cell_value(0, col)[:8] == 'Recharge':
                    if ws.cell_value(0, col)[-3:] == 'Max':
                        rec_col[0] = col
                    else:
                        rec_col[1] = col
                elif ws.cell_value(0, col)[:9] == 'Discharge':
                    if ws.cell_value(0, col)[-3:] == 'Max':
                        dis_col[0] = col
                    else:
                        dis_col[1] = col
                elif ws.cell_value(0, col)[:9] == 'Parasitic':
                    par_col = col
                elif ws.cell_value(0, col)[:9] == 'Wait Time':
                    wait_col = col
                elif ws.cell_value(0, col)[:11] == 'Warmup Time':
                    warm_col = col
            strt_row = 1
        else:
            self.setStatus('Not a ' + self.file_labels[it] + ' worksheet.')
            return
        try:
            cat_col = cat_col
        except:
            self.setStatus('Not a ' + self.file_labels[it] + ' worksheet.')
            return
        self.constraints = {}
        for row in range(strt_row, ws.nrows):
            if wait_col >= 0:
                min_run_time = ws.cell_value(row, wait_col)
            if warm_col >= 0:
                warm_time = ws.cell_value(row, warm_col)
            self.constraints[str(ws.cell_value(row, 0))] = Constraint(str(ws.cell_value(row, 0)),
                                     str(ws.cell_value(row, cat_col)),
                                     ws.cell_value(row, cap_col[0]), ws.cell_value(row, cap_col[1]),
                                     ws.cell_value(row, ramp_col[0]), ws.cell_value(row, ramp_col[1]),
                                     ws.cell_value(row, rec_col[0]), ws.cell_value(row, rec_col[1]),
                                     ws.cell_value(row, dis_col[0]), ws.cell_value(row, dis_col[1]),
                                     ws.cell_value(row, par_col), min_run_time, warm_time)
        return

    def getGenerators(self, ws):
        if ws is None:
            self.generators = {}
            self.generators['<name>'] = Facility('<name>', 0, '<constraint>', 0., 0., 0., 0.)
            return
        if ws.cell_value(0, 0) != 'Name':
            self.setStatus('Not a ' + self.file_labels[G] + ' worksheet.')
            return
        for col in range(ws.ncols):
            if ws.cell_value(0, col)[:8] == 'Dispatch' or ws.cell_value(0, col) == 'Order':
                ord_col = col
            elif ws.cell_value(0, col) == 'Constraint':
                con_col = col
            elif ws.cell_value(0, col)[:8] == 'Capacity':
                cap_col = col
            elif ws.cell_value(0, col)[:7] == 'Initial':
                ini_col = col
            elif ws.cell_value(0, col) == 'LCOE CF':
                lcc_col = col
            elif ws.cell_value(0, col)[:4] == 'LCOE':
                lco_col = col
            elif ws.cell_value(0, col)[:9] == 'Emissions':
                emi_col = col
        try:
            lco_col = lco_col
        except:
            self.setStatus('Not a ' + self.file_labels[G] + ' worksheet.')
            return
        self.generators = {}
        for row in range(1, ws.nrows):
            self.generators[str(ws.cell_value(row, 0))] = Facility(str(ws.cell_value(row, 0)),
                                     ws.cell_value(row, ord_col), str(ws.cell_value(row, con_col)),
                                     ws.cell_value(row, cap_col), ws.cell_value(row, lco_col),
                                     ws.cell_value(row, lcc_col), ws.cell_value(row, emi_col),
                                     initial=ws.cell_value(row, ini_col))
        return

    def getoptimisation(self, ws):
        if ws is None:
            self.optimisation = {}
            self.optimisation['<name>'] = Optimisation('<name>', 'None', None)
            return
        if ws.cell_value(0, 0) != 'Name':
            self.setStatus('Not an ' + self.file_labels[O] + ' worksheet.')
            return
        cols = ['Name', 'Approach', 'Values', 'Capacity Max', 'Capacity Min',
                'Capacity Step', 'Capacities']
        coln = [-1] * len(cols)
        for col in range(ws.ncols):
            try:
                i = cols.index(ws.cell_value(0, col))
                coln[i] = col
            except:
                pass
        if coln[0] < 0:
            self.setStatus('Not an ' + self.file_labels[O] + ' worksheet.')
            return
        self.optimisation = {}
        for row in range(1, ws.nrows):
            tech = ws.cell_value(row, 0)
            if coln[2] > 0: # values format
                self.optimisation[tech] = Optimisation(tech,
                                     ws.cell_value(row, coln[1]),
                                     ws.cell_value(row, coln[2]))
            else:
                if ws.cell_value(row, coln[1]) == 'Discrete': # fudge values format
                    self.optimisation[tech] = Optimisation(tech,
                                         ws.cell_value(row, coln[1]),
                                         ws.cell_value(row, coln[-1]))
                else:
                    self.optimisation[tech] = Optimisation(tech, '', '')
                    for col in range(1, len(coln)):
                        if coln[col] > 0:
                            attr = cols[col].lower().replace(' ', '_')
                            setattr(self.optimisation[tech], attr,
                                    ws.cell_value(row, coln[col]))
            try:
                self.optimisation[tech].capacity = self.generators[tech].capacity
            except:
                pass
        return

    def setOrder(self):
        self.order.clear()
        self.ignore.clear()
        if self.generators is None:
            order = ['Storage', 'Biomass', 'PHS', 'Gas', 'CCG1', 'Other', 'Coal']
            for stn in order:
                self.order.addItem(stn)
        else:
            order = []
            zero = []
            for key, value in self.generators.items():
                if value.capacity == 0:
                    continue
                try:
                    o = int(value.order)
                    if o > 0:
                        while len(order) <= o:
                            order.append([])
                        order[o - 1].append(key)
                    elif o == 0:
                        zero.append(key)
                except:
                    pass
            order.append(zero)
            for cat in order:
                for stn in cat:
                    self.order.addItem(stn)

    def pmClicked(self):
        self.setStatus(self.sender().text() + ' processing started')
        if self.sender().text() == 'Powermatch': # detailed spreadsheet?
            option = 'P'
        else:
            if self.sender().text() == 'Optimise': # do optimisation?
                option = 'O'
                self.optExit = False #??
            else:
                option = 'S'
        if option != 'O':
            self.progressbar.setMinimum(0)
            self.progressbar.setMaximum(10)
            self.progressbar.setHidden(False)
        err_msg = ''
        if self.constraints is None:
            try:
                ts = xlrd.open_workbook(self.get_filename(self.files[C].text()))
                ws = ts.sheet_by_name(self.sheets[C].currentText())
                self.getConstraints(ws)
                ts.release_resources()
                del ts
            except FileNotFoundError:
                err_msg = 'Constraints file not found - ' + self.files[C].text()
                self.getConstraints(None)
            except:
                err_msg = 'Error accessing Constraints'
                self.getConstraints(None)
        if self.generators is None:
            try:
                ts = xlrd.open_workbook(self.get_filename(self.files[G].text()))
                ws = ts.sheet_by_name(self.sheets[G].currentText())
                self.getGenerators(ws)
                ts.release_resources()
                del ts
            except FileNotFoundError:
                if err_msg != '':
                    err_msg += ' nor Generators - ' + self.files[G].text()
                else:
                    err_msg = 'Generators file not found - ' + self.files[G].text()
                self.getGenerators(None)
            except:
                if err_msg != '':
                    err_msg += ' and Generators'
                else:
                    err_msg = 'Error accessing Generators'
                self.getGenerators(None)
        if option == 'O' and self.optimisation is None:
            try:
                ts = xlrd.open_workbook(self.get_filename(self.files[O].text()))
                ws = ts.sheet_by_name(self.sheets[O].currentText())
                self.getoptimisation(ws)
                ts.release_resources()
                del ts
                if self.optimisation is None:
                    if err_msg != '':
                        err_msg += ' not an Optimisation worksheet'
                    else:
                        err_msg = 'Not an optimisation worksheet'
            except FileNotFoundError:
                if err_msg != '':
                    err_msg += ' nor Optimisation - ' + self.files[O].text()
                else:
                    err_msg = 'Optimisation file not found - ' + self.files[O].text()
            except:
                if err_msg != '':
                    err_msg += ' and Optimisation'
                else:
                    err_msg = 'Error accessing Optimisation'
            if self.optimisation is None:
                self.getoptimisation(None)
        if err_msg != '':
            self.setStatus(err_msg)
            return
        pm_data_file = self.get_filename(self.files[D].text())
        if pm_data_file[-5:] != '.xlsx': #xlsx format only
            self.setStatus('Not a Powermatch data spreadsheet')
            self.progressbar.setHidden(True)
            return
        try:
            ts = oxl.load_workbook(pm_data_file)
        except FileNotFoundError:
            self.setStatus('Data file not found - ' + self.files[D].text())
            return
        except:
            self.setStatus('Error accessing Data file - ' + self.files[D].text())
            return
        ws = ts.active
        top_row = ws.max_row - 8760
        if ws.cell(row=top_row, column=1).value != 'Hour' or ws.cell(row=top_row, column=2).value != 'Period':
            self.setStatus('Not a Powermatch data spreadsheet')
            self.progressbar.setHidden(True)
            return
        typ_row = top_row - 1
        gen_row = typ_row
        while typ_row > 0:
            if ws.cell(row=typ_row, column=1).value[:9] == 'Generated':
                gen_row = typ_row
            if ws.cell(row=typ_row, column=3).value in tech_names:
                break
            typ_row -= 1
        else:
            self.setStatus('no suitable data')
            return
        icap_row = typ_row + 1
        while icap_row < top_row:
            if ws.cell(row=icap_row, column=1).value[:8] == 'Capacity':
                break
            icap_row += 1
        else:
            self.setStatus('no capacity data')
            return
        year = ws.cell(row=top_row + 1, column=2).value[:4]
        pmss_details = {}
        pmss_data = []
        re_order = [] # order for re technology
        dispatch_order = [] # order for dispatchable technology
        load_col = -1
        for col in range(3, ws.max_column + 1):
            try:
                valu = ws.cell(row=typ_row, column=col).value.replace('-','')
                i = tech_names.index(valu)
            except:
                continue
            if tech_names[i] == 'Load':
                load_col = len(pmss_data)
                typ = 'L'
                capacity = 0
            else:
                try:
                    capacity = float(ws.cell(row=icap_row, column=col).value)
                except:
                    continue
                if capacity <= 0:
                    continue
                typ = 'R'
            pmss_details[tech_names[i]] = [capacity, typ, len(pmss_data), 1]
            pmss_data.append([])
            re_order.append(tech_names[i])
            for row in range(top_row + 1, ws.max_row + 1):
                pmss_data[-1].append(ws.cell(row=row, column=col).value)
        pmss_details['Load'][0] = sum(pmss_data[load_col])
        do_adjust = False
        if option == 'O':
            for itm in range(self.order.count()):
                gen = self.order.item(itm).text()
                try:
                    if self.generators[gen].constraint in self.constraints and \
                      self.constraints[self.generators[gen].constraint].category == 'Generator':
                        typ = 'G'
                    else:
                        typ = 'S'
                except:
                    continue
                dispatch_order.append(gen)
                pmss_details[gen] = [self.generators[gen].capacity, typ, -1, 1]
            self.optClicked(year, option, pmss_details, pmss_data, re_order, dispatch_order,
                            None, None)
            return
        if self.adjust.isChecked():
            generated = 0
            for row in range(top_row + 1, ws.max_row + 1):
                generated = generated + ws.cell(row=row, column=3).value
            adjustin = []
            for col in range(3, ws.max_column + 1):
                try:
                    valu = ws.cell(row=typ_row, column=col).value.replace('-','')
                    i = tech_names.index(valu)
                except:
                    break
                if valu == 'Load':
                    adjustin.append([tech_names[i], '', generated])
                else:
                    try:
                        typ = self.constraints[tech_names[i]].category
                        adjustin.append([tech_names[i], typ,
                                        float(ws.cell(row=icap_row, column=col).value)])
                    except:
                        try:
                            adjustin.append([tech_names[i], '',
                                            float(ws.cell(row=icap_row, column=col).value)])
                        except:
                            pass
            for i in range(self.order.count()):
                if self.generators[self.order.item(i).text()].capacity > 0:
                    who = self.order.item(i).text()
                    try:
                        typ = self.constraints[who].category
                    except:
                        typ = ''
                    adjustin.append([who, typ, self.generators[who].capacity])
            adjust = Adjustments(self, adjustin, self.adjustby, self.adjust_cap,
                                 save_folder=self.scenarios)
            adjust.exec_()
            if adjust.getValues() is None:
                self.setStatus('Execution aborted.')
                self.progressbar.setHidden(True)
                return
            self.adjustby = adjust.getValues()
            self.updated = True
            do_adjust = True
        ts.close()
        self.progressbar.setValue(1)
        if self.files[R].text() == '':
            i = pm_data_file.find('/')
            if i >= 0:
                data_file = pm_data_file[i + 1:]
            else:
                data_file = pm_data_file
            data_file = data_file.replace('data', 'results')
        else:
            data_file = self.get_filename(self.files[R].text())
        self.progressbar.setValue(2)
        for itm in range(self.order.count()):
            gen = self.order.item(itm).text()
            if do_adjust:
                try:
                    if self.adjustby[gen] <= 0:
                        continue
                except:
                    pass
            try:
                if self.generators[gen].constraint in self.constraints and \
                  self.constraints[self.generators[gen].constraint].category == 'Generator':
                    typ = 'G'
                else:
                    typ = 'S'
            except:
                continue
            dispatch_order.append(gen)
            pmss_details[gen] = [self.generators[gen].capacity, typ, -1, 1]
        if do_adjust:
            for gen, value in self.adjustby.items():
                if value != 1:
                     try:
                         pmss_details[gen][3] = value
                     except:
                         pass
        self.doDispatch(year, option, pmss_details, pmss_data, re_order, dispatch_order,
                        pm_data_file, data_file)

    def doDispatch(self, year, option, pmss_details, pmss_data, re_order, dispatch_order,
                   pm_data_file, data_file):
        def format_period(per):
            hr = per % 24
            day = int((per - hr) / 24)
            mth = 0
            while day > the_days[mth] - 1:
                day -= the_days[mth]
                mth += 1
            return '{}-{:02d}-{:02d} {:02d}:00'.format(year, mth+1, day+1, hr)

    # The "guts" of Powermatch processing. Have a single calculation algorithm
    # for Summary, Powermatch (detail), and Optimise. The detail makes it messy
        the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        headers = ['Facility', 'Capacity\n(Gen, MW;\nStor, MWh)', 'Subtotal\n(MWh)', 'CF', 'Cost\n($/yr)',
                   'LCOE\n($/MWh)', 'Emissions\n(tCO2e)','Reference\nLCOE','Reference\nCF']
        if self.surplus_sign < 0:
            sf_test = ['>', '<']
            sf_sign = ['+', '-']
        else:
            sf_test = ['<', '>']
            sf_sign = ['-', '+']
        sp_cols = []
        sp_cap = []
        shortfall = [0.] * 8760
        start_time = time.time()
        for gen in re_order:
            col = pmss_details[gen][2]
            multiplier = pmss_details[gen][3]
            if multiplier == 1:
                if gen == 'Load':
                    for h in range(len(pmss_data[col])):
                        shortfall[h] += pmss_data[col][h]
                else:
                    for h in range(len(pmss_data[col])):
                        shortfall[h] -= pmss_data[col][h]
            else:
                if gen == 'Load':
                    for h in range(len(pmss_data[col])):
                        shortfall[h] += pmss_data[col][h] * multiplier
                else:
                    for h in range(len(pmss_data[col])):
                        shortfall[h] -= pmss_data[col][h] * multiplier
        if option == 'P':
            ds = oxl.Workbook()
            ns = ds.active
            ns.title = 'Detail'
            normal = oxl.styles.Font(name='Arial')
            bold = oxl.styles.Font(name='Arial', bold=True)
            ss = ds.create_sheet('Summary', 0)
            re_sum = '=('
            sto_sum = ''
            not_sum = ''
            cap_row = 1
            ns.cell(row=cap_row, column=2).value = 'Capacity (MW/MWh)' #headers[1].replace('\n', ' ')
            ss.row_dimensions[3].height = 40
            ss.cell(row=3, column=1).value = headers[0]
            ss.cell(row=3, column=2).value = headers[1]
            ini_row = 2
            ns.cell(row=ini_row, column=2).value = 'Initial Capacity'
            sum_row = 3
            ns.cell(row=sum_row, column=2).value = headers[2].replace('\n', ' ')
            ss.cell(row=3, column=3).value = headers[2]
            cf_row = 4
            ns.cell(row=cf_row, column=2).value = headers[3].replace('\n', ' ')
            ss.cell(row=3, column=4).value = headers[3]
            cost_row = 5
            ns.cell(row=cost_row, column=2).value = headers[4].replace('\n', ' ')
            ss.cell(row=3, column=5).value = headers[4]
            lcoe_row = 6
            ns.cell(row=lcoe_row, column=2).value = headers[5].replace('\n', ' ')
            ss.cell(row=3, column=6).value = headers[5]
            emi_row = 7
            ns.cell(row=emi_row, column=2).value = headers[6].replace('\n', ' ')
            ss.cell(row=3, column=7).value = headers[6]
            ss.cell(row=3, column=8).value = headers[7]
            ss.cell(row=3, column=9).value = headers[8]
            ss_row = 3
            fall_row = 8
            ns.cell(row=fall_row, column=2).value = 'Shortfall periods'
            max_row = 9
            ns.cell(row=max_row, column=2).value = 'Maximum (MW/MWh)'
            hrs_row = 10
            ns.cell(row=hrs_row, column=2).value = 'Hours of usage'
            what_row = 11
            hrows = 12
            ns.cell(row=what_row, column=1).value = 'Hour'
            ns.cell(row=what_row, column=2).value = 'Period'
            ns.cell(row=what_row, column=3).value = 'Load'
            ns.cell(row=sum_row, column=3).value = '=SUM(' + ss_col(3) + str(hrows) + \
                                                   ':' + ss_col(3) + str(hrows + 8759) + ')'
            ns.cell(row=sum_row, column=3).number_format = '#,##0'
            ns.cell(row=max_row, column=3).value = '=MAX(' + ss_col(3) + str(hrows) + \
                                                   ':' + ss_col(3) + str(hrows + 8759) + ')'
            ns.cell(row=max_row, column=3).number_format = '#,##0.00'
            o = 4
            col = 3
            # hour, period
            for row in range(hrows, 8760 + hrows):
                ns.cell(row=row, column=1).value = row - hrows + 1
                ns.cell(row=row, column=2).value = format_period(row - hrows)
            # and load
            load_col = pmss_details['Load'][2]
            if pmss_details['Load'][3] == 1:
                for row in range(hrows, 8760 + hrows):
                    ns.cell(row=row, column=3).value = pmss_data[load_col][row - hrows]
                    ns.cell(row=row, column=col).number_format = '#,##0.00'
            else:
                for row in range(hrows, 8760 + hrows):
                    ns.cell(row=row, column=3).value = pmss_data[load_col][row - hrows] * \
                            pmss_details['Load'][3]
                    ns.cell(row=row, column=col).number_format = '#,##0.00'
            # here we're processing renewables (so no storage)
            for gen in re_order:
                if gen == 'Load':
                    continue
                if pmss_details[gen][3] <= 0:
                        continue
                col += 1
                sp_cols.append(gen)
                sp_cap.append(pmss_details[gen][0] * pmss_details[gen][3])
                ss_row += 1
                ns.cell(row=what_row, column=col).value = gen
                ss.cell(row=ss_row, column=1).value = '=Detail!' + ss_col(col) + str(what_row)
                ns.cell(row=cap_row, column=col).value = sp_cap[-1]
                ns.cell(row=cap_row, column=col).number_format = '#,##0.00'
                ss.cell(row=ss_row, column=2).value = '=Detail!' + ss_col(col) + str(cap_row)
                ss.cell(row=ss_row, column=2).number_format = '#,##0.00'
                ns.cell(row=sum_row, column=col).value = '=SUM(' + ss_col(col) \
                        + str(hrows) + ':' + ss_col(col) + str(hrows + 8759) + ')'
                ns.cell(row=sum_row, column=col).number_format = '#,##0'
                ss.cell(row=ss_row, column=3).value = '=Detail!' + ss_col(col) + str(sum_row)
                ss.cell(row=ss_row, column=3).number_format = '#,##0'
                re_sum += 'C' + str(ss_row) + '+'
                ns.cell(row=cf_row, column=col).value = '=IF(' + ss_col(col) + '1>0,' + \
                        ss_col(col) + '3/' + ss_col(col) + '1/8760,"")'
                ns.cell(row=cf_row, column=col).number_format = '#,##0.00'
                ss.cell(row=ss_row, column=4).value = '=Detail!' + ss_col(col) + str(cf_row)
                ss.cell(row=ss_row, column=4).number_format = '#,##0.00'
                if gen not in self.generators:
                    continue
                if self.generators[gen].lcoe > 0:
                    ns.cell(row=cost_row, column=col).value = '=IF(' + ss_col(col) + str(cf_row) + \
                            '>0,' + ss_col(col) + str(sum_row) + '*Summary!H' + str(ss_row) + \
                            '*Summary!I' + str(ss_row) + '/' + ss_col(col) + str(cf_row) + ',0)'
                    ns.cell(row=cost_row, column=col).number_format = '$#,##0'
                    if self.remove_cost:
                        ss.cell(row=ss_row, column=5).value = '=IF(Detail!' + ss_col(col) + str(sum_row) \
                                + '>0,Detail!' + ss_col(col) + str(cost_row) + ',"")'
                    else:
                        ss.cell(row=ss_row, column=5).value = '=Detail!' + ss_col(col) + str(cost_row)
                    ss.cell(row=ss_row, column=5).number_format = '$#,##0'
                    ns.cell(row=lcoe_row, column=col).value = '=IF(AND(' + ss_col(col) + str(cf_row) + '>0,' \
                            + ss_col(col) + str(cap_row) + '>0),' + ss_col(col) + str(cost_row) + '/8760/' \
                            + ss_col(col) + str(cf_row) +'/' + ss_col(col) + str(cap_row) + ',"")'
                    ns.cell(row=lcoe_row, column=col).number_format = '$#,##0.00'
                    ss.cell(row=ss_row, column=6).value = '=Detail!' + ss_col(col) + str(lcoe_row)
                    ss.cell(row=ss_row, column=6).number_format = '$#,##0.00'
                if self.generators[gen].emissions > 0:
                    ns.cell(row=emi_row, column=col).value = '=' + ss_col(col) + str(sum_row) \
                            + '*' + str(self.generators[gen].emissions)
                    ns.cell(row=emi_row, column=col).number_format = '#,##0'
                    if self.remove_cost:
                        ss.cell(row=ss_row, column=7).value = '=IF(Detail!' + ss_col(col) + str(sum_row) \
                                + '>0,Detail!' + ss_col(col) + str(emi_row) + ',"")'
                    else:
                        ss.cell(row=ss_row, column=7).value = '=Detail!' + ss_col(col) + str(emi_row)
                    ss.cell(row=ss_row, column=7).number_format = '#,##0'
                ss.cell(row=ss_row, column=8).value = self.generators[gen].lcoe
                ss.cell(row=ss_row, column=8).number_format = '$#,##0.00'
                ss.cell(row=ss_row, column=9).value = self.generators[gen].lcoe_cf
                ss.cell(row=ss_row, column=9).number_format = '#,##0.00'
                ns.cell(row=max_row, column=col).value = '=MAX(' + ss_col(col) + str(hrows) + \
                                               ':' + ss_col(col) + str(hrows + 8759) + ')'
                ns.cell(row=max_row, column=col).number_format = '#,##0.00'
                ns.cell(row=hrs_row, column=col).value = '=COUNTIF(' + ss_col(col) + str(hrows) + \
                                               ':' + ss_col(col) + str(hrows + 8759) + ',">0")'
                ns.cell(row=hrs_row, column=col).number_format = '#,##0'
                di = pmss_details[gen][2]
                if pmss_details[gen][3] == 1:
                    for row in range(hrows, 8760 + hrows):
                        ns.cell(row=row, column=col).value = pmss_data[di][row - hrows]
                        ns.cell(row=row, column=col).number_format = '#,##0.00'
                else:
                    for row in range(hrows, 8760 + hrows):
                        ns.cell(row=row, column=col).value = pmss_data[di][row - hrows] * \
                                                             pmss_details[gen][3]
                        ns.cell(row=row, column=col).number_format = '#,##0.00'
                shrt_col = col + 1
            ns.cell(row=fall_row, column=shrt_col).value = '=COUNTIF(' + ss_col(shrt_col) \
                            + str(hrows) + ':' + ss_col(shrt_col) + str(hrows + 8759) + \
                            ',"' + sf_test[0] + '0")'
            ns.cell(row=fall_row, column=shrt_col).number_format = '#,##0'
            ns.cell(row=what_row, column=shrt_col).value = 'Shortfall (' + sf_sign[0] \
                    + ') /\nSurplus (' + sf_sign[1] + ')'
            ns.cell(row=max_row, column=shrt_col).value = '=MAX(' + ss_col(shrt_col) + str(hrows) + \
                                           ':' + ss_col(shrt_col) + str(hrows + 8759) + ')'
            ns.cell(row=max_row, column=shrt_col).number_format = '#,##0.00'
            for col in range(3, shrt_col + 1):
                ns.cell(row=what_row, column=col).alignment = oxl.styles.Alignment(wrap_text=True,
                        vertical='bottom', horizontal='center')
                ns.cell(row=row, column=shrt_col).value = shortfall[row - hrows] * -self.surplus_sign
                for col in range(3, shrt_col + 1):
                    ns.cell(row=row, column=col).number_format = '#,##0.00'
            for row in range(hrows, 8760 + hrows):
                ns.cell(row=row, column=shrt_col).value = shortfall[row - hrows] * -self.surplus_sign
                ns.cell(row=row, column=col).number_format = '#,##0.00'
            col = shrt_col + 1
        else:
            sp_data = []
            sp_load = 0.
            hrows = 10
            for gen in re_order:
                if gen == 'Load':
                    sp_load = sum(pmss_data[pmss_details[gen][2]]) * pmss_details[gen][3]
                    continue
                sp_data.append([gen, pmss_details[gen][0] * pmss_details[gen][3],
                               0., '', '', '', '', '', ''])
                g = sum(pmss_data[pmss_details[gen][2]]) * pmss_details[gen][3]
                sp_data[-1][2] += g
                # if ignore not used
        #storage? = []
        if option != 'O' and option != '1':
            self.progressbar.setValue(3)
        storage_names = []
        # find any minimum generation for generators
        short_taken = {}
        short_taken_tot = 0
        for gen in dispatch_order:
            if pmss_details[gen][1] == 'G': # generators
                if self.constraints[self.generators[gen].constraint].capacity_min != 0:
                    try:
                        short_taken[gen] = self.generators[gen].capacity * pmss_details[gen][3] * \
                            self.constraints[self.generators[gen].constraint].capacity_min
                    except:
                        short_taken[gen] = self.generators[gen].capacity * \
                            self.constraints[self.generators[gen].constraint].capacity_min
                    short_taken_tot += short_taken[gen]
                    for row in range(8760):
                        shortfall[row] = shortfall[row] - short_taken[gen]
        for gen in dispatch_order:
         #   min_after = [0, 0, -1, 0, 0, 0] # initial, low balance, period, final, low after, period
         #  Min_after is there to see if storage is as full at the end as at the beginning
            try:
                capacity = self.generators[gen].capacity * pmss_details[gen][3]
            except:
                capacity = self.generators[gen].capacity
            if self.generators[gen].constraint in self.constraints and \
              self.constraints[self.generators[gen].constraint].category == 'Storage': # storage
                storage_names.append(gen)
                storage = [0., 0., 0., 0.] # capacity, initial, min level, max drain
                storage[0] = capacity
                if option == 'P':
                    ns.cell(row=cap_row, column=col + 2).value = capacity
                    ns.cell(row=cap_row, column=col + 2).number_format = '#,##0.00'
                try:
                    storage[1] = self.generators[gen].initial * pmss_details[gen][3]
                except:
                    storage[1] = self.generators[gen].initial
          #      min_after[0] = storage[1]
           #     min_after[1] = storage[1]
            #    min_after[2] = 0
                if self.constraints[self.generators[gen].constraint].capacity_min > 0:
                    storage[2] = capacity * self.constraints[self.generators[gen].constraint].capacity_min
                if self.constraints[self.generators[gen].constraint].capacity_max > 0:
                    storage[3] = capacity * self.constraints[self.generators[gen].constraint].capacity_max
                else:
                    storage[3] = capacity
                recharge = [0., 0.] # cap, loss
                if self.constraints[self.generators[gen].constraint].recharge_max > 0:
                    recharge[0] = capacity * self.constraints[self.generators[gen].constraint].recharge_max
                else:
                    recharge[0] = capacity
                if self.constraints[self.generators[gen].constraint].recharge_loss > 0:
                    recharge[1] = self.constraints[self.generators[gen].constraint].recharge_loss
                discharge = [0., 0.] # cap, loss
                if self.constraints[self.generators[gen].constraint].discharge_max > 0:
                    discharge[0] = capacity * self.constraints[self.generators[gen].constraint].discharge_max
                if self.constraints[self.generators[gen].constraint].discharge_loss > 0:
                    discharge[1] = self.constraints[self.generators[gen].constraint].discharge_loss
                if self.constraints[self.generators[gen].constraint].parasitic_loss > 0:
                    parasite = self.constraints[self.generators[gen].constraint].parasitic_loss / 24.
                else:
                    parasite = 0.
                in_run = [False, False]
                min_run_time = self.constraints[self.generators[gen].constraint].min_run_time
                in_run[0] = True # start off in_run
                if min_run_time > 0 and self.generators[gen].initial == 0:
                    in_run[0] = False
                warm_time = self.constraints[self.generators[gen].constraint].warm_time
                storage_carry = storage[1] # self.generators[gen].initial
                if option == 'P':
                    ns.cell(row=ini_row, column=col + 2).value = storage_carry
                    ns.cell(row=ini_row, column=col + 2).number_format = '#,##0.00'
                storage_bal = []
                storage_can = 0.
                for row in range(8760):
                    storage_loss = 0.
                    storage_losses = 0.
                    if storage_carry > 0:
                        loss = storage_carry * parasite
                        # for later: record parasitic loss
                        storage_carry = storage_carry - loss
                        storage_losses -= loss
                    if shortfall[row] < 0:  # excess generation
                        if min_run_time > 0:
                            in_run[0] = False
                        if warm_time > 0:
                            in_run[1] = False
                        can_use = - (storage[0] - storage_carry) * (1 / (1 - recharge[1]))
                        if can_use < 0: # can use some
                            if shortfall[row] > can_use:
                                can_use = shortfall[row]
                            if can_use < - recharge[0] * (1 / (1 - recharge[1])):
                                can_use = - recharge[0]
                        else:
                            can_use = 0.
                        # for later: record recharge loss
                        storage_losses += can_use * recharge[1]
                        storage_carry -= (can_use * (1 - recharge[1]))
                        shortfall[row] -= can_use
                    else: # shortfall
                        if min_run_time > 0 and shortfall[row] > 0:
                            if not in_run[0]:
                                if row + min_run_time <= 8759:
                                    for i in range(row + 1, row + min_run_time + 1):
                                        if shortfall[i] <= 0:
                                            break
                                    else:
                                        in_run[0] = True
                        if in_run[0]:
                            can_use = shortfall[row] * (1 / (1 - discharge[1]))
                            can_use = min(can_use, discharge[0])
                            if can_use > storage_carry - storage[2]:
                                can_use = storage_carry - storage[2]
                            if warm_time > 0 and not in_run[1]:
                                in_run[1] = True
                                can_use = can_use * (1 - warm_time)
                        else:
                            can_use = 0
                        if can_use > 0:
                            storage_loss = can_use * discharge[1]
                            storage_losses -= storage_loss
                            storage_carry -= can_use
                            can_use = can_use - storage_loss
                            shortfall[row] -= can_use
                            if storage_carry < 0:
                                storage_carry = 0
                        else:
                            can_use = 0.
               #         if storage_carry < min_after[1]:
                #            min_after[1] = storage_carry
                 #           min_after[2] = row
                  #      if shortfall[row] > min_after[4]:
                    #        min_after[4] = shortfall[row]
                     #       min_after[5] = row
                      #  min_after[3] = storage_carry
                    storage_bal.append(storage_carry)
                    if option == 'P':
                        if can_use > 0:
                            ns.cell(row=row + hrows, column=col).value = 0
                            ns.cell(row=row + hrows, column=col + 2).value = can_use * self.surplus_sign
                        else:
                            ns.cell(row=row + hrows, column=col).value = can_use * -self.surplus_sign
                            ns.cell(row=row + hrows, column=col + 2).value = 0
                        ns.cell(row=row + hrows, column=col + 1).value = storage_losses
                        ns.cell(row=row + hrows, column=col + 3).value = storage_carry
                        ns.cell(row=row + hrows, column=col + 4).value = (shortfall[row] + short_taken_tot) * -self.surplus_sign
                        for ac in range(5):
                            ns.cell(row=row + hrows, column=col + ac).number_format = '#,##0.00'
                            ns.cell(row=max_row, column=col + ac).value = '=MAX(' + ss_col(col + ac) + \
                                    str(hrows) + ':' + ss_col(col + ac) + str(hrows + 8759) + ')'
                            ns.cell(row=max_row, column=col + ac).number_format = '#,##0.00'
     #                   ns.cell(row=hrs_row, column=col + 1).value = '=COUNTIF(' + ss_col(col + 1) + \
      #                          str(hrows) + ':' + ss_col(col + 1) + str(hrows + 8759) + ',">0")'
       #                 ns.cell(row=hrs_row, column=col + 1).number_format = '#,##0'
        #                ns.cell(row=hrs_row, column=col + 2).value = '=' + ss_col(col + 1) + \
         #                       str(hrs_row) + '/8760'
          #              ns.cell(row=hrs_row, column=col + 2).number_format = '#,##0.0%'
                    else:
                        if can_use > 0:
                            storage_can += can_use
                if option == 'P':
                    ns.cell(row=sum_row, column=col).value = '=SUMIF(' + ss_col(col) + \
                            str(hrows) + ':' + ss_col(col) + str(hrows + 8759) + ',">0")'
                    ns.cell(row=sum_row, column=col).number_format = '#,##0'
                    ns.cell(row=sum_row, column=col + 1).value = '=SUMIF(' + ss_col(col + 1) + \
                            str(hrows) + ':' + ss_col(col + 1) + str(hrows + 8759) + ',"<0")'
                    ns.cell(row=sum_row, column=col + 1).number_format = '#,##0'
                    ns.cell(row=sum_row, column=col + 2).value = '=SUMIF(' + ss_col(col + 2) + \
                            str(hrows) + ':' + ss_col(col + 2) + str(hrows + 8759) + ',">0")'
                    ns.cell(row=sum_row, column=col + 2).number_format = '#,##0'
                    ns.cell(row=cf_row, column=col + 2).value = '=IF(' + ss_col(col + 2) + '1>0,' + \
                                                    ss_col(col + 2) + '3/' + ss_col(col + 2) + '1/8760,"")'
                    ns.cell(row=cf_row, column=col + 2).number_format = '#,##0.00'
                    ns.cell(row=max_row, column=col).value = '=MAX(' + ss_col(col) + \
                            str(hrows) + ':' + ss_col(col) + str(hrows + 8759) + ')'
                    ns.cell(row=max_row, column=col).number_format = '#,##0.00'
                    ns.cell(row=hrs_row, column=col + 2).value = '=COUNTIF(' + ss_col(col + 2) + \
                            str(hrows) + ':' + ss_col(col + 2) + str(hrows + 8759) + ',">0")'
                    ns.cell(row=hrs_row, column=col + 2).number_format = '#,##0'
                    ns.cell(row=hrs_row, column=col + 3).value = '=' + ss_col(col + 2) + \
                            str(hrs_row) + '/8760'
                    ns.cell(row=hrs_row, column=col + 3).number_format = '#,##0.0%'
                    col += 5
                else:
                    sp_data.append([gen, storage[0], storage_can, '', '', '', '', '', ''])
            else: # generator
                if self.constraints[self.generators[gen].constraint].capacity_max > 0:
                    cap_capacity = capacity * self.constraints[self.generators[gen].constraint].capacity_max
                else:
                    cap_capacity = capacity
                if gen in short_taken.keys():
                    for row in range(8760):
                        shortfall[row] = shortfall[row] + short_taken[gen]
                    short_taken_tot -= short_taken[gen]
                    min_gen = short_taken[gen]
                else:
                    min_gen = 0
                if option == 'P':
                    ns.cell(row=cap_row, column=col).value = capacity
                    ns.cell(row=cap_row, column=col).number_format = '#,##0.00'
     #               min_after[4] = shortfall[0]
      #              min_after[5] = 0
                    for row in range(8760):
                        if shortfall[row] >= 0: # shortfall?
                            if shortfall[row] >= cap_capacity:
                                shortfall[row] = shortfall[row] - cap_capacity
                                ns.cell(row=row + hrows, column=col).value = cap_capacity
                            elif shortfall[row] < min_gen:
                                ns.cell(row=row + hrows, column=col).value = min_gen
                                shortfall[row] -= min_gen
                            else:
                                ns.cell(row=row + hrows, column=col).value = shortfall[row]
                                shortfall[row] = 0
                    #        if shortfall[row] > min_after[4]:
                     #           min_after[4] = shortfall[row]
                      #          min_after[5] = row
                        else:
                            shortfall[row] -= min_gen
                            ns.cell(row=row + hrows, column=col).value = min_gen
                        ns.cell(row=row + hrows, column=col + 1).value = (shortfall[row] + short_taken_tot) * -self.surplus_sign
                        ns.cell(row=row + hrows, column=col).number_format = '#,##0.00'
                        ns.cell(row=row + hrows, column=col + 1).number_format = '#,##0.00'
                    ns.cell(row=sum_row, column=col).value = '=SUM(' + ss_col(col) + str(hrows) + \
                            ':' + ss_col(col) + str(hrows + 8759) + ')'
                    ns.cell(row=sum_row, column=col).number_format = '#,##0'
                    ns.cell(row=cf_row, column=col).value = '=IF(' + ss_col(col) + '1>0,' + \
                                                ss_col(col) + '3/' + ss_col(col) + '1/8760,"")'
                    ns.cell(row=cf_row, column=col).number_format = '#,##0.00'
                    ns.cell(row=max_row, column=col).value = '=MAX(' + ss_col(col) + \
                            str(hrows) + ':' + ss_col(col) + str(hrows + 8759) + ')'
                    ns.cell(row=max_row, column=col).number_format = '#,##0.00'
                    ns.cell(row=hrs_row, column=col).value = '=COUNTIF(' + ss_col(col) + \
                            str(hrows) + ':' + ss_col(col) + str(hrows + 8759) + ',">0")'
                    ns.cell(row=hrs_row, column=col).number_format = '#,##0'
                    ns.cell(row=hrs_row, column=col + 1).value = '=' + ss_col(col) + \
                            str(hrs_row) + '/8760'
                    ns.cell(row=hrs_row, column=col + 1).number_format = '#,##0.0%'
                    col += 2
                else:
                    gen_can = 0.
              #      min_after[4] = shortfall[0]
               #     min_after[5] = 0
                    for row in range(8760):
                        if shortfall[row] >= 0: # shortfall?
                            if shortfall[row] >= cap_capacity:
                                shortfall[row] = shortfall[row] - cap_capacity
                                gen_can += cap_capacity
                            elif shortfall[row] < min_gen:
                                gen_can += min_gen
                                shortfall[row] -= min_gen
                            else:
                                gen_can += shortfall[row]
                                shortfall[row] = 0
                #            if shortfall[row] > min_after[4]:
                 #               min_after[4] = shortfall[row]
                  #              min_after[5] = row
                        else:
                            gen_can += min_gen
                            shortfall[row] -= min_gen # ??
                    sp_data.append([gen, capacity, gen_can, '', '', '', '', '', ''])
        if option != 'O' and option != '1':
            self.progressbar.setValue(4)
        if option != 'P':
         #   if min_after[2] >= 0:
          #      min_after[2] = format_period(min_after[2])
           # else:
            #    min_after[3] = ''
           # min_after[5] = format_period(min_after[5])
            cap_sum = 0.
            gen_sum = 0.
            re_sum = 0.
            ff_sum = 0.
            sto_sum = 0.
            cost_sum = 0.
            co2_sum = 0.
            for sp in range(len(sp_data)):
                if sp_data[sp][1] > 0:
                    cap_sum += sp_data[sp][1]
                    sp_data[sp][3] = sp_data[sp][2] / sp_data[sp][1] / 8760
                gen_sum += sp_data[sp][2]
                gen = sp_data[sp][0]
                if gen in tech_names:
                    re_sum += sp_data[sp][2]
                if gen not in self.generators:
                    continue
                if gen in storage_names:
                    sto_sum += sp_data[sp][2]
                elif gen not in tech_names:
                    ff_sum += sp_data[sp][2]
                if self.generators[gen].lcoe > 0:
                    if self.remove_cost and sp_data[sp][2] == 0:
                        sp_data[sp][4] = 0
                        continue
                    if self.generators[gen].lcoe_cf > 0:
                        lcoe_cf = self.generators[gen].lcoe_cf
                    else:
                        lcoe_cf = sp_data[sp][3]
                    sp_data[sp][4] = self.generators[gen].lcoe * lcoe_cf * 8760 * sp_data[sp][1]
                    if sp_data[sp][1] > 0 and sp_data[sp][3] > 0:
                        sp_data[sp][5] = sp_data[sp][4] / 8760 / sp_data[sp][3] / sp_data[sp][1]
                    cost_sum += sp_data[sp][4]
                    sp_data[sp][7] = self.generators[gen].lcoe
                    sp_data[sp][8] = lcoe_cf
                if self.generators[gen].emissions > 0:
                    sp_data[sp][6] = sp_data[sp][2] * self.generators[gen].emissions
                    co2_sum += sp_data[sp][6]
            if cap_sum > 0:
                cs = gen_sum / cap_sum / 8760
            else:
                cs = ''
            if gen_sum > 0:
                gs = cost_sum / gen_sum
                gsw = cost_sum / sp_load # corrected LCOE
            else:
                gs = ''
                gsw = ''
            sf_sums = [0., 0., 0.]
            for sf in range(len(shortfall)):
                if shortfall[sf] > 0:
                    sf_sums[0] += shortfall[sf]
                    sf_sums[2] += pmss_data[0][sf] * pmss_details['Load'][3]
                else:
                    sf_sums[1] += shortfall[sf]
                    sf_sums[2] += pmss_data[0][sf] * pmss_details['Load'][3]
            sp_data.append(['Total', cap_sum, gen_sum, cs, cost_sum, gs, co2_sum])
            if option == 'O' or option == '1':
                op_load_tot = pmss_details['Load'][0] * pmss_details['Load'][3]
             #   if (sf_sums[2] - sf_sums[0]) / op_load_tot < 1:
              #      lcoe = 500
               # el
                if self.corrected_lcoe:
                    lcoe = gsw # target is corrected lcoe
                else:
                    lcoe = gs
                if gen_sum == 0:
                    re_pct = 0
                    load_pct = 0
                else:
                    load_pct = (sf_sums[2] - sf_sums[0]) / op_load_tot
                    try:
                        non_re = gen_sum - re_sum - sto_sum
                        re_pct = (op_load_tot - non_re) / op_load_tot
                        re_pct = re_pct * load_pct # ???
                      #  re_pct = re_sum / gen_sum
                    except:
                        re_pct = 0
                multi_value = {'lcoe': lcoe, #lcoe. lower better
                    'load_pct': (sf_sums[2] - sf_sums[0]) / op_load_tot, #load met. 100% better
                    'surplus_pct': -sf_sums[1] / op_load_tot, #surplus. lower better
                    're_pct': re_pct, # RE pct. higher better
                    'cost': cost_sum, # cost. lower better
                    'co2': co2_sum} # CO2. lower better
                if option == 'O':
                    if multi_value['lcoe'] == '':
                        multi_value['lcoe'] = 0
                    return multi_value, sp_data, None
                else:
                    extra = [gsw, op_load_tot, sto_sum, re_sum, re_pct, sf_sums]
                    return multi_value, sp_data, extra
            if self.corrected_lcoe:
                sp_data.append(['Corrected LCOE', '', '', '', '', gsw, ''])
            if self.carbon_price > 0:
                cc = co2_sum * self.carbon_price
                if self.corrected_lcoe:
                    cs = (cost_sum + cc) / sp_load
                else:
                    cs = (cost_sum + cc) / gen_sum
                if self.carbon_price == int(self.carbon_price):
                   cp = str(int(self.carbon_price))
                else:
                   cp = '{:.2f}'.format(self.carbon_price)
                sp_data.append(['Carbon Cost ($' + cp + ')', '', '', '', cc, cs])
            sp_data.append(['RE %age', round(re_sum * 100. / gen_sum, 1)])
            if sto_sum > 0:
                sp_data.append(['RE %age with Storage', round((re_sum + sto_sum) * 100. / gen_sum, 1)])
            sp_data.append(' ')
            sp_data.append('Load Analysis')
            pct = '{:.1%})'.format((sf_sums[2] - sf_sums[0]) / sp_load)
            sp_data.append(['Load met (' + pct, '', sf_sums[2] - sf_sums[0]])
            pct = '{:.1%})'.format(sf_sums[0] / sp_load)
            sp_data.append(['Shortfall (' + pct, '', sf_sums[0]])
            sp_data.append(['Total Load', '', sp_load])
            pct = '{:.1%})'.format( -sf_sums[1] / sp_load)
            sp_data.append(['Surplus (' + pct, '', -sf_sums[1]])
            sp_data.append(['RE %age of Total Load', round((sp_load - sf_sums[0] - ff_sum) * \
                           100. / sp_load, 1)])
       #     sp_data.append(' ')
        #    try:
      #          if min_after[2] >= 0:
       #             sp_data.append(['Storage Initial', round(min_after[0], 2)])
        #            sp_data.append(['Storage Minimum ' + min_after[2], round(min_after[1], 2)])
         #           sp_data.append(['Storage Final', round(min_after[3], 2)])
          #  except:
           #     pass
            max_short = [0, 0]
            for h in range(len(shortfall)):
                if shortfall[h] > max_short[1]:
                    max_short[0] = h
                    max_short[1] = shortfall[h]
            if max_short[1] > 0:
                sp_data.append(['Largest Shortfall ' + format_period(max_short[0]),
                                round(max_short[1], 2)])
            sp_data.append(' ')
            adjusted = False
            if self.adjust.isChecked():
                for gen, value in iter(sorted(pmss_details.items())):
                    if value[3] != 1:
                        if not adjusted:
                            adjusted = True
                            sp_data.append('Generators Adjustments:')
                        sp_data.append([gen, value[3]])
            list(map(list, list(zip(*sp_data))))
            sp_pts = [0, 2, 0, 2, 0, 2, 0, 2, 2]
            self.setStatus(self.sender().text() + ' completed')
            dialog = displaytable.Table(sp_data, title=self.sender().text(), fields=headers,
                     save_folder=self.scenarios, sortby='', decpts=sp_pts)
            dialog.exec_()
            self.progressbar.setValue(10)
            self.progressbar.setHidden(True)
            self.progressbar.setValue(0)
            return # finish if not detailed spreadsheet
        col = shrt_col + 1
        for gen in dispatch_order:
            ss_row += 1
            if self.constraints[self.generators[gen].constraint].category == 'Storage':
                nc = 2
                ns.cell(row=what_row, column=col).value = 'Charge\n' + gen
                ns.cell(row=what_row, column=col).alignment = oxl.styles.Alignment(wrap_text=True,
                        vertical='bottom', horizontal='center')
                ns.cell(row=what_row, column=col + 1).value = gen + '\nLosses'
                ns.cell(row=what_row, column=col + 1).alignment = oxl.styles.Alignment(wrap_text=True,
                        vertical='bottom', horizontal='center')
                is_storage = True
                sto_sum += '+C' + str(ss_row)
            else:
                nc = 0
                is_storage = False
                not_sum += '-C' + str(ss_row)
            ns.cell(row=what_row, column=col + nc).value = gen
            ss.cell(row=ss_row, column=1).value = '=Detail!' + ss_col(col + nc) + str(what_row)
            ss.cell(row=ss_row, column=2).value = '=Detail!' + ss_col(col + nc) + str(cap_row)
            ss.cell(row=ss_row, column=2).number_format = '#,##0.00'
            ss.cell(row=ss_row, column=3).value = '=Detail!' + ss_col(col + nc) + str(sum_row)
            ss.cell(row=ss_row, column=3).number_format = '#,##0'
            ss.cell(row=ss_row, column=4).value = '=Detail!' + ss_col(col + nc) + str(cf_row)
            ss.cell(row=ss_row, column=4).number_format = '#,##0.00'
            if self.generators[gen].lcoe > 0:
                capacity = self.generators[gen].capacity
                if self.adjust.isChecked():
                    try:
                        capacity = self.generators[gen].capacity * pmss_details[gen][3]
                    except:
                        pass
                ns.cell(row=cost_row, column=col + nc).value = '=IF(' + ss_col(col + nc) + str(cf_row) + \
                        '>0,' + ss_col(col + nc) + str(sum_row) + '*Summary!H' + str(ss_row) + \
                        '*Summary!I' + str(ss_row) + '/' + ss_col(col + nc) + str(cf_row) + ',0)'
                ns.cell(row=cost_row, column=col + nc).number_format = '$#,##0'
                if self.remove_cost:
                    ss.cell(row=ss_row, column=5).value = '=IF(Detail!' + ss_col(col + nc) + str(sum_row) \
                            + '>0,Detail!' + ss_col(col + nc) + str(cost_row) + ',"")'
                else:
                    ss.cell(row=ss_row, column=5).value = '=Detail!' + ss_col(col + nc) + str(cost_row)
                ss.cell(row=ss_row, column=5).number_format = '$#,##0'
                ns.cell(row=lcoe_row, column=col + nc).value = '=IF(AND(' + ss_col(col + nc) + str(cf_row) + '>0,' \
                            + ss_col(col + nc) + str(cap_row) + '>0),' + ss_col(col + nc) + str(cost_row) + '/8760/' \
                            + ss_col(col + nc) + str(cf_row) + '/' + ss_col(col + nc) + str(cap_row)+  ',"")'
                ns.cell(row=lcoe_row, column=col + nc).number_format = '$#,##0.00'
                ss.cell(row=ss_row, column=6).value = '=Detail!' + ss_col(col + nc) + str(lcoe_row)
                ss.cell(row=ss_row, column=6).number_format = '$#,##0.00'
            if self.generators[gen].emissions > 0:
                ns.cell(row=emi_row, column=col + nc).value = '=' + ss_col(col + nc) + str(sum_row) \
                        + '*' + str(self.generators[gen].emissions)
                ns.cell(row=emi_row, column=col + nc).number_format = '#,##0'
                if self.remove_cost:
                    ss.cell(row=ss_row, column=7).value = '=IF(Detail!' + ss_col(col + nc) + str(sum_row) \
                            + '>0,Detail!' + ss_col(col + nc) + str(emi_row) + ',"")'
                else:
                    ss.cell(row=ss_row, column=7).value = '=Detail!' + ss_col(col + nc) + str(emi_row)
                ss.cell(row=ss_row, column=7).number_format = '#,##0'
            ss.cell(row=ss_row, column=8).value = self.generators[gen].lcoe
            ss.cell(row=ss_row, column=8).number_format = '$#,##0.00'
            if self.generators[gen].lcoe_cf == 0:
                ss.cell(row=ss_row, column=9).value = '=D' + str(ss_row)
            else:
                ss.cell(row=ss_row, column=9).value = self.generators[gen].lcoe_cf
            ss.cell(row=ss_row, column=9).number_format = '#,##0.00'
            ns.cell(row=what_row, column=col + nc).alignment = oxl.styles.Alignment(wrap_text=True,
                    vertical='bottom', horizontal='center')
            ns.cell(row=what_row, column=col + nc + 1).alignment = oxl.styles.Alignment(wrap_text=True,
                    vertical='bottom', horizontal='center')
            if is_storage:
             #   ns.cell(row=what_row, column=col + 1).value = gen
                ns.cell(row=what_row, column=col + 3).value = gen + '\nBalance'
                ns.cell(row=what_row, column=col + 3).alignment = oxl.styles.Alignment(wrap_text=True,
                        vertical='bottom', horizontal='center')
                ns.cell(row=what_row, column=col + 4).value = 'After\n' + gen
                ns.cell(row=what_row, column=col + 4).alignment = oxl.styles.Alignment(wrap_text=True,
                        vertical='bottom', horizontal='center')
                ns.cell(row=fall_row, column=col + 4).value = '=COUNTIF(' + ss_col(col + 4) \
                        + str(hrows) + ':' + ss_col(col + 4) + str(hrows + 8759) + \
                        ',"' + sf_test[0] + '0")'
                ns.cell(row=fall_row, column=col + 4).number_format = '#,##0'
                col += 5
            else:
                ns.cell(row=what_row, column=col + 1).value = 'After\n' + gen
                ns.cell(row=fall_row, column=col + 1).value = '=COUNTIF(' + ss_col(col + 1) \
                        + str(hrows) + ':' + ss_col(col + 1) + str(hrows + 8759) + \
                        ',"' + sf_test[0] + '0")'
                ns.cell(row=fall_row, column=col + 1).number_format = '#,##0'
                col += 2
        if is_storage:
            ns.cell(row=emi_row, column=col - 2).value = '=MIN(' + ss_col(col - 2) + str(hrows) + \
                    ':' + ss_col(col - 2) + str(hrows + 8759) + ')'
            ns.cell(row=emi_row, column=col - 2).number_format = '#,##0.00'
    #    ns.cell(row=emi_row, column=col - 1).value = '=MIN(' + ss_col(col - 1) + str(hrows) + \
     #           ':' + ss_col(col - 1) + str(hrows + 8759) + ')'
      #  ns.cell(row=emi_row, column=col - 1).number_format = '#,##0.00'
        for column_cells in ns.columns:
            length = 0
            value = ''
            row = 0
            sum_value = 0
            do_sum = False
            for cell in column_cells:
                if cell.row >= hrows:
                    if do_sum:
                        try:
                            sum_value += abs(cell.value)
                        except:
                            pass
                    else:
                        try:
                            value = str(round(cell.value, 2))
                            if len(value) > length:
                                length = len(value) + 2
                        except:
                            pass
                elif cell.row > 0:
                    if str(cell.value)[0] != '=':
                        values = str(cell.value).split('\n')
                        for value in values:
                            if cell.row == cost_row:
                                valf = value.split('.')
                                alen = int(len(valf[0]) * 1.6)
                            else:
                                alen = len(value) + 2
                            if alen > length:
                                length = alen
                    else:
                        if cell.value[1:4] == 'SUM':
                            do_sum = True
            if sum_value > 0:
                alen = len(str(int(sum_value))) * 1.5
                if alen > length:
                    length = alen
            if isinstance(cell.column, int):
                cel = ss_col(cell.column)
            else:
                cel = cell.column
            ns.column_dimensions[cel].width = max(length, 10)
        ns.column_dimensions['A'].width = 6
        ns.column_dimensions['B'].width = 21
        st_row = hrows + 8760
        st_col = col
        for row in range(1, st_row):
            for col in range(1, st_col):
                try:
                    ns.cell(row=row, column=col).font = normal
                except:
                    pass
        self.progressbar.setValue(6)
        ns.row_dimensions[what_row].height = 30
        ns.freeze_panes = 'C' + str(hrows)
        ns.activeCell = 'C' + str(hrows)
        ss.cell(row=1, column=1).value = 'Powermatch - Summary'
        ss.cell(row=1, column=1).font = bold
        ss_row +=1
        for col in range(1, 10):
            ss.cell(row=3, column=col).font = bold
            ss.cell(row=ss_row, column=col).font = bold
        ss.cell(row=ss_row, column=2).value = '=SUM(B4:B' + str(ss_row - 1) + ')'
        ss.cell(row=ss_row, column=2).number_format = '#,##0.00'
        ss.cell(row=ss_row, column=3).value = '=SUM(C4:C' + str(ss_row - 1) + ')'
        ss.cell(row=ss_row, column=3).number_format = '#,##0'
        ss.cell(row=ss_row, column=4).value = '=C' + str(ss_row) + '/B' + str(ss_row) + '/8760'
        ss.cell(row=ss_row, column=4).number_format = '#,##0.00'
        ss.cell(row=ss_row, column=5).value = '=SUM(E4:E' + str(ss_row - 1) + ')'
        ss.cell(row=ss_row, column=5).number_format = '$#,##0'
        ss.cell(row=ss_row, column=6).value = '=E' + str(ss_row) + '/C' + str(ss_row)
        ss.cell(row=ss_row, column=6).number_format = '$#,##0.00'
        ss.cell(row=ss_row, column=7).value = '=SUM(G4:G' + str(ss_row - 1) + ')'
        ss.cell(row=ss_row, column=7).number_format = '#,##0'
        ss.cell(row=ss_row, column=1).value = 'Total'
        if self.corrected_lcoe:
            ss_row +=1
            ss.cell(row=ss_row, column=1).value = 'Corrected LCOE'
            ss.cell(row=ss_row, column=1).font = bold
        lcoe_row = ss_row
        for column_cells in ss.columns:
            length = 0
            value = ''
            for cell in column_cells:
                if str(cell.value)[0] != '=':
                    values = str(cell.value).split('\n')
                    for value in values:
                        if len(value) + 1 > length:
                            length = len(value) + 1
            if isinstance(cell.column, int):
                cel = ss_col(cell.column)
            else:
                cel = cell.column
            if cell.column == 'E':
                ss.column_dimensions[cel].width = max(length, 10) * 2.
            else:
                ss.column_dimensions[cel].width = max(length, 10) * 1.2
        last_col = ss_col(ns.max_column)
        r = 1
        if self.corrected_lcoe:
            r += 1
        if self.carbon_price > 0:
            ss_row += 1
            if self.carbon_price == int(self.carbon_price):
               cp = str(int(self.carbon_price))
            else:
               cp = '{:.2f}'.format(self.carbon_price)
            ss.cell(row=ss_row, column=1).value = 'Carbon Cost ($/tCO2e)'
            ss.cell(row=ss_row, column=2).value = self.carbon_price
            ss.cell(row=ss_row, column=2).number_format = '$#,##0.00'
            ss.cell(row=ss_row, column=5).value = '=G' + str(ss_row - r) + '*B' + str(ss_row)
            ss.cell(row=ss_row, column= 5).number_format = '$#,##0'
            if not self.corrected_lcoe:
                ss.cell(row=ss_row, column=6).value = '=(E' + str(ss_row - r) + \
                        '+E'  + str(ss_row) + ')/C' + str(ss_row - r)
                ss.cell(row=ss_row, column=6).number_format = '$#,##0.00'
            r += 1
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'RE %age'
        ss.cell(row=ss_row, column=3).value = re_sum[:-1] + ')/C' + str(ss_row - r)
        ss.cell(row=ss_row, column=3).number_format = '#,##0.0%'
        # if storage
        if sto_sum != '':
            ss_row += 1
            ss.cell(row=ss_row, column=1).value = 'RE %age with storage'
            ss.cell(row=ss_row, column=3).value = re_sum[:-1] + sto_sum + ')/C' + str(ss_row - r - 1)
            ss.cell(row=ss_row, column=3).number_format = '#,##0.0%'
        ss_row += 2
        ss.cell(row=ss_row, column=1).value = 'Load Analysis'
        ss.cell(row=ss_row, column=1).font = bold
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Load met'
        lm_row = ss_row
        if self.surplus_sign < 0:
            addsub = ')+C'
        else:
            addsub = ')-C'
        ss.cell(row=ss_row, column=3).value = '=SUMIF(Detail!' + last_col + str(hrows) + ':Detail!' \
            + last_col + str(hrows + 8759) + ',"' + sf_test[0] + '=0",Detail!C' + str(hrows) \
            + ':Detail!C' + str(hrows + 8759) + ')+SUMIF(Detail!' + last_col + str(hrows) + ':Detail!' \
            + last_col + str(hrows + 8759) + ',"' + sf_test[1] + '0",Detail!C' + str(hrows) + ':Detail!C' \
            + str(hrows + 8759) + addsub + str(ss_row + 1)
        ss.cell(row=ss_row, column=3).value = '=Detail!C' + str(sum_row) + '-C' + str(ss_row + 1)
        ss.cell(row=ss_row, column=3).number_format = '#,##0'
        ss.cell(row=ss_row, column=4).value = '=C' + str(ss_row) + '/C' + str(ss_row + 2)
        ss.cell(row=ss_row, column=4).number_format = '#,##0.0%'
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Shortfall'
        sf_text = 'SUMIF(Detail!' + last_col + str(hrows) + ':Detail!' + last_col \
                  + str(hrows + 8759) + ',"' + sf_test[0] + '0",Detail!' + last_col \
                  + str(hrows) + ':Detail!' + last_col + str(hrows + 8759) + ')'
        if self.surplus_sign > 0:
            ss.cell(row=ss_row, column=3).value = '=-' + sf_text
        else:
            ss.cell(row=ss_row, column=3).value = '=' + sf_text
        ss.cell(row=ss_row, column=3).number_format = '#,##0'
        ss.cell(row=ss_row, column=4).value = '=C' + str(ss_row) + '/C' + str(ss_row + 1)
        ss.cell(row=ss_row, column=4).number_format = '#,##0.0%'
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Total Load'
        ss.cell(row=ss_row, column=1).font = bold
        ss.cell(row=ss_row, column=3).value = '=SUM(C' + str(ss_row - 2) + ':C' + str(ss_row - 1) + ')'
        ss.cell(row=ss_row, column=3).number_format = '#,##0'
        ss.cell(row=ss_row, column=3).font = bold
        # values for corrected LCOE and Carbon Cost LCOE
        if self.corrected_lcoe:
            ss.cell(row=lcoe_row, column=6).value = '=F' + str(lcoe_row-1) + '*C' + str(lcoe_row-1) + \
                    '/C' + str(ss_row)
            ss.cell(row=lcoe_row, column=6).number_format = '$#,##0.00'
            ss.cell(row=lcoe_row, column=6).font = bold
            if self.carbon_price > 0:
                ss.cell(row=lcoe_row + 1, column=6).value = '=(E' + str(lcoe_row - 1) + \
                        '+E'  + str(lcoe_row + 1) + ')/C' + str(ss_row)
                ss.cell(row=lcoe_row + 1, column=6).number_format = '$#,##0.00'
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Surplus'
        sf_text = 'SUMIF(Detail!' + last_col + str(hrows) + ':Detail!' + last_col \
                  + str(hrows + 8759) + ',"' + sf_test[1] + '0",Detail!' + last_col + str(hrows) \
                  + ':Detail!' + last_col + str(hrows + 8759) + ')'
        if self.surplus_sign < 0:
            ss.cell(row=ss_row, column=3).value = '=-' + sf_text
        else:
            ss.cell(row=ss_row, column=3).value = '=' + sf_text
        ss.cell(row=ss_row, column=3).number_format = '#,##0'
        ss.cell(row=ss_row, column=4).value = '=C' + str(ss_row) + '/C' + str(ss_row - 1)
        ss.cell(row=ss_row, column=4).number_format = '#,##0.0%'
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'RE %age of Total Load'
        ss.cell(row=ss_row, column=1).font = bold
        if not_sum == '':
            ss.cell(row=ss_row, column=3).value = re_sum[:-1] + '-C' + str(ss_row - 1) + ')/C' + \
                                                  str(ss_row - 2)
        else:
            ss.cell(row=ss_row, column=3).value = '=(C' + str(lm_row) + not_sum + ')/C' + \
                                                  str(ss_row - 2)
        ss.cell(row=ss_row, column=3).number_format = '#,##0.0%'
        ss.cell(row=ss_row, column=3).font = bold
     #   if min_after[4] > 0:
       #     ss_row += 1
        #    ss.cell(row=ss_row, column=1).value = 'Shortfall Minimum:'
         #   ss.cell(row=ss_row, column=2).value = '=Detail!B' + str(hrows + min_after[5])
          #  ss.cell(row=ss_row, column=3).value = '=Detail!' + last_col + str(hrows + min_after[5])
           # ss.cell(row=ss_row, column=3).number_format = '#,##0.00'
        max_short = [0, 0]
        for h in range(len(shortfall)):
            if shortfall[h] > max_short[1]:
                max_short[0] = h
                max_short[1] = shortfall[h]
        if max_short[1] > 0:
            ss_row += 1
            ss.cell(row=ss_row, column=1).value = 'Largest Shortfall:'
            ss.cell(row=ss_row, column=2).value = '=Detail!B' + str(hrows + max_short[0])
            ss.cell(row=ss_row, column=3).value = '=Detail!' + last_col + str(hrows + max_short[0])
            ss.cell(row=ss_row, column=3).number_format = '#,##0.00'
        ss_row += 2
        ss.cell(row=ss_row, column=1).value = 'Data sources:'
        ss.cell(row=ss_row, column=1).font = bold
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Scenarios folder'
        ss.cell(row=ss_row, column=2).value = self.scenarios
        ss.merge_cells('B' + str(ss_row) + ':I' + str(ss_row))
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Powermatch data file'
        if pm_data_file[: len(self.scenarios)] == self.scenarios:
            pm_data_file = pm_data_file[len(self.scenarios):]
        ss.cell(row=ss_row, column=2).value = pm_data_file
        ss.merge_cells('B' + str(ss_row) + ':I' + str(ss_row))
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Constraints worksheet'
        ss.cell(row=ss_row, column=2).value = str(self.files[C].text()) \
               + '.' + str(self.sheets[C].currentText())
        ss.merge_cells('B' + str(ss_row) + ':I' + str(ss_row))
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Facility worksheet'
        ss.cell(row=ss_row, column=2).value = str(self.files[G].text()) \
               + '.' + str(self.sheets[G].currentText())
        ss.merge_cells('B' + str(ss_row) + ':I' + str(ss_row))
        self.progressbar.setValue(7)
        try:
            if self.adjust.isChecked():
                adjusted = ''
                for key, value in iter(sorted(pmss_details.items())):
                    if value != 1:
                        adjusted += key + ': {:.4f}'.format(value[3]).rstrip('0') + '; '
                if len(adjusted) > 0:
                    ss_row += 1
                    ss.cell(row=ss_row, column=1).value = 'Inputs adjusted'
                    ss.cell(row=ss_row, column=2).value = adjusted[:-2]
                    ss.merge_cells('B' + str(ss_row) + ':I' + str(ss_row))
        except:
            pass
        for row in range(1, ss_row + 1):
            for col in range(1, 10):
                try:
                    if ss.cell(row=row, column=col).font.name != 'Arial':
                        ss.cell(row=row, column=col).font = normal
                except:
                    pass
        ss.freeze_panes = 'B4'
        ss.activeCell = 'B4'
        ds.save(data_file)
        self.progressbar.setValue(10)
        j = data_file.rfind('/')
        data_file = data_file[j + 1:]
        msg = '%s created (%.2f seconds)' % (data_file, time.time() - start_time)
        msg = '%s created.' % data_file
        self.setStatus(msg)
        self.progressbar.setHidden(True)
        self.progressbar.setValue(0)

    def show_ProgressBar(self, maximum, msg, title):
        if self.opt_progressbar is None:
            self.opt_progressbar = ProgressBar(maximum=maximum, msg=msg, title=title)
            self.opt_progressbar.setWindowModality(QtCore.Qt.WindowModal)
         #   self.opt_progressbar.progress.connect(self.opt_progressbar.progress)
         #   self.opt_progressbar.range.connect(self.opt_progressbar.range)
            self.opt_progressbar.show()
            self.opt_progressbar.setVisible(False)
            self.activateWindow()
        else:
            self.opt_progressbar.barRange(0, maximum, msg=msg)

    def show_FloatStatus(self):
        if not self.log_status:
            return
        if self.floatstatus is None:
            self.floatstatus = FloatStatus(self, self.scenarios, None, program='Powermatch')
            self.floatstatus.setWindowModality(QtCore.Qt.WindowModal)
            self.floatstatus.setWindowFlags(self.floatstatus.windowFlags() |
                         QtCore.Qt.WindowSystemMenuHint |
                         QtCore.Qt.WindowMinMaxButtonsHint)
            self.floatstatus.procStart.connect(self.getStatus)
       #     self.floatstatus.log.connect(self.floatstatus.log)
            self.floatstatus.show()
            self.activateWindow()

    def setStatus(self, text):
        if self.log.text() == text:
            return
        self.log.setText(text)
        if text == '':
            return
        if self.floatstatus and self.log_status:
            self.floatstatus.log(text)
            QtWidgets.QApplication.processEvents()

    @QtCore.pyqtSlot(str)
    def getStatus(self, text):
        if text == 'goodbye':
            self.floatstatus = None

    def exit(self):
        self.updated = False
        if self.floatstatus is not None:
            self.floatstatus.exit()
        self.close()

    def optClicked(self, in_year, in_option, in_pmss_details, in_pmss_data, in_re_order,
                   in_dispatch_order, pm_data_file, data_file):

        def create_starting_population(individuals, chromosome_length):
            # Set up an initial array of all zeros
            population = np.zeros((individuals, chromosome_length))
            # Loop through each row (individual)
            for i in range(individuals):
                # Choose a random number of ones to create
                ones = random.randint(0, chromosome_length)
                # Change the required number of zeros to ones
                population[i, 0:ones] = 1
                # Sfuffle row
                np.random.shuffle(population[i])
            return population

        def select_individual_by_tournament(population, *argv):
            # Get population size
            population_size = len(population)

            # Pick individuals for tournament
            fighter = [0, 0]
            fighter_fitness = [0, 0]
            fighter[0] = random.randint(0, population_size-1)
            fighter[1] = random.randint(0, population_size-1)

            # Get fitness score for each
            if len(argv) == 1:
                fighter_fitness[0] = argv[0][fighter[0]]
                fighter_fitness[1] = argv[0][fighter[1]]
            else:
                for arg in argv:
                    min1 = min(arg)
                    max1 = max(arg)
                    for f in range(len(fighter)):
                        try:
                            fighter_fitness[f] += (arg[f] - min1) / (max1 - min1)
                        except:
                            pass
            # Identify individual with lowest score
            # Fighter 1 will win if score are equal
            if fighter_fitness[0] <= fighter_fitness[1]:
                winner = fighter[0]
            else:
                winner = fighter[1]

            # Return the chromsome of the winner
            return population[winner, :]

        def breed_by_crossover(parent_1, parent_2, points=2):
            # Get length of chromosome
            chromosome_length = len(parent_1)

            # Pick crossover point, avoiding ends of chromsome
            if points == 1:
                crossover_point = random.randint(1,chromosome_length-1)
            # Create children. np.hstack joins two arrays
                child_1 = np.hstack((parent_1[0:crossover_point],
                                     parent_2[crossover_point:]))
                child_2 = np.hstack((parent_2[0:crossover_point],
                                     parent_1[crossover_point:]))
            else: # only do 2 at this    stage
                crossover_point_1 = random.randint(1, chromosome_length - 2)
                crossover_point_2 = random.randint(crossover_point_1 + 1, chromosome_length - 1)
                child_1 = np.hstack((parent_1[0:crossover_point_1],
                                     parent_2[crossover_point_1:crossover_point_2],
                                     parent_1[crossover_point_2:]))
                child_2 = np.hstack((parent_2[0:crossover_point_1],
                                     parent_1[crossover_point_1:crossover_point_2],
                                     parent_2[crossover_point_2:]))
            # Return children
            return child_1, child_2

        def randomly_mutate_population(population, mutation_probability):
            # Apply random mutation
            random_mutation_array = np.random.random(size=(population.shape))
            random_mutation_boolean = random_mutation_array <= mutation_probability
        #    random_mutation_boolean[0][:] = False # keep the best multi and lcoe
       #     random_mutation_boolean[1][:] = False
            population[random_mutation_boolean] = np.logical_not(population[random_mutation_boolean])
            # Return mutation population
            return population

        def calculate_fitness(population):
            lcoe_fitness_scores = [] # scores = LCOE values
            multi_fitness_scores = [] # scores = multi-variable weight
            multi_values = [] # values for each of the six variables
            if len(population) == 1:
                option = '1'
            else:
                option = 'O'
            if self.debug:
                self.popn += 1
                self.chrom = 0
            for chromosome in population:
                # now get random amount of generation per technology (both RE and non-RE)
                for gen, value in opt_order.items():
                    capacity = value[2]
                    for c in range(value[0], value[1]):
                        if chromosome[c]:
                            capacity = capacity + capacities[c]
                    pmss_details[gen][3] = capacity / pmss_details[gen][0]
                multi_value, op_data, extra = self.doDispatch(year, option, pmss_details, pmss_data, re_order,
                                              dispatch_order, pm_data_file, data_file)
                lcoe_fitness_scores.append(multi_value['lcoe'])
                multi_values.append(multi_value)
                multi_fitness_scores.append(calc_weight(multi_value))
                if self.debug:
                    self.chrom += 1
                    line = str(self.popn) + ',' + str(self.chrom) + ','
                    for gen, value in opt_order.items():
                        try:
                            line += str(pmss_details[gen][0] * pmss_details[gen][3]) + ','
                        except:
                            line += ','
                    for key in self.targets.keys():
                        try:
                            line += '{:.3f},'.format(multi_value[key])
                        except:
                            line += multi_value[key] + ','
                    line += '{:.5f},'.format(multi_fitness_scores[-1])
                    self.db_file.write(line + '\n')
            # alternative approach to calculating fitness
            multi_fitness_scores1 = []
            maxs = {}
            mins = {}
            tgts = {}
            for key in multi_value.keys():
                if key[-4:] == '_pct':
                    tgts[key] = abs(self.targets[key][2] - self.targets[key][3])
                else:
                    maxs[key] = 0
                    mins[key] = -1
                    for popn in multi_values:
                        try:
                            tgt = abs(self.targets[key][2] - popn[key])
                        except:
                            continue
                        if tgt > maxs[key]:
                            maxs[key] = tgt
                        if mins[key] < 0 or tgt < mins[key]:
                            mins[key] = tgt
            for popn in multi_values:
                weight = 0
                for key, value in multi_value.items():
                    if self.targets[key][1] <= 0:
                        continue
                    try:
                        tgt = abs(self.targets[key][2] - popn[key])
                    except:
                        continue
                    if key[-4:] == '_pct':
                        if tgts[key] != 0:
                            if tgt > tgts[key]:
                                weight += 1 * self.targets[key][1]
                            else:
                                try:
                                    weight += 1 - ((tgt / tgts[key]) * self.targets[key][1])
                                except:
                                    pass
                    else:
                        try:
                            weight += 1 - (((maxs[key] - tgt) / (maxs[key] - mins[key])) \
                                      * self.targets[key][1])
                        except:
                            pass
                multi_fitness_scores1.append(weight)
            if len(population) == 1: # return the table for best chromosome
            # a real fudge for the moment
                cap_sum, gen_sum, cs, cost_sum, gs, co2_sum = op_data[-1][1:]
                gsw, op_load_tot, sto_sum, re_sum, re_pct, sf_sums = extra
                if self.corrected_lcoe:
                    op_data.append(['Corrected LCOE', '', '', '', '', gsw, ''])
                if self.carbon_price > 0:
                    cc = co2_sum * self.carbon_price
                    if self.corrected_lcoe:
                        cs = (cost_sum + cc) / op_load_tot
                    else:
                        cs = (cost_sum + cc) / gen_sum
                    if self.carbon_price == int(self.carbon_price):
                       cp = str(int(self.carbon_price))
                    else:
                       cp = '{:.2f}'.format(self.carbon_price)
                    op_data.append(['Carbon Cost ($' + cp + ')', '', '', '', cc, cs])
                try:
                    op_data.append(['RE %age', round(re_sum * 100. / gen_sum, 1)])
                except:
                    pass
                if sto_sum > 0:
                    op_data.append(['RE %age with Storage', round((re_sum + sto_sum) * 100. / gen_sum, 1)])
                op_data.append(' ')
                op_data.append('Load Analysis')
                pct = '{:.1%})'.format((sf_sums[2] - sf_sums[0]) / op_load_tot)
                op_data.append(['Load met (' + pct, '', sf_sums[2] - sf_sums[0],])
                pct = '{:.1%})'.format(sf_sums[0] / op_load_tot)
                op_data.append(['Shortfall (' + pct, '', sf_sums[0]])
                op_data.append(['Total Load', '', op_load_tot])
                pct = '{:.1%})'.format( -sf_sums[1] / op_load_tot)
                op_data.append(['Surplus (' + pct, '', -sf_sums[1]])
                op_data.append(['RE %age of load', round(re_pct * 100., 2)])
                return op_data, multi_values
            else:
                return lcoe_fitness_scores, multi_fitness_scores, multi_values

        def optQuitClicked(event):
            self.optExit = True
            optDialog.close()

        def chooseClicked(event):
            self.opt_choice = self.sender().text()
            chooseDialog.close()

        def calc_weight(multi_value, calc=0):
            weight = [0., 0.]
            if calc == 0:
                for key, value in self.targets.items():
                    if multi_value[key] == '':
                        continue
                    if value[1] <= 0:
                        continue
                    if value[2] == value[3]: # wants specific target
                        if multi_value[key] == value[2]:
                            w = 0
                        else:
                            w = 1
                    elif value[2] > value[3]: # wants higher target
                        if multi_value[key] > value[2]: # high no weight
                            w = 0.
                        elif multi_value[key] < value[3]: # low maximum weight
                            w = 1.
                        else:
                            w = 1 - (multi_value[key] - value[3]) / (value[2] - value[3])
                    else: # lower target
                        if multi_value[key] == -1 or multi_value[key] > value[3]: # high maximum weight
                            w = 1.
                        elif multi_value[key] < value[2]: # low no weight
                            w = 0.
                        else:
                            w = multi_value[key] / (value[3] - value[2])
                    weight[0] += w * value[1]
            elif calc == 1:
                for key, value in self.targets.items():
                    if multi_value[key] == '':
                        continue
                    if value[1] <= 0:
                        continue
                    if multi_value[key] < 0:
                        w = 1
                    elif value[2] == value[3]: # wants specific target
                        if multi_value[key] == value[2]:
                            w = 0
                        else:
                            w = 1
                    else: # target range
                        w = min(abs(value[2] - multi_value[key]) / abs(value[2] - value[3]), 1)
                    weight[1] += w * value[1]
            return weight[calc]

        def plot_multi(multi_best, multi_order, title):
            data = [[], [], []]
            max_amt = [0., 0.]
            for multi in multi_best:
                max_amt[0] = max(max_amt[0], multi['cost'])
                max_amt[1] = max(max_amt[1], multi['co2'])
            pwr_chr = ['', '']
            divisor = [1., 1.]
            pwr_chrs = ' KMBTPEZY'
            for m in range(2):
                for pwr in range(len(pwr_chrs) - 1, -1, -1):
                    if max_amt[m] > pow(10, pwr * 3):
                        pwr_chr[m] = pwr_chrs[pwr]
                        divisor[m] = 1. * pow(10, pwr * 3)
                        break
            self.targets['cost'][5] = self.targets['cost'][5].replace('pwr_chr', pwr_chr[0])
            self.targets['co2'][5] = self.targets['co2'][5].replace('pwr_chr', pwr_chr[1])
            for multi in multi_best:
                for axis in range(3): # only three axes in plot
                    if multi_order[axis] == 'cost':
                        data[axis].append(multi[multi_order[axis]] / divisor[0]) # cost
                    elif multi_order[axis] == 'co2':
                        data[axis].append(multi[multi_order[axis]] / divisor[1]) # co2
                    elif multi_order[axis][-4:] == '_pct': # percentage
                        data[axis].append(multi[multi_order[axis]] * 100.)
                    else:
                        data[axis].append(multi[multi_order[axis]])
            fig = plt.figure(title + QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), '_yyyy-MM-dd_hhmm'))
            mx = fig.gca(projection='3d')
            plt.title('\n' + title.title() + '\n')
            try:
                surf = mx.scatter(data[0], data[1], data[2], picker=1) # enable picking a point
            except:
                print('(2800)', data[0], data[1], data[2])
                return
            mx.xaxis.set_major_formatter(FormatStrFormatter(self.targets[multi_order[0]][5]))
            mx.yaxis.set_major_formatter(FormatStrFormatter(self.targets[multi_order[1]][5]))
            mx.zaxis.set_major_formatter(FormatStrFormatter(self.targets[multi_order[2]][5]))
            mx.set_xlabel(self.targets[multi_order[0]][6])
            mx.set_ylabel(self.targets[multi_order[1]][6])
            mx.set_zlabel(self.targets[multi_order[2]][6])
            zp = ZoomPanX()
            f = zp.zoom_pan(mx, base_scale=1.2, annotate=True)
            plt.show()
            if zp.datapoint is not None: # user picked a point
                if self.more_details:
                    for p in zp.datapoint:
                        msg = 'Generation ' + str(p[0]) + ': ' + \
                              self.targets[multi_order[0]][6].replace('%', '%%') + ': ' + \
                              self.targets[multi_order[0]][5] + '; ' + \
                              self.targets[multi_order[1]][6].replace('%', '%%') + ': ' + \
                              self.targets[multi_order[1]][5] + '; ' + \
                              self.targets[multi_order[2]][6].replace('%', '%%') + ': ' + \
                              self.targets[multi_order[2]][5]
                        msg = msg % (p[1], p[2], p[3])
                        self.setStatus(msg)
            return zp.datapoint

#       optClicked mainline starts here
        year = in_year
        option = in_option
        pmss_details = dict(in_pmss_details)
        pmss_data = in_pmss_data[:]
        re_order = in_re_order[:]
        dispatch_order = in_dispatch_order[:]
        if self.optimise_debug:
            self.debug = True
        else:
            self.debug = False
        check = ''
        for gen in re_order:
            if gen == 'Load':
                continue
            if gen not in self.optimisation.keys():
                check += gen + ', '
        for gen in dispatch_order:
            if gen not in self.optimisation.keys():
                check += gen + ', '
        if check != '':
            check = check[:-2]
            self.setStatus('Key Error: Missing Optimisation entries for: ' + check)
            return
        self.optExit = False
        self.setStatus('Optimise processing started')
        err_msg = ''
        optExit = False
        optDialog = QtWidgets.QDialog()
        grid = QtWidgets.QGridLayout()
        grid.addWidget(QtWidgets.QLabel('Adjust load'), 0, 0)
        optLoad = QtWidgets.QDoubleSpinBox()
        optLoad.setRange(-1, self.adjust_cap)
        optLoad.setDecimals(3)
        optLoad.setSingleStep(.1)
        optLoad.setValue(pmss_details['Load'][3])
        rw = 0
        grid.addWidget(optLoad, rw, 1)
        grid.addWidget(QtWidgets.QLabel('Multiplier for input Load'), rw, 2, 1, 3)
        rw += 1
        grid.addWidget(QtWidgets.QLabel('Population size'), rw, 0)
        optPopn = QtWidgets.QSpinBox()
        optPopn.setRange(10, 500)
        optPopn.setSingleStep(10)
        optPopn.setValue(self.optimise_population)
        optPopn.valueChanged.connect(self.changes)
        grid.addWidget(optPopn, rw, 1)
        grid.addWidget(QtWidgets.QLabel('Size of population'), rw, 2, 1, 3)
        rw += 1
        grid.addWidget(QtWidgets.QLabel('No. of generations'), rw, 0, 1, 3)
        optGenn = QtWidgets.QSpinBox()
        optGenn.setRange(10, 500)
        optGenn.setSingleStep(10)
        optGenn.setValue(self.optimise_generations)
        optGenn.valueChanged.connect(self.changes)
        grid.addWidget(optGenn, rw, 1)
        grid.addWidget(QtWidgets.QLabel('Number of generations (iterations)'), rw, 2, 1, 3)
        rw += 1
        grid.addWidget(QtWidgets.QLabel('Mutation probability'), rw, 0)
        optMutn = QtWidgets.QDoubleSpinBox()
        optMutn.setRange(0, 1)
        optMutn.setDecimals(4)
        optMutn.setSingleStep(0.001)
        optMutn.setValue(self.optimise_mutation)
        optMutn.valueChanged.connect(self.changes)
        grid.addWidget(optMutn, rw, 1)
        grid.addWidget(QtWidgets.QLabel('Add in mutation'), rw, 2, 1, 3)
        rw += 1
        grid.addWidget(QtWidgets.QLabel('Exit if stable'), rw, 0)
        optStop = QtWidgets.QSpinBox()
        optStop.setRange(0, 50)
        optStop.setSingleStep(10)
        optStop.setValue(self.optimise_stop)
        optStop.valueChanged.connect(self.changes)
        grid.addWidget(optStop, rw, 1)
        grid.addWidget(QtWidgets.QLabel('Exit if LCOE/weight remains the same after this many iterations'),
                       rw, 2, 1, 3)
        rw += 1
        grid.addWidget(QtWidgets.QLabel('Optimisation choice'), rw, 0)
        optCombo = QtWidgets.QComboBox()
        choices = ['LCOE', 'Multi', 'Both']
        for choice in choices:
            optCombo.addItem(choice)
            if choice == self.optimise_choice:
                optCombo.setCurrentIndex(optCombo.count() - 1)
        grid.addWidget(optCombo, rw, 1)
        grid.addWidget(QtWidgets.QLabel('Choose type of optimisation'),
                       rw, 2, 1, 3)
        rw += 1
        # for each variable name
        grid.addWidget(QtWidgets.QLabel('Variable'), rw, 0)
        grid.addWidget(QtWidgets.QLabel('Weight'), rw, 1)
        grid.addWidget(QtWidgets.QLabel('Better'), rw, 2)
        grid.addWidget(QtWidgets.QLabel('Worse'), rw, 3)
        rw += 1
        ndx = grid.count()
        for key in self.targets.keys():
            self.targets[key][4] = ndx
            ndx += 4
        for key, value in self.targets.items():
            if value[2] == value[3]:
                ud = '(=)'
            elif value[2] < 0:
                ud = '(<html>&uarr;</html>)'
            elif value[3] < 0 or value[3] > value[2]:
                ud = '(<html>&darr;</html>)'
            else:
                ud = '(<html>&uarr;</html>)'
            grid.addWidget(QtWidgets.QLabel(value[0] + ': ' + ud), rw, 0)
            weight = QtWidgets.QDoubleSpinBox()
            weight.setRange(0, 1)
            weight.setDecimals(2)
            weight.setSingleStep(0.05)
            weight.setValue(value[1])
            grid.addWidget(weight, rw, 1)
            if key[-4:] == '_pct':
                minim = QtWidgets.QDoubleSpinBox()
                minim.setRange(-.1, 1.)
                minim.setDecimals(2)
                minim.setSingleStep(0.1)
                minim.setValue(value[2])
                grid.addWidget(minim, rw, 2)
                maxim = QtWidgets.QDoubleSpinBox()
                maxim.setRange(-.1, 1.)
                maxim.setDecimals(2)
                maxim.setSingleStep(0.1)
                maxim.setValue(value[3])
                grid.addWidget(maxim, rw, 3)
            else:
                minim = QtWidgets.QLineEdit()
                minim.setValidator(QtGui.QDoubleValidator())
                minim.validator().setDecimals(2)
                minim.setText(str(value[2]))
                grid.addWidget(minim, rw, 2)
                maxim = QtWidgets.QLineEdit()
                maxim.setValidator(QtGui.QDoubleValidator())
                maxim.validator().setDecimals(2)
                maxim.setText(str(value[3]))
                grid.addWidget(maxim, rw, 3)
            rw += 1
        quit = QtWidgets.QPushButton('Quit', self)
        grid.addWidget(quit, rw, 0)
        quit.clicked.connect(optQuitClicked)
        show = QtWidgets.QPushButton('Proceed', self)
        grid.addWidget(show, rw, 1)
        show.clicked.connect(optDialog.close)
        optDialog.setLayout(grid)
        optDialog.setWindowTitle('Choose Optimisation Parameters')
        optDialog.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        optDialog.exec_()
        if self.optExit: # a fudge to exit
            self.setStatus('Execution aborted.')
            return
        # check we have optimisation entries for generators and storage
        # update any changes to targets
        self.optimise_choice = optCombo.currentText()
        for key in self.targets.keys():
            weight = grid.itemAt(self.targets[key][4] + 1).widget()
            self.targets[key][1] = weight.value()
            minim = grid.itemAt(self.targets[key][4] + 2).widget()
            try:
                self.targets[key][2] = minim.value()
            except:
                self.targets[key][2] = float(minim.text())
            maxim = grid.itemAt(self.targets[key][4] + 3).widget()
            try:
                self.targets[key][3] = maxim.value()
            except:
                self.targets[key][3] = float(maxim.text())
        updates = {}
        lines = []
        lines.append('optimise_choice=' + self.optimise_choice)
        lines.append('optimise_generations=' + str(self.optimise_generations))
        lines.append('optimise_mutation=' + str(self.optimise_mutation))
        lines.append('optimise_population=' + str(self.optimise_population))
        lines.append('optimise_stop=' + str(self.optimise_stop))
        if self.optimise_choice == 'LCOE':
            do_lcoe = True
            do_multi = False
        elif self.optimise_choice == 'Multi':
            do_lcoe = False
            do_multi = True
        else:
            do_lcoe = True
            do_multi = True
        multi_order = []
        for key, value in self.targets.items():
            line = 'optimise_{}={:.2f},{:.2f},{:.2f}'.format(key, value[1], value[2], value[3])
            lines.append(line)
            multi_order.append('{:.2f}{}'.format(value[1], key))
        updates['Powermatch'] = lines
        SaveIni(updates)
        multi_order.sort(reverse=True)
    #    multi_order = multi_order[:3] # get top three weighted variables - but I want them all
        multi_order = [o[4:] for o in multi_order]
        self.adjust_re = True
        orig_load = []
        load_col = -1
        orig_tech = {}
        orig_capacity = {}
        opt_order = {} # rely on it being processed in added order
        # each entry = [first entry in chrom, last entry, minimum capacity]
        # first get original renewables generation from data sheet
        for gen in re_order:
            if gen == 'Load':
                continue
            opt_order[gen] = [0, 0, 0]
        # now add scheduled generation
        for gen in dispatch_order:
            opt_order[gen] = [0, 0, 0]
        capacities = []
        for gen in opt_order.keys():
            opt_order[gen][0] = len(capacities) # first entry
            try:
                if self.optimisation[gen].approach == 'Discrete':
                    capacities.extend(self.optimisation[gen].capacities)
                    opt_order[gen][1] = len(capacities) # last entry
                elif self.optimisation[gen].approach == 'Range':
                    if self.optimisation[gen].capacity_max == self.optimisation[gen].capacity_min:
                        capacities.extend([0])
                        opt_order[gen][1] = len(capacities)
                        opt_order[gen][2] = self.optimisation[gen].capacity_min
                        continue
                    ctr = int((self.optimisation[gen].capacity_max - self.optimisation[gen].capacity_min) / \
                              self.optimisation[gen].capacity_step)
                    if ctr < 1:
                        self.setStatus("Error with Optimisation table entry for '" + gen + "'")
                        return
                    capacities.extend([self.optimisation[gen].capacity_step] * ctr)
                    tot = self.optimisation[gen].capacity_step * ctr + self.optimisation[gen].capacity_min
                    if tot < self.optimisation[gen].capacity_max:
                        capacities.append(self.optimisation[gen].capacity_max - tot)
                    opt_order[gen][1] = len(capacities)
                    opt_order[gen][2] = self.optimisation[gen].capacity_min
                else:
                    opt_order[gen][1] = len(capacities)
            except KeyError as err:
                self.setStatus('Key Error: No Optimisation entry for ' + str(err))
                opt_order[gen] = [len(capacities), len(capacities) + 5, 0]
                capacities.extend([pmss_details[gen][0] / 5.] * 5)
            except ZeroDivisionError as err:
                self.setStatus('Zero capacity: ' + gen + ' ignored')
            except:
                err = str(sys.exc_info()[0]) + ',' + str(sys.exc_info()[1]) + ',' + gen + ',' \
                      + str(opt_order[gen])
                self.setStatus('Error: ' + str(err))
                return
        # chromosome = [1] * int(len(capacities) / 2) + [0] * (len(capacities) - int(len(capacities) / 2))
        # we have the original data - from here down we can do our multiple optimisations
        # Set general parameters
        self.setStatus('Optimisation choice is ' + self.optimise_choice)
        chromosome_length = len(capacities)
        self.setStatus(f'Chromosome length: {chromosome_length}; {pow(2, chromosome_length):,} permutations')
        self.optimise_population = optPopn.value()
        population_size = self.optimise_population
        self.optimise_generations = optGenn.value()
        maximum_generation = self.optimise_generations
        self.optimise_mutation = optMutn.value()
        self.optimise_stop = optStop.value()
        lcoe_scores = []
        multi_scores = []
        multi_values = []
     #   if do_lcoe:
      #      lcoe_target = 0. # aim for this LCOE
        if do_multi:
            multi_best = [] # list of six variables for best weight
            multi_best_popn = [] # list of chromosomes for best weight
        self.show_ProgressBar(maximum=optGenn.value(), msg='Process generations', title='SIREN - Powermatch Progress')
        self.opt_progressbar.setVisible(True)
        start_time = time.time()
        # Create starting population
        self.opt_progressbar.barProgress(1, 'Processing generation 1')
        QtCore.QCoreApplication.processEvents()
        population = create_starting_population(population_size, chromosome_length)
        # calculate best score(s) in starting population
        # if do_lcoe best_score = lowest lcoe
        # if do_multi best_multi = lowest weight and if not do_lcoe best_score also = best_weight
        if self.debug:
            filename = self.scenarios + 'opt_debug_' + \
                       QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(),
                       'yyyy-MM-dd_hhmm') + '.csv'
            self.db_file = open(filename, 'w')
            line0 = 'Popn,Chrom,'
            line1 = 'Weights,,'
            line2 = 'Targets,,'
            line3 = 'Range,' + str(population_size) + ','
            for gen, value in opt_order.items():
                 line0 += gen + ','
                 line1 += ','
                 line2 += ','
                 line3 += ','
            for key in self.targets.keys():
                 line0 += key + ','
                 line1 += str(self.targets[key][1]) + ','
                 line2 += str(self.targets[key][2]) + ','
                 if key[-4:] == '_pct':
                     line3 += str(abs(self.targets[key][2] - self.targets[key][3])) + ','
                 else:
                     line3 += ','
            line0 += 'Score'
            self.db_file.write(line0 + '\n' + line1 + '\n' + line2 + '\n' + line3 + '\n')
            self.popn = 0
            self.chrom = 0
        lcoe_scores, multi_scores, multi_values = calculate_fitness(population)
        if do_lcoe:
            try:
                best_score = np.min(lcoe_scores)
            except:
                print('(3133)', lcoe_scores)
            best_ndx = lcoe_scores.index(best_score)
            lowest_chrom = population[best_ndx]
            self.setStatus('Starting LCOE: $%.2f' % best_score)
        if do_multi:
            if self.more_details: # display starting population ?
                pick = plot_multi(multi_values, multi_order, 'starting population')
            # want maximum from first round to set base upper limit
            for key in self.targets.keys():
                if self.targets[key][2] < 0: # want a maximum from first round
                    setit = 0
                    for multi in multi_values:
                        setit = max(multi[key], setit)
                    self.targets[key][2] = setit
                if self.targets[key][3] < 0: # want a maximum from first round
                    setit = 0
                    for multi in multi_values:
                        setit = max(multi[key], setit)
                    self.targets[key][3] = setit
            # now we can find the best weighted result - lowest is best
            best_multi = np.min(multi_scores)
            best_mndx = multi_scores.index(best_multi)
            multi_lowest_chrom = population[best_mndx]
            multi_best_popn.append(multi_lowest_chrom)
            multi_best.append(multi_values[best_mndx])
            self.setStatus('Starting Weight: %.4f' % best_multi)
            multi_best_weight = best_multi
            if not do_lcoe:
                best_score = best_multi
            last_multi_score = best_multi
            lowest_multi_score = best_multi
            mud = '='
        # Add starting best score to progress tracker
        best_score_progress = [best_score]
        best_ctr = 1
        last_score = best_score
        lowest_score = best_score
        lud = '='
        # Now we'll go through the generations of genetic algorithm
        for generation in range(1, maximum_generation):
            lcoe_status = ''
            multi_status = ''
            if do_lcoe:
                lcoe_status = ' %s $%.2f ;' % (lud, best_score)
            if do_multi:
                multi_status = ' %s %.4f ;' % (mud, best_multi)
            tim = (time.time() - start_time)
            if tim < 60:
                tim = ' (%s%s %.1f secs)' % (lcoe_status, multi_status, tim)
            else:
                tim = ' (%s%s %.2f mins)' % (lcoe_status, multi_status, tim / 60.)
            self.opt_progressbar.barProgress(generation + 1,
                'Processing generation ' + str(generation + 1) + tim)
            QtWidgets.QApplication.processEvents()
            if not self.opt_progressbar.be_open:
                break
        # Create an empty list for new population
            new_population = []
        # Using elitism approach include best individual
            if do_lcoe:
                new_population.append(lowest_chrom)
            if do_multi:
                new_population.append(multi_lowest_chrom)
            # Create new population generating two children at a time
            if do_lcoe:
                if do_multi:
                    for i in range(int(population_size/2)):
                        parent_1 = select_individual_by_tournament(population, lcoe_scores,
                                                                   multi_scores)
                        parent_2 = select_individual_by_tournament(population, lcoe_scores,
                                                                   multi_scores)
                        child_1, child_2 = breed_by_crossover(parent_1, parent_2)
                        new_population.append(child_1)
                        new_population.append(child_2)
                else:
                    for i in range(int(population_size/2)):
                        parent_1 = select_individual_by_tournament(population, lcoe_scores)
                        parent_2 = select_individual_by_tournament(population, lcoe_scores)
                        child_1, child_2 = breed_by_crossover(parent_1, parent_2)
                        new_population.append(child_1)
                        new_population.append(child_2)
            else:
                for i in range(int(population_size/2)):
                    parent_1 = select_individual_by_tournament(population, multi_scores)
                    parent_2 = select_individual_by_tournament(population, multi_scores)
                    child_1, child_2 = breed_by_crossover(parent_1, parent_2)
                    new_population.append(child_1)
                    new_population.append(child_2)
            # get back to original size (after elitism adds)
            if do_lcoe:
                new_population.pop()
            if do_multi:
                new_population.pop()
            # Replace the old population with the new one
            population = np.array(new_population)
            if self.optimise_mutation > 0:
                population = randomly_mutate_population(population, self.optimise_mutation)
            # Score best solution, and add to tracker
            lcoe_scores, multi_scores, multi_values = calculate_fitness(population)
            if do_lcoe:
                best_lcoe = np.min(lcoe_scores)
                best_ndx = lcoe_scores.index(best_lcoe)
                best_score = best_lcoe
            # now we can find the best weighted result - lowest is best
            if do_multi:
                best_multi = np.min(multi_scores)
                best_mndx = multi_scores.index(best_multi)
                multi_lowest_chrom = population[best_mndx]
                multi_best_popn.append(multi_lowest_chrom)
                multi_best.append(multi_values[best_mndx])
           #     if multi_best_weight > best_multi:
                multi_best_weight = best_multi
                if not do_lcoe:
                    best_score = best_multi
            best_score_progress.append(best_score)
            if best_score < lowest_score:
                lowest_score = best_score
                if do_lcoe:
                    lowest_chrom = population[best_ndx]
                else: #(do_multi only)
                    multi_lowest_chrom = population[best_mndx]
            if self.optimise_stop > 0:
                if best_score == last_score:
                    best_ctr += 1
                    if best_ctr >= self.optimise_stop:
                        break
                else:
                    last_score = best_score
                    best_ctr = 1
            last_score = best_score
            if do_lcoe:
                if best_score == best_score_progress[-2]:
                    lud = '='
                elif best_score < best_score_progress[-2]:
                    lud = '<html>&darr;</html>'
                else:
                    lud = '<html>&uarr;</html>'
            if do_multi:
                if best_multi == last_multi_score:
                    mud = '='
                elif best_multi < last_multi_score:
                    mud = '<html>&darr;</html>'
                else:
                    mud = '<html>&uarr;</html>'
                last_multi_score = best_multi
        if self.debug:
            try:
                self.db_file.close()
                optimiseDebug(self.db_file.name)
                os.remove(self.db_file.name)
            except:
                pass
            self.debug = False
        self.opt_progressbar.setVisible(False)
        self.opt_progressbar.close()
        tim = (time.time() - start_time)
        if tim < 60:
            tim = '%.1f secs)' % tim
        else:
            tim = '%.2f mins)' % (tim / 60.)
        msg = 'Optimise completed (%0d generations; %s' % (generation + 1, tim)
        if best_score > lowest_score:
            msg += ' Try more generations.'
        # we'll keep two or three to save re-calculating_fitness
        op_data = [[], [], []]
        score_data = [None, None, None]
        if do_lcoe:
            op_data[0], score_data[0] = calculate_fitness([lowest_chrom])
        if do_multi:
            op_data[1], score_data[1] = calculate_fitness([multi_lowest_chrom])
        self.setStatus(msg)
        QtWidgets.QApplication.processEvents()
        self.progressbar.setHidden(True)
        self.progressbar.setValue(0)
        # GA has completed required generation
        if do_lcoe:
            self.setStatus('Final LCOE: $%.2f' % best_score)
            fig = 'optimise_lcoe'
            titl = 'Optimise LCOE using Genetic Algorithm'
            ylbl = 'Best LCOE ($/MWh)'
        else:
            fig = 'optimise_multi'
            titl = 'Optimise Multi using Genetic Algorithm'
            ylbl = 'Best Weight'
        if do_multi:
            self.setStatus('Final Weight: %.4f' % multi_best_weight)
        # Plot progress
        x = list(range(1, len(best_score_progress)+ 1))
        matplotlib.rcParams['savefig.directory'] = self.scenarios
        plt.figure(fig + QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(),
                   '_yyyy-MM-dd_hhmm'))
        lx = plt.subplot(111)
        plt.title(titl)
        lx.plot(x, best_score_progress)
        lx.set_xlabel('Optimise Cycle (' + str(len(best_score_progress)) + ' generations)')
        lx.set_ylabel(ylbl)
        zp = ZoomPanX()
        f = zp.zoom_pan(lx, base_scale=1.2, annotate=True)
        plt.show()
        if do_multi:
            pick = plot_multi(multi_best, multi_order, 'best of each generation')
        else:
            pick = None
        headers = ['Facility', 'Capacity (MW/MWh)', 'Subtotal (MWh)', 'CF', 'Cost ($/yr)',
                   'LCOE ($/MWh)', 'Emissions (tCO2e)','Reference LCOE','Reference CF']
        op_pts = [0, 3, 0, 2, 0, 2, 0, 2, 2]
        if self.more_details:
            if do_lcoe:
                list(map(list, list(zip(*op_data[0]))))
                dialog = displaytable.Table(op_data[0], title=self.sender().text(), fields=headers,
                         save_folder=self.scenarios, sortby='', decpts=op_pts)
                dialog.exec_()
                del dialog
            if do_multi:
                list(map(list, list(zip(*op_data[1]))))
                dialog = displaytable.Table(op_data[1], title='Multi_' + self.sender().text(), fields=headers,
                         save_folder=self.scenarios, sortby='', decpts=op_pts)
                dialog.exec_()
                del dialog
        # now I'll display the resulting capacities for LCOE, lowest weight, picked
        # now get random amount of generation per technology (both RE and non-RE)
        its = {}
        for gen, value in opt_order.items():
            its[gen] = []
        chrom_hdrs = []
        chroms = []
        ndxes = []
        if do_lcoe:
            chrom_hdrs = ['Lowest LCOE']
            chroms = [lowest_chrom]
            ndxes = [0]
        if do_multi:
            chrom_hdrs.append('Lowest Weight')
            chroms.append(multi_lowest_chrom)
            ndxes.append(1)
        if pick is not None:
            # at present I'll calculate the best weight for the chosen picks. Could actually present all for user choice
            picks = []
            if len(pick) == 1:
                multi_lowest_chrom = multi_best_popn[pick[0][0]]
            else:
                for pck in pick:
                    picks.append(multi_best_popn[pck[0]])
                a, b, c = calculate_fitness(picks)
                best_multi = np.min(b)
                best_mndx = b.index(best_multi)
                multi_lowest_chrom = picks[best_mndx]
            op_data[2], score_data[2] = calculate_fitness([multi_lowest_chrom])
            if self.more_details:
                list(map(list, list(zip(*op_data[2]))))
                dialog = displaytable.Table(op_data[2], title='Pick_' + self.sender().text(), fields=headers,
                         save_folder=self.scenarios, sortby='', decpts=op_pts)
                dialog.exec_()
                del dialog
            chrom_hdrs.append('Your pick')
            chroms.append(multi_lowest_chrom)
            ndxes.append(2)
        for chromosome in chroms:
            for gen, value in opt_order.items():
                capacity = opt_order[gen][2]
                for c in range(value[0], value[1]):
                    if chromosome[c]:
                        capacity = capacity + capacities[c]
                its[gen].append(capacity / pmss_details[gen][0])
        chooseDialog = QtWidgets.QDialog()
        hbox = QtWidgets.QHBoxLayout()
        grid = [QtWidgets.QGridLayout()]
        label = QtWidgets.QLabel('<b>Facility</b>')
        label.setAlignment(QtCore.Qt.AlignCenter)
        grid[0].addWidget(label, 0, 0)
        for h in range(len(chrom_hdrs)):
            grid.append(QtWidgets.QGridLayout())
            label = QtWidgets.QLabel('<b>' + chrom_hdrs[h] + '</b>')
            label.setAlignment(QtCore.Qt.AlignCenter)
            grid[-1].addWidget(label, 0, 0, 1, 2)
        rw = 1
        for key, value in its.items():
            grid[0].addWidget(QtWidgets.QLabel(key), rw, 0)
            for h in range(len(chrom_hdrs)):
                label = QtWidgets.QLabel('{:.2f}'.format(value[h]))
                label.setAlignment(QtCore.Qt.AlignRight)
                grid[h + 1].addWidget(label, rw, 0)
                label = QtWidgets.QLabel('({:,.2f})'.format(value[h] * pmss_details[key][0]))
                label.setAlignment(QtCore.Qt.AlignRight)
                grid[h + 1].addWidget(label, rw, 1)
            rw += 1
        max_amt = [0., 0.]
        if do_lcoe:
            max_amt[0] = score_data[0][0]['cost']
            max_amt[1] = score_data[0][0]['co2']
        if do_multi:
            for multi in multi_best:
                max_amt[0] = max(max_amt[0], multi['cost'])
                max_amt[1] = max(max_amt[1], multi['co2'])
        pwr_chr = ['', '']
        divisor = [1., 1.]
        pwr_chrs = ' KMBTPEZY'
        for m in range(2):
            for pwr in range(len(pwr_chrs) - 1, -1, -1):
                if max_amt[m] > pow(10, pwr * 3):
                    pwr_chr[m] = pwr_chrs[pwr]
                    divisor[m] = 1. * pow(10, pwr * 3)
                    break
        self.targets['cost'][5] = self.targets['cost'][5].replace('pwr_chr', pwr_chr[0])
        self.targets['co2'][5] = self.targets['co2'][5].replace('pwr_chr', pwr_chr[1])
        for key in multi_order:
            lbl = QtWidgets.QLabel('<i>' + self.targets[key][0] + '</i>')
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            grid[0].addWidget(lbl, rw, 0)
            for h in range(len(chrom_hdrs)):
                if key == 'cost':
                    amt = score_data[ndxes[h]][0][key] / divisor[0] # cost
                elif key == 'co2':
                    amt = score_data[ndxes[h]][0][key] / divisor[1] # co2
                elif key[-4:] == '_pct': # percentage
                    amt = score_data[ndxes[h]][0][key] * 100.
                else:
                    amt = score_data[ndxes[h]][0][key]
                txt = '<i>' + self.targets[key][5] + '</i>'
                try:
                    label = QtWidgets.QLabel(txt % amt)
                except:
                    label = QtWidgets.QLabel('?')
                    print('(3456)', key, txt, amt)
                label.setAlignment(QtCore.Qt.AlignCenter)
                grid[h + 1].addWidget(label, rw, 0, 1, 2)
            rw += 1
        cshow = QtWidgets.QPushButton('Quit', self)
        grid[0].addWidget(cshow)
        cshow.clicked.connect(chooseDialog.close)
        for h in range(len(chrom_hdrs)):
            button = QtWidgets.QPushButton(chrom_hdrs[h], self)
            grid[h + 1].addWidget(button, rw, 0, 1, 2)
            button.clicked.connect(chooseClicked) #(chrom_hdrs[h]))
        for gri in grid:
            frame = QtWidgets.QFrame()
            frame.setFrameStyle(QtWidgets.QFrame.Box)
            frame.setLineWidth(1)
            frame.setLayout(gri)
            hbox.addWidget(frame)
   #     grid.addWidget(show, rw, 1)
    #    show.clicked.connect(optDialog.close)
        chooseDialog.setLayout(hbox)
        chooseDialog.setWindowTitle('Choose Optimal Generator Mix')
        chooseDialog.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
     #  this is a big of a kluge but I couldn't get it to behave
        self.opt_choice = ''
        chooseDialog.exec_()
        del chooseDialog
        try:
            h = chrom_hdrs.index(self.opt_choice)
        except:
            return
        op_data[h], score_data[h] = calculate_fitness([chroms[h]]) # make it current
        msg = chrom_hdrs[h] + ': '
        for key in multi_order[:3]:
            msg += self.targets[key][0] + ': '
            if key == 'cost':
                amt = score_data[h][0][key] / divisor[0] # cost
            elif key == 'co2':
                amt = score_data[h][0][key] / divisor[1] # co2
            elif key[-4:] == '_pct': # percentage
                amt = score_data[h][0][key] * 100.
            else:
                amt = score_data[h][0][key]
            txt = self.targets[key][5]
            txt = txt % amt
            msg += txt + '; '
        self.setStatus(msg)
        list(map(list, list(zip(*op_data[h]))))
        dialog = displaytable.Table(op_data[h], title='Chosen_' + self.sender().text(), fields=headers,
                 save_folder=self.scenarios, sortby='', decpts=op_pts)
        dialog.exec_()
        del dialog
        if self.adjust.isChecked():
            self.adjustby = {}
            for gen, value in iter(sorted(pmss_details.items())):
                if value[3] != 1:
                    self.adjustby[gen] = value[3]
        return

if "__main__" == __name__:
    app = QtWidgets.QApplication(sys.argv)
    ex = powerMatch()
    app.exec_()
    app.deleteLater()
    sys.exit()
